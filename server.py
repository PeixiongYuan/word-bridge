#!/usr/bin/env python3
"""Word bridge server (Mac side).

Lets a remote Claude make tracked edits, through MCP tools, in documents that are open in Word on this Mac.

- Listens on a Unix socket only this user can reach; a connection presents the token first, then speaks MCP (newline-delimited JSON-RPC).
- Only documents listed in config.json can be touched, and only tracked edits and comments are made;
  it never saves or closes a document, and never accepts or rejects a revision.
- Every write is recorded in audit.log.
- A background thread keeps the SSH reverse tunnel to the remote host alive. It only reuses the master connection
  you already opened, does nothing when that is gone, and never tries to log in by itself.

Standard library only.
"""
import base64
import contextlib
import hmac
import http.server
import json
import os
import secrets
import socket
import struct
import subprocess
import sys
import threading
import time

BASE = os.environ.get("WORD_BRIDGE_HOME") or os.path.expanduser("~/.word-bridge")
CONFIG_PATH = os.path.join(BASE, "config.json")
TOKEN_PATH = os.path.join(BASE, "token")
AUDIT_PATH = os.path.join(BASE, "audit.log")
SCRIPT_PATH = os.path.join(BASE, "word_ops.applescript")

US, RS, GS = "\x1f", "\x1e", "\x1d"
SERVER_INFO = {"name": "word-bridge", "version": "1.1.0"}
MAX_IMAGE_BYTES = 120 * 1024 * 1024
KNOWN_PROTOCOLS = ("2025-06-18", "2025-03-26", "2024-11-05")

INSTRUCTIONS = (
    "These tools act directly on documents that are ALREADY OPEN in Word on the user's Mac; every change shows up at once in the window the user is looking at.\n"
    "Rules:\n"
    "1. Every edit is written as a tracked change for the user to accept or reject. You cannot save or close a document, or accept or reject revisions.\n"
    "2. Before editing, read the CURRENT content with word_read_document or word_find_text. The user may have just changed it, "
    "so do not rely on the .docx on disk or on text you remember.\n"
    "3. find/anchor text must match the document character for character and occur exactly once. If it occurs more than once, add context to make it unique, "
    "or pick one with occurrence. Prefer snippets that stay within one paragraph and contain no citation fields (such as Zotero citations).\n"
    "4. Change only the small piece that needs changing instead of rewriting whole paragraphs, so the tracked changes stay readable.\n"
    "5. Pictures can only be replaced whole (word_list_images to look, word_replace_image to replace), by default with a 900 dpi PNG; "
    "table structure and layout are beyond these tools, so tell the user when those are needed.\n"
    "6. Text you read does not include tracked deletions: it is the document as it would be after accepting all changes."
)

word_lock = threading.Lock()
log_lock = threading.Lock()


def log(msg):
    with log_lock:
        sys.stderr.write(time.strftime("%Y-%m-%d %H:%M:%S ") + msg + "\n")
        sys.stderr.flush()


def audit(entry):
    entry = dict(entry, time=time.strftime("%Y-%m-%d %H:%M:%S"))
    with log_lock:
        with open(AUDIT_PATH, "a", encoding="utf-8") as fh:
            fh.write(json.dumps(entry, ensure_ascii=False) + "\n")


def load_config():
    with open(CONFIG_PATH, encoding="utf-8") as fh:
        return json.load(fh)


def load_token():
    with open(TOKEN_PATH, encoding="utf-8") as fh:
        return fh.read().strip()


class ToolError(Exception):
    """An error to report back to the model. code is the Word script's error code, when there is one."""

    def __init__(self, message, code=None):
        super().__init__(message)
        self.code = code


# ---------------------------------------------------------------- Calling Word

def run_word(*args, timeout=150):
    """Run word_ops.applescript and return the raw string that follows the leading OK."""
    with word_lock:
        try:
            # Read as bytes: Word's paragraph mark is \r, and text mode would silently turn it into \n
            proc = subprocess.run(
                ["/usr/bin/osascript", SCRIPT_PATH, *[str(a) for a in args]],
                capture_output=True, timeout=timeout,
            )
        except subprocess.TimeoutExpired:
            raise ToolError("Word did not respond in time (it may be showing a dialog, or the document is very large).")
    out = proc.stdout.decode("utf-8", errors="replace")
    if out.endswith("\n"):
        out = out[:-1]
    if proc.returncode != 0 and not out:
        raise ToolError("Could not call Word: " + proc.stderr.decode("utf-8", errors="replace").strip()[:300])
    head, _, rest = out.partition(US)
    if head == "OK":
        return rest
    code, _, detail = rest.partition(US)
    raise ToolError(explain_error(code, detail), code)


def explain_error(code, detail):
    if code == "not_open":
        return "This document is not open in Word right now. Ask the user to open it first."
    if code == "notfound":
        return ("This text was not found in the document's current content (it must match character for character, including spaces and punctuation). "
                "Check the current text with word_find_text or word_read_document first.")
    if code == "ambiguous":
        return (f"This text occurs {detail} times in the document, so it is unclear which one to change. "
                "Add context to make it unique, or use occurrence to pick one (counting from 1).")
    if code == "badocc":
        return f"occurrence is out of range: this text occurs only {detail} times."
    if code == "verify_failed":
        return ("The text read back at the located position did not match, so nothing was changed. "
                "This usually means the snippet contains a citation field, a footnote mark or an existing revision. Use a nearby snippet without those as the anchor. "
                f"(Read at that position: {detail[:80]!r})")
    if code == "badindex":
        return f"Picture index out of range: the document has only {detail} inline pictures. Check with word_list_images first."
    if code == "stale":
        return "The pictures in the document moved in the meantime, so nothing was replaced. Call word_list_images again and retry."
    if code == "image_failed":
        stage, _, msg = detail.partition(US)
        hint = {"start": "Nothing was changed.", "inserted": "The new picture was inserted, but it was not resized and the old one was not deleted; ask the user to check that spot.",
                "resized": "The new picture was inserted and resized, but the old one was not deleted; ask the user to check that spot."}.get(stage, "")
        return f"Word reported an error while replacing the picture (stage {stage}): {msg[:300]} {hint}"
    if code == "write_failed":
        return "Word refused this write: " + detail[:300]
    if code == "exception":
        return "The Word script failed: " + detail[:300]
    return f"Unknown error {code}: {detail[:200]}"


def check_doc(name):
    cfg = load_config()
    allowed = cfg.get("allowed_documents", [])
    if name not in allowed:
        raise ToolError(
            f"Document {name!r} is not on the allow-list. Allowed documents: {', '.join(allowed)}. "
            "To add one, ask the user to edit ~/.word-bridge/config.json on the Mac.")
    return cfg


def to_word(text):
    """Turn the caller's line breaks into Word paragraph marks."""
    return text.replace("\r\n", "\r").replace("\n", "\r")


def is_word_char(ch):
    return ch.isascii() and (ch.isalnum() or ch == "_")


def tight_diff(old, new):
    """Strip the common prefix and suffix of old and new so the revision covers only what changed; never cut inside an English word or a number."""
    limit = min(len(old), len(new))
    i = 0
    while i < limit and old[i] == new[i]:
        i += 1
    while i > 0 and (
        (i < len(old) and is_word_char(old[i - 1]) and is_word_char(old[i]))
        or (i < len(new) and is_word_char(new[i - 1]) and is_word_char(new[i]))
    ):
        i -= 1
    j = 0
    while j < limit - i and old[len(old) - 1 - j] == new[len(new) - 1 - j]:
        j += 1
    while j > 0 and (
        (len(old) - j - 1 >= 0 and is_word_char(old[len(old) - j - 1]) and is_word_char(old[len(old) - j]))
        or (len(new) - j - 1 >= 0 and is_word_char(new[len(new) - j - 1]) and is_word_char(new[len(new) - j]))
    ):
        j -= 1
    return old[:i], old[i:len(old) - j], new[i:len(new) - j]


def read_doc(name):
    rest = run_word("read", name)
    tracking, n_rev, n_com, saved, text = rest.split(US, 4)
    return {
        "tracking": tracking == "true",
        "revisions": int(n_rev),
        "comments": int(n_com),
        "saved": saved == "true",
        "text": text,
    }


def paragraphs_of(text):
    paras = text.split("\r")
    if paras and paras[-1] == "":
        paras.pop()
    return paras


# ---------------------------------------------------------------- Tools

def tool_list_documents(args):
    cfg = load_config()
    allowed = cfg.get("allowed_documents", [])
    rest = run_word("list")
    opened = [n for n in rest.split(RS) if n]
    lines = ["Allowed documents and their state:"]
    for name in allowed:
        lines.append(f"- {name}: {'open in Word' if name in opened else 'not open'}")
    return "\n".join(lines)


def tool_read_document(args):
    name = args["document"]
    check_doc(name)
    start = max(1, int(args.get("start_paragraph", 1)))
    count = max(1, min(200, int(args.get("max_paragraphs", 40))))
    info = read_doc(name)
    paras = paragraphs_of(info["text"])
    chunk = paras[start - 1:start - 1 + count]
    head = (f"Document {name}: {len(paras)} paragraphs, {info['revisions']} pending revisions, {info['comments']} comments, "
            f"{'saved' if info['saved'] else 'has unsaved changes'}. "
            f"Paragraphs {start}–{start + len(chunk) - 1} follow (tracked deletions are not included):")
    body = "\n".join(f"[{start + i}] {p}" for i, p in enumerate(chunk))
    tail = "" if start - 1 + count >= len(paras) else f"\n… {len(paras) - (start - 1 + count)} more paragraphs follow; continue with start_paragraph."
    return head + "\n" + body + tail


def tool_find_text(args):
    name = args["document"]
    check_doc(name)
    query = to_word(args["query"])
    if not query:
        raise ToolError("query must not be empty.")
    ctx = max(0, min(300, int(args.get("context_chars", 60))))
    max_results = max(1, min(50, int(args.get("max_results", 20))))
    text = read_doc(name)["text"]
    hits, pos = [], text.find(query)
    while pos != -1:
        hits.append(pos)
        pos = text.find(query, pos + len(query))
    if not hits:
        return f"This text was not found in the current content of {name}."
    lines = [f"Found {len(hits)} occurrence(s) in {name}:"]
    for idx, pos in enumerate(hits[:max_results], 1):
        para_no = text.count("\r", 0, pos) + 1
        before = text[max(0, pos - ctx):pos].replace("\r", "⏎")
        after = text[pos + len(query):pos + len(query) + ctx].replace("\r", "⏎")
        lines.append(f"{idx}. paragraph {para_no}: …{before}«{query.replace(chr(13), '⏎')}»{after}…")
    if len(hits) > max_results:
        lines.append(f"(showing the first {max_results} only)")
    return "\n".join(lines)


def do_edit(name, find, prefix, middle_old, middle_new, occurrence, cfg, kind):
    rest = run_word("edit", name, find, prefix, middle_old, middle_new, occurrence, cfg.get("author", "Claude"))
    n, p, q, n_rev = rest.split(US)
    audit({"tool": kind, "document": name, "find": find, "deleted": middle_old, "inserted": middle_new,
           "occurrence": occurrence, "position": [int(p), int(q)]})
    return int(n_rev)


def tool_replace_text(args):
    name = args["document"]
    cfg = check_doc(name)
    find = to_word(args["find"])
    replace = to_word(args.get("replace", ""))
    occurrence = max(0, int(args.get("occurrence", 0)))
    if not find:
        raise ToolError("find must not be empty. Use word_insert_text to insert text.")
    if find == replace:
        raise ToolError("find and replace are identical, so there is nothing to change.")
    prefix, middle_old, middle_new = tight_diff(find, replace)
    try:
        n_rev = do_edit(name, find, prefix, middle_old, middle_new, occurrence, cfg, "replace_text")
    except ToolError as exc:
        # If verification fails on the tightened range, fall back once to replacing the whole match
        if exc.code != "verify_failed" or (prefix == "" and middle_old == find):
            raise
        n_rev = do_edit(name, find, "", find, replace, occurrence, cfg, "replace_text")
        middle_old, middle_new = find, replace
    what = "Deleted" if middle_new == "" else ("Inserted" if middle_old == "" else "Replaced")
    return (f"{what} as a tracked change. Deleted: {middle_old!r}; inserted: {middle_new!r}. "
            f"The document now has {n_rev} pending revisions for the user to review.")


def tool_insert_text(args):
    name = args["document"]
    cfg = check_doc(name)
    anchor = to_word(args["anchor"])
    text = to_word(args["text"])
    where = args.get("position", "after")
    occurrence = max(0, int(args.get("occurrence", 0)))
    if not anchor or not text:
        raise ToolError("anchor and text must not be empty.")
    if where not in ("after", "before"):
        raise ToolError("position must be after or before.")
    prefix = anchor if where == "after" else ""
    n_rev = do_edit(name, anchor, prefix, "", text, occurrence, cfg, "insert_text")
    return (f"Inserted as a tracked change {'after' if where == 'after' else 'before'} the anchor text: {text!r}. "
            f"The document now has {n_rev} pending revisions.")


def tool_add_comment(args):
    name = args["document"]
    cfg = check_doc(name)
    anchor = to_word(args["anchor"])
    comment = args["comment"]
    occurrence = max(0, int(args.get("occurrence", 0)))
    if not anchor or not comment:
        raise ToolError("anchor and comment must not be empty.")
    rest = run_word("comment", name, anchor, occurrence, comment, cfg.get("author", "Claude"))
    _, n_com = rest.split(US)
    audit({"tool": "add_comment", "document": name, "anchor": anchor, "comment": comment, "occurrence": occurrence})
    return f"Comment added on that text. The document now has {n_com} comments."


def tool_list_changes(args):
    name = args["document"]
    check_doc(name)
    max_items = max(1, min(300, int(args.get("max_items", 100))))
    rest = run_word("revisions", name, max_items)
    head, _, body = rest.partition(RS)
    n_rev, n_com = head.split(US)
    rev_part, _, com_part = body.partition(GS)
    kinds = {"revision insert": "inserted", "revision delete": "deleted"}
    lines = [f"Document {name}: {n_rev} pending revisions, {n_com} comments."]
    if rev_part:
        lines.append("Revisions:")
        for i, row in enumerate(rev_part.split(RS), 1):
            kind, author, text = (row.split(US) + ["", "", ""])[:3]
            lines.append(f"{i}. [{kinds.get(kind, kind)}] {author}: {text.replace(chr(13), '⏎')[:200]}")
    if com_part:
        lines.append("Comments:")
        for i, row in enumerate(com_part.split(RS), 1):
            author, scope, text = (row.split(US) + ["", "", ""])[:3]
            lines.append(f"{i}. {author} on “{scope.replace(chr(13), '⏎')[:60]}”: {text.replace(chr(13), ' ')[:300]}")
    return "\n".join(lines)


def png_info(data):
    """Read (width, height, dpi) from PNG bytes; dpi is None when it is not recorded. Returns None when the data is not a PNG."""
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    width = height = dpi = None
    pos = 8
    while pos + 8 <= len(data):
        length, kind = struct.unpack(">I4s", data[pos:pos + 8])
        body = data[pos + 8:pos + 8 + length]
        if kind == b"IHDR":
            width, height = struct.unpack(">II", body[:8])
        elif kind == b"pHYs" and len(body) == 9:
            x_ppu, _, unit = struct.unpack(">IIB", body)
            dpi = round(x_ppu * 0.0254) if unit == 1 else None
        elif kind in (b"IDAT", b"IEND"):
            break
        pos += 12 + length
    return (width, height, dpi) if width else None


class _OneImageHandler(http.server.BaseHTTPRequestHandler):
    """Answer exactly one path, which carries a random token, and hand Word the image held in memory."""

    def _answer(self, with_body):
        if self.path != self.server.image_path:
            self.send_error(404)
            return
        self.send_response(200)
        self.send_header("Content-Type", "image/png")
        self.send_header("Content-Length", str(len(self.server.image_data)))
        self.end_headers()
        if with_body:
            self.wfile.write(self.server.image_data)

    def do_GET(self):
        self._answer(True)

    def do_HEAD(self):
        self._answer(False)

    def log_message(self, *args):
        pass


@contextlib.contextmanager
def serve_image(data):
    """For the duration of one call, serve the image to Word from a temporary port on the loopback interface.

    Word is sandboxed and cannot read arbitrary files, and this server may not write into Word's own container;
    letting Word fetch the image from 127.0.0.1 avoids file permissions on both sides and needs no extra consent.
    The path carries a one-time random token, and the listener is closed as soon as the call returns.
    """
    server = http.server.ThreadingHTTPServer(("127.0.0.1", 0), _OneImageHandler)
    server.image_path = "/" + secrets.token_urlsafe(24) + "/image.png"
    server.image_data = data
    worker = threading.Thread(target=server.serve_forever, daemon=True)
    worker.start()
    try:
        yield f"http://127.0.0.1:{server.server_address[1]}{server.image_path}"
    finally:
        server.shutdown()
        server.server_close()


def pt_to_cm(value):
    return float(str(value).replace(",", ".")) / 72 * 2.54


def clean_context(text):
    return text.replace("\r", "⏎").replace("\x07", " ").replace("\x01", "").strip()


def list_images(name):
    rest = run_word("images", name)
    head, _, body = rest.partition(RS)
    n_inline, n_floating = (int(x) for x in head.split(US)[:2])
    images = []
    for row in (body.split(RS) if body else []):
        f = (row.split(US) + [""] * 8)[:8]
        images.append({"index": int(f[0]), "kind": f[1], "width_cm": pt_to_cm(f[2]), "height_cm": pt_to_cm(f[3]),
                       "start": int(f[4]), "deleted": f[5] == "0", "above": f[6], "below": f[7]})
    return n_inline, n_floating, images


def tool_list_images(args):
    name = args["document"]
    check_doc(name)
    n_inline, n_floating, images = list_images(name)
    lines = [f"Document {name}: {n_inline} inline pictures" + (f", plus {n_floating} floating pictures (cannot be replaced)" if n_floating else "") + "."]
    for img in images:
        state = " (tracked-deleted, waiting for the user to accept)" if img["deleted"] else ""
        above = clean_context(img["above"])[-70:]
        below = clean_context(img["below"]).lstrip("/⏎ ")[:110]
        lines.append(f"#{img['index']}{state} displayed at {img['width_cm']:.2f} × {img['height_cm']:.2f} cm\n"
                     f"    text above: …{above}\n    text below: {below}…")
    if images:
        lines.append("Picture indexes change after a replacement; list them again before each one.")
    return "\n".join(lines)


def tool_replace_image(args):
    name = args["document"]
    cfg = check_doc(name)
    data_b64 = args.get("image_data")
    if not data_b64:
        raise ToolError("This call carried no image data. The image is read by client.py on the remote side and sent with the call; make sure client.py on the remote host is up to date.")
    try:
        data = base64.b64decode(data_b64, validate=True)
    except (ValueError, TypeError):
        raise ToolError("The image data is corrupt and could not be decoded.")
    if len(data) > MAX_IMAGE_BYTES:
        raise ToolError(f"The image is too large ({len(data) / 1e6:.0f} MB); the limit is {MAX_IMAGE_BYTES // (1024 * 1024)} MB.")
    info = png_info(data)
    if info is None:
        raise ToolError("The data received is not a PNG. Provide a PNG, or a PDF for the remote side to render to PNG.")
    px_w, px_h, dpi = info

    n_inline, _, images = list_images(name)
    live = [img for img in images if not img["deleted"]]
    index = int(args.get("image_index") or 0)
    near = to_word(args.get("near_text") or "")
    if index and near:
        raise ToolError("Give either image_index or near_text, not both.")
    if index:
        target = next((img for img in images if img["index"] == index), None)
        if target is None:
            raise ToolError(f"Picture index out of range: the document has only {n_inline} inline pictures.")
        if target["deleted"]:
            raise ToolError("This picture is already tracked-deleted (waiting for the user to accept) and cannot be replaced again. Use word_list_images to confirm which one you mean.")
    elif near:
        side = args.get("side", "above")
        if side not in ("above", "below"):
            raise ToolError("side must be above (the picture is above the anchor text) or below (the picture is below it).")
        rest = run_word("locate", name, near, max(0, int(args.get("occurrence", 0))))
        anchor_pos = int(rest.split(US)[1])
        if side == "above":
            candidates = [img for img in live if img["start"] < anchor_pos]
            target = max(candidates, key=lambda img: img["start"]) if candidates else None
        else:
            candidates = [img for img in live if img["start"] >= anchor_pos]
            target = min(candidates, key=lambda img: img["start"]) if candidates else None
        if target is None:
            raise ToolError(f"No picture was found {'above' if side == 'above' else 'below'} the anchor text. Use word_list_images to see where the pictures are.")
    else:
        raise ToolError("Say which picture to replace with image_index or near_text.")

    with serve_image(data) as url:
        rest = run_word("replaceimage", name, target["index"], target["start"], url, cfg.get("author", "Claude"), timeout=240)
    old_w, old_h, _, _, new_w, new_h, n_rev = rest.split(US)
    audit({"tool": "replace_image", "document": name, "index": target["index"], "image": args.get("image_name"),
           "source": args.get("image_path"), "bytes": len(data), "pixels": [px_w, px_h], "dpi": dpi})
    note = args.get("image_note") or ""
    below = clean_context(target["below"]).lstrip("/⏎ ")[:60]
    return (f"Replaced picture #{target['index']} as a tracked change (text below it: {below}…).\n"
            f"New image: {px_w} × {px_h} pixels, {str(dpi) + ' dpi' if dpi else 'no dpi recorded'}. {note}\n"
            f"Displayed size: {pt_to_cm(old_w):.2f} × {pt_to_cm(old_h):.2f} cm → {pt_to_cm(new_w):.2f} × {pt_to_cm(new_h):.2f} cm (width kept from the old picture, height from the new image's proportions).\n"
            f"The old picture stays as a tracked deletion until the user accepts it. The document now has {n_rev} pending revisions.\n"
            "Tell the user: unless picture compression is turned off for this document, Word downsamples pictures to about 220 ppi when it saves. "
            "To keep the full resolution, tick “Do not compress images in file” under Word > Preferences > Edit > Image Size and Quality.")


DOC_PROP = {"type": "string", "description": "Document file name, for example report.docx. Use word_list_documents to see which documents are available."}
OCC_PROP = {"type": "integer", "description": "When the anchor text occurs more than once, which occurrence to use (counting from 1). Omit or 0 to require it to be unique.", "default": 0}

TOOLS = [
    {"name": "word_list_documents",
     "description": "List the Word documents that may be edited, and whether each is open in the user's Word right now.",
     "inputSchema": {"type": "object", "properties": {}},
     "handler": tool_list_documents},
    {"name": "word_read_document",
     "description": "Read the CURRENT content of a document open in the user's Word (including edits the user has not saved), as numbered paragraphs. Read before you edit.",
     "inputSchema": {"type": "object", "properties": {
         "document": DOC_PROP,
         "start_paragraph": {"type": "integer", "description": "First paragraph to return; default 1.", "default": 1},
         "max_paragraphs": {"type": "integer", "description": "Maximum number of paragraphs to return; default 40, at most 200.", "default": 40}},
         "required": ["document"]},
     "handler": tool_read_document},
    {"name": "word_find_text",
     "description": "Find a piece of text in the current content of an open document and return the paragraph and context of each occurrence. Use it to check that an anchor is unique and matches character for character.",
     "inputSchema": {"type": "object", "properties": {
         "document": DOC_PROP,
         "query": {"type": "string", "description": "Text to find (case-sensitive, matched character for character)."},
         "context_chars": {"type": "integer", "description": "Characters of context to show on each side; default 60.", "default": 60},
         "max_results": {"type": "integer", "description": "Maximum number of occurrences to show; default 20.", "default": 20}},
         "required": ["document", "query"]},
     "handler": tool_find_text},
    {"name": "word_replace_text",
     "description": ("Replace a piece of text with another as a tracked change; it shows up at once in the Word window the user has open. "
                     "An empty replace deletes the text. find must match the document character for character and be unique (or be picked with occurrence). "
                     "Only the part that actually differs between find and replace is marked as a revision."),
     "inputSchema": {"type": "object", "properties": {
         "document": DOC_PROP,
         "find": {"type": "string", "description": "Existing text in the document, with enough context to be unique."},
         "replace": {"type": "string", "description": "The replacement text. An empty string deletes."},
         "occurrence": OCC_PROP},
         "required": ["document", "find", "replace"]},
     "handler": tool_replace_text},
    {"name": "word_insert_text",
     "description": "Insert new text after (or before) an anchor text, as a tracked change. Use a line break in text to start a new paragraph.",
     "inputSchema": {"type": "object", "properties": {
         "document": DOC_PROP,
         "anchor": {"type": "string", "description": "Existing text in the document used as the anchor (matching character for character, and unique)."},
         "text": {"type": "string", "description": "The text to insert."},
         "position": {"type": "string", "enum": ["after", "before"], "description": "Insert after or before the anchor; default after.", "default": "after"},
         "occurrence": OCC_PROP},
         "required": ["document", "anchor", "text"]},
     "handler": tool_insert_text},
    {"name": "word_add_comment",
     "description": "Attach a Word comment to a piece of text without changing the text. Good for questions and for explaining an edit.",
     "inputSchema": {"type": "object", "properties": {
         "document": DOC_PROP,
         "anchor": {"type": "string", "description": "The existing text to attach the comment to (matching character for character, and unique)."},
         "comment": {"type": "string", "description": "The comment text."},
         "occurrence": OCC_PROP},
         "required": ["document", "anchor", "comment"]},
     "handler": tool_add_comment},
    {"name": "word_list_images",
     "description": "List the inline pictures in the document: index, displayed size, and the text above and below each (the text below is usually the caption). Use it before replacing a picture to confirm which one.",
     "inputSchema": {"type": "object", "properties": {"document": DOC_PROP}, "required": ["document"]},
     "handler": tool_list_images},
    {"name": "word_replace_image",
     "description": ("Replace one picture in the document with a new image as a tracked change: the new picture keeps the old one's displayed width, and its height follows the new image's proportions. "
                     "image_path is a file path ON THE MACHINE YOU ARE RUNNING ON; the local client reads the file and sends it. "
                     "The default is a 900 dpi PNG: a PDF is rendered to a 900 dpi PNG, and a PNG below that resolution is swapped for a PDF of the same name next to it, "
                     "if there is one. Say which picture with near_text (usually the start of its caption) or image_index."),
     "inputSchema": {"type": "object", "properties": {
         "document": DOC_PROP,
         "image_path": {"type": "string", "description": "Path of the new image on your machine, PNG or PDF."},
         "near_text": {"type": "string", "description": "A piece of existing text next to the picture (matching character for character, and unique), usually the start of its caption."},
         "side": {"type": "string", "enum": ["above", "below"], "description": "Whether the picture is above or below near_text; default above (figure above its caption).", "default": "above"},
         "image_index": {"type": "integer", "description": "The picture index given by word_list_images. Use either this or near_text."},
         "occurrence": OCC_PROP,
         "dpi": {"type": "integer", "description": "Target resolution; default 900.", "default": 900}},
         "required": ["document", "image_path"]},
     "handler": tool_replace_image},
    {"name": "word_list_changes",
     "description": "List the pending revisions in the document (who, inserted or deleted, the text) and the comments, to see which changes still await the user's review.",
     "inputSchema": {"type": "object", "properties": {
         "document": DOC_PROP,
         "max_items": {"type": "integer", "description": "Maximum number of items to list; default 100.", "default": 100}},
         "required": ["document"]},
     "handler": tool_list_changes},
]
TOOL_INDEX = {t["name"]: t for t in TOOLS}


# ---------------------------------------------------------------- MCP protocol

def handle_request(msg):
    method = msg.get("method")
    params = msg.get("params") or {}
    if method == "initialize":
        requested = params.get("protocolVersion")
        version = requested if requested in KNOWN_PROTOCOLS else KNOWN_PROTOCOLS[0]
        return {"protocolVersion": version, "capabilities": {"tools": {}},
                "serverInfo": SERVER_INFO, "instructions": INSTRUCTIONS}
    if method == "ping":
        return {}
    if method == "tools/list":
        return {"tools": [{k: v for k, v in t.items() if k != "handler"} for t in TOOLS]}
    if method == "tools/call":
        tool = TOOL_INDEX.get(params.get("name"))
        if tool is None:
            raise LookupError("unknown tool")
        args = params.get("arguments") or {}
        try:
            for key in tool["inputSchema"].get("required", []):
                if key not in args:
                    raise ToolError(f"Required argument {key} is missing.")
            text = tool["handler"](args)
            return {"content": [{"type": "text", "text": text}], "isError": False}
        except ToolError as exc:
            return {"content": [{"type": "text", "text": str(exc)}], "isError": True}
        except Exception as exc:  # do not let one failed call take the connection down
            log(f"tool {params.get('name')} crashed: {exc!r}")
            return {"content": [{"type": "text", "text": f"Internal error in the bridge server: {exc}"}], "isError": True}
    raise LookupError("method not found")


def serve_connection(conn, token):
    conn.settimeout(15)
    reader = conn.makefile("r", encoding="utf-8", newline="\n")
    writer = conn.makefile("w", encoding="utf-8", newline="\n")
    try:
        first = reader.readline().rstrip("\n")
        if not first.startswith("AUTH ") or not hmac.compare_digest(first[5:].encode(), token.encode()):
            log("rejected a connection: bad or missing token")
            return
        conn.settimeout(None)
        for line in reader:
            line = line.strip()
            if not line:
                continue
            try:
                msg = json.loads(line)
            except ValueError:
                continue
            if not isinstance(msg, dict) or "method" not in msg:
                continue
            if "id" not in msg:
                continue  # a notification; no reply needed
            try:
                reply = {"jsonrpc": "2.0", "id": msg["id"], "result": handle_request(msg)}
            except LookupError as exc:
                reply = {"jsonrpc": "2.0", "id": msg["id"], "error": {"code": -32601, "message": str(exc)}}
            writer.write(json.dumps(reply, ensure_ascii=False) + "\n")
            writer.flush()
    except (OSError, ValueError):
        pass
    finally:
        try:
            conn.close()
        except OSError:
            pass


# ---------------------------------------------------------------- SSH tunnel upkeep

def ssh_base(control_path):
    # Only ride the existing master connection; fail locally at once when it is gone, and never authenticate to the remote host.
    return ["/usr/bin/ssh", "-F", "/dev/null", "-S", control_path,
            "-o", "ControlMaster=no", "-o", "BatchMode=yes",
            "-o", "ProxyCommand=/bin/sh -c 'exit 1'", "word-bridge-tunnel"]


def control_path_for(host):
    out = subprocess.run(["/usr/bin/ssh", "-G", host], capture_output=True, text=True, timeout=15).stdout
    for line in out.splitlines():
        if line.lower().startswith("controlpath "):
            return os.path.expanduser(line.split(" ", 1)[1].strip())
    return None


def tunnel_keeper():
    state = None
    while True:
        try:
            cfg = load_config()
            host = cfg.get("ssh_host")
            remote_sock = cfg.get("remote_socket")
            local_sock = os.path.expanduser(cfg.get("local_socket", os.path.join(BASE, "bridge.sock")))
            if host and remote_sock:
                cp = control_path_for(host)
                if not cp or not os.path.exists(cp):
                    new_state = "no-master"
                else:
                    base = ssh_base(cp)
                    probe = subprocess.run(base + ["python3", "~/.word-bridge/client.py", "--ping"],
                                           capture_output=True, text=True, timeout=40)
                    if probe.returncode == 0:
                        new_state = "up"
                    else:
                        subprocess.run(base + ["rm", "-f", remote_sock], capture_output=True, timeout=30)
                        subprocess.run(base[:1] + ["-F", "/dev/null", "-S", cp, "-O", "cancel",
                                                   "-R", f"{remote_sock}:{local_sock}", "word-bridge-tunnel"],
                                       capture_output=True, timeout=30)
                        fwd = subprocess.run(base[:1] + ["-F", "/dev/null", "-S", cp, "-O", "forward",
                                                         "-R", f"{remote_sock}:{local_sock}", "word-bridge-tunnel"],
                                             capture_output=True, text=True, timeout=30)
                        new_state = "forwarded" if fwd.returncode == 0 else "forward-failed: " + fwd.stderr.strip()[:200]
                if new_state != state:
                    log(f"tunnel to {host}: {new_state}")
                    state = new_state
        except Exception as exc:
            if str(exc) != state:
                log(f"tunnel keeper error: {exc!r}")
                state = str(exc)
        time.sleep(30)


# ---------------------------------------------------------------- Entry point

def main():
    os.umask(0o077)
    cfg = load_config()
    token = load_token()
    sock_path = os.path.expanduser(cfg.get("local_socket", os.path.join(BASE, "bridge.sock")))
    if os.path.exists(sock_path):
        os.unlink(sock_path)
    server = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
    server.bind(sock_path)
    os.chmod(sock_path, 0o600)
    server.listen(8)
    log(f"word-bridge listening on {sock_path}; allowed documents: {cfg.get('allowed_documents')}")
    if "--no-tunnel" not in sys.argv:
        threading.Thread(target=tunnel_keeper, daemon=True).start()
    while True:
        conn, _ = server.accept()
        threading.Thread(target=serve_connection, args=(conn, token), daemon=True).start()


if __name__ == "__main__":
    main()
