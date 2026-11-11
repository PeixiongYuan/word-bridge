#!/usr/bin/env python3
"""Word bridge server (Mac side).

Lets a remote Claude make tracked edits, through MCP tools, in documents that are open in Word on this Mac.

- Listens on a Unix socket only this user can reach; a connection presents the token first, then speaks MCP (newline-delimited JSON-RPC).
- Only documents listed in config.json can be touched (by file name, or pinned to one file by full path or URL),
  and only tracked edits and comments are made;
  it never saves or closes a document, and never accepts or rejects a revision.
- Every write is recorded in audit.log.
- A background thread keeps the SSH reverse tunnel to the remote host alive. It only reuses the master connection
  you already opened, does nothing when that is gone, and never tries to log in by itself.

Standard library only.
"""
import base64
import contextlib
import difflib
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
import unicodedata
import urllib.parse

BASE = os.environ.get("WORD_BRIDGE_HOME") or os.path.expanduser("~/.word-bridge")
CONFIG_PATH = os.path.join(BASE, "config.json")
TOKEN_PATH = os.path.join(BASE, "token")
AUDIT_PATH = os.path.join(BASE, "audit.log")
SCRIPT_PATH = os.path.join(BASE, "word_ops.applescript")

US, RS, GS = "\x1f", "\x1e", "\x1d"
SERVER_INFO = {"name": "word-bridge", "version": "1.4.0"}
MAX_IMAGE_BYTES = 120 * 1024 * 1024
KNOWN_PROTOCOLS = ("2025-06-18", "2025-03-26", "2024-11-05")

INSTRUCTIONS = (
    "These tools act directly on documents that are ALREADY OPEN in Word on the user's Mac; every change shows up at once in the window the user is looking at.\n"
    "Rules:\n"
    "1. Every edit is written as a tracked change for the user to accept or reject. You cannot save or close a document, or accept or reject revisions.\n"
    "2. Before editing, read the CURRENT content with word_read_document or word_find_text. The user may have just changed it, "
    "so do not rely on the .docx on disk or on text you remember.\n"
    "3. find/anchor text must match the document character for character and occur exactly once. If it occurs more than once, add context to make it unique, "
    "or pick one with occurrence. Prefer snippets that stay within one paragraph.\n"
    "4. Change only the small piece that needs changing instead of rewriting whole paragraphs, so the tracked changes stay readable.\n"
    "5. Pictures: word_list_images shows the inline pictures, word_replace_image replaces one whole, and word_insert_image adds a new one "
    "next to an anchor text (for a figure, in front of the start of its caption). The default is a 900 dpi PNG. "
    "In text you read, a picture shows up as a single / character. "
    "Table structure and layout are beyond these tools, so tell the user when those are needed.\n"
    "6. Text you read does not include tracked deletions: it is the document as it would be after accepting all changes.\n"
    "7. Fields (citations such as Zotero's, the bibliography, cross-references, a table of contents, hyperlinks) are protected. Text around a citation can be edited "
    "and the citation stays intact; an edit that would rewrite or delete a field is refused unless you pass allow_field_changes. "
    "A citation shows up in the text as its number or label, for example the 3,4 in 'levels3,4.'; keep it unchanged in replace.\n"
    "8. Every write is checked by reading the document back, and the result shows the passage as it now reads. "
    "If a result carries a WARNING, stop editing that document and tell the user."
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

    def __init__(self, message, code=None, detail=""):
        super().__init__(message)
        self.code = code
        self.detail = detail


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
    raise ToolError(explain_error(code, detail), code, detail)


def explain_error(code, detail):
    if code == "not_open":
        return "This document is not open in Word right now. Ask the user to open it first."
    if code == "duplicate_name":
        return DUPLICATE_REFUSAL
    if code == "wrong_file":
        return WRONG_FILE_REFUSAL
    if code == "lost_document":
        return LOST_REFUSAL
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
                "This can happen when the snippet has a footnote mark, a combining character or an existing revision at its edge. Use a nearby snippet as the anchor. "
                f"(Read at that position: {detail[:80]!r})")
    if code == "in_field":
        return ("This spot is inside a field (a citation, the bibliography, a cross-reference, a table of contents or a hyperlink), so nothing was changed. "
                "Word regenerates what a field shows, and an edit there would be lost or would break the field. Use text outside the field instead. "
                "To change the field's own text on purpose, call word_replace_text with allow_field_changes=true.")
    if code == "field_in_range":
        return FIELD_REFUSAL
    if code == "bad_segments":
        return "Internal error: the parts of this edit overlap. Nothing was changed."
    if code == "badindex":
        return f"Picture index out of range: the document has only {detail} inline pictures. Check with word_list_images first."
    if code == "stale":
        return "The pictures in the document moved in the meantime, so nothing was replaced. Call word_list_images again and retry."
    if code == "image_failed":
        stage, _, msg = detail.partition(US)
        hint = {"start": "Nothing was changed.", "inserted": "The new picture was inserted, but it was not resized and the old one was not deleted; ask the user to check that spot.",
                "resized": "The new picture was inserted and resized, but the old one was not deleted; ask the user to check that spot."}.get(stage, "")
        return f"Word reported an error while replacing the picture (stage {stage}): {msg[:300]} {hint}"
    if code == "insert_image_failed":
        stage, _, msg = detail.partition(US)
        hint = {"start": "Nothing was changed.",
                "marks": "The paragraph mark for the picture was inserted, but the picture itself was not; ask the user to check that spot.",
                "inserted": "The picture was inserted, but it could not be found again to set its size; ask the user to check that spot.",
                "sized": "The picture was inserted and sized."}.get(stage, "")
        return f"Word reported an error while inserting the picture (stage {stage}): {msg[:300]} {hint}"
    if code == "write_failed":
        done, _, msg = detail.partition(US)
        partial = "" if done in ("", "0") else f" {done} part(s) of this edit had already been written; ask the user to check that passage."
        return "Word refused this write: " + msg[:300] + partial
    if code == "exception":
        return "The Word script failed: " + detail[:300]
    return f"Unknown error {code}: {detail[:200]}"


FIELD_REFUSAL = (
    "This edit would rewrite or delete a field (a citation, a cross-reference or a hyperlink) that lies inside the text to replace, so nothing was changed. "
    "Keep the field's visible text (for a citation, its number or label) unchanged and in the same order in replace, or edit the text before and after it in separate calls. "
    "To delete or overwrite the field on purpose, repeat the call with allow_field_changes=true.")


DUPLICATE_REFUSAL = (
    "Two documents with this file name are open in Word at the same time. Word's scripting interface cannot tell them apart reliably, "
    "so nothing was read or changed. Ask the user to close the copy that should not be edited, then retry.")
WRONG_FILE_REFUSAL = (
    "A document with this file name is open in Word, but it is not the file on the allow-list: same name, different location. "
    "Nothing was read or changed. Ask the user to open the allowed file instead, or to update allowed_documents in ~/.word-bridge/config.json on the Mac.")


LOST_REFUSAL = (
    "This document is open, but Word's scripting interface has lost track of it, which happens after two documents with the same file name "
    "were open and one of them was closed. Nothing was read or changed. Ask the user to close this document and open it again.")


def is_location(entry):
    """An allow-list entry that pins one file: a path or a URL. Anything else is a bare file name."""
    return "/" in entry or entry.startswith("~")


def entry_name(entry):
    """The file name an allow-list entry stands for."""
    if "://" in entry:
        return urllib.parse.unquote(entry.rstrip("/").rsplit("/", 1)[-1])
    if is_location(entry):
        return os.path.basename(os.path.expanduser(entry))
    return entry


def same_location(entry, place):
    """Whether an allow-list path or URL and a location reported by Word mean the same file."""
    def norm(text):
        return unicodedata.normalize("NFC", text)
    if not place:
        return False
    if "://" in entry or "://" in place:
        return norm(urllib.parse.unquote(entry)).rstrip("/") == norm(urllib.parse.unquote(place)).rstrip("/")
    wanted = os.path.expanduser(entry)
    try:
        if os.path.exists(wanted) and os.path.exists(place):
            return os.path.samefile(wanted, place)
    except OSError:
        pass
    return norm(os.path.realpath(wanted)) == norm(os.path.realpath(place))


def open_documents():
    """The documents open in Word, as {file name: [locations]}. A name with more than one location is open more than once.
    A location of None marks a document that Word can no longer find by name."""
    rest = run_word("list")
    doc_part, _, win_part = rest.partition(GS)
    counts, by_doc, by_window = {}, {}, {}
    for row in doc_part.split(RS):
        if row:
            name, place, reachable = (row.split(US) + ["", "1"])[:3]
            counts[name] = counts.get(name, 0) + 1
            by_doc.setdefault(name, []).append(place if reachable == "1" else None)
    for row in win_part.split(RS):
        if row:
            name, _, place = row.partition(US)
            if place not in by_window.setdefault(name, []):
                by_window[name].append(place)
    opened = {}
    for name, n in counts.items():
        if n == 1:
            opened[name] = by_doc[name][:1]
        else:
            # By name, Word reports the same document n times; the windows show the real locations
            places = by_window.get(name, [])
            opened[name] = places if len(places) >= n else places + [""] * (n - len(places))
    return opened


def doc_state(name, entries, opened):
    """How the allowed document of this name stands in Word: (state, location). States: closed, duplicate, wrong_file, ok."""
    places = opened.get(name, [])
    if not places:
        return "closed", ""
    if len(places) > 1:
        return "duplicate", ""
    place = places[0]
    if place is None:
        return "lost", ""
    if any(not is_location(e) or same_location(e, place) for e in entries):
        return "ok", place
    return "wrong_file", place


def check_doc(requested):
    """Check a tool's document argument against the allow-list and against what is open in Word.

    Returns (config, document spec for word_ops). The spec carries the location seen now, so that the script
    refuses if a different document of the same name has taken its place by the time it runs."""
    cfg = load_config()
    allowed = cfg.get("allowed_documents", [])
    name = entry_name(requested) if is_location(requested) else requested
    entries = [e for e in allowed if entry_name(e) == name]
    if is_location(requested):
        entries = [e for e in entries if is_location(e) and same_location(e, os.path.expanduser(requested))]
    if not entries:
        names = sorted({entry_name(e) for e in allowed})
        raise ToolError(
            f"Document {requested!r} is not on the allow-list. Allowed documents: {', '.join(names)}. "
            "To add one, ask the user to edit ~/.word-bridge/config.json on the Mac.")
    state, place = doc_state(name, entries, open_documents())
    if state == "closed":
        raise ToolError(explain_error("not_open", ""), "not_open")
    if state == "duplicate":
        raise ToolError(DUPLICATE_REFUSAL, "duplicate_name")
    if state == "wrong_file":
        raise ToolError(WRONG_FILE_REFUSAL, "wrong_file")
    if state == "lost":
        raise ToolError(LOST_REFUSAL, "lost_document")
    return cfg, name + GS + place


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
    opened = open_documents()
    notes = {"closed": "not open", "ok": "open in Word",
             "duplicate": "open more than once under this name; the user has to close one copy before it can be read or edited",
             "wrong_file": "a document of this name is open, but it is a different file from the allowed one",
             "lost": "open, but Word's scripting interface has lost track of it; the user has to close it and open it again"}
    lines = ["Allowed documents and their state:"]
    for name in dict.fromkeys(entry_name(e) for e in allowed):
        state, _ = doc_state(name, [e for e in allowed if entry_name(e) == name], opened)
        lines.append(f"- {name}: {notes[state]}")
    return "\n".join(lines)


def tool_read_document(args):
    name = args["document"]
    _, spec = check_doc(name)
    start = max(1, int(args.get("start_paragraph", 1)))
    count = max(1, min(200, int(args.get("max_paragraphs", 40))))
    info = read_doc(spec)
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
    _, spec = check_doc(name)
    query = to_word(args["query"])
    if not query:
        raise ToolError("query must not be empty.")
    ctx = max(0, min(300, int(args.get("context_chars", 60))))
    max_results = max(1, min(50, int(args.get("max_results", 20))))
    text = read_doc(spec)["text"]
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


def split_around_fields(find, replace, fields):
    """Split the change from find to replace into parts that leave the fields inside find untouched.

    fields: (visible text between the start of find and the field's result, the result text) for each field.
    Returns a list of (prefix, old, new), or None when replace does not keep every field's text in place."""
    spans = []
    for lead, shown in fields:
        a, b = len(lead), len(lead) + len(shown)
        if not shown or find[:a] != lead or find[a:b] != shown or (spans and a < spans[-1][1]):
            return None
        spans.append((a, b))
    blocks = difflib.SequenceMatcher(None, find, replace, autojunk=False).get_matching_blocks()
    mapped = []
    for a, b in spans:
        hit = next((blk for blk in blocks if blk.a <= a and b <= blk.a + blk.size), None)
        if hit is None:
            return None
        new_a = hit.b + (a - hit.a)
        mapped.append((a, b, new_a, new_a + (b - a)))
    mapped.append((len(find), len(find), len(replace), len(replace)))
    segments = []
    old_at = new_at = 0
    for a, b, new_a, new_b in mapped:
        if new_a < new_at:
            return None
        old_piece, new_piece = find[old_at:a], replace[new_at:new_a]
        if old_piece != new_piece:
            lead, middle_old, middle_new = tight_diff(old_piece, new_piece)
            segments.append((find[:old_at] + lead, middle_old, middle_new))
        old_at, new_at = b, new_b
    return segments or None


def run_edit(spec, find, segments, occurrence, cfg, expect_after, allow_fields):
    flat = [part for seg in segments for part in seg]
    rest = run_word("edit", spec, find, occurrence, cfg.get("author", "Claude"), "1" if allow_fields else "0",
                    expect_after, len(segments), *flat)
    n, positions, n_rev, post, ctx_a, ctx_b = rest.split(US, 5)
    return {"positions": [[int(x) for x in pair.split("-")] for pair in positions.split(",") if pair],
            "revisions": int(n_rev), "verified": post == "ok", "before": ctx_a, "after": ctx_b}


def apply_edit(spec, find, replace, segments, occurrence, cfg, kind, allow_fields=False):
    """Write the segments as tracked changes. Fields are protected first; allow_fields is only used when that fails."""
    kept_fields = 0
    try:
        res = run_edit(spec, find, segments, occurrence, cfg, replace, False)
    except ToolError as exc:
        if exc.code == "field_in_range":
            fields = [tuple((row.split(GS) + [""])[:2]) for row in exc.detail.split(RS) if row]
            split = split_around_fields(find, replace, fields)
            if split:
                segments, kept_fields = split, len(fields)
                res = run_edit(spec, find, segments, occurrence, cfg, replace, False)
            elif allow_fields:
                res = run_edit(spec, find, segments, occurrence, cfg, replace, True)
            else:
                raise
        elif exc.code == "in_field" and allow_fields:
            res = run_edit(spec, find, segments, occurrence, cfg, replace, True)
        else:
            raise
    name, _, place = spec.partition(GS)
    audit({"tool": kind, "document": name, "location": place, "find": find, "occurrence": occurrence,
           "changes": [{"deleted": old, "inserted": new, "position": pos}
                       for (_, old, new), pos in zip(segments, res["positions"])],
           "verified": res["verified"]})
    res["segments"] = segments
    res["kept_fields"] = kept_fields
    return res


def visible(text):
    return text.replace("\r", "⏎")


def readback_note(res, new_text):
    """What the passage reads like after the write, or a warning when the read-back did not match."""
    if not res["verified"]:
        return ("\nWARNING: the change was written, but when the document was read back its text did not match what this edit should have produced. "
                "The change may have landed in the wrong place, or the user may have been typing at the same moment. "
                "Stop editing this document, look for the new text with word_find_text, and tell the user.")
    shown = new_text if len(new_text) <= 320 else new_text[:150] + " … " + new_text[-150:]
    return f"\nVerified by reading the document back. The passage now reads: …{visible(res['before'])}«{visible(shown)}»{visible(res['after'])}…"


def tool_replace_text(args):
    name = args["document"]
    cfg, spec = check_doc(name)
    find = to_word(args["find"])
    replace = to_word(args.get("replace", ""))
    occurrence = max(0, int(args.get("occurrence", 0)))
    allow_fields = bool(args.get("allow_field_changes", False))
    if not find:
        raise ToolError("find must not be empty. Use word_insert_text to insert text.")
    if find == replace:
        raise ToolError("find and replace are identical, so there is nothing to change.")
    prefix, middle_old, middle_new = tight_diff(find, replace)
    try:
        res = apply_edit(spec, find, replace, [(prefix, middle_old, middle_new)], occurrence, cfg, "replace_text", allow_fields)
    except ToolError as exc:
        # If verification fails on the tightened range, fall back once to replacing the whole match
        if exc.code != "verify_failed" or (prefix == "" and middle_old == find):
            raise
        res = apply_edit(spec, find, replace, [("", find, replace)], occurrence, cfg, "replace_text", allow_fields)
    segments = res["segments"]
    if len(segments) == 1:
        _, middle_old, middle_new = segments[0]
        what = "Deleted" if middle_new == "" else ("Inserted" if middle_old == "" else "Replaced")
        head = f"{what} as a tracked change. Deleted: {middle_old!r}; inserted: {middle_new!r}."
    else:
        parts = "; ".join(f"({i}) deleted {old!r}, inserted {new!r}" for i, (_, old, new) in enumerate(segments, 1))
        head = (f"Changed {len(segments)} separate pieces as tracked changes, leaving the {res['kept_fields']} field(s) "
                f"(citations or links) between them untouched: {parts}.")
    return (f"{head} The document now has {res['revisions']} pending revisions for the user to review."
            + readback_note(res, replace))


def tool_insert_text(args):
    name = args["document"]
    cfg, spec = check_doc(name)
    anchor = to_word(args["anchor"])
    text = to_word(args["text"])
    where = args.get("position", "after")
    occurrence = max(0, int(args.get("occurrence", 0)))
    if not anchor or not text:
        raise ToolError("anchor and text must not be empty.")
    if where not in ("after", "before"):
        raise ToolError("position must be after or before.")
    prefix = anchor if where == "after" else ""
    now_reads = anchor + text if where == "after" else text + anchor
    res = apply_edit(spec, anchor, now_reads, [(prefix, "", text)], occurrence, cfg, "insert_text")
    return (f"Inserted as a tracked change {'after' if where == 'after' else 'before'} the anchor text: {text!r}. "
            f"The document now has {res['revisions']} pending revisions." + readback_note(res, now_reads))


def tool_add_comment(args):
    name = args["document"]
    cfg, spec = check_doc(name)
    anchor = to_word(args["anchor"])
    comment = args["comment"]
    occurrence = max(0, int(args.get("occurrence", 0)))
    if not anchor or not comment:
        raise ToolError("anchor and comment must not be empty.")
    rest = run_word("comment", spec, anchor, occurrence, comment, cfg.get("author", "Claude"))
    _, n_com = rest.split(US)
    audit({"tool": "add_comment", "document": name, "location": spec.partition(GS)[2], "anchor": anchor, "comment": comment, "occurrence": occurrence})
    return f"Comment added on that text. The document now has {n_com} comments."


def tool_list_changes(args):
    name = args["document"]
    _, spec = check_doc(name)
    max_items = max(1, min(300, int(args.get("max_items", 100))))
    rest = run_word("revisions", spec, max_items)
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


def list_images(spec):
    rest = run_word("images", spec)
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
    _, spec = check_doc(name)
    n_inline, n_floating, images = list_images(spec)
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


COMPRESSION_NOTE = (
    "Tell the user: unless picture compression is turned off for this document, Word downsamples pictures to about 220 ppi when it saves. "
    "To keep the full resolution, tick “Do not compress images in file” under Word > Preferences > Edit > Image Size and Quality.")


def image_from_call(args):
    """The PNG that client.py attached to the call: (bytes, width in pixels, height in pixels, dpi or None)."""
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
    return (data,) + info


def tool_insert_image(args):
    name = args["document"]
    cfg, spec = check_doc(name)
    anchor = to_word(args.get("anchor") or "")
    where = args.get("position", "before")
    occurrence = max(0, int(args.get("occurrence", 0)))
    own_paragraph = bool(args.get("own_paragraph", True))
    if not anchor:
        raise ToolError("anchor must not be empty: give a piece of existing text next to where the picture should go, for example the start of its caption.")
    if where not in ("after", "before"):
        raise ToolError("position must be before or after.")
    try:
        width_cm = float(args.get("width_cm") or 0)
    except (TypeError, ValueError):
        raise ToolError("width_cm must be a number.")
    if width_cm and not 0.5 <= width_cm <= 55:
        raise ToolError("width_cm must be between 0.5 and 55.")
    data, px_w, px_h, dpi = image_from_call(args)
    with serve_image(data) as url:
        rest = run_word("insertimage", spec, anchor, occurrence, where, url, cfg.get("author", "Claude"),
                        f"{width_cm * 72 / 2.54:.2f}", "1" if own_paragraph else "0", timeout=240)
    index, nat_w, _, new_w, new_h, n_rev, post, n_lead, n_trail, ctx_a, ctx_b = rest.split(US, 10)
    verified = post == "ok"
    audit({"tool": "insert_image", "document": name, "location": spec.partition(GS)[2], "anchor": anchor, "where": where,
           "index": int(index), "image": args.get("image_name"), "source": args.get("image_path"), "bytes": len(data),
           "pixels": [px_w, px_h], "dpi": dpi, "verified": verified})
    note = args.get("image_note") or ""
    fitted = "" if width_cm or pt_to_cm(new_w) >= pt_to_cm(nat_w) - 0.02 else " It was scaled down to fit the width of the text area."
    placing = "in a paragraph of its own" if own_paragraph else "inline with the text"
    if verified:
        check = (f"Verified by reading the document back. Around the picture (shown as «/») the text now reads: "
                 f"…{visible(ctx_a)}{'⏎' * int(n_lead)}«/»{'⏎' * int(n_trail)}{visible(ctx_b)}…")
    else:
        check = ("WARNING: the picture was inserted, but when the document was read back its text did not match what this insertion should have produced. "
                 "It may have landed in the wrong place, or the user may have been typing at the same moment. "
                 "Stop editing this document, check with word_list_images, and tell the user.")
    return (f"Inserted a new picture (now #{index}) as a tracked change {where} the anchor text, {placing}.\n"
            f"New image: {px_w} × {px_h} pixels, {str(dpi) + ' dpi' if dpi else 'no dpi recorded'}. {note}\n"
            f"Displayed size: {pt_to_cm(new_w):.2f} × {pt_to_cm(new_h):.2f} cm.{fitted} The document now has {n_rev} pending revisions.\n"
            f"{check}\n{COMPRESSION_NOTE}")


def tool_replace_image(args):
    name = args["document"]
    cfg, spec = check_doc(name)
    data, px_w, px_h, dpi = image_from_call(args)

    n_inline, _, images = list_images(spec)
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
        rest = run_word("locate", spec, near, max(0, int(args.get("occurrence", 0))))
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
        rest = run_word("replaceimage", spec, target["index"], target["start"], url, cfg.get("author", "Claude"), timeout=240)
    old_w, old_h, _, _, new_w, new_h, n_rev = rest.split(US)
    audit({"tool": "replace_image", "document": name, "location": spec.partition(GS)[2], "index": target["index"], "image": args.get("image_name"),
           "source": args.get("image_path"), "bytes": len(data), "pixels": [px_w, px_h], "dpi": dpi})
    note = args.get("image_note") or ""
    below = clean_context(target["below"]).lstrip("/⏎ ")[:60]
    return (f"Replaced picture #{target['index']} as a tracked change (text below it: {below}…).\n"
            f"New image: {px_w} × {px_h} pixels, {str(dpi) + ' dpi' if dpi else 'no dpi recorded'}. {note}\n"
            f"Displayed size: {pt_to_cm(old_w):.2f} × {pt_to_cm(old_h):.2f} cm → {pt_to_cm(new_w):.2f} × {pt_to_cm(new_h):.2f} cm (width kept from the old picture, height from the new image's proportions).\n"
            f"The old picture stays as a tracked deletion until the user accepts it. The document now has {n_rev} pending revisions.\n"
            + COMPRESSION_NOTE)


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
                     "Only the part that actually differs between find and replace is marked as a revision. "
                     "A citation or other field inside find is left untouched as long as replace keeps its visible text; the result shows the passage as read back from Word."),
     "inputSchema": {"type": "object", "properties": {
         "document": DOC_PROP,
         "find": {"type": "string", "description": "Existing text in the document, with enough context to be unique."},
         "replace": {"type": "string", "description": "The replacement text. An empty string deletes."},
         "occurrence": OCC_PROP,
         "allow_field_changes": {"type": "boolean", "description": "Set to true only to delete or overwrite a field (a citation, cross-reference or hyperlink) covered by find on purpose. Default false: such an edit is refused.", "default": False}},
         "required": ["document", "find", "replace"]},
     "handler": tool_replace_text},
    {"name": "word_insert_text",
     "description": ("Insert new text after (or before) an anchor text, as a tracked change. Use a line break in text to start a new paragraph. "
                     "Next to a citation or other field the text goes outside the field; the result shows the passage as read back from Word."),
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
    {"name": "word_insert_image",
     "description": ("Insert a NEW picture next to an anchor text, as a tracked change. For a figure, use the start of its caption as the anchor with position before: "
                     "the picture then gets a paragraph of its own above the caption. "
                     "image_path is a file path ON THE MACHINE YOU ARE RUNNING ON; the local client reads the file and sends it. "
                     "The default is a 900 dpi PNG: a PDF is rendered to a 900 dpi PNG, and a PNG below that resolution is swapped for a PDF of the same name next to it, "
                     "if there is one. Without width_cm the picture keeps its natural size, scaled down if needed to fit the width of the text area. "
                     "To swap a picture that is already there, use word_replace_image instead."),
     "inputSchema": {"type": "object", "properties": {
         "document": DOC_PROP,
         "image_path": {"type": "string", "description": "Path of the image on your machine, PNG or PDF."},
         "anchor": {"type": "string", "description": "A piece of existing text next to where the picture should go (matching character for character, and unique), usually the start of the figure's caption."},
         "position": {"type": "string", "enum": ["before", "after"], "description": "Put the picture before or after the anchor; default before (figure above its caption).", "default": "before"},
         "occurrence": OCC_PROP,
         "width_cm": {"type": "number", "description": "Displayed width in centimetres; the height follows the image's proportions. Omit to keep the natural size, limited to the width of the text area."},
         "own_paragraph": {"type": "boolean", "description": "Give the picture a paragraph of its own (default true). Set false to place it inline within the text.", "default": True},
         "dpi": {"type": "integer", "description": "Target resolution; default 900.", "default": 900}},
         "required": ["document", "image_path", "anchor"]},
     "handler": tool_insert_image},
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

def print_open_documents():
    """For the user: the open documents and the location string that pins each one in allowed_documents."""
    opened = open_documents()
    if not opened:
        print("No document is open in Word.")
        return
    print("Documents open in Word. To pin an allow-list entry to one file, put its location in allowed_documents instead of the bare name.")
    for name, places in opened.items():
        note = "   (open more than once: the bridge refuses this name until one copy is closed)" if len(places) > 1 else ""
        print(f"\n  {name}{note}")
        for place in places:
            print("      " + ("(Word has lost track of this document; close it and open it again)" if place is None else place or "(not saved yet)"))


def main():
    if "--list-open" in sys.argv:
        try:
            print_open_documents()
        except ToolError as exc:
            sys.exit(str(exc))
        return
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
    log(f"word-bridge listening on {sock_path}; allowed documents: {[entry_name(e) for e in cfg.get('allowed_documents', [])]}")
    if "--no-tunnel" not in sys.argv:
        threading.Thread(target=tunnel_keeper, daemon=True).start()
    while True:
        conn, _ = server.accept()
        threading.Thread(target=serve_connection, args=(conn, token), daemon=True).start()


if __name__ == "__main__":
    main()
