#!/usr/bin/env python3
"""Word bridge client (remote host side).

Claude Code starts it as a stdio MCP server:
    claude mcp add word-bridge -- python3 ~/.word-bridge/client.py

It forwards each request through a Unix socket inside the SSH tunnel to the bridge server on the Mac and hands the reply back unchanged.
It does not exit while the tunnel is down: the tool list is answered from a local cache, tool calls return an explanation, and it reconnects by itself.

Standard library only.  Usage: client.py [--ping] [--socket PATH]
"""
import base64
import json
import os
import shutil
import socket
import struct
import subprocess
import sys
import tempfile
import zlib

BASE = os.environ.get("WORD_BRIDGE_HOME") or os.path.expanduser("~/.word-bridge")
TOKEN_PATH = os.path.join(BASE, "token")
CACHE_PATH = os.path.join(BASE, "tools_cache.json")
RESPONSE_TIMEOUT = 300
DEFAULT_DPI = 900
MAX_IMAGE_BYTES = 120 * 1024 * 1024
MAX_PIXELS = 250 * 1000 * 1000
PNG_MAGIC = b"\x89PNG\r\n\x1a\n"

DOWN_MESSAGE = (
    "The Word bridge cannot reach the user's Mac right now. Possible causes: the Mac is asleep or offline, the SSH connection from the Mac to this host has dropped, "
    "or the bridge server on the Mac is not running. Tell the user to re-establish the SSH connection from the Mac to this host; "
    "the bridge recovers by itself within about half a minute, then retry. Until then, do not fall back to rewriting the .docx file directly."
)


def socket_path():
    if "--socket" in sys.argv:
        return sys.argv[sys.argv.index("--socket") + 1]
    return os.environ.get("WORD_BRIDGE_SOCKET", os.path.join(BASE, "bridge.sock"))


class Link:
    """One connection to the bridge server on the Mac; reconnects on the next request if it has dropped."""

    def __init__(self):
        self.sock = None
        self.reader = None

    def connect(self):
        with open(TOKEN_PATH, encoding="utf-8") as fh:
            token = fh.read().strip()
        sock = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        sock.settimeout(10)
        sock.connect(socket_path())
        sock.sendall(("AUTH " + token + "\n").encode("utf-8"))
        sock.settimeout(RESPONSE_TIMEOUT)
        self.sock = sock
        self.reader = sock.makefile("r", encoding="utf-8", newline="\n")

    def close(self):
        try:
            if self.sock is not None:
                self.sock.close()
        except OSError:
            pass
        self.sock = None
        self.reader = None

    def request(self, line, expect_reply):
        """Send one line and, when a reply is expected, return the reply line. Raises OSError on failure."""
        for attempt in (1, 2):
            try:
                if self.sock is None:
                    self.connect()
                self.sock.sendall((line + "\n").encode("utf-8"))
                if not expect_reply:
                    return None
                reply = self.reader.readline()
                if not reply:
                    raise OSError("connection closed by the bridge")
                return reply.rstrip("\n")
            except (OSError, ValueError):
                self.close()
                if attempt == 2:
                    raise
        return None


class ImageError(Exception):
    pass


def png_info(data):
    """(width, height, dpi); dpi is None when the file does not record it."""
    if data[:8] != PNG_MAGIC:
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


def png_with_dpi(data, dpi):
    """Return the PNG with the given dpi written into it (replacing or adding the pHYs chunk); the pixels are unchanged."""
    ppm = round(dpi / 0.0254)
    body = struct.pack(">IIB", ppm, ppm, 1)
    phys = struct.pack(">I", 9) + b"pHYs" + body + struct.pack(">I", zlib.crc32(b"pHYs" + body) & 0xffffffff)
    out = [data[:8]]
    pos = 8
    placed = False
    while pos + 8 <= len(data):
        length, kind = struct.unpack(">I4s", data[pos:pos + 8])
        chunk = data[pos:pos + 12 + length]
        if kind == b"pHYs":
            pos += 12 + length
            continue
        out.append(chunk)
        if kind == b"IHDR" and not placed:
            out.append(phys)
            placed = True
        pos += 12 + length
    return b"".join(out)


def find_renderer():
    for tool in ("pdftocairo", "pdftoppm"):
        path = shutil.which(tool) or os.path.expanduser("~/.local/bin/" + tool)
        if os.path.isfile(path) and os.access(path, os.X_OK):
            return tool, path
    return None, None


def render_pdf(pdf_path, dpi):
    """Render the first page of a PDF to a PNG at the given dpi and return the bytes."""
    tool, exe = find_renderer()
    if exe is None:
        raise ImageError("Neither pdftocairo nor pdftoppm is available on this machine, so the PDF cannot be rendered to PNG. Provide a PNG instead.")
    with tempfile.TemporaryDirectory(prefix="word-bridge-") as tmp:
        prefix = os.path.join(tmp, "page")
        cmd = [exe, "-png", "-r", str(dpi), "-f", "1", "-l", "1", "-singlefile", pdf_path, prefix]
        try:
            proc = subprocess.run(cmd, capture_output=True, timeout=240)
        except subprocess.TimeoutExpired:
            raise ImageError(f"Rendering {os.path.basename(pdf_path)} with {tool} timed out.")
        out_path = prefix + ".png"
        if proc.returncode != 0 or not os.path.isfile(out_path):
            raise ImageError(f"{tool} failed to render the PDF: " + proc.stderr.decode("utf-8", "replace").strip()[:200])
        with open(out_path, "rb") as fh:
            data = fh.read()
    return png_with_dpi(data, dpi)


def prepare_image(args):
    """Read image_path, turn it into a PNG by the default rules, and add the data to the call arguments. Returns the note shown to the model."""
    raw = args.get("image_path")
    if not raw:
        raise ImageError("image_path is missing.")
    try:
        dpi = int(args.get("dpi") or DEFAULT_DPI)
    except (TypeError, ValueError):
        raise ImageError("dpi must be an integer.")
    if not 72 <= dpi <= 2400:
        raise ImageError("dpi must be between 72 and 2400.")
    path = os.path.abspath(os.path.expanduser(raw))
    if not os.path.isfile(path):
        raise ImageError(f"Image file not found: {path}")
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        data = render_pdf(path, dpi)
        note = f"Rendered from {os.path.basename(path)} to a {dpi} dpi PNG."
    elif ext == ".png":
        with open(path, "rb") as fh:
            data = fh.read()
        info = png_info(data)
        if info is None:
            raise ImageError(f"{os.path.basename(path)} is not a valid PNG file.")
        have = info[2]
        sibling = os.path.splitext(path)[0] + ".pdf"
        if (have is None or have < dpi) and os.path.isfile(sibling) and find_renderer()[1]:
            data = render_pdf(sibling, dpi)
            note = (f"The PNG given is at {str(have) + ' dpi' if have else 'an unrecorded resolution'}, below {dpi} dpi; "
                    f"rendered {os.path.basename(sibling)}, the PDF of the same name, to a {dpi} dpi PNG instead.")
        elif have is None or have < dpi:
            note = (f"Note: the PNG given is at {str(have) + ' dpi' if have else 'an unrecorded resolution'}, below the default {dpi} dpi, "
                    "and there is no PDF of the same name next to it to render from, so it was used as is.")
        else:
            note = f"Used the PNG as given ({have} dpi)."
    else:
        raise ImageError("Only PNG and PDF are supported. Export the image as a PNG (900 dpi by default) or provide a vector PDF.")
    info = png_info(data)
    if info is None:
        raise ImageError("The resulting image is not a valid PNG.")
    if info[0] * info[1] > MAX_PIXELS:
        raise ImageError(f"The image is {info[0]} × {info[1]} pixels, above the limit. Lower the dpi or crop the page.")
    if len(data) > MAX_IMAGE_BYTES:
        raise ImageError(f"The image file is {len(data) / 1e6:.0f} MB, above the limit.")
    args["image_data"] = base64.b64encode(data).decode("ascii")
    args["image_name"] = os.path.basename(path)
    args["image_path"] = path
    args["image_note"] = note
    return note


def fallback(msg):
    """Answer locally while the tunnel is down, so that Claude Code does not mark this MCP server as failed."""
    method = msg.get("method")
    params = msg.get("params") or {}
    if method == "initialize":
        result = {"protocolVersion": params.get("protocolVersion", "2024-11-05"),
                  "capabilities": {"tools": {}},
                  "serverInfo": {"name": "word-bridge", "version": "1.0.0"},
                  "instructions": "Word bridge: the tunnel is down for now; tool calls will return an explanation."}
    elif method == "ping":
        result = {}
    elif method == "tools/list":
        try:
            with open(CACHE_PATH, encoding="utf-8") as fh:
                result = json.load(fh)
        except (OSError, ValueError):
            result = {"tools": []}
    elif method == "tools/call":
        result = {"content": [{"type": "text", "text": DOWN_MESSAGE}], "isError": True}
    else:
        return {"jsonrpc": "2.0", "id": msg.get("id"), "error": {"code": -32601, "message": "method not found"}}
    return {"jsonrpc": "2.0", "id": msg.get("id"), "result": result}


def ping():
    link = Link()
    try:
        reply = link.request(json.dumps({"jsonrpc": "2.0", "id": 1, "method": "ping"}), True)
        ok = "result" in json.loads(reply)
    except (OSError, ValueError):
        ok = False
    link.close()
    print("ok" if ok else "unreachable")
    return 0 if ok else 1


def main():
    if "--ping" in sys.argv:
        sys.exit(ping())
    link = Link()
    for raw in sys.stdin:
        line = raw.strip()
        if not line:
            continue
        try:
            msg = json.loads(line)
        except ValueError:
            continue
        if not isinstance(msg, dict):
            continue
        expect_reply = "id" in msg and "method" in msg
        params = msg.get("params") if isinstance(msg.get("params"), dict) else {}
        if msg.get("method") == "tools/call" and params.get("name") == "word_replace_image":
            # The image is on this machine: read it and turn it into a PNG here, then send it to the Mac with the call
            try:
                prepare_image(params.setdefault("arguments", {}))
                line = json.dumps(msg, ensure_ascii=False)
            except ImageError as exc:
                if expect_reply:
                    err = {"jsonrpc": "2.0", "id": msg.get("id"),
                           "result": {"content": [{"type": "text", "text": str(exc)}], "isError": True}}
                    sys.stdout.write(json.dumps(err, ensure_ascii=False) + "\n")
                    sys.stdout.flush()
                continue
        try:
            reply = link.request(line, expect_reply)
            if expect_reply and msg.get("method") == "tools/list":
                try:
                    parsed = json.loads(reply)
                    if "result" in parsed:
                        with open(CACHE_PATH, "w", encoding="utf-8") as fh:
                            json.dump(parsed["result"], fh, ensure_ascii=False)
                except (OSError, ValueError):
                    pass
        except (OSError, ValueError):
            reply = json.dumps(fallback(msg), ensure_ascii=False) if expect_reply else None
        if reply is not None:
            sys.stdout.write(reply + "\n")
            sys.stdout.flush()
    link.close()


if __name__ == "__main__":
    main()
