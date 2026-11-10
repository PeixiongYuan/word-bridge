#!/usr/bin/env python3
"""Word 桥接客户端（远程主机端）。

Claude Code 把它当作一个 stdio 的 MCP 服务来启动：
    claude mcp add word-bridge -- python3 ~/.word-bridge/client.py

它把每个请求经 SSH 隧道里的 Unix 套接字转给 Mac 上的桥接服务，再把回复原样交还。
隧道暂时不通时不会退出：工具列表用本地缓存作答，工具调用返回一条说明，隧道恢复后自动重连。

只依赖 Python 标准库。  用法：client.py [--ping] [--socket PATH]
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
    "Word 桥接当前连不上用户的 Mac。可能的原因：Mac 休眠或断网、Mac 到这台主机的 SSH 连接断了、"
    "或 Mac 上的桥接服务没在运行。请告诉用户：在 Mac 上重新建立到这台主机的 SSH 连接，"
    "桥接会在半分钟内自动恢复；然后再重试这次操作。在此之前不要改用直接改写 docx 文件的办法。"
)


def socket_path():
    if "--socket" in sys.argv:
        return sys.argv[sys.argv.index("--socket") + 1]
    return os.environ.get("WORD_BRIDGE_SOCKET", os.path.join(BASE, "bridge.sock"))


class Link:
    """到 Mac 桥接服务的一条连接，断了就在下次请求时重连。"""

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
        """发送一行；需要回复时返回回复行。失败时抛 OSError。"""
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
    """(宽, 高, dpi)；dpi 没有记录时为 None。"""
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
    """返回写入了指定 dpi 的 PNG（替换或新增 pHYs 块），像素不变。"""
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
    """把 PDF 的第一页渲染成指定 dpi 的 PNG，返回字节。"""
    tool, exe = find_renderer()
    if exe is None:
        raise ImageError("这台机器上没有 pdftocairo 或 pdftoppm，无法把 PDF 渲染成 PNG。请直接提供 PNG。")
    with tempfile.TemporaryDirectory(prefix="word-bridge-") as tmp:
        prefix = os.path.join(tmp, "page")
        cmd = [exe, "-png", "-r", str(dpi), "-f", "1", "-l", "1", "-singlefile", pdf_path, prefix]
        try:
            proc = subprocess.run(cmd, capture_output=True, timeout=240)
        except subprocess.TimeoutExpired:
            raise ImageError(f"用 {tool} 渲染 {os.path.basename(pdf_path)} 超时。")
        out_path = prefix + ".png"
        if proc.returncode != 0 or not os.path.isfile(out_path):
            raise ImageError(f"{tool} 渲染失败：" + proc.stderr.decode("utf-8", "replace").strip()[:200])
        with open(out_path, "rb") as fh:
            data = fh.read()
    return png_with_dpi(data, dpi)


def prepare_image(args):
    """读取 image_path，按默认规则得到 PNG，并把数据写进调用参数。返回给模型看的说明。"""
    raw = args.get("image_path")
    if not raw:
        raise ImageError("缺少 image_path。")
    try:
        dpi = int(args.get("dpi") or DEFAULT_DPI)
    except (TypeError, ValueError):
        raise ImageError("dpi 必须是整数。")
    if not 72 <= dpi <= 2400:
        raise ImageError("dpi 应在 72 到 2400 之间。")
    path = os.path.abspath(os.path.expanduser(raw))
    if not os.path.isfile(path):
        raise ImageError(f"找不到图片文件：{path}")
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        data = render_pdf(path, dpi)
        note = f"由 {os.path.basename(path)} 渲染为 {dpi} dpi 的 PNG。"
    elif ext == ".png":
        with open(path, "rb") as fh:
            data = fh.read()
        info = png_info(data)
        if info is None:
            raise ImageError(f"{os.path.basename(path)} 不是有效的 PNG 文件。")
        have = info[2]
        sibling = os.path.splitext(path)[0] + ".pdf"
        if (have is None or have < dpi) and os.path.isfile(sibling) and find_renderer()[1]:
            data = render_pdf(sibling, dpi)
            note = (f"所给 PNG 的分辨率为 {have if have else '未记录'} dpi，低于 {dpi} dpi；"
                    f"已改用同名的 {os.path.basename(sibling)} 渲染为 {dpi} dpi 的 PNG。")
        elif have is None or have < dpi:
            note = (f"注意：所给 PNG 的分辨率为 {have if have else '未记录'} dpi，低于默认的 {dpi} dpi，"
                    "旁边也没有可用来重新渲染的同名 PDF，已按原样使用。")
        else:
            note = f"使用所给的 PNG（{have} dpi）。"
    else:
        raise ImageError("只支持 PNG 或 PDF。请把图片导出为 PNG（默认 900 dpi）或提供矢量 PDF。")
    info = png_info(data)
    if info is None:
        raise ImageError("得到的图片不是有效的 PNG。")
    if info[0] * info[1] > MAX_PIXELS:
        raise ImageError(f"图片有 {info[0]} × {info[1]} 像素，超过上限。请降低 dpi 或裁剪页面。")
    if len(data) > MAX_IMAGE_BYTES:
        raise ImageError(f"图片文件有 {len(data) / 1e6:.0f} MB，超过上限。")
    args["image_data"] = base64.b64encode(data).decode("ascii")
    args["image_name"] = os.path.basename(path)
    args["image_path"] = path
    args["image_note"] = note
    return note


def fallback(msg):
    """隧道不通时本地作答，保证 Claude Code 不会把这个 MCP 服务判为失败。"""
    method = msg.get("method")
    params = msg.get("params") or {}
    if method == "initialize":
        result = {"protocolVersion": params.get("protocolVersion", "2024-11-05"),
                  "capabilities": {"tools": {}},
                  "serverInfo": {"name": "word-bridge", "version": "1.0.0"},
                  "instructions": "Word 桥接：隧道暂时不通，工具调用会返回说明。"}
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
            # 图片在这台机器上：先在本地读取并处理成 PNG，再随调用一起发给 Mac
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
