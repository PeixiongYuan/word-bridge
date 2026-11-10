#!/usr/bin/env python3
"""Word 桥接客户端（远程主机端）。

Claude Code 把它当作一个 stdio 的 MCP 服务来启动：
    claude mcp add word-bridge -- python3 ~/.word-bridge/client.py

它把每个请求经 SSH 隧道里的 Unix 套接字转给 Mac 上的桥接服务，再把回复原样交还。
隧道暂时不通时不会退出：工具列表用本地缓存作答，工具调用返回一条说明，隧道恢复后自动重连。

只依赖 Python 标准库。  用法：client.py [--ping] [--socket PATH]
"""
import json
import os
import socket
import sys

BASE = os.environ.get("WORD_BRIDGE_HOME") or os.path.expanduser("~/.word-bridge")
TOKEN_PATH = os.path.join(BASE, "token")
CACHE_PATH = os.path.join(BASE, "tools_cache.json")
RESPONSE_TIMEOUT = 200

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
