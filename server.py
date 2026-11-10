#!/usr/bin/env python3
"""Word 桥接服务（Mac 端）。

让远程的 Claude 通过 MCP 工具，在本机 Word 里已打开的文档上做带修订痕迹的修改。

- 监听一个仅本人可访问的 Unix 套接字，连接后先校验口令，再走 MCP（按行分隔的 JSON-RPC）。
- 只允许操作 config.json 里列出的文档；只做修订模式的修改和批注；
  从不保存、关闭文档，也从不接受或拒绝修订。
- 所有写操作记入 audit.log。
- 另有一个后台线程负责维持到远程主机的 SSH 反向隧道：只复用已登录的主连接，
  主连接断开时什么也不做，绝不自己发起登录。

只依赖 Python 标准库。
"""
import hmac
import json
import os
import socket
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
SERVER_INFO = {"name": "word-bridge", "version": "1.0.0"}
KNOWN_PROTOCOLS = ("2025-06-18", "2025-03-26", "2024-11-05")

INSTRUCTIONS = (
    "这些工具直接操作用户 Mac 上 Word 里【已经打开】的文档，改动会立刻出现在用户眼前的窗口里。\n"
    "规则：\n"
    "1. 所有修改都以修订模式写入，由用户逐条接受或拒绝；你不能保存、关闭文档，也不能接受/拒绝修订。\n"
    "2. 修改前先用 word_read_document 或 word_find_text 读取【当前】内容——用户可能刚改过，"
    "不要依赖磁盘上的 docx 或你记忆里的旧文本。\n"
    "3. find/anchor 必须与文档里的文字逐字符一致，并且只出现一次；出现多次时加长上下文使其唯一，"
    "或用 occurrence 指定第几处。尽量选不跨段落、不含引文域（如 Zotero 引用）的片段。\n"
    "4. 每次只改需要改的那一小段，不要整段重写，这样修订痕迹才清楚。\n"
    "5. 插图、表格结构、版式调整不在这些工具的能力范围内；需要时请告诉用户，在文档关闭后改文件。\n"
    "6. 读到的文本不包含已被修订删除的文字（即“接受全部修订后”的样子）。"
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
    pass


# ---------------------------------------------------------------- Word 调用

def run_word(*args, timeout=150):
    """调用 word_ops.applescript，返回去掉 OK 之后的原始字符串。"""
    with word_lock:
        try:
            # 用字节模式读取：Word 的段落标记是 \r，文本模式会把它悄悄换成 \n
            proc = subprocess.run(
                ["/usr/bin/osascript", SCRIPT_PATH, *[str(a) for a in args]],
                capture_output=True, timeout=timeout,
            )
        except subprocess.TimeoutExpired:
            raise ToolError("Word 在限定时间内没有响应（可能正弹着对话框，或文档很大）。")
    out = proc.stdout.decode("utf-8", errors="replace")
    if out.endswith("\n"):
        out = out[:-1]
    if proc.returncode != 0 and not out:
        raise ToolError("调用 Word 失败：" + proc.stderr.decode("utf-8", errors="replace").strip()[:300])
    head, _, rest = out.partition(US)
    if head == "OK":
        return rest
    code, _, detail = rest.partition(US)
    raise ToolError(explain_error(code, detail))


def explain_error(code, detail):
    if code == "not_open":
        return "这份文档目前没有在 Word 里打开。请用户先打开它。"
    if code == "notfound":
        return ("在文档当前内容里找不到这段文字（需要逐字符一致，包括空格和标点）。"
                "请先用 word_find_text 或 word_read_document 核对当前文本。")
    if code == "ambiguous":
        return (f"这段文字在文档里出现了 {detail} 次，无法确定改哪一处。"
                "请加长片段使其唯一，或用 occurrence 指定第几处（从 1 开始）。")
    if code == "badocc":
        return f"occurrence 超出范围：这段文字只出现了 {detail} 次。"
    if code == "verify_failed":
        return ("定位后复核不一致，为安全起见没有做任何修改。"
                "这通常是因为片段里夹着引文域、脚注标记或已有的修订。请换一段不含这些内容的文字作为定位依据。"
                f"（该位置实际读到：{detail[:80]!r}）")
    if code == "write_failed":
        return "Word 拒绝了这次写入：" + detail[:300]
    if code == "exception":
        return "Word 脚本出错：" + detail[:300]
    return f"未知错误 {code}: {detail[:200]}"


def check_doc(name):
    cfg = load_config()
    allowed = cfg.get("allowed_documents", [])
    if name not in allowed:
        raise ToolError(
            f"文档 {name!r} 不在允许列表里。允许的文档：{', '.join(allowed)}。"
            "需要增加时请让用户修改 Mac 上的 ~/.word-bridge/config.json。")
    return cfg


def to_word(text):
    """调用方的换行统一成 Word 的段落标记。"""
    return text.replace("\r\n", "\r").replace("\n", "\r")


def is_word_char(ch):
    return ch.isascii() and (ch.isalnum() or ch == "_")


def tight_diff(old, new):
    """去掉 old/new 的公共前后缀，让修订只覆盖真正变化的部分；不在英文单词或数字中间下刀。"""
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


# ---------------------------------------------------------------- 工具实现

def tool_list_documents(args):
    cfg = load_config()
    allowed = cfg.get("allowed_documents", [])
    rest = run_word("list")
    opened = [n for n in rest.split(RS) if n]
    lines = ["允许操作的文档及其状态："]
    for name in allowed:
        lines.append(f"- {name}：{'已在 Word 中打开' if name in opened else '未打开'}")
    return "\n".join(lines)


def tool_read_document(args):
    name = args["document"]
    check_doc(name)
    start = max(1, int(args.get("start_paragraph", 1)))
    count = max(1, min(200, int(args.get("max_paragraphs", 40))))
    info = read_doc(name)
    paras = paragraphs_of(info["text"])
    chunk = paras[start - 1:start - 1 + count]
    head = (f"文档 {name}：共 {len(paras)} 段，待处理修订 {info['revisions']} 处，批注 {info['comments']} 条，"
            f"{'已保存' if info['saved'] else '有未保存的修改'}。"
            f"以下是第 {start}–{start + len(chunk) - 1} 段（不含已被修订删除的文字）：")
    body = "\n".join(f"[{start + i}] {p}" for i, p in enumerate(chunk))
    tail = "" if start - 1 + count >= len(paras) else f"\n……后面还有 {len(paras) - (start - 1 + count)} 段，用 start_paragraph 继续读。"
    return head + "\n" + body + tail


def tool_find_text(args):
    name = args["document"]
    check_doc(name)
    query = to_word(args["query"])
    if not query:
        raise ToolError("query 不能为空。")
    ctx = max(0, min(300, int(args.get("context_chars", 60))))
    max_results = max(1, min(50, int(args.get("max_results", 20))))
    text = read_doc(name)["text"]
    hits, pos = [], text.find(query)
    while pos != -1:
        hits.append(pos)
        pos = text.find(query, pos + len(query))
    if not hits:
        return f"在 {name} 的当前内容里没有找到这段文字。"
    lines = [f"在 {name} 中找到 {len(hits)} 处："]
    for idx, pos in enumerate(hits[:max_results], 1):
        para_no = text.count("\r", 0, pos) + 1
        before = text[max(0, pos - ctx):pos].replace("\r", "⏎")
        after = text[pos + len(query):pos + len(query) + ctx].replace("\r", "⏎")
        lines.append(f"{idx}. 第 {para_no} 段：…{before}【{query.replace(chr(13), '⏎')}】{after}…")
    if len(hits) > max_results:
        lines.append(f"（只显示前 {max_results} 处）")
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
        raise ToolError("find 不能为空。要插入文字请用 word_insert_text。")
    if find == replace:
        raise ToolError("find 和 replace 完全相同，没有需要修改的内容。")
    prefix, middle_old, middle_new = tight_diff(find, replace)
    try:
        n_rev = do_edit(name, find, prefix, middle_old, middle_new, occurrence, cfg, "replace_text")
    except ToolError as exc:
        # 精简后的区间复核失败时，退回到整段替换再试一次
        if "复核不一致" not in str(exc) or (prefix == "" and middle_old == find):
            raise
        n_rev = do_edit(name, find, "", find, replace, occurrence, cfg, "replace_text")
        middle_old, middle_new = find, replace
    what = "删除" if middle_new == "" else ("插入" if middle_old == "" else "替换")
    return (f"已以修订模式{what}。删除：{middle_old!r}；插入：{middle_new!r}。"
            f"文档现有 {n_rev} 处待处理修订，等用户审阅。")


def tool_insert_text(args):
    name = args["document"]
    cfg = check_doc(name)
    anchor = to_word(args["anchor"])
    text = to_word(args["text"])
    where = args.get("position", "after")
    occurrence = max(0, int(args.get("occurrence", 0)))
    if not anchor or not text:
        raise ToolError("anchor 和 text 都不能为空。")
    if where not in ("after", "before"):
        raise ToolError("position 只能是 after 或 before。")
    prefix = anchor if where == "after" else ""
    n_rev = do_edit(name, anchor, prefix, "", text, occurrence, cfg, "insert_text")
    return (f"已以修订模式在定位文字{'之后' if where == 'after' else '之前'}插入：{text!r}。"
            f"文档现有 {n_rev} 处待处理修订。")


def tool_add_comment(args):
    name = args["document"]
    cfg = check_doc(name)
    anchor = to_word(args["anchor"])
    comment = args["comment"]
    occurrence = max(0, int(args.get("occurrence", 0)))
    if not anchor or not comment:
        raise ToolError("anchor 和 comment 都不能为空。")
    rest = run_word("comment", name, anchor, occurrence, comment, cfg.get("author", "Claude"))
    _, n_com = rest.split(US)
    audit({"tool": "add_comment", "document": name, "anchor": anchor, "comment": comment, "occurrence": occurrence})
    return f"已在该文字上添加批注。文档现有 {n_com} 条批注。"


def tool_list_changes(args):
    name = args["document"]
    check_doc(name)
    max_items = max(1, min(300, int(args.get("max_items", 100))))
    rest = run_word("revisions", name, max_items)
    head, _, body = rest.partition(RS)
    n_rev, n_com = head.split(US)
    rev_part, _, com_part = body.partition(GS)
    kinds = {"revision insert": "插入", "revision delete": "删除"}
    lines = [f"文档 {name}：待处理修订 {n_rev} 处，批注 {n_com} 条。"]
    if rev_part:
        lines.append("修订：")
        for i, row in enumerate(rev_part.split(RS), 1):
            kind, author, text = (row.split(US) + ["", "", ""])[:3]
            lines.append(f"{i}. [{kinds.get(kind, kind)}] {author}：{text.replace(chr(13), '⏎')[:200]}")
    if com_part:
        lines.append("批注：")
        for i, row in enumerate(com_part.split(RS), 1):
            author, scope, text = (row.split(US) + ["", "", ""])[:3]
            lines.append(f"{i}. {author} 批注于“{scope.replace(chr(13), '⏎')[:60]}”：{text.replace(chr(13), ' ')[:300]}")
    return "\n".join(lines)


DOC_PROP = {"type": "string", "description": "文档文件名，例如 report.docx。先用 word_list_documents 查看可用的文档。"}
OCC_PROP = {"type": "integer", "description": "定位文字出现多次时，指定改第几处（从 1 开始）。省略或 0 表示必须唯一。", "default": 0}

TOOLS = [
    {"name": "word_list_documents",
     "description": "列出允许操作的 Word 文档，以及它们此刻是否已在用户的 Word 里打开。",
     "inputSchema": {"type": "object", "properties": {}},
     "handler": tool_list_documents},
    {"name": "word_read_document",
     "description": "读取用户 Word 里已打开文档的【当前】内容（包含用户尚未保存的修改），按段落编号返回。修改前先读。",
     "inputSchema": {"type": "object", "properties": {
         "document": DOC_PROP,
         "start_paragraph": {"type": "integer", "description": "从第几段开始，默认 1。", "default": 1},
         "max_paragraphs": {"type": "integer", "description": "最多返回多少段，默认 40，上限 200。", "default": 40}},
         "required": ["document"]},
     "handler": tool_read_document},
    {"name": "word_find_text",
     "description": "在已打开文档的当前内容里查找一段文字，返回每处所在段落和上下文。用来确认定位文字是否唯一、逐字符是否一致。",
     "inputSchema": {"type": "object", "properties": {
         "document": DOC_PROP,
         "query": {"type": "string", "description": "要查找的文字（区分大小写，逐字符匹配）。"},
         "context_chars": {"type": "integer", "description": "每处前后显示多少字符，默认 60。", "default": 60},
         "max_results": {"type": "integer", "description": "最多显示多少处，默认 20。", "default": 20}},
         "required": ["document", "query"]},
     "handler": tool_find_text},
    {"name": "word_replace_text",
     "description": ("以修订模式把一段文字替换成另一段，改动立刻显示在用户打开的 Word 窗口里。"
                     "replace 为空字符串即删除。find 必须与文档逐字符一致且唯一（或用 occurrence 指定）。"
                     "只会把 find 与 replace 之间真正不同的部分标为修订。"),
     "inputSchema": {"type": "object", "properties": {
         "document": DOC_PROP,
         "find": {"type": "string", "description": "文档里现有的文字，带足够上下文以保证唯一。"},
         "replace": {"type": "string", "description": "替换后的文字。空字符串表示删除。"},
         "occurrence": OCC_PROP},
         "required": ["document", "find", "replace"]},
     "handler": tool_replace_text},
    {"name": "word_insert_text",
     "description": "以修订模式在一段定位文字之后（或之前）插入新文字。要另起一段时在 text 里用换行符。",
     "inputSchema": {"type": "object", "properties": {
         "document": DOC_PROP,
         "anchor": {"type": "string", "description": "文档里现有的、用来定位的文字（逐字符一致且唯一）。"},
         "text": {"type": "string", "description": "要插入的文字。"},
         "position": {"type": "string", "enum": ["after", "before"], "description": "插在定位文字之后还是之前，默认 after。", "default": "after"},
         "occurrence": OCC_PROP},
         "required": ["document", "anchor", "text"]},
     "handler": tool_insert_text},
    {"name": "word_add_comment",
     "description": "在一段文字上添加 Word 批注（不改正文）。适合提出疑问、说明修改理由。",
     "inputSchema": {"type": "object", "properties": {
         "document": DOC_PROP,
         "anchor": {"type": "string", "description": "批注要挂在哪段现有文字上（逐字符一致且唯一）。"},
         "comment": {"type": "string", "description": "批注内容。"},
         "occurrence": OCC_PROP},
         "required": ["document", "anchor", "comment"]},
     "handler": tool_add_comment},
    {"name": "word_list_changes",
     "description": "列出文档里尚未处理的修订（谁、插入还是删除、内容）和批注，用来了解哪些改动还在等用户审阅。",
     "inputSchema": {"type": "object", "properties": {
         "document": DOC_PROP,
         "max_items": {"type": "integer", "description": "最多列出多少条，默认 100。", "default": 100}},
         "required": ["document"]},
     "handler": tool_list_changes},
]
TOOL_INDEX = {t["name"]: t for t in TOOLS}


# ---------------------------------------------------------------- MCP 协议

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
                    raise ToolError(f"缺少必填参数 {key}。")
            text = tool["handler"](args)
            return {"content": [{"type": "text", "text": text}], "isError": False}
        except ToolError as exc:
            return {"content": [{"type": "text", "text": str(exc)}], "isError": True}
        except Exception as exc:  # 不让单次调用的意外拖垮连接
            log(f"tool {params.get('name')} crashed: {exc!r}")
            return {"content": [{"type": "text", "text": f"桥接服务内部错误：{exc}"}], "isError": True}
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
                continue  # 通知，不需要回复
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


# ---------------------------------------------------------------- SSH 隧道维护

def ssh_base(control_path):
    # 只走已有的主连接；主连接不在时在本地立即失败，绝不向远程主机发起认证。
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


# ---------------------------------------------------------------- 入口

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
