"""stdio MCP server for SolidWorks COM automation (parametric box modeling with
volume verification, SLDPRT/STEP export). Tools: mcp__solidworks__solidworks_box.
Requires Windows, pywin32, and SolidWorks installed."""
import json
import os
import sys
import traceback

LIB = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "lib")
if LIB not in sys.path:
    sys.path.insert(0, LIB)

os.environ.setdefault("PYTHONDONTWRITEBYTECODE", "1")

# Windows 控制台默认 GBK：强制 UTF-8，否则中文报错文本会炸掉 JSON-RPC。
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


def _out(sub=""):
    """输出目录：LAB_OUT_DIR > <cwd>/lab-out（不可写则逐级回退，跨工作区只读 cwd 不崩溃）。"""
    def _try(base):
        d = os.path.join(base, sub) if sub else base
        try:
            os.makedirs(d, exist_ok=True)
            probe = os.path.join(d, ".write-test")
            open(probe, "w").close()
            os.remove(probe)
            return d
        except OSError:
            return None

    env = os.environ.get("LAB_OUT_DIR")
    for cand in ([env] if env else []) + [os.path.join(os.getcwd(), "lab-out"),
                                          os.path.expandvars(r"%LOCALAPPDATA%\lab-out"),
                                          os.path.join(os.path.expanduser("~"), "Desktop", "ds", "lab-out")]:
        d = _try(cand)
        if d:
            return d
    import tempfile
    return tempfile.mkdtemp(prefix="lab-out-")


def tool_sw_box(args):
    """SolidWorks 长方体: {w,h,d}(mm), export_step?:bool"""
    from sw_helper import connect, new_part, box
    sw, styp = connect()
    model, mtyp, ext, skmgr, featmgr = new_part(sw, styp)
    w, h, d = float(args["w"]), float(args["h"]), float(args["d"])
    vol = box(mtyp, ext, skmgr, featmgr, w, h, d)
    out = {"volume_mm3": round(vol * 1e9, 1), "expect_mm3": w * h * d}
    sld = os.path.join(_out("solidworks"), f"box_{int(w)}x{int(h)}x{int(d)}.SLDPRT")
    mtyp.SaveAs3(sld, 0, 2)
    out["sldprt"] = sld
    if args.get("export_step"):
        step = sld.rsplit(".", 1)[0] + ".step"
        errs, warn = 0, 0
        ext.SaveAs(step, 0, 2, None, errs, warn)
        out["step"] = step if os.path.exists(step) else None
    return out



TOOL_SPECS = {
"solidworks_box": {
        "description": ("Model a parametric rectangular block in SolidWorks (mm), verify volume via "
                        "mass properties, save SLDPRT and optionally STEP. Args: w, h, d, export_step(bool)."),
        "inputSchema": {
            "type": "object",
            "properties": {
                "w": {"type": "number"}, "h": {"type": "number"}, "d": {"type": "number"},
                "export_step": {"type": "boolean"},
            },
            "required": ["w", "h", "d"],
        },
        "fn": tool_sw_box,
    }
}


# ---------------------------------------------------------------- stdio ----
def _send(msg):
    body = json.dumps(msg, ensure_ascii=False)
    sys.stdout.write(body + "\n")
    sys.stdout.flush()


def main():
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
        except json.JSONDecodeError:
            continue
        rid = req.get("id")
        method = req.get("method", "")
        params = req.get("params") or {}

        if method == "initialize":
            _send({"jsonrpc": "2.0", "id": rid, "result": {
                "protocolVersion": params.get("protocolVersion", "2024-11-05"),
                "capabilities": {"tools": {}},
                "serverInfo": {"name": "solidworks", "version": "1.0.0"},
            }})
        elif method.startswith("notifications/"):
            continue
        elif method == "ping":
            _send({"jsonrpc": "2.0", "id": rid, "result": {}})
        elif method == "tools/list":
            tools = [{"name": n, "description": s["description"], "inputSchema": s["inputSchema"]}
                     for n, s in TOOL_SPECS.items()]
            _send({"jsonrpc": "2.0", "id": rid, "result": {"tools": tools}})
        elif method == "tools/call":
            name = params.get("name")
            spec = TOOL_SPECS.get(name)
            if spec is None:
                _send({"jsonrpc": "2.0", "id": rid, "result": {
                    "isError": True, "content": [{"type": "text", "text": f"unknown tool {name}"}]}})
                continue
            try:
                out = spec["fn"](params.get("arguments") or {})
                _send({"jsonrpc": "2.0", "id": rid, "result": {
                    "content": [{"type": "text", "text": json.dumps(out, ensure_ascii=False)}]}})
            except Exception as exc:
                tb = traceback.format_exc(limit=4)
                _send({"jsonrpc": "2.0", "id": rid, "result": {
                    "isError": True, "content": [{"type": "text", "text": f"{exc}\n{tb}"}]}})
        else:
            _send({"jsonrpc": "2.0", "id": rid,
                   "error": {"code": -32601, "message": f"method not found: {method}"}})


if __name__ == "__main__":
    main()
