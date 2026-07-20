"""
设计模式 VBScript 测试运行器
  - classicASPcode/*.vbs  -> cscript.exe //nologo //E:vbscript
  - axonASPcode/*.asp    -> axonasp-cli.exe --run
用法:  python run_tests.py
输出:  test_report.md
"""
import os, subprocess, time, datetime

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
CLASSIC_DIR  = os.path.join(BASE_DIR, "classicASPcode")
AXON_DIR     = os.path.join(BASE_DIR, "axonASPcode")
REPORT_FILE  = os.path.join(BASE_DIR, "test_report.md")

CSCRIPT      = r"C:\Windows\System32\cscript.exe"
WSCRIPT      = r"C:\Windows\System32\wscript.exe"
AXON_CLI     = r"C:\axonasp\axonasp-cli.exe"
TIMEOUT      = 30   # seconds


# ── runners ──────────────────────────────────────────────────────
def _ensure_vbscript_engine():
    """某些 Win11 LTSC 默认 vbscript 引擎没注册，这里尝试重新注册"""
    import ctypes
    try:
        ctypes.windll.vbscript.RegistrationFree()
    except Exception:
        pass
    # 用 regsvr32 注册（不可用时静默失败）
    for dll in ("vbscript.dll",):
        try:
            subprocess.run(
                ["regsvr32", "/s", f"C:\\Windows\\System32\\{dll}"],
                capture_output=True, timeout=10,
            )
        except Exception:
            pass


def run_classic(filepath):
    """用 cscript 运行 .vbs 文件"""
    _ensure_vbscript_engine()
    try:
        t0 = time.monotonic()
        r  = subprocess.run(
            [CSCRIPT, "//nologo", "//E:vbscript", filepath],
            capture_output=True, text=True, timeout=TIMEOUT,
            encoding="gbk", errors="replace",
        )
        # cscript.exe 即使有运行时错误 returncode 也可能是 0
        # 所以必须额外检查 stderr 中是否含"错误"字样
        has_error = False
        err_text = r.stderr.strip()
        if err_text:
            lower = err_text.lower()
            if "错误" in err_text or "error" in lower or "错误:" in err_text:
                has_error = True
        success = (r.returncode == 0) and (not has_error)
        return {
            "success": success,
            "output":  r.stdout.strip(),
            "error":   err_text,
            "duration": round(time.monotonic() - t0, 3),
        }
    except subprocess.TimeoutExpired:
        return {"success": False, "output": "", "error": "TIMEOUT", "duration": TIMEOUT}
    except Exception as e:
        return {"success": False, "output": "", "error": str(e), "duration": 0}


def run_axon(filepath):
    """用 axonasp-cli --run 运行 .asp 文件"""
    try:
        t0 = time.monotonic()
        r  = subprocess.run(
            [AXON_CLI, "--run", filepath],
            capture_output=True, text=True, timeout=TIMEOUT,
            encoding="utf-8", errors="replace",
        )
        return {
            "success": r.returncode == 0,
            "output":  r.stdout.strip(),
            "error":   r.stderr.strip(),
            "duration": round(time.monotonic() - t0, 3),
        }
    except subprocess.TimeoutExpired:
        return {"success": False, "output": "", "error": "TIMEOUT", "duration": TIMEOUT}
    except Exception as e:
        return {"success": False, "output": "", "error": str(e), "duration": 0}


# ── test runner ──────────────────────────────────────────────────
def run_all():
    results = []

    # ClassicASP
    print("=" * 60)
    print("  ClassicASP Tests (cscript.exe)")
    print("=" * 60)
    if os.path.isdir(CLASSIC_DIR):
        for fn in sorted(os.listdir(CLASSIC_DIR)):
            if not fn.lower().endswith(".vbs"):
                continue
            fp = os.path.join(CLASSIC_DIR, fn)
            print(f"\n  >>> {fn}")
            r = run_classic(fp)
            r["filename"] = fn
            r["type"]     = "ClassicASP"
            results.append(r)
            _print_result(r)
    else:
        print("  [WARN] classicASPcode/ not found")

    # AxonASP
    print("\n" + "=" * 60)
    print("  AxonASP Tests (axonasp-cli.exe)")
    print("=" * 60)
    if os.path.isdir(AXON_DIR):
        for fn in sorted(os.listdir(AXON_DIR)):
            if not fn.lower().endswith(".asp"):
                continue
            fp = os.path.join(AXON_DIR, fn)
            print(f"\n  >>> {fn}")
            r = run_axon(fp)
            r["filename"] = fn
            r["type"]     = "AxonASP"
            results.append(r)
            _print_result(r)
    else:
        print("  [WARN] axonASPcode/ not found")

    return results


def _print_result(r):
    tag = "PASS" if r["success"] else "FAIL"
    print(f"  [{tag}] {r['duration']}s")
    if r["output"]:
        for line in r["output"].splitlines():
            print(f"       | {line}")
    if not r["success"] and r["error"]:
        for line in r["error"].splitlines():
            print(f"       ! {line}")


# ── report ───────────────────────────────────────────────────────
def generate_report(results):
    classic = [r for r in results if r["type"] == "ClassicASP"]
    axon    = [r for r in results if r["type"] == "AxonASP"]

    c_pass = sum(1 for r in classic if r["success"])
    a_pass = sum(1 for r in axon    if r["success"])
    c_avg  = sum(r["duration"] for r in classic) / len(classic) if classic else 0
    a_avg  = sum(r["duration"] for r in axon)    / len(axon)    if axon    else 0

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    lines = [
        f"# Design Pattern VBScript Test Report",
        f"",
        f"Generated: {now}",
        f"",
        f"## Summary",
        f"",
        f"| Engine     | Total | Pass | Fail | Avg Time |",
        f"|------------|-------|------|------|----------|",
        f"| ClassicASP | {len(classic):5d} | {c_pass:4d} | {len(classic)-c_pass:4d} | {c_avg:.3f}s  |",
        f"| AxonASP    | {len(axon):5d} | {a_pass:4d} | {len(axon)-a_pass:4d} | {a_avg:.3f}s  |",
        f"",
    ]

    # ClassicASP detail
    lines += ["## ClassicASP Details", ""]
    for r in classic:
        s = "PASS" if r["success"] else "FAIL"
        lines.append(f"- **{r['filename']}** : {s} ({r['duration']}s)")
        if r["output"]:
            for l in r["output"].splitlines():
                lines.append(f"  - `{l}`")
        if not r["success"] and r["error"]:
            for l in r["error"].splitlines():
                lines.append(f"  - ERR: `{l}`")
    lines.append("")

    # AxonASP detail
    lines += ["## AxonASP Details", ""]
    for r in axon:
        s = "PASS" if r["success"] else "FAIL"
        lines.append(f"- **{r['filename']}** : {s} ({r['duration']}s)")
        if r["output"]:
            for l in r["output"].splitlines():
                lines.append(f"  - `{l}`")
        if not r["success"] and r["error"]:
            for l in r["error"].splitlines():
                lines.append(f"  - ERR: `{l}`")
    lines.append("")

    # Fix suggestions
    fixes = []
    for r in results:
        if r["success"]:
            continue
        fn = r["filename"]
        if fn == "04_建造者模式.vbs":
            fixes.append(f"- **{fn}**: 变量名 `director` 与类名 `Director` 冲突（VBScript 不区分大小写）。建议改为 `myDirector`。")
        elif fn == "05_原型模式.vbs":
            fixes.append(f"- **{fn}**: `ReDim copy.Skills(ub)` 语法错误。`ReDim` 只能用于数组变量，不能用于对象属性。建议改为临时数组赋值方式。")
        elif fn == "06_代理模式.vbs":
            fixes.append(f"- **{fn}**: `m_RealImage` 未初始化时值为 `Empty`，不是 `Nothing`。`If m_RealImage Is Nothing` 在 `Empty` 上会报'缺少对象'。建议在 `Class_Initialize` 中加 `Set m_RealImage = Nothing`。")
        elif fn == "15_模板方法模式.vbs":
            fixes.append(f"- **{fn}**: `PDFMiner` / `CSVMinder` 没有 `MineData` 方法。VBScript 不支持 `Extends` 继承，需要在这些类中各自实现 `MineData`，或改用组合模式包含一个 `DataMiner` 实例。")
        elif fn == "17_责任链模式.vbs":
            fixes.append(f"- **{fn}**: `m_Next` 未初始化时值为 `Empty`，不是 `Nothing`。`If m_Next Is Nothing` 在 `Empty` 上会报'缺少对象'。建议在 `Class_Initialize` 中加 `Set m_Next = Nothing`。")
        elif fn == "20_中介者模式.asp":
            fixes.append(f"- **{fn}**: Object required（对象引用为空）。可能是接口方法调用时对象未初始化。")
        else:
            fixes.append(f"- **{fn}**: 运行时/编译错误，请检查代码。")

    if fixes:
        lines += ["## 修复建议", ""]
        lines += fixes
        lines.append("")

    report = "\n".join(lines)
    with open(REPORT_FILE, "w", encoding="utf-8") as f:
        f.write(report)
    print(f"\n[REPORT] Saved to {REPORT_FILE}")
    return report


# ── main ─────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 60)
    print("  Design Pattern VBScript Test Framework")
    print("=" * 60)
    results = run_all()
    generate_report(results)
    print("\n" + "=" * 60)
    print("  All tests completed")
    print("=" * 60)