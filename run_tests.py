"""
设计模式 VBScript + VB.NET 测试运行器
  - classicASPcode/*.vbs      -> cscript.exe //nologo //E:vbscript
  - axonAspModernCode/*.asp   -> axonasp-cli.exe --run   (AxonASP-Modern)
  - axonAspClassicCode/*.asp  -> axonasp-cli.exe --run   (AxonASP-Classic, 传统语法)
  - aspPycode/*.asp           -> python asppycli.py
  - vbNetcode/*.vb            -> dotnet (build 后直接运行产物)
用法:  python run_tests.py
输出:  test_report.md
"""
import os, re, subprocess, time, datetime, tempfile, shutil

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
CLASSIC_DIR  = os.path.join(BASE_DIR, "classicASPcode")
AXON_MODERN_DIR  = os.path.join(BASE_DIR, "axonAspModernCode")
ASPPY_DIR    = os.path.join(BASE_DIR, "aspPycode")
VBNET_DIR    = os.path.join(BASE_DIR, "vbNetcode")
AXON_CLASSIC_DIR = os.path.join(BASE_DIR, "axonAspClassicCode")
REPORT_FILE  = os.path.join(BASE_DIR, "test_report.md")

CSCRIPT      = r"C:\Windows\System32\cscript.exe"
WSCRIPT      = r"C:\Windows\System32\wscript.exe"
AXON_CLI     = r"C:\Users\jeffr\Documents\GitHub\axonasp\axonasp-cli.exe"
ASPPY_CLI    = r"C:\Users\jeffr\Documents\GitHub\ASPPY\asppycli.py"
DOTNET       = r"dotnet"
TIMEOUT      = 60   # seconds


# ── runners ──────────────────────────────────────────────────────
def _ensure_vbscript_engine():
    """某些 Win11 LTSC 默认 vbscript 引擎没注册，这里尝试重新注册"""
    import ctypes
    try:
        ctypes.windll.vbscript.RegistrationFree()
    except Exception:
        pass
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


def run_asppy(filepath):
    """用 python asppycli.py 运行 .asp 文件"""
    try:
        t0 = time.monotonic()
        r  = subprocess.run(
            ["python", ASPPY_CLI, filepath],
            capture_output=True, text=True, timeout=TIMEOUT,
            encoding="utf-8", errors="replace",
        )
        # ASPPY exit code: 0=OK(status<400), 1=server error(>=400), 2=usage error
        success = (r.returncode == 0)
        # asppycli: rendered body to stdout, errors may also appear in stderr
        # 如果 stderr 里有明确的编译/运行错误信息，标记为失败
        err_text = r.stderr.strip()
        if err_text:
            lower = err_text.lower()
            if "error" in lower or "exception" in lower or "traceback" in lower:
                success = False
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


def run_vbnet(filepath):
    """用 dotnet CLI 编译+运行 .vb 文件（.NET Core / .NET 5+）。
    创建临时控制台项目，**先编译（不计入 duration）**，再单独运行（只计运行时）。
    """
    workdir = tempfile.mkdtemp(prefix="vbnet_")
    try:
        # 把 .vb 文件复制为 Program.vb
        import shutil as _sh
        target_vb = os.path.join(workdir, "Program.vb")
        _sh.copy2(filepath, target_vb)

        # 写 .vbproj
        vbproj = os.path.join(workdir, "App.vbproj")
        with open(vbproj, "w", encoding="utf-8") as f:
            f.write(
                '<?xml version="1.0" encoding="utf-8"?>\n'
                '<Project Sdk="Microsoft.NET.Sdk">\n'
                '  <PropertyGroup>\n'
                '    <OutputType>Exe</OutputType>\n'
                '    <TargetFramework>net10.0</TargetFramework>\n'
                '    <RootNamespace>PatternApp</RootNamespace>\n'
                '    <OptionStrict>Off</OptionStrict>\n'
                '    <OptionExplicit>On</OptionExplicit>\n'
                '    <Nullable>Disable</Nullable>\n'
                '  </PropertyGroup>\n'
                '</Project>\n'
            )

        # 1) 先 build（不计入 duration）。如果编译失败就直接返回错误。
        build = subprocess.run(
            [DOTNET, "build", vbproj, "--nologo", "--verbosity", "quiet"],
            capture_output=True, text=True, timeout=TIMEOUT,
            encoding="utf-8", errors="replace",
            cwd=workdir,
        )
        if build.returncode != 0:
            # 合并 build 的 stderr/stdout 作为错误
            combined_err = (build.stderr.strip() + "\n" + build.stdout.strip()).strip()
            return {
                "success": False,
                "output": "",
                "error": "[BUILD] " + (combined_err if combined_err else f"exit code = {build.returncode}"),
                "duration": 0.0,
            }

        # 2) 定位 build 产物：优先 App.exe（apphost），退化为 App.dll
        #    预期位置: <workdir>/bin/Debug/net10.0/App.exe 或 App.dll
        app_exe = None
        app_dll = None
        for root, _dirs, files in os.walk(os.path.join(workdir, "bin")):
            for fn in files:
                if fn.lower() == "app.exe":
                    app_exe = os.path.join(root, fn)
                elif fn.lower() == "app.dll":
                    app_dll = os.path.join(root, fn)
        runner_cmd = None
        if app_exe and os.path.isfile(app_exe):
            runner_cmd = [app_exe]
        elif app_dll and os.path.isfile(app_dll):
            runner_cmd = [DOTNET, "exec", app_dll]
        else:
            return {
                "success": False,
                "output": "",
                "error": "[BUILD] build succeeded but cannot find App.exe / App.dll under bin/",
                "duration": 0.0,
            }

        # 3) 运行阶段：**只计这段时间**（不含 MSBuild/SDK 启动开销）
        t0 = time.monotonic()
        r = subprocess.run(
            runner_cmd,
            capture_output=True, text=True, timeout=TIMEOUT,
            encoding="utf-8", errors="replace",
            cwd=workdir,
        )

        success = (r.returncode == 0)
        combined_err = r.stderr.strip()
        if not success and not combined_err:
            combined_err = f"exit code = {r.returncode}"
        return {
            "success": success,
            "output":  r.stdout.strip(),
            "error":   combined_err,
            "duration": round(time.monotonic() - t0, 3),
        }
    except subprocess.TimeoutExpired:
        return {"success": False, "output": "", "error": "TIMEOUT", "duration": TIMEOUT}
    except Exception as e:
        return {"success": False, "output": "", "error": str(e), "duration": 0}
    finally:
        try:
            shutil.rmtree(workdir, ignore_errors=True)
        except Exception:
            pass


# ── output comparison (baseline = cscript) ───────────────────────
def _normalize(text):
    """规范化输出便于跨引擎比对：去除所有空白（换行/缩进/空格）。
    原因：cscript 按行输出（缩进在行首），AxonASP/ASPPY 连续输出（缩进在中间），
    空白只是展示格式，比对实质内容。"""
    return re.sub(r"\s+", "", text)


def _check_match(r, baseline):
    """r 与基准比对，设置 r["status"]: PASS / MISMATCH / FAIL"""
    if not r["success"]:
        r["status"] = "FAIL"
        return
    stem = os.path.splitext(r["filename"])[0]
    exp  = baseline.get(stem)
    if exp is not None and _normalize(r["output"]) != exp:
        r["status"]   = "MISMATCH"
        r["expected"] = exp
        r["actual"]   = _normalize(r["output"])
    else:
        r["status"] = "PASS"


# ── test runner ──────────────────────────────────────────────────
def run_all():
    results = []
    baseline = {}   # 文件名主干 -> 规范化的 cscript 输出

    # ClassicASP（基准）
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
            r["status"]   = "PASS" if r["success"] else "FAIL"
            if r["success"]:
                baseline[os.path.splitext(fn)[0]] = _normalize(r["output"])
            results.append(r)
            _print_result(r)
    else:
        print("  [WARN] classicASPcode/ not found")

    # AxonASP-Modern（跑 Axon 增强语法的代码）
    print("\n" + "=" * 60)
    print("  AxonASP-Modern Tests (axonasp-cli.exe)")
    print("=" * 60)
    if os.path.isdir(AXON_MODERN_DIR):
        for fn in sorted(os.listdir(AXON_MODERN_DIR)):
            if not fn.lower().endswith(".asp"):
                continue
            fp = os.path.join(AXON_MODERN_DIR, fn)
            print(f"\n  >>> {fn}")
            r = run_axon(fp)
            r["filename"] = fn
            r["type"]     = "AxonASPModern"
            _check_match(r, baseline)
            results.append(r)
            _print_result(r)
    else:
        print("  [WARN] axonASPcode/ not found")

    # AxonASP-Classic（跑传统 VBScript 语法的代码）
    print("\n" + "=" * 60)
    print("  AxonASP-Classic Tests (axonasp-cli.exe, classic syntax)")
    print("=" * 60)
    if os.path.isdir(AXON_CLASSIC_DIR):
        for fn in sorted(os.listdir(AXON_CLASSIC_DIR)):
            if not fn.lower().endswith(".asp"):
                continue
            fp = os.path.join(AXON_CLASSIC_DIR, fn)
            print(f"\n  >>> {fn}")
            r = run_axon(fp)
            r["filename"] = fn
            r["type"]     = "AxonASPClassic"
            _check_match(r, baseline)
            results.append(r)
            _print_result(r)
    else:
        print("  [WARN] axonAspClassicCode/ not found")

    # ASPPY
    print("\n" + "=" * 60)
    print("  ASPPY Tests (python asppycli.py)")
    print("=" * 60)
    if os.path.isdir(ASPPY_DIR):
        for fn in sorted(os.listdir(ASPPY_DIR)):
            if not fn.lower().endswith(".asp"):
                continue
            fp = os.path.join(ASPPY_DIR, fn)
            print(f"\n  >>> {fn}")
            r = run_asppy(fp)
            r["filename"] = fn
            r["type"]     = "ASPPY"
            _check_match(r, baseline)
            results.append(r)
            _print_result(r)
    else:
        print("  [WARN] aspPycode/ not found")

    # VB.NET
    print("\n" + "=" * 60)
    print("  VB.NET Tests (dotnet run)")
    print("=" * 60)
    if os.path.isdir(VBNET_DIR):
        for fn in sorted(os.listdir(VBNET_DIR)):
            if not fn.lower().endswith(".vb"):
                continue
            fp = os.path.join(VBNET_DIR, fn)
            print(f"\n  >>> {fn}")
            r = run_vbnet(fp)
            r["filename"] = fn
            r["type"]     = "VBNET"
            _check_match(r, baseline)
            results.append(r)
            _print_result(r)
    else:
        print("  [WARN] vbNetcode/ not found")

    return results


def _print_result(r):
    tag = r.get("status", "PASS" if r["success"] else "FAIL")
    print(f"  [{tag}] {r['duration']}s")
    if r["output"]:
        for line in r["output"].splitlines():
            print(f"       | {line}")
    if r.get("status") == "MISMATCH":
        print(f"       ! expected: {r['expected']}")
        print(f"       ! actual  : {r['actual']}")
    if not r["success"] and r["error"]:
        for line in r["error"].splitlines():
            print(f"       ! {line}")


# ── report ───────────────────────────────────────────────────────
def generate_report(results):
    classic = [r for r in results if r["type"] == "ClassicASP"]
    axon    = [r for r in results if r["type"] == "AxonASPModern"]
    axonc   = [r for r in results if r["type"] == "AxonASPClassic"]
    asppy   = [r for r in results if r["type"] == "ASPPY"]
    vbnet   = [r for r in results if r["type"] == "VBNET"]

    def _cnt(lst, st):  # count by status
        return sum(1 for r in lst if r.get("status") == st)

    c_pass, c_mis, c_fail = _cnt(classic,"PASS"), _cnt(classic,"MISMATCH"), _cnt(classic,"FAIL")
    a_pass, a_mis, a_fail = _cnt(axon,"PASS"),   _cnt(axon,"MISMATCH"),   _cnt(axon,"FAIL")
    ac_pass, ac_mis, ac_fail = _cnt(axonc,"PASS"), _cnt(axonc,"MISMATCH"), _cnt(axonc,"FAIL")
    s_pass, s_mis, s_fail = _cnt(asppy,"PASS"),  _cnt(asppy,"MISMATCH"),  _cnt(asppy,"FAIL")
    v_pass, v_mis, v_fail = _cnt(vbnet,"PASS"),  _cnt(vbnet,"MISMATCH"),  _cnt(vbnet,"FAIL")
    c_avg  = sum(r["duration"] for r in classic) / len(classic) if classic else 0
    a_avg  = sum(r["duration"] for r in axon)    / len(axon)    if axon    else 0
    ac_avg = sum(r["duration"] for r in axonc)   / len(axonc)   if axonc   else 0
    s_avg  = sum(r["duration"] for r in asppy)   / len(asppy)   if asppy   else 0
    v_avg  = sum(r["duration"] for r in vbnet)   / len(vbnet)   if vbnet   else 0

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    lines = [
        f"# Design Pattern VBScript + VB.NET Test Report",
        f"",
        f"Generated: {now}",
        f"",
        f"输出比对基准：ClassicASP (cscript)。MISMATCH = 无报错但输出与基准不一致（静默错误）。",
        f"",
        f"## Summary",
        f"",
        f"| Engine         | Total | Pass | Mismatch | Fail | Avg Time |",
        f"|----------------|-------|------|----------|------|----------|",
        f"| ClassicASP     | {len(classic):5d} | {c_pass:4d} | {c_mis:8d} | {c_fail:4d} | {c_avg:.3f}s  |",
        f"| AxonASP-Modern | {len(axon):5d} | {a_pass:4d} | {a_mis:8d} | {a_fail:4d} | {a_avg:.3f}s  |",
        f"| AxonASP-Classic| {len(axonc):5d} | {ac_pass:4d} | {ac_mis:8d} | {ac_fail:4d} | {ac_avg:.3f}s  |",
        f"| ASPPY          | {len(asppy):5d} | {s_pass:4d} | {s_mis:8d} | {s_fail:4d} | {s_avg:.3f}s  |",
        f"| VB.NET         | {len(vbnet):5d} | {v_pass:4d} | {v_mis:8d} | {v_fail:4d} | {v_avg:.3f}s  |",
        f"",
    ]

    for section_title, section_list in [
        ("ClassicASP", classic), ("AxonASP-Modern", axon),
        ("AxonASP-Classic", axonc), ("ASPPY", asppy), ("VB.NET", vbnet),
    ]:
        lines += [f"## {section_title} Details", ""]
        for r in section_list:
            s = r.get("status", "PASS" if r["success"] else "FAIL")
            lines.append(f"- **{r['filename']}** : {s} ({r['duration']}s)")
            if r["output"]:
                for l in r["output"].splitlines():
                    lines.append(f"  - `{l}`")
            if r.get("status") == "MISMATCH":
                lines.append(f"  - EXPECTED: `{r['expected']}`")
                lines.append(f"  - ACTUAL  : `{r['actual']}`")
            if not r["success"] and r["error"]:
                for l in r["error"].splitlines():
                    lines.append(f"  - ERR: `{l}`")
        lines.append("")

    # Fix suggestions
    fixes = []
    for r in results:
        if r.get("status", "PASS" if r["success"] else "FAIL") == "PASS":
            continue
        fn = r["filename"]
        if r.get("status") == "MISMATCH":
            fixes.append(f"- **{fn}** ({r['type']}): 输出与 cscript 基准不一致（静默错误，无报错但结果不对）。")
        elif r["type"] == "VBNET":
            fixes.append(f"- **{fn}**: VB.NET 编译或运行错误，请检查代码（类/模块结构、Sub Main、引用）。")
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
    print("  Design Pattern VBScript + VB.NET Test Framework")
    print("=" * 60)
    results = run_all()
    generate_report(results)
    print("\n" + "=" * 60)
    print("  All tests completed")
    print("=" * 60)
