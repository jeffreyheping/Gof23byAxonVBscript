"""
性能测试运行器（独立脚本，在 run_tests.py 之后按需运行）
  - 不拆代码，直接跑 run_tests.py 生成的现有目录
  - classicASPcode/*.vbs      -> cscript.exe
  - axonAspModernCode/*.asp   -> axonasp-cli.exe
  - axonAspClassicCode/*.asp  -> axonasp-cli.exe
  - aspPycode/*.asp           -> python asppycli.py
  - vbNetcode/*.vb            -> dotnet（每文件只 build 一次，重复运行产物计时）
用法:
  python perf_tests.py               # 每文件重复 10/20/30 次
  python perf_tests.py 5             # 只跑 5 次
  python perf_tests.py 5 ASPPY       # 只跑指定引擎（不区分大小写子串匹配，可多个）
输出:  perf_report.md
"""
import os, sys, time, shutil, tempfile, statistics, datetime, subprocess

import run_tests as rt   # 复用 runner 与路径常量

REPORT_FILE = os.path.join(rt.BASE_DIR, "perf_report.md")

ENGINES = [
    # (显示名, 目录, 扩展名, runner)；VBNET 特殊处理（build 一次，重复运行）
    ("ClassicASP",      rt.CLASSIC_DIR,      ".vbs", rt.run_classic),
    ("AxonASP-Modern",  rt.AXON_MODERN_DIR,  ".asp", rt.run_axon),
    ("AxonASP-Classic", rt.AXON_CLASSIC_DIR, ".asp", rt.run_axon),
    ("ASPPY",           rt.ASPPY_DIR,        ".asp", rt.run_asppy),
    ("VBNET",           rt.VBNET_DIR,        ".vb",  None),
]


# ── VB.NET：build 一次，重复运行 ──────────────────────────────────
def vbnet_build_once(filepath):
    """编译一次，返回 (runner_cmd, workdir)；失败返回 (None, None)"""
    workdir = tempfile.mkdtemp(prefix="vbnetperf_")
    try:
        target_vb = os.path.join(workdir, "Program.vb")
        shutil.copy2(filepath, target_vb)
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
        build = subprocess.run(
            [rt.DOTNET, "build", vbproj, "--nologo", "--verbosity", "quiet"],
            capture_output=True, text=True, timeout=rt.TIMEOUT,
            encoding="utf-8", errors="replace", cwd=workdir,
        )
        if build.returncode != 0:
            return None, None
        app_exe = app_dll = None
        for root, _dirs, files in os.walk(os.path.join(workdir, "bin")):
            for fn in files:
                if fn.lower() == "app.exe":
                    app_exe = os.path.join(root, fn)
                elif fn.lower() == "app.dll":
                    app_dll = os.path.join(root, fn)
        if app_exe:
            return [app_exe], workdir
        if app_dll:
            return [rt.DOTNET, "exec", app_dll], workdir
        return None, None
    except Exception:
        return None, None


def vbnet_run(runner_cmd, workdir):
    """运行已编译产物一次，返回耗时（秒）"""
    t0 = time.monotonic()
    subprocess.run(
        runner_cmd, capture_output=True, text=True, timeout=rt.TIMEOUT,
        encoding="utf-8", errors="replace", cwd=workdir,
    )
    return time.monotonic() - t0


# ── 主流程 ────────────────────────────────────────────────────────
def main(reps_list, engine_filters=None):
    print("=" * 60)
    print(f"  Performance Tests  reps={reps_list}"
          + (f"  engines={engine_filters}" if engine_filters else ""))
    print("=" * 60)

    # data[engine][reps] = 所有单次耗时列表；file_avg[engine][fn][reps] = 平均
    data     = {}
    file_avg = {}

    for name, d, ext, runner in ENGINES:
        if engine_filters and not any(f in name.lower() for f in engine_filters):
            continue
        if not os.path.isdir(d):
            print(f"\n  [WARN] {d} not found, skip {name}")
            continue

        files = sorted(f for f in os.listdir(d) if f.lower().endswith(ext))
        print("\n" + "=" * 60)
        print(f"  {name}  ({len(files)} files)")
        print("=" * 60)

        data[name]     = {reps: [] for reps in reps_list}
        file_avg[name] = {}

        for fn in files:
            fp = os.path.join(d, fn)
            cmd = wd = None
            if name == "VBNET":
                cmd, wd = vbnet_build_once(fp)
                if cmd is None:
                    print(f"  [SKIP] {fn} (build failed)")
                    continue
            file_avg[name][fn] = {}
            parts = []
            for reps in reps_list:
                ds = []
                for _ in range(reps):
                    if name == "VBNET":
                        ds.append(vbnet_run(cmd, wd))
                    else:
                        ds.append(runner(fp)["duration"])
                data[name][reps].extend(ds)
                file_avg[name][fn][reps] = sum(ds) / len(ds) if ds else 0
                parts.append(f"x{reps} avg={file_avg[name][fn][reps]:.3f}s")
            print(f"  {fn:<28} " + "  ".join(parts))
            if name == "VBNET":
                shutil.rmtree(wd, ignore_errors=True)

    generate_report(data, file_avg, reps_list)
    print("\n" + "=" * 60)
    print("  Performance tests completed")
    print("=" * 60)


def generate_report(data, file_avg, reps_list):
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        "# Performance Test Report",
        "",
        f"Generated: {now}",
        "",
        "计时口径与 run_tests.py 一致：每次运行 = 进程启动 + 编译 + 执行。",
        "VB.NET 为例外：每文件只 build 一次（不计入），计时只含运行产物。",
        "",
        "## Summary（全部单次运行汇总 / all runs pooled）",
        "",
        "| Engine | Reps | Runs | Total | Avg | Median | Min | Max | vs ClassicASP |",
        "|---|---|---|---|---|---|---|---|---|",
    ]

    # 基准：ClassicASP 各 reps 的平均（用于比值）
    base_avg = {}
    for reps in reps_list:
        ds = data.get("ClassicASP", {}).get(reps, [])
        base_avg[reps] = sum(ds) / len(ds) if ds else None

    for name in data:
        for reps in reps_list:
            ds = data[name].get(reps, [])
            if not ds:
                continue
            avg  = sum(ds) / len(ds)
            med  = statistics.median(ds)
            ratio = f"{avg / base_avg[reps]:.2f}x" if base_avg.get(reps) else "-"
            lines.append(
                f"| {name} | {reps} | {len(ds)} | {sum(ds):.1f}s | {avg:.3f}s "
                f"| {med:.3f}s | {min(ds):.3f}s | {max(ds):.3f}s | {ratio} |"
            )
    lines.append("")

    for name in data:
        lines += [f"## {name} Per-File", "",
                  "| File | " + " | ".join(f"avg x{r}" for r in reps_list) + " |",
                  "|---" * (len(reps_list) + 1) + "|"]
        for fn, avgs in file_avg[name].items():
            cells = " | ".join(
                f"{avgs[r]:.3f}s" if r in avgs else "-" for r in reps_list
            )
            lines.append(f"| {fn} | {cells} |")
        lines.append("")

    with open(REPORT_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"\n[REPORT] Saved to {REPORT_FILE}")


if __name__ == "__main__":
    args  = sys.argv[1:]
    reps  = sorted({int(a) for a in args if a.isdigit() and int(a) > 0}) or [10, 20, 30]
    flts  = [a.lower() for a in args if not a.isdigit()]
    main(reps, flts or None)
