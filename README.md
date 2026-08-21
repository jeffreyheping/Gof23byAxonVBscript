# Gof23byAxonVBscript / 23 GoF Design Patterns in VBScript

用 VBScript（传统语法 + AxonASP 增强语法）实现全部 23 个 GoF 设计模式，并提供跨 5 种运行引擎的自动化兼容性测试。

All 23 GoF design patterns implemented in VBScript (classic syntax + AxonASP enhanced syntax), with automated compatibility tests across 5 runtime engines.

## 目录结构 / Project Structure

| 目录 / Directory | 说明 / Description |
|---|---|
| `classicASPcode/` | 传统 VBScript 版（.vbs，cscript 运行）/ Classic VBScript (.vbs, run by cscript) |
| `axonAspModernCode/` | AxonASP 增强语法版（.asp）/ AxonASP enhanced syntax (.asp) |
| `axonAspClassicCode/` | 传统语法的 AxonASP 版（.asp）/ Classic-syntax code run on AxonASP (.asp) |
| `aspPycode/` | 传统语法版，供 ASPPY（Python 实现）运行 / Classic-syntax code for ASPPY (Python-based engine) |
| `vbNetcode/` | VB.NET 对照版（.vb）/ VB.NET reference implementations (.vb) |
| `byChapterMDcn/` | 按章节拆分的中文文档 / Chinese docs split by chapter |

## 工具脚本 / Tool Scripts

- `extract_code.py` — 从总 MD 文档拆出各引擎代码并注入 `Option Explicit` 等适配代码 / Extracts per-engine code from the master MD and injects `Option Explicit` etc.
- `run_tests.py` — 跑全套测试（含输出比对，以 cscript 为基准）并生成 `test_report.md` / Runs the full test suite (with output comparison against the cscript baseline) and generates `test_report.md`
- `perf_tests.py` — 性能测试：各引擎重复运行 10/20/30 次，生成 `perf_report.md` / Performance tests: repeat runs (10/20/30) per engine, generates `perf_report.md`
- `merge_chapters.py` — 合并章节 MD / Merges chapter MDs

用法 / Usage:

```bash
python extract_code.py    # 拆代码 / extract code (auto-picks the latest MD)
python run_tests.py       # 全套测试 / run all tests
python perf_tests.py      # 性能测试（在 run_tests.py 之后）/ perf tests (after run_tests.py)
```

## 最新测试结果 / Latest Test Results

Windows 11，每文件平均耗时（进程启动 + 编译 + 执行）/ Windows 11, avg time per file (process start + compile + execute)。
输出比对基准为 cscript：Mismatch = 无报错但输出与基准不一致（静默错误）/ Output compared against the cscript baseline: Mismatch = no error but silent wrong output.

| 引擎 / Engine | 通过 / Pass | 不一致 / Mismatch | 失败 / Fail | 平均耗时 / Avg Time |
|---|---|---|---|---|
| AxonASP-Classic | 23 | 0 | 0 | 0.105s |
| AxonASP-Modern | 23 | 0 | 0 | 0.131s |
| ClassicASP (cscript) | 23 | 0 | 0 | 0.140s |
| ASPPY | 23 | 0 | 0 | 0.302s |
| VB.NET (dotnet) | 23 | 0 | 0 | 0.255s |

**5 引擎 × 23 模式 = 115 用例全部通过** / All 115 test cases (5 engines × 23 patterns) pass.



完整结果见 [test_report.md](test_report.md) / Full results in [test_report.md](test_report.md)

## 性能测试 / Performance Tests

每文件重复运行 10/20/30 次（共 6900 次运行），下表取 30 次档的平均值 / Each file run 10/20/30 times repeatedly (6,900 runs total); the table below shows the 30-rep averages:

| 引擎 / Engine | Avg | Median | Min | Max | vs cscript |
|---|---|---|---|---|---|
| AxonASP-Modern | 0.073s | 0.072s | 0.066s | 0.116s | **0.80x** |
| AxonASP-Classic | 0.075s | 0.073s | 0.067s | 0.229s | 0.83x |
| VB.NET (dotnet)* | 0.089s | 0.086s | 0.075s | 0.201s | 0.98x |
| ClassicASP (cscript) | 0.091s | 0.090s | 0.084s | 0.130s | 1.00x |
| ASPPY | 0.218s | 0.212s | 0.188s | 0.393s | 2.40x |

\* VB.NET 首轮有冷启动开销（x10 档 0.120s），跑热后稳定在 0.089s / VB.NET has cold-start overhead on the first round (0.120s at x10), settling at 0.089s once warmed up.

计时口径与 run_tests.py 一致：每次运行 = 进程启动 + 编译 + 执行；VB.NET 为例外（每文件只 build 一次不计入，计时只含运行产物）/ Timing is consistent with run_tests.py: each run = process start + compile + execute; VB.NET is the exception (built once per file, excluded from timing; only artifact execution is timed).

完整数据见 [perf_report.md](perf_report.md) / Full data in [perf_report.md](perf_report.md)

## 相关项目 / Related Projects

- [AxonASP](https://github.com/guimaraeslucas/axonasp) — Go 实现的现代 ASP/VBScript 引擎 / A modern ASP/VBScript engine written in Go
- [ASPPY](https://github.com/PieterCooreman/ASPPY) — Python 实现的 ASP 引擎 / A Python-based ASP engine
