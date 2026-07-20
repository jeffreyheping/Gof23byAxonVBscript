import subprocess

r = subprocess.run(
    [r"C:\axonasp\axonasp-cli.exe", "--run", r"axonASPcode\test_arr2.asp"],
    capture_output=True, text=True, encoding="utf-8"
)
print("returncode:", r.returncode)
print("stdout:", repr(r.stdout))
print("stderr:", repr(r.stderr))
