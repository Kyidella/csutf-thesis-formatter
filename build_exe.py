"""一条命令把 app.py 打包成单文件 exe。

    python build_exe.py            打包
    python build_exe.py --clean    先清掉上次的 build/dist 再打包

为什么要先把规则文件复制一份：仓库里的规则文件名是中文
（`中南林科大_硕士_学术学位_理工类.yaml`），中文路径进 PyInstaller 的
`--add-data` 是打包环节最容易翻车的地方。所以复制成 `staged_rules/default.yaml`
（纯 ASCII）再打包。仓库里仍然只有那一份中文名的 YAML，不产生第二份要同步的副本。

运行时怎么找规则：按扩展名在 `rules/` 里找第一个 .yaml
（见 gui_worker.default_rules_path），代码里不写死文件名。

产物体积：Qt 的 DLL 多，单文件大概 80～150MB，第一次启动要解压，会等几秒。

两种形态：

  onefile  一个 exe，自用方便。但**每次启动都要把自己解压到 %TEMP% 下的
           _MEIxxxxxx**：启动慢几秒、杀软容易拦，被强杀还会留下解压残留。
  onedir   一个文件夹（约 200 个文件）。启动即用、不写临时目录、杀软友好。
           分享时打成 zip，对方解压即用。

    python build_exe.py --mode onefile    # 默认
    python build_exe.py --mode onedir     # 文件夹 + 自动打 zip
    python build_exe.py --mode both       # 两个都出

分发包里会一并放进 LICENSE / THIRD_PARTY_NOTICES.md / LICENSE-LGPLv3.txt——
**这是 LGPLv3 的要求**（exe 含 Qt 的二进制，分发就得附许可），别删。
"""

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
RULES_DIR = ROOT / "rules"
STAGE = ROOT / "build" / "staged_rules"
NAME = "ThesisFormatter"

# 这些 Qt 模块 app.py 一个都用不到，排掉能小几十 MB。
# 排错了会让 exe 起来就崩，所以打完必须真的运行一次。
EXCLUDES = [
    "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets",
    "PySide6.QtWebEngineQuick", "PySide6.QtQuick", "PySide6.QtQml",
    "PySide6.Qt3DCore", "PySide6.Qt3DRender", "PySide6.QtMultimedia",
    "PySide6.QtMultimediaWidgets", "PySide6.QtCharts",
    "PySide6.QtDataVisualization", "PySide6.QtQuick3D",
    "PySide6.QtBluetooth", "PySide6.QtNfc", "PySide6.QtPositioning",
    "PySide6.QtSerialPort", "PySide6.QtSql", "PySide6.QtTest",
    "tkinter", "unittest", "pydoc_data",
]


def stage_rules():
    """把仓库里的规则复制成 ASCII 文件名，供 --add-data 使用。"""
    if STAGE.exists():
        shutil.rmtree(STAGE)
    STAGE.mkdir(parents=True)
    src = sorted(RULES_DIR.glob("*.yaml"))
    if not src:
        sys.exit(f"没找到规则文件：{RULES_DIR}")
    for p in src:
        shutil.copy2(p, STAGE / "default.yaml")
        print(f"  规则 {p.name} → {STAGE.name}/default.yaml")
    return STAGE


def _exe_path():
    return ROOT / "dist" / (f"{NAME}.exe" if sys.platform == "win32" else NAME)


def check_not_running():
    """打包前先看旧程序是不是还开着。

    Windows 不允许删/覆盖正在运行的程序，否则 PyInstaller 会甩一堆
    PermissionError 的堆栈，看不出真正的原因（实测踩到）。
    """
    exe = _exe_path()
    if not exe.exists():
        return
    try:
        exe.rename(exe)          # 能改名就说明没被占用，改回去
        exe.rename(exe)
    except OSError:
        sys.exit(f"「{exe.name}」正在运行，先把它关掉再打包。\n"
                 f"（Windows 不允许覆盖正在运行的程序）")


def clean():
    for d in (ROOT / "build", ROOT / "dist"):
        if not d.exists():
            continue
        try:
            shutil.rmtree(d)
        except PermissionError:
            sys.exit(f"删不掉 {d.name}/ —— 多半是打包好的程序还开着，先关掉窗口。")
        print(f"  清掉 {d.name}/")


# 分发时要一并给出的文件（LGPL 合规：exe 里含 Qt 二进制，必须附许可）
LEGAL_FILES = ("LICENSE", "THIRD_PARTY_NOTICES.md", "LICENSE-LGPLv3.txt")


def _run_pyinstaller(onefile):
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--windowed",
        "--onefile" if onefile else "--onedir",
        "--name", NAME,
        f"--add-data={STAGE}{';' if sys.platform == 'win32' else ':'}rules",
    ]
    for m in EXCLUDES:
        cmd += ["--exclude-module", m]
    cmd.append("app.py")
    print("\n  " + " ".join(cmd) + "\n")
    if subprocess.run(cmd, cwd=ROOT).returncode != 0:
        sys.exit("PyInstaller 失败")


def _copy_legal(dst_dir):
    for name in LEGAL_FILES:
        src = ROOT / name
        if src.exists():
            shutil.copy2(src, dst_dir / name)
        else:
            print(f"  ！缺 {name}——分发前必须补上（LGPL 要求）")


def build_onedir():
    _run_pyinstaller(onefile=False)
    folder = ROOT / "dist" / NAME
    if not folder.exists():
        sys.exit("没看到 onedir 产物")
    _copy_legal(folder)
    zip_base = ROOT / "dist" / f"{NAME}-免安装版"
    out = shutil.make_archive(str(zip_base), "zip",
                              root_dir=folder.parent, base_dir=NAME)
    mb = Path(out).stat().st_size / (1024 ** 2)
    print(f"\n  打包完成：{out}（{mb:.0f} MB，解压后双击 {NAME}.exe）")
    print("  解压即用：不写临时目录、启动快、不容易被杀软拦。")


def build_onefile():
    _run_pyinstaller(onefile=True)
    exe = _exe_path()
    if not exe.exists():
        sys.exit("没看到产物文件，检查 dist/")
    mb = exe.stat().st_size / (1024 ** 2)
    print(f"\n  打包完成：{exe}（{mb:.0f} MB）")
    print("  注意：单文件每次启动会解压到 %TEMP% 下的 _MEIxxxxxx；"
          "被强杀会留残留，正常关窗口会自己清。")


def build(do_clean, mode="onefile"):
    check_not_running()
    if do_clean:
        clean()
    stage_rules()

    if mode == "onedir":
        build_onedir()
    elif mode == "both":
        build_onedir()
        build_onefile()
    else:
        build_onefile()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--clean", action="store_true", help="先清掉 build/dist")
    ap.add_argument("--mode", choices=("onefile", "onedir", "both"),
                    default="onefile",
                    help="单文件 / 文件夹 / 两个都出（默认单文件）")
    args = ap.parse_args()
    build(args.clean, args.mode)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
