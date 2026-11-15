# -*- coding: utf-8 -*-
"""NX 菜单入口：一键启动原始 GongCha.exe。"""

import os
import NXOpen


def find_gongcha_exe():
    """在当前项目及其父目录中的 mytool-1021 文件夹内查找 GongCha.exe。"""
    current = os.path.abspath(os.path.dirname(__file__))
    candidates = []

    for _ in range(6):
        folder = os.path.join(current, "mytool-1021")
        if os.path.isdir(folder):
            direct = os.path.join(folder, "GongCha.exe")
            if os.path.isfile(direct):
                candidates.append(direct)
            for root, _dirs, files in os.walk(folder):
                for name in files:
                    if name.lower() == "gongcha.exe":
                        candidates.append(os.path.join(root, name))
        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent

    if candidates:
        return os.path.abspath(candidates[0])
    raise FileNotFoundError("在 mytool-1021 文件夹中找不到 GongCha.exe")


def main():
    try:
        exe_path = find_gongcha_exe()

        # Windows 下直接启动独立程序，不阻塞 NX 主线程。
        os.startfile(exe_path)
    except Exception as exc:
        try:
            NXOpen.UI.GetUI().NXMessageBox.Show(
                "公差查询手册",
                NXOpen.NXMessageBox.DialogType.Error,
                str(exc),
            )
        except Exception:
            pass


if __name__ == "__main__":
    main()
