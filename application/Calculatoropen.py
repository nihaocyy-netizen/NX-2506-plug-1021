# -*- coding: utf-8 -*-
"""NX 菜单入口：打开计算器。"""

import os
import NXOpen


def find_calculator():
    """优先查找 mytool-1021 中的计算器程序，找不到则使用系统计算器。"""
    current = os.path.abspath(os.path.dirname(__file__))
    names = {"计算器.exe", "calculator.exe", "calc.exe"}

    for _ in range(6):
        folder = os.path.join(current, "mytool-1021")
        if os.path.isdir(folder):
            for root, _dirs, files in os.walk(folder):
                for name in files:
                    if name.lower() in {item.lower() for item in names}:
                        return os.path.join(root, name)
        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent

    return "calc.exe"


def main():
    try:
        os.startfile(find_calculator())
    except Exception as exc:
        try:
            NXOpen.UI.GetUI().NXMessageBox.Show(
                "计算器",
                NXOpen.NXMessageBox.DialogType.Error,
                str(exc),
            )
        except Exception:
            pass


if __name__ == "__main__":
    main()
