"""清除半自动出图进度；可从任意安装目录运行，无需配套 DLX。"""
import os
import tempfile

import NXOpen


# 必须与 Semiautodwg.py 使用相同的进度文件。
PROGRESS_FILE = os.path.join(tempfile.gettempdir(), "nx_batch_progress.txt")


def reset_progress(path):
    """直接删除避免先检查再删除的竞争；不存在时视为无需重置。"""
    try:
        os.remove(path)
    except FileNotFoundError:
        return False
    return True


def main():
    ui = NXOpen.UI.GetUI()
    try:
        removed = reset_progress(PROGRESS_FILE)
    except PermissionError as error:
        ui.NXMessageBox.Show(
            "重置失败", NXOpen.NXMessageBox.DialogType.Error,
            "无法清除进度文件，可能被占用或没有删除权限。\n"
            "请等待批量脚本执行结束后重试。\n\n文件：{}\n详情：{}".format(PROGRESS_FILE, error))
        return
    except OSError as error:
        ui.NXMessageBox.Show(
            "重置失败", NXOpen.NXMessageBox.DialogType.Error,
            "清除进度文件失败。\n文件：{}\n详情：{}".format(PROGRESS_FILE, error))
        return

    if removed:
        message = "批量任务进度已清除。\n再次运行半自动出图，可重新选择实体。"
    else:
        message = "当前没有批量任务进度。\n可直接运行半自动出图选择实体。"
    ui.NXMessageBox.Show("重置进度", NXOpen.NXMessageBox.DialogType.Information, message)


if __name__ == "__main__":
    main()
