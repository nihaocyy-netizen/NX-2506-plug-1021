# 在 Windows 资源管理器中打开当前 NX 工作零件所在的文件夹
# 使用方法：在 NX 中打开已保存的零件，再运行本脚本。

import os

import NXOpen


def show_message(ui, message, is_error=False):
    dialog_type = (NXOpen.NXMessageBox.DialogType.Error if is_error
                   else NXOpen.NXMessageBox.DialogType.Information)
    ui.NXMessageBox.Show("打开文件所在文件夹", dialog_type, message)


def main():
    session = NXOpen.Session.GetSession()
    ui = NXOpen.UI.GetUI()
    part = session.Parts.Work
    if part is None:
        part = session.Parts.Display
    if part is None:
        show_message(ui, "请先打开零件文件。")
        return

    file_path = part.FullPath
    if not file_path:
        show_message(ui, "当前零件尚未保存，请先保存后再打开所在文件夹。")
        return

    # Teamcenter 管理模式中的零件标识不一定对应本地文件路径。
    if not os.path.isabs(file_path):
        show_message(ui, "当前零件没有可用的本地文件路径：\n" + file_path)
        return
    folder_path = os.path.dirname(file_path)
    if not os.path.isdir(folder_path):
        show_message(ui, "文件所在文件夹不存在或无法访问：\n" + folder_path,
                     is_error=True)
        return

    try:
        os.startfile(folder_path)
    except OSError as error:
        show_message(ui, "无法打开文件夹：\n{}\n\n{}".format(folder_path, error),
                     is_error=True)


if __name__ == "__main__":
    main()
