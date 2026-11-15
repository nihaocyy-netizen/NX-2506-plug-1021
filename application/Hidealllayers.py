# NX：一键关闭所有非工作层，保留当前工作层。
# 工作层不能隐藏；Ctrl+Z 撤销本次图层操作。
import NXOpen
import NXOpen.Layer


def main():
    session = NXOpen.Session.GetSession()
    ui = NXOpen.UI.GetUI()
    part = session.Parts.Display
    if part is None:
        ui.NXMessageBox.Show("图层全关", NXOpen.NXMessageBox.DialogType.Information,
                             "请先打开零件。")
        return

    mark = session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "图层全关")
    try:
        layers = part.Layers
        work_layer = layers.WorkLayer
        # 使用单层接口，跳过不能隐藏的工作层。
        for number in range(1, 257):
            if number != work_layer:
                layers.SetState(number, NXOpen.Layer.State.Hidden, False)
    except Exception as error:
        session.UndoToMark(mark, "")
        ui.NXMessageBox.Show("图层全关", NXOpen.NXMessageBox.DialogType.Error,
                             "关闭图层失败，已撤销本次修改：\n" + str(error))


if __name__ == "__main__":
    main()
