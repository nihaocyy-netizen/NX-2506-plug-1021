# NX：一键打开全部 256 个图层，保留当前工作层。
# 运行后所有非工作层均可见、可选；Ctrl+Z 撤销。
import NXOpen
import NXOpen.Layer


def main():
    session = NXOpen.Session.GetSession()
    ui = NXOpen.UI.GetUI()
    part = session.Parts.Display
    if part is None:
        ui.NXMessageBox.Show("图层全开", NXOpen.NXMessageBox.DialogType.Information,
                             "请先打开零件。")
        return

    mark = session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "图层全开")
    try:
        layers = part.Layers
        work_layer = layers.WorkLayer
        # SetStates 需要 StateCollection，不能直接传入 Python list。
        # 使用已核对的单层接口，跳过不允许直接修改状态的工作层。
        for number in range(1, 257):
            if number != work_layer:
                layers.SetState(number, NXOpen.Layer.State.Selectable, False)
    except Exception as error:
        session.UndoToMark(mark, "")
        ui.NXMessageBox.Show("图层全开", NXOpen.NXMessageBox.DialogType.Error,
                             "打开图层失败，已撤销本次修改：\n" + str(error))


if __name__ == "__main__":
    main()
