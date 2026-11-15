# NX：切换到该方向最近的有物体图层，自动跳过空层，不循环。
# 支持 Ctrl+Z 撤销，原工作层的对象仍保留显示。
import NXOpen


def main():
    session = NXOpen.Session.GetSession()
    ui = NXOpen.UI.GetUI()
    # GetAllObjectsOnLayer 要求查询当前显示零件。
    part = session.Parts.Display
    if part is None:
        ui.NXMessageBox.Show("上一图层", NXOpen.NXMessageBox.DialogType.Information,
                             "请先打开零件。")
        return

    mark = None
    try:
        layers = part.Layers
        current_layer = layers.WorkLayer
        target_layer = None
        for number in range(current_layer + (-1), 0, -1):
            # 隐藏的对象也算有物体；跳过没有可显示对象的层。
            if any(isinstance(obj, NXOpen.DisplayableObject)
                   for obj in layers.GetAllObjectsOnLayer(number)):
                target_layer = number
                break
        if target_layer is None:
            return
        mark = session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "切换上一工作图层")
        layers.WorkLayer = target_layer
    except Exception as error:
        if mark is not None:
            session.UndoToMark(mark, "")
        ui.NXMessageBox.Show("上一图层", NXOpen.NXMessageBox.DialogType.Error,
                             "切换失败，已撤销本次修改：\n" + str(error))


if __name__ == "__main__":
    main()
