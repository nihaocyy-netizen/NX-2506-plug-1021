"""仅显示选定对象；隐藏其他实体、曲线、点、草图、基准和注释等。"""
import NXOpen


def collect_display_objects(obj, result, visited):
    if obj is None:
        return
    tag = obj.Tag
    if tag in visited:
        return
    visited.add(tag)
    # NX 不支持独立显示实体面/边，选择面或边时保留所属实体。
    if isinstance(obj, (NXOpen.Face, NXOpen.Edge)):
        collect_display_objects(obj.GetBody(), result, visited)
        return
    if isinstance(obj, NXOpen.DisplayableObject):
        result.append(obj)
        return
    for name in ("GetBodies", "GetResultBodies", "GetBody", "GetResultBody",
                 "GetEntities", "GetFaces", "GetEdges"):
        method = getattr(obj, name, None)
        if not callable(method):
            continue
        try:
            entities = method()
        except Exception:
            continue
        if not entities:
            continue
        if isinstance(entities, NXOpen.NXObject):
            entities = [entities]
        before = len(result)
        for entity in entities:
            collect_display_objects(entity, result, visited)
        if len(result) > before:
            return


def get_operation_objects(ui, title):
    """支持先选对象再执行，以及先点按钮再选择；取消时不执行。"""
    manager = ui.SelectionManager
    count = manager.GetNumSelectedObjects()
    if count:
        return [manager.GetSelectedTaggedObject(i) for i in range(count)]
    response, objects = manager.SelectTaggedObjects(
        "请选择要操作的物体，可多选或框选，确认后执行", title,
        NXOpen.Selection.SelectionScope.AnyInAssembly,
        True, False)
    if response != NXOpen.Selection.Response.Ok:
        return []
    return list(objects)


def main():
    session = NXOpen.Session.GetSession()
    ui = NXOpen.UI.GetUI()
    if session.Parts.Work is None:
        ui.NXMessageBox.Show("仅显示", NXOpen.NXMessageBox.DialogType.Information, "请先打开零件。")
        return
    selection = ui.SelectionManager
    selected_objects = get_operation_objects(ui, "仅显示")
    if not selected_objects:
        return
    keep = []
    visited = set()
    for obj in selected_objects:
        collect_display_objects(obj, keep, visited)
    if not keep:
        ui.NXMessageBox.Show("仅显示", NXOpen.NXMessageBox.DialogType.Information,
                             "请先选择要保留显示的物体、曲线、点或特征。")
        return

    mark = session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "仅显示选定对象")
    try:
        # 使用 NX 的全部类型隐藏，而非仅遍历 Bodies，涵盖非实体元素。
        session.DisplayManager.HideByType(
            NXOpen.DisplayManager.ShowHideType.All,
            NXOpen.DisplayManager.ShowHideScope.AnyInAssembly)
        session.DisplayManager.UnblankObjects(keep)
        # 恢复被隐藏的父组件，以便装配中的选定对象可见。
        parents = []
        parent_tags = set()
        for obj in keep:
            parent = getattr(obj, "OwningComponent", None)
            while parent is not None and parent.Tag not in parent_tags:
                parent_tags.add(parent.Tag)
                parents.append(parent)
                parent = parent.Parent
        if parents:
            session.DisplayManager.UnblankObjects(parents)
    except Exception as error:
        try:
            session.UndoToMark(mark, "")
        except Exception as undo_error:
            ui.NXMessageBox.Show("仅显示失败", NXOpen.NXMessageBox.DialogType.Error,
                                 "{}\n撤销失败，请检查显示状态：{}".format(error, undo_error))
            return
        ui.NXMessageBox.Show("仅显示失败", NXOpen.NXMessageBox.DialogType.Error, str(error))
        return

    selection.ClearGlobalSelectionList()
    for obj in keep:
        obj.Unhighlight()


if __name__ == "__main__":
    main()
