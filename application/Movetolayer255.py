# 在 NX 中将选定对象一键移动至 255 层
# 先选择实体、曲线等对象或特征，再运行；Ctrl+Z 可撤销。
import NXOpen

TARGET_LAYER = 255


def collect_display_objects(selected_obj, result, visited):
    """收集可显示对象；如果选中的是特征，则展开其生成的实体。"""
    tag = getattr(selected_obj, "Tag", None)
    if tag in visited:
        return
    if tag is not None:
        visited.add(tag)

    if isinstance(selected_obj, NXOpen.DisplayableObject):
        result.append(selected_obj)
        return

    # 不同特征类型提供的输出接口不同：BodyFeature 通常要用
    # GetBodies()，曲线/草图类特征则常用 GetEntities()、GetFaces()
    # 或 GetEdges()。依次尝试这些接口，兼容更多 NX 特征。
    for method_name in (
        "GetBodies", "GetResultBodies", "GetBody", "GetResultBody",
        "GetEntities", "GetFaces", "GetEdges"
    ):
        method = getattr(selected_obj, method_name, None)
        if method is None:
            continue
        try:
            entities = method()
        except Exception:
            continue
        if entities:
            # GetBody()/GetResultBody() 可能返回单个对象，其余接口通常返回集合。
            if isinstance(entities, NXOpen.NXObject):
                entities = (entities,)
            previous_count = len(result)
            for entity in entities:
                collect_display_objects(entity, result, visited)
            if len(result) > previous_count:
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
        ui.NXMessageBox.Show("移动至255层", NXOpen.NXMessageBox.DialogType.Information,
                             "请先打开零件。")
        return

    selection = ui.SelectionManager
    selected_objects = get_operation_objects(ui, "移动至255层")
    if not selected_objects:
        return
    objects = []
    visited = set()
    for obj in selected_objects:
        collect_display_objects(obj, objects, visited)

    if not objects:
        ui.NXMessageBox.Show("移动至255层", NXOpen.NXMessageBox.DialogType.Information,
                             "请先选择需要移动的对象或特征。")
        return

    mark = session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "移动选定对象至255层")
    modification = None
    try:
        modification = session.DisplayManager.NewDisplayModification()
        modification.NewLayer = TARGET_LAYER
        modification.Apply(objects)
    except Exception as error:
        session.UndoToMark(mark, "")
        ui.NXMessageBox.Show("移动至255层", NXOpen.NXMessageBox.DialogType.Error,
                             "移动失败，已撤销本次修改：\n" + str(error))
    else:
        selection.ClearGlobalSelectionList()
        for obj in objects:
            obj.Unhighlight()
    finally:
        if modification is not None:
            modification.Dispose()


if __name__ == "__main__":
    main()