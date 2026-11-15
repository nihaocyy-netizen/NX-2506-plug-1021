# NX2506 选中对象一键设置透明度
# 使用方法：先在NX里点选/框选对象，再运行本脚本

import NXOpen

# ========== 透明度配置 ==========
TRANSLUCENCY = 50  # 透明度：0=完全不透明, 100=完全透明
# ==============================

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
    theSession = NXOpen.Session.GetSession()
    theUI = NXOpen.UI.GetUI()
    workPart = theSession.Parts.Work

    if workPart is None:
        return

    selected_objects = get_operation_objects(theUI, "设置透明")
    if not selected_objects:
        return
    display_objs = []
    visited = set()
    for obj in selected_objects:
        collect_display_objects(obj, display_objs, visited)

    if len(display_objs) == 0:
        lw = theSession.ListingWindow
        lw.Open()
        lw.WriteLine("[提示] 选中的对象无法找到可设置透明度的几何体，请选择实体、面或特征")
        return

    # ===== 记录撤销点（改完不满意按Ctrl+Z撤回）=====
    theSession.SetUndoMark(
        NXOpen.Session.MarkVisibility.Visible,
        "设置对象透明度"
    )

    # 自动启用 NX 会话的半透明显示总开关。
    theSession.Preferences.VisualizationVisualPreferences.Translucency = True

    # 设置透明度
    dm = theSession.DisplayManager.NewDisplayModification()
    try:
        # NX 透明度作用于面，选择实体/特征时必须同时应用到所有面。
        dm.ApplyToAllFaces = True
        dm.NewTranslucency = TRANSLUCENCY
        dm.Apply(display_objs)
    finally:
        dm.Dispose()

    # 操作成功后清除选择与高亮，直接显示透明度效果。
    theUI.SelectionManager.ClearGlobalSelectionList()
    for obj in display_objs:
        obj.Unhighlight()

    # 刷新当前显示视图，让半透明立即生效。
    displayPart = theSession.Parts.Display
    if displayPart is not None:
        displayPart.ModelingViews.WorkView.Regenerate()

    # 输出结果
    lw = theSession.ListingWindow
    lw.Open()
    #lw.WriteLine(f"已设置 {len(display_objs)} 个选中对象的透明度为 {TRANSLUCENCY}%")
    #lw.WriteLine("")
    #lw.WriteLine("💡 改完不满意？直接按 Ctrl+Z 就能撤回")

if __name__ == '__main__':
    main()
