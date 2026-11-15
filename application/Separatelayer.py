# NX2506 选中实体自动分到不同图层
# 使用方法：先在NX里选好实体，再运行本脚本
# 不满意按 Ctrl+Z 撤回

import NXOpen

# ========== 配置 ==========
START_LAYER = 150   # 从第几层开始放
LAYER_STEP = 1    # 每个实体间隔几层
# ==========================

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

    selected_objects = get_operation_objects(theUI, "实体自动分层")
    if not selected_objects:
        return
    display_objs = []
    visited = set()
    for obj in selected_objects:
        collect_display_objects(obj, display_objs, visited)
    solid_bodies = [obj for obj in display_objs
                    if isinstance(obj, NXOpen.Body) and obj.IsSolidBody]

    if len(solid_bodies) == 0:
        lw = theSession.ListingWindow
        lw.Open()
        lw.WriteLine("[提示] 选中的对象里没有实体，请选择实体或生成实体的特征再运行")
        return

    # 记录撤销点
    theSession.SetUndoMark(
        NXOpen.Session.MarkVisibility.Visible,
        "实体分配到图层"
    )

    # 把每个实体分到不同图层
    layer_list = []
    for idx, body in enumerate(solid_bodies):
        target_layer = START_LAYER + idx * LAYER_STEP
        dm = theSession.DisplayManager.NewDisplayModification()
        dm.NewLayer = target_layer
        dm.Apply([body])
        dm.Dispose()
        layer_list.append(target_layer)

    # 完成后取消选择，清除黄色高亮。
    theUI.SelectionManager.ClearGlobalSelectionList()
    for obj in display_objs:
        obj.Unhighlight()

    # 输出结果
    lw = theSession.ListingWindow
    lw.Open()
    lw.WriteLine(f"已将 {len(solid_bodies)} 个实体分配到不同图层：")
    for i, layer in enumerate(layer_list):
        lw.WriteLine(f"  实体{i+1} → 第 {layer} 层")
    lw.WriteLine("")
    lw.WriteLine("💡 不满意？直接按 Ctrl+Z 撤回")

if __name__ == '__main__':
    main()
