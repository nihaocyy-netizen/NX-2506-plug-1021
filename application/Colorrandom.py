# NX2506 选中实体一键随机改色
# 使用方法：先在NX里选好实体，再运行本脚本
# 不满意按 Ctrl+Z 撤回

import NXOpen
import random

# ========== 配置 ==========
COLOR_RANGE = 100  # 随机颜色范围：1~200（NX标准色号）
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
            if isinstance(entities, NXOpen.NXObject):
                entities = (entities,)
            for entity in entities:
                collect_display_objects(entity, result, visited)
            if result:
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

    selected_objects = get_operation_objects(theUI, "随机改色")
    if not selected_objects:
        return
    display_objs = []
    visited = set()
    for obj in selected_objects:
        collect_display_objects(obj, display_objs, visited)

    if len(display_objs) == 0:
        lw = theSession.ListingWindow
        lw.Open()
        lw.WriteLine("[提示] 选中的对象无法找到可改色实体，请选择实体/面/曲线或特征")
        return

    # 记录撤销点
    theSession.SetUndoMark(
        NXOpen.Session.MarkVisibility.Visible,
        "随机改选中实体颜色"
    )

    # 给每个选中对象随机分配颜色
    for obj in display_objs:
        random_color = random.randint(1, COLOR_RANGE)
        dm = theSession.DisplayManager.NewDisplayModification()
        dm.ApplyToAllFaces = True
        dm.NewColor = random_color
        dm.Apply([obj])
        dm.Dispose()

    # 改色完成后清除全局选择及高亮，直接显示实际颜色。
    theUI.SelectionManager.ClearGlobalSelectionList()
    for obj in display_objs:
        obj.Unhighlight()

    # 输出结果
    lw = theSession.ListingWindow
    lw.Open()

if __name__ == '__main__':
    main()


