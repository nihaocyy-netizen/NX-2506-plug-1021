#==============================================================================
# 最终修复版 export_body01.py
# 修复：Ok大小写、MessageBox传整型数字、dlx绝对路径
#==============================================================================
import NXOpen
import NXOpen.BlockStyler
import NXOpen.UF
import os
import tempfile

PROGRESS_FILE = os.path.join(tempfile.gettempdir(), "nx_batch_progress.txt")
# 优先使用脚本同目录的配套界面，支持整个工具目录移动部署。
APPLICATION_DIR = os.path.dirname(os.path.abspath(__file__))
DLX_ABS_PATH = os.path.join(APPLICATION_DIR, "Semiautodwg.dlx")
DEFAULT_DLX_PATH = DLX_ABS_PATH


def find_dialog_file():
    candidates = list(dict.fromkeys((DLX_ABS_PATH, DEFAULT_DLX_PATH)))
    for path in candidates:
        if os.path.isfile(path):
            return path
    raise FileNotFoundError("未找到配套界面文件，请将 Semiautodwg.dlx 放到 application 目录。\n已检查：\n" + "\n".join(candidates))


MSG_INFO = 1
MSG_ERR = 2

class export_body01:
    theSession = None
    theUI = None
    theUf = None
    workPart = None
    lw = None
    _selected_body_tags = None

    SnapPointTypesEnabled_UserDefined = 1 << 0
    SnapPointTypesEnabled_Inferred = 1 << 1
    SnapPointTypesEnabled_ScreenPosition = 1 << 2
    SnapPointTypesEnabled_EndPoint = 1 << 3
    SnapPointTypesEnabled_MidPoint = 1 << 4
    SnapPointTypesEnabled_ControlPoint = 1 << 5
    SnapPointTypesEnabled_Intersection = 1 << 6
    SnapPointTypesEnabled_ArcCenter = 1 << 7
    SnapPointTypesEnabled_QuadrantPoint = 1 << 8
    SnapPointTypesEnabled_ExistingPoint = 1 << 9
    SnapPointTypesEnabled_PointonCurve = 1 << 10
    SnapPointTypesEnabled_PointonSurface = 1 << 11
    SnapPointTypesEnabled_PointConstructor = 1 << 12
    SnapPointTypesEnabled_TwocurveIntersection = 1 << 13
    SnapPointTypesEnabled_TangentPoint = 1 << 14
    SnapPointTypesEnabled_Poles = 1 << 15
    SnapPointTypesEnabled_BoundedGridPoint = 1 << 16
    SnapPointTypesEnabled_FacetVertexPoint = 1 << 17
    SnapPointTypesEnabled_DefiningPoint = 1 << 18

    SnapPointTypesOnByDefault_EndPoint = 1 << 3
    SnapPointTypesOnByDefault_MidPoint = 1 << 4
    SnapPointTypesOnByDefault_ControlPoint = 1 << 5
    SnapPointTypesOnByDefault_Intersection = 1 << 6
    SnapPointTypesOnByDefault_ArcCenter = 1 << 7
    SnapPointTypesOnByDefault_QuadrantPoint = 1 << 8
    SnapPointTypesOnByDefault_ExistingPoint = 1 << 9
    SnapPointTypesOnByDefault_PointonCurve = 1 << 10
    SnapPointTypesOnByDefault_PointonSurface = 1 << 11
    SnapPointTypesOnByDefault_PointConstructor = 1 << 12
    SnapPointTypesOnByDefault_BoundedGridPoint = 1 << 16
    SnapPointTypesOnByDefault_FacetVertexPoint = 1 << 17

    def __init__(self):
        try:
            self.theSession = NXOpen.Session.GetSession()
            self.theUI = NXOpen.UI.GetUI()
            self.theUf = NXOpen.UF.UFSession.GetUFSession()
            self.workPart = self.theSession.Parts.Work
            self.lw = self.theSession.ListingWindow
            self.lw.Open()
            self.lw.WriteLine(f"进度文件路径：{PROGRESS_FILE}")
            self.theDlxFileName = find_dialog_file()
            self.lw.WriteLine(f"DLX完整路径：{self.theDlxFileName}")
            self.theDialog = self.theUI.CreateDialog(self.theDlxFileName)
            self.theDialog.AddOkHandler(self.ok_cb)
            self.theDialog.AddUpdateHandler(self.update_cb)
            self.theDialog.AddInitializeHandler(self.initialize_cb)
            self.theDialog.AddDialogShownHandler(self.dialogShown_cb)

            self.selectObject0 = None
            self._selected_body_tags = None
        except Exception as ex:
            raise ex

    def Launch(self):
        dialogResponse = NXOpen.BlockStyler.BlockDialog.DialogResponse.Invalid
        try:
            dialogResponse = self.theDialog.Launch()
        except Exception as ex:
            if self.theUI:
                self.theUI.NXMessageBox.Show("Block Styler", str(ex), MSG_ERR)
        return dialogResponse

    def Dispose(self):
        if self.theDialog is not None:
            self.theDialog.Dispose()
            self.theDialog = None
        if self.lw is not None:
            self.lw.Close()

    def _show_msg(self, title, msg, msg_type=MSG_INFO):
        if self.theUI is None:
            return
        self.theUI.NXMessageBox.Show(title, msg, msg_type)

    def initialize_cb(self):
        try:
            self.selectObject0 = self.theDialog.TopBlock.FindBlock("selectObject0")
        except Exception as ex:
            self._show_msg("初始化错误", str(ex), MSG_ERR)

    def dialogShown_cb(self):
        pass

    def update_cb(self, block):
        try:
            if block == self.selectObject0:
                pass
        except Exception as ex:
            self._show_msg("update_cb异常", str(ex), MSG_ERR)
        return 0

    def ok_cb(self):
        errorCode = 0
        try:
            sel_objs = self.selectObject0.GetSelectedObjects()
            tag_list = []
            for obj in sel_objs:
                typ, subtyp = self.theUf.Obj.AskTypeAndSubtype(obj.Tag)
                if typ == 70 and subtyp == 0:
                    tag_list.append(obj.Tag)
            if len(tag_list) <= 0:
                self._show_msg("提示", "未选择任何实体！", MSG_INFO)
                self._selected_body_tags = None
                return 1
            self._selected_body_tags = tag_list
        except Exception as ex:
            self._show_msg("读取选择异常", str(ex), MSG_ERR)
            self._selected_body_tags = None
            errorCode = 1
        return errorCode

    def run_next_export(self):
        lw = self.lw
        theUf = self.theUf
        all_bodies = list(self.workPart.Bodies)

        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                lines = f.readlines()
                start_idx = int(lines[0].strip())
                tag_list = [int(x.strip()) for x in lines[1:] if x.strip()]
            lw.WriteLine(f"[Debug] start_idx={start_idx} , 选中实体总数={len(tag_list)}")
        except Exception as e:
            lw.WriteLine(f"[Error]读取进度失败:{str(e)}")
            return

        if start_idx >= len(tag_list):
            lw.WriteLine("[Done] ✅全部实体处理完成!")
            self._show_msg("完成", "所有实体任务结束", MSG_INFO)
            if os.path.exists(PROGRESS_FILE):
                os.remove(PROGRESS_FILE)
                lw.WriteLine("[Info]已自动删除进度文件")
            return

        current_body = None
        for b in all_bodies:
            if b.Tag == tag_list[start_idx]:
                current_body = b
                break
        if current_body is None:
            lw.WriteLine(f"[Skip]实体 {start_idx+1} 找不到，跳过")
            next_idx = start_idx + 1
            with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
                f.write(f"{next_idx}\n")
                for t in tag_list:
                    f.write(f"{t}\n")
            self.run_next_export()
            return

        self.theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, f"Batch show body {start_idx+1}")
        for b in all_bodies:
            theUf.Obj.SetBlankStatus(b.Tag, 1)
        theUf.Obj.SetBlankStatus(current_body.Tag, 0)

        next_idx = start_idx + 1
        with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
            f.write(f"{next_idx}\n")
            for t in tag_list:
                f.write(f"{t}\n")

        lw.WriteLine("")
        lw.WriteLine(f">>> Body {start_idx+1}/{len(tag_list)} 已单独显示")
        lw.WriteLine("    完成当前实体操作后，再次运行本脚本切换下一个实体")
        if next_idx >= len(tag_list):
            lw.WriteLine("    (这是最后一个实体！)")


def main():
    theexport_body01 = None
    ui_global = None
    try:
        ui_global = NXOpen.UI.GetUI()
        theexport_body01 = export_body01()
        app = theexport_body01
        if app.workPart is None:
            app._show_msg("错误", "请先打开PRT零件！", MSG_ERR)
            return

        if os.path.exists(PROGRESS_FILE):
            try:
                with open(PROGRESS_FILE, "r", encoding="utf-8") as fp:
                    lines = fp.readlines()
                    start_idx = int(lines[0].strip())
                    total = len(lines)-1
                app.lw.WriteLine(f"[Debug]检测到进度: idx={start_idx}, total={total}")
                if start_idx >= total:
                    app.lw.WriteLine("[Info]废弃进度，自动删除，重新开始")
                    os.remove(PROGRESS_FILE)
                    ret = app.Launch()
                    # 修复：.Ok 大小写
                    if ret == NXOpen.BlockStyler.BlockDialog.DialogResponse.Ok and app._selected_body_tags is not None:
                        with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
                            f.write("0\n")
                            for t in app._selected_body_tags:
                                f.write(f"{t}\n")
                        app.run_next_export()
                else:
                    app.lw.WriteLine("[Info]继续未完成任务")
                    app.run_next_export()
            except Exception:
                app.lw.WriteLine("[Info]进度文件损坏，删除，重新选择实体")
                os.remove(PROGRESS_FILE)
                ret = app.Launch()
                if ret == NXOpen.BlockStyler.BlockDialog.DialogResponse.Ok and app._selected_body_tags is not None:
                    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
                        f.write("0\n")
                        for t in app._selected_body_tags:
                            f.write(f"{t}\n")
                    app.run_next_export()
        else:
            app.lw.WriteLine("[Info]无进度文件，弹出块UI选择实体窗口")
            ret = app.Launch()
            if ret == NXOpen.BlockStyler.BlockDialog.DialogResponse.Ok and app._selected_body_tags is not None:
                with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
                    f.write("0\n")
                    for t in app._selected_body_tags:
                        f.write(f"{t}\n")
                app.run_next_export()

    except Exception as ex:
        if ui_global is not None:
            ui_global.NXMessageBox.Show("脚本异常", str(ex), MSG_ERR)
    finally:
        if theexport_body01 is not None:
            theexport_body01.Dispose()
            theexport_body01 = None

if __name__ == '__main__':
    main()
