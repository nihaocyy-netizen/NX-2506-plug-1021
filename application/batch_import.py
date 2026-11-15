# -*- coding: utf-8 -*-
r"""
批量导入多个文件：每个文件各自成为一个独立部件。

支持格式：
    .prt                 直接打开为显示部件（已在会话中打开则切换为活动显示）
    .stp / .step         STEP AP214 导入到新建部件
    .igs / .iges         IGES 导入到新建部件
    .x_t / .x_b          Parasolid 导入到新建部件

行为约定：
    1. 多选使用 Windows 系统文件对话框（ctypes 调 comdlg32.GetOpenFileNameW），
       不依赖 tkinter —— NX 内嵌 Python 不一定带 tkinter。
    2. 每个文件一个独立部件；不把几何导入到当前工作部件里。
    3. 绝不覆盖磁盘上的用户文件：.prt 输入只打开、只显示、不另存；
       转换类（STEP/IGES/Parasolid）另存为源文件同目录下的同名 .prt，
       同名已存在则加 _2/_3。
    4. 单个文件失败只记录、继续下一个，不中断整批。
    5. 进度写 NX 信息窗口（ListingWindow），日志写插件目录 logs\batch_import.log。
    6. 取消选择即静默退出。

注意：不要在本文件里定义 startup(argc, args)。
startup 目录下的 .py 会被 NX 当成“启动应用”导入；一旦它有 startup 函数，
之后点菜单用 journal 播放同一模块时 nxpythonldr 会崩（syslog 里是 signal 11，
界面弹空白的 Journal execution error）。没有 startup 时 NX 只会在 syslog 记一条
“no attribute startup”的提示，功能正常。
"""

import os
import re
import time
import tempfile
import ctypes

import NXOpen

# ================= 多语言 =================
# UGII_LANG 的常见取值：simpl_chinese / trad_chinese / english
# 未知或未设置时回落英文；三张表缺项时逐项回落英文。


def _detect_lang():
    value = (os.environ.get("UGII_LANG") or "").strip().lower()
    if value in ("simpl_chinese", "simplified_chinese", "zh_cn", "zh-cn", "chs",
                 "schinese", "zh_hans"):
        return "zh_cn"
    if value in ("trad_chinese", "traditional_chinese", "zh_tw", "zh-tw", "cht",
                 "chinese", "zh_hant"):
        return "zh_tw"
    return "en"


LANG = _detect_lang()

MSG = {
    "en": {
        "imp.title": "Select files to import",
        "imp.filter": "Supported files",
        "imp.filter_all": "All files",
        "imp.ui.title": "Select files to import",
        "imp.ui.tree_label": "Folders (click to select, double-click to expand)",
        "imp.ui.list_label": "*.prt / *.stp / *.igs / *.x_t files in current folder (Ctrl/Shift multi-select, double-click to add)",
        "imp.ui.col_name": "Name",
        "imp.ui.col_size": "Size (KB)",
        "imp.ui.this_pc": "This PC",
        "imp.ui.denied": "Cannot access this folder (permission denied or invalid path):",
        "imp.ui.denied_hint": "The folder does not exist or access is denied.",
        "imp.ui.empty_dir": "(No importable file in this folder)",
        "imp.ui.grid_label": "Selected files ({0})",
        "imp.ui.col_type": "Type",
        "imp.ui.col_path": "Path",
        "imp.ui.btn_add_file": "+ Add file",
        "imp.ui.btn_clear": "Clear",
        "imp.ui.btn_cancel": "Cancel",
        "imp.ui.btn_ok": "Import",
        "imp.ui.note_group": "How the import works",
        "imp.ui.note_text": "Each file is opened as its own part and all of them stay loaded in the NX session. .prt is only opened. STEP/IGES/Parasolid is also saved as <source>.prt next to the source file (_2/_3 if the name exists).",
        "imp.ui.msg_remove_title": "Remove file",
        "imp.ui.msg_remove_text": "Remove this file from the list?",
        "imp.ui.msg_empty_title": "Nothing selected",
        "imp.ui.msg_empty_one": "Pick a folder on the left, then double-click a file (or select it and click Add Files).",
        "imp.ui.msg_empty_two": "The selected-file list is empty, nothing to import.",
        "imp.err.dialog": "Cannot open the system file dialog: {0}",
        "imp.confirm.title": "Batch Import",
        "imp.confirm.body": "{0} file(s) to import:\n\n{1}",
        "imp.confirm.more": "\n... and {0} more (see the log)",
        "imp.confirm.proceed_log": "Each file becomes an independent part. .prt is opened only; STEP/IGES/Parasolid are saved as <source>.prt next to the source file (_2/_3 if the name exists).",
        "imp.sum.title": "Batch Import Finished",
        "imp.sum.body": "Succeeded {0} / Failed {1} / Skipped {2}\n\n{3}\n\nLog: {4}",
        "imp.sum.fail_item": "- {0}: {1}",
        "imp.sum.no_fail": "(no failures)",
        "imp.sum.more_fail": "... and {0} more failure(s), see the log",
        "imp.sum.skipped": "{0} file(s) skipped (duplicate or missing), see the log.",
        "imp.sum.disk": "Saved to disk {0} / memory only {1}",
        "imp.sum.temp_logs": "Failed items keep translator logs in: {0}",
        "imp.log.header": "Batch import started: {0} file(s); log {1}",
        "imp.log.file": "[{0}/{1}] {2}",
        "imp.log.skip_dup": "  Skipped (duplicate): {0}",
        "imp.log.skip_missing": "  Skipped (file not found): {0}",
        "imp.log.prt_open": "  Open .prt: {0}",
        "imp.log.prt_already": "  Already open in this session, switched to active display",
        "imp.log.prt_done": "  Done {0:.1f}s -> part {1}",
        "imp.log.new_part": "  New display part: {0}",
        "imp.log.import_begin": "  Import {0} -> part {1}",
        "imp.log.import_done": "  Done {0:.1f}s",
        "imp.log.saveas_start": "  Save as: {0}",
        "imp.log.saveas_ok": "  Saved as: {0}",
        "imp.log.saveas_rename": "  Target exists, renamed to: {0}",
        "imp.err.saveas": "  Save failed ({0}): {1}; geometry kept in memory (part not closed)",
        "imp.err.saveas_missing": "  Save failed: target file not written {0}",
        "imp.log.load_status": "  Part load status: {0}",
        "imp.log.fail": "  Failed: {0}",
        "imp.log.fit_fail": "  Fit view failed (ignored): {0}",
        "imp.log.destroy_fail": "  Destroy importer failed (ignored): {0}",
        "imp.log.work_part": "  Work part: {0}",
        "imp.log.close_empty": "  Empty part closed: {0}",
        "imp.log.close_fail": "  Cannot close the empty part (kept in session): {0}",
        "imp.log.import_to": "  ImportTo not set (continuing with default work part): {0}",
        "imp.log.prep": "  Importer prepared: {0}; OutputFile={1}",
        "imp.log.prep_fail": "  Cannot prepare importer (continuing with defaults): {0}",
        "imp.log.settings_missing": "  STEP214 settings file not found, using the importer defaults",
        "imp.log.sew_probe": "  {0} importer default SewSurfaces = {1}",
        "imp.log.sew_probe_fail": "  Cannot read {0} importer SewSurfaces: {1}",
        "imp.log.sew_set": "  {0} importer SewSurfaces set to True",
        "imp.log.sew_set_fail": "  Cannot set {0} importer SewSurfaces: {1}",
        "imp.log.sew_unavailable": "  {0} importer has no SewSurfaces property, skipped",
        "imp.log.temp_clean": "  Removed {0} translator log(s) of this run from: {1}",
        "imp.log.temp_clean_old": "  Removed {0} translator log(s) older than {1} day(s)",
        "imp.log.summary": "Summary: succeeded {0} / failed {1} / skipped {2}",
        "imp.log.log_fallback": "Plugin folder not writable, log redirected to: {0}",
        "imp.log.no_confirm": "Confirmation dialog unavailable, continuing",
        "imp.log.open_mode": "  Open mode: {0}",
        "imp.log.keep_open": "  Work/display switch: {0}",
        "imp.log.parts_check": "  Session check: {0} part(s) expected, {1} still loaded",
        "imp.log.parts_restored": "  Re-opened {0} part(s) that the display switch had closed",
        "imp.log.parts_missing": "  Parts no longer in the session and cannot be re-opened: {0}",
        "imp.log.restore_fail": "  Re-open failed ({0}): {1}",
        "imp.reason.open": "Cannot open part: {0}",
        "imp.reason.import": "Import failed: {0}",
        "imp.reason.no_part": "Cannot create new part: {0}",
        "imp.reason.work_part": "Cannot set the new part as work part: {0}",
        "imp.reason.ext": "Unsupported file type",
        "imp.msg.cancelled": "Cancelled, nothing imported.",
        "imp.msg.no_files": "No importable file.",
    },
    "zh_cn": {
        "imp.title": "选择要导入的文件",
        "imp.filter": "支持的文件",
        "imp.filter_all": "所有文件",
        "imp.ui.title": "选择要导入的文件",
        "imp.ui.tree_label": "文件夹（单击选中，双击展开）",
        "imp.ui.list_label": "当前文件夹中的 .prt / .stp / .igs / .x_t（可 Ctrl/Shift 多选，双击加入）",
        "imp.ui.col_name": "文件名",
        "imp.ui.col_size": "大小(KB)",
        "imp.ui.this_pc": "此电脑",
        "imp.ui.denied": "无法访问该目录（权限不足或路径无效）：",
        "imp.ui.denied_hint": "文件夹不存在或没有访问权限。",
        "imp.ui.empty_dir": "（该目录下没有可导入的文件）",
        "imp.ui.grid_label": "已选文件清单（共 {0} 个）",
        "imp.ui.col_type": "类型",
        "imp.ui.col_path": "路径",
        "imp.ui.btn_add_file": "＋ 加入文件",
        "imp.ui.btn_clear": "清空",
        "imp.ui.btn_cancel": "取消",
        "imp.ui.btn_ok": "确认导入",
        "imp.ui.note_group": "导入说明",
        "imp.ui.note_text": "每个文件各自打开为一个独立部件，全部保留在 NX 会话中；.prt 只打开不另存，STEP/IGES/Parasolid 会另存为源文件同目录下的同名 .prt（同名自动加 _2/_3）。",
        "imp.ui.msg_remove_title": "移除文件",
        "imp.ui.msg_remove_text": "从已选清单中移除这个文件？",
        "imp.ui.msg_empty_title": "还没有选择文件",
        "imp.ui.msg_empty_one": "请先在左侧选择文件夹，再双击文件（或选中后点「添加文件」）加入清单。",
        "imp.ui.msg_empty_two": "已选文件清单为空，无法导入。",
        "imp.err.dialog": "无法打开系统文件选择框：{0}",
        "imp.confirm.title": "批量导入",
        "imp.confirm.body": "共 {0} 个文件待导入：\n\n{1}",
        "imp.confirm.more": "\n…… 其余 {0} 个见日志",
        "imp.confirm.proceed_log": "确认后每个文件各自打开为一个独立部件；.prt 只打开不另存，STEP/IGES/Parasolid 会另存为源文件同目录下的同名 .prt（同名自动加 _2/_3）。",
        "imp.sum.title": "批量导入完成",
        "imp.sum.body": "成功 {0} / 失败 {1} / 跳过 {2}\n\n{3}\n\n日志：{4}",
        "imp.sum.fail_item": "· {0}：{1}",
        "imp.sum.no_fail": "（没有失败项）",
        "imp.sum.more_fail": "…… 另有 {0} 个失败项，见日志",
        "imp.sum.skipped": "另有 {0} 个文件被跳过（重复选择或文件不存在），详见日志。",
        "imp.sum.disk": "已落盘 {0} / 仅内存 {1}",
        "imp.sum.temp_logs": "失败项的转换器日志保留在：{0}",
        "imp.log.header": "开始批量导入：共 {0} 个文件；日志 {1}",
        "imp.log.file": "[{0}/{1}] {2}",
        "imp.log.skip_dup": "  跳过（重复选择）：{0}",
        "imp.log.skip_missing": "  跳过（文件不存在）：{0}",
        "imp.log.prt_open": "  打开 .prt：{0}",
        "imp.log.prt_already": "  该部件已在会话中打开，切换为活动显示",
        "imp.log.prt_done": "  完成 {0:.1f}s -> 部件 {1}",
        "imp.log.new_part": "  新建显示部件：{0}",
        "imp.log.import_begin": "  导入 {0} -> 部件 {1}",
        "imp.log.import_done": "  完成 {0:.1f}s",
        "imp.log.saveas_start": "  开始另存：{0}",
        "imp.log.saveas_ok": "  另存成功：{0}",
        "imp.log.saveas_rename": "  同名文件已存在，改名为：{0}",
        "imp.err.saveas": "  落盘失败（{0}）：{1}；几何保留在内存，未关闭部件",
        "imp.err.saveas_missing": "  落盘失败：目标文件未生成 {0}",
        "imp.log.load_status": "  部件加载状态：{0}",
        "imp.log.fail": "  失败：{0}",
        "imp.log.fit_fail": "  拟合视图失败（忽略）：{0}",
        "imp.log.destroy_fail": "  销毁导入器失败（忽略）：{0}",
        "imp.log.work_part": "  工作部件：{0}",
        "imp.log.close_empty": "  已关闭空部件：{0}",
        "imp.log.close_fail": "  空部件关闭失败（保留在会话中）：{0}",
        "imp.log.import_to": "  ImportTo 设置失败（按默认工作部件继续）：{0}",
        "imp.log.prep": "  已补齐导入器参数：{0}；OutputFile={1}",
        "imp.log.prep_fail": "  导入器参数补齐失败（按默认继续）：{0}",
        "imp.log.settings_missing": "  未找到 STEP214 设置文件，按导入器默认值处理",
        "imp.log.sew_probe": "  {0} 导入器当前 SewSurfaces 默认值 = {1}",
        "imp.log.sew_probe_fail": "  读取 {0} 导入器 SewSurfaces 失败：{1}",
        "imp.log.sew_set": "  {0} 导入器 SewSurfaces 已设为 True",
        "imp.log.sew_set_fail": "  设置 {0} 导入器 SewSurfaces 失败：{1}",
        "imp.log.sew_unavailable": "  {0} 导入器没有 SewSurfaces 属性，已跳过",
        "imp.log.temp_clean": "  已清理本次产生的 {0} 个转换器日志：{1}",
        "imp.log.temp_clean_old": "  已清理 {1} 天前的 {0} 个转换器日志",
        "imp.log.summary": "汇总：成功 {0} / 失败 {1} / 跳过 {2}",
        "imp.log.log_fallback": "插件目录不可写，日志改写到：{0}",
        "imp.log.no_confirm": "无法弹出确认框，按继续处理",
        "imp.log.open_mode": "  打开方式：{0}",
        "imp.log.keep_open": "  工作/显示切换方式：{0}",
        "imp.log.parts_check": "  会话核对：应有 {0} 个部件，当前仍在会话中 {1} 个",
        "imp.log.parts_restored": "  已重新打开被显示切换关闭的部件 {0} 个",
        "imp.log.parts_missing": "  以下部件已不在会话中且无法重新打开：{0}",
        "imp.log.restore_fail": "  重新打开失败（{0}）：{1}",
        "imp.reason.open": "无法打开部件：{0}",
        "imp.reason.import": "导入失败：{0}",
        "imp.reason.no_part": "未能创建新部件：{0}",
        "imp.reason.work_part": "无法把新部件设为工作部件：{0}",
        "imp.reason.ext": "不支持的文件类型",
        "imp.msg.cancelled": "已取消，未导入任何文件。",
        "imp.msg.no_files": "没有可导入的文件。",
    },
    "zh_tw": {
        "imp.title": "選擇要匯入的檔案",
        "imp.filter": "支援的檔案",
        "imp.filter_all": "所有檔案",
        "imp.ui.title": "選擇要匯入的檔案",
        "imp.ui.tree_label": "資料夾（單擊選取，雙擊展開）",
        "imp.ui.list_label": "目前資料夾中的 .prt / .stp / .igs / .x_t（可 Ctrl/Shift 多選，雙擊加入）",
        "imp.ui.col_name": "檔案名稱",
        "imp.ui.col_size": "大小(KB)",
        "imp.ui.this_pc": "本機",
        "imp.ui.denied": "無法存取此目錄（權限不足或路徑無效）：",
        "imp.ui.denied_hint": "資料夾不存在或沒有存取權限。",
        "imp.ui.empty_dir": "（此目錄下沒有可匯入的檔案）",
        "imp.ui.grid_label": "已選檔案清單（共 {0} 個）",
        "imp.ui.col_type": "類型",
        "imp.ui.col_path": "路徑",
        "imp.ui.btn_add_file": "＋ 加入檔案",
        "imp.ui.btn_clear": "清空",
        "imp.ui.btn_cancel": "取消",
        "imp.ui.btn_ok": "確認匯入",
        "imp.ui.note_group": "匯入說明",
        "imp.ui.note_text": "每個檔案各自開啟為一個獨立部件，全部保留在 NX 會話中；.prt 只開啟不另存，STEP/IGES/Parasolid 會另存為源檔案同目錄下的同名 .prt（同名自動加 _2/_3）。",
        "imp.ui.msg_remove_title": "移除檔案",
        "imp.ui.msg_remove_text": "從已選清單中移除這個檔案？",
        "imp.ui.msg_empty_title": "還沒有選擇檔案",
        "imp.ui.msg_empty_one": "請先在左側選擇資料夾，再雙擊檔案（或選中後點「加入檔案」）加入清單。",
        "imp.ui.msg_empty_two": "已選檔案清單為空，無法匯入。",
        "imp.err.dialog": "無法開啟系統檔案選擇框：{0}",
        "imp.confirm.title": "批次匯入",
        "imp.confirm.body": "共 {0} 個檔案待匯入：\n\n{1}",
        "imp.confirm.more": "\n…… 其餘 {0} 個見日誌",
        "imp.confirm.proceed_log": "確認後每個檔案各自開啟為一個獨立部件；.prt 只開啟不另存，STEP/IGES/Parasolid 會另存為源檔案同目錄下的同名 .prt（同名自動加 _2/_3）。",
        "imp.sum.title": "批次匯入完成",
        "imp.sum.body": "成功 {0} / 失敗 {1} / 略過 {2}\n\n{3}\n\n日誌：{4}",
        "imp.sum.fail_item": "· {0}：{1}",
        "imp.sum.no_fail": "（沒有失敗項）",
        "imp.sum.more_fail": "…… 另有 {0} 個失敗項，見日誌",
        "imp.sum.skipped": "另有 {0} 個檔案被略過（重複選擇或檔案不存在），詳見日誌。",
        "imp.sum.disk": "已落盤 {0} / 僅記憶體 {1}",
        "imp.sum.temp_logs": "失敗項的轉換器日誌保留在：{0}",
        "imp.log.header": "開始批次匯入：共 {0} 個檔案；日誌 {1}",
        "imp.log.file": "[{0}/{1}] {2}",
        "imp.log.skip_dup": "  略過（重複選擇）：{0}",
        "imp.log.skip_missing": "  略過（檔案不存在）：{0}",
        "imp.log.prt_open": "  開啟 .prt：{0}",
        "imp.log.prt_already": "  該部件已在會話中開啟，切換為活動顯示",
        "imp.log.prt_done": "  完成 {0:.1f}s -> 部件 {1}",
        "imp.log.new_part": "  新建顯示部件：{0}",
        "imp.log.import_begin": "  匯入 {0} -> 部件 {1}",
        "imp.log.import_done": "  完成 {0:.1f}s",
        "imp.log.saveas_start": "  開始另存：{0}",
        "imp.log.saveas_ok": "  另存成功：{0}",
        "imp.log.saveas_rename": "  同名檔案已存在，改名為：{0}",
        "imp.err.saveas": "  落盤失敗（{0}）：{1}；幾何保留在記憶體，未關閉部件",
        "imp.err.saveas_missing": "  落盤失敗：目標檔案未產生 {0}",
        "imp.log.load_status": "  部件載入狀態：{0}",
        "imp.log.fail": "  失敗：{0}",
        "imp.log.fit_fail": "  擬合視圖失敗（忽略）：{0}",
        "imp.log.destroy_fail": "  銷毀匯入器失敗（忽略）：{0}",
        "imp.log.work_part": "  工作部件：{0}",
        "imp.log.close_empty": "  已關閉空部件：{0}",
        "imp.log.close_fail": "  空部件關閉失敗（保留於會話中）：{0}",
        "imp.log.import_to": "  ImportTo 設定失敗（依預設工作部件繼續）：{0}",
        "imp.log.prep": "  已補齊匯入器參數：{0}；OutputFile={1}",
        "imp.log.prep_fail": "  匯入器參數補齊失敗（依預設繼續）：{0}",
        "imp.log.settings_missing": "  未找到 STEP214 設定檔，依匯入器預設值處理",
        "imp.log.sew_probe": "  {0} 匯入器目前 SewSurfaces 預設值 = {1}",
        "imp.log.sew_probe_fail": "  讀取 {0} 匯入器 SewSurfaces 失敗：{1}",
        "imp.log.sew_set": "  {0} 匯入器 SewSurfaces 已設為 True",
        "imp.log.sew_set_fail": "  設定 {0} 匯入器 SewSurfaces 失敗：{1}",
        "imp.log.sew_unavailable": "  {0} 匯入器沒有 SewSurfaces 屬性，已略過",
        "imp.log.temp_clean": "  已清理本次產生的 {0} 個轉換器日誌：{1}",
        "imp.log.temp_clean_old": "  已清理 {1} 天前的 {0} 個轉換器日誌",
        "imp.log.summary": "彙總：成功 {0} / 失敗 {1} / 略過 {2}",
        "imp.log.log_fallback": "外掛目錄不可寫，日誌改寫到：{0}",
        "imp.log.no_confirm": "無法彈出確認框，按繼續處理",
        "imp.log.open_mode": "  開啟方式：{0}",
        "imp.log.keep_open": "  工作/顯示切換方式：{0}",
        "imp.log.parts_check": "  會話核對：應有 {0} 個部件，目前仍在會話中 {1} 個",
        "imp.log.parts_restored": "  已重新開啟被顯示切換關閉的部件 {0} 個",
        "imp.log.parts_missing": "  以下部件已不在會話中且無法重新開啟：{0}",
        "imp.log.restore_fail": "  重新開啟失敗（{0}）：{1}",
        "imp.reason.open": "無法開啟部件：{0}",
        "imp.reason.import": "匯入失敗：{0}",
        "imp.reason.no_part": "未能建立新部件：{0}",
        "imp.reason.work_part": "無法把新部件設為工作部件：{0}",
        "imp.reason.ext": "不支援的檔案類型",
        "imp.msg.cancelled": "已取消，未匯入任何檔案。",
        "imp.msg.no_files": "沒有可匯入的檔案。",
    },
}


def T(key, *args):
    """取当前语言文案；该语言缺项时回落英文。"""
    table = MSG.get(LANG) or {}
    text = table.get(key)
    if text is None:
        text = MSG["en"].get(key, key)
    if args:
        return text.format(*args)
    return text


# 非 Unicode 控制台（英文 cp437、繁中 cp950 等）打印中文会抛 UnicodeEncodeError。
try:
    import sys
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

# ================= 常量 =================
SUPPORTED_PRT = (".prt",)
SUPPORTED_STEP = (".stp", ".step")
SUPPORTED_IGES = (".igs", ".iges")
SUPPORTED_PARASOLID = (".x_t", ".x_b")
ALL_EXTENSIONS = SUPPORTED_PRT + SUPPORTED_STEP + SUPPORTED_IGES + SUPPORTED_PARASOLID
FILTER_PATTERN = ";".join("*" + ext for ext in ALL_EXTENSIONS)

MAX_NAME_LEN = 80
LAST_DIR_FILE = "import_last_dir.txt"

# 自动缝合开关：默认关，保持与改动前完全一致的导入结果。
# 只在 True 时才写 importer.SewSurfaces；False 时仅打印一次转换器里的当前默认值。
SEW_SURFACES = False
# 会话启动时清理 translator 临时日志的保留天数（只清 .log，不动 .prt 中间产物）
TEMP_LOG_DAYS = 7

_log_handle = None
_trace_handle = None
_sew_probed = set()          # 已探查过 SewSurfaces 默认值的转换器类型（每种只记一次日志）


# ================= 日志 =================

def plugin_dir():
    r"""插件根目录（脚本在 startup\ 下，取它的上一层；被 exec 执行时回落当前目录）"""
    try:
        here = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()
    if os.path.basename(here).lower() == "startup":
        parent = os.path.dirname(here)
        if parent:
            return parent
    return here


def log(msg):
    """同时打印到 NX 信息窗口/控制台和日志文件"""
    text = str(msg)
    try:
        print(text)
    except Exception:
        pass
    if _log_handle:
        try:
            _log_handle.write(text + "\n")
            _log_handle.flush()
        except Exception:
            pass


def trace(step, extra=""):
    """行级打点：绝对时间戳 + 立即落盘（fsync）。

    用于定位“NX 弹工作进行中但无进展”这类卡死：日志停在哪个 step，
    就说明卡在下一步调用上。无 trace 文件时静默跳过。
    """
    if not _trace_handle:
        return
    try:
        now = time.time()
        stamp = "%s.%03d" % (time.strftime("%H:%M:%S", time.localtime(now)),
                             int(now * 1000) % 1000)
        _trace_handle.write("%s | %s | %s\n" % (stamp, step, extra))
        _trace_handle.flush()
        os.fsync(_trace_handle.fileno())
    except Exception:
        pass


def _force_remove(path):
    try:
        os.remove(path)
    except PermissionError:
        try:
            os.chmod(path, 0o666)
            os.remove(path)
        except Exception:
            pass
    except Exception:
        pass


def _rotate_log(path, max_bytes=2 * 1024 * 1024, keep=3):
    """日志超过上限时轮转：当前 -> .1，旧的依次后移，最多保留 keep 份"""
    try:
        if (not os.path.exists(path)) or os.path.getsize(path) < max_bytes:
            return
        oldest = "%s.%d" % (path, keep)
        if os.path.exists(oldest):
            _force_remove(oldest)
        for index in range(keep - 1, 0, -1):
            src = "%s.%d" % (path, index)
            if os.path.exists(src):
                os.replace(src, "%s.%d" % (path, index + 1))
        os.replace(path, path + ".1")
    except Exception:
        pass


def open_log():
    r"""日志放到插件目录 logs\；插件目录不可写时回落 %TEMP%"""
    global _log_handle, _trace_handle
    try:
        log_dir = os.path.join(plugin_dir(), "logs")
        if not os.path.isdir(log_dir):
            os.makedirs(log_dir)
        log_path = os.path.join(log_dir, "batch_import.log")
        _rotate_log(log_path)
        _log_handle = open(log_path, "w", encoding="utf-8-sig")
        try:
            _trace_handle = open(os.path.join(log_dir, "batch_import_trace.log"),
                                 "a", encoding="utf-8")
            _trace_handle.write("\n===== run %s =====\n" % time.strftime("%Y-%m-%d %H:%M:%S"))
            _trace_handle.flush()
        except Exception:
            _trace_handle = None
        return log_path
    except Exception:
        import tempfile
        fallback = os.path.join(tempfile.gettempdir(),
                                "nx_batch_import_%d.log" % os.getpid())
        _log_handle = open(fallback, "w", encoding="utf-8-sig")
        try:
            _trace_handle = open(os.path.join(tempfile.gettempdir(),
                                              "nx_batch_import_trace_%d.log" % os.getpid()),
                                 "a", encoding="utf-8")
        except Exception:
            _trace_handle = None
        return fallback


def _last_dir_path():
    return os.path.join(plugin_dir(), "logs", LAST_DIR_FILE)


def read_last_dir():
    try:
        with open(_last_dir_path(), "r", encoding="utf-8") as handle:
            value = handle.read().strip()
        if value and os.path.isdir(value):
            return value
    except Exception:
        pass
    return ""


def write_last_dir(folder):
    try:
        path = _last_dir_path()
        parent = os.path.dirname(path)
        if parent and not os.path.isdir(parent):
            os.makedirs(parent)
        with open(path, "w", encoding="utf-8") as handle:
            handle.write(folder)
    except Exception:
        pass


# ================= Windows 系统多选对话框 =================
# 不依赖 tkinter：NX 内嵌 Python 未必带。直接调 comdlg32 的 GetOpenFileNameW。

OFN_READONLY = 0x00000001
OFN_OVERWRITEPROMPT = 0x00000002
OFN_HIDEREADONLY = 0x00000004
OFN_NOCHANGEDIR = 0x00000008
OFN_PATHMUSTEXIST = 0x00000800
OFN_FILEMUSTEXIST = 0x00001000
OFN_ALLOWMULTISELECT = 0x00000200
OFN_EXPLORER = 0x00080000

_FILE_BUFFER_CHARS = 32768


class _OPENFILENAMEW(ctypes.Structure):
    _fields_ = [
        ("lStructSize", ctypes.c_uint32),
        ("hwndOwner", ctypes.c_void_p),
        ("hInstance", ctypes.c_void_p),
        ("lpstrFilter", ctypes.c_wchar_p),
        ("lpstrCustomFilter", ctypes.c_wchar_p),
        ("nMaxCustFilter", ctypes.c_uint32),
        ("nFilterIndex", ctypes.c_uint32),
        ("lpstrFile", ctypes.c_wchar_p),
        ("nMaxFile", ctypes.c_uint32),
        ("lpstrFileTitle", ctypes.c_wchar_p),
        ("nMaxFileTitle", ctypes.c_uint32),
        ("lpstrInitialDir", ctypes.c_wchar_p),
        ("lpstrTitle", ctypes.c_wchar_p),
        ("Flags", ctypes.c_uint32),
        ("nFileOffset", ctypes.c_uint16),
        ("nFileExtension", ctypes.c_uint16),
        ("lpstrDefExt", ctypes.c_wchar_p),
        ("lCustData", ctypes.c_void_p),
        ("lpfnHook", ctypes.c_void_p),
        ("lpTemplateName", ctypes.c_wchar_p),
        ("pvReserved", ctypes.c_void_p),
        ("dwReserved", ctypes.c_uint32),
        ("FlagsEx", ctypes.c_uint32),
    ]


def _file_filter():
    return "%s (*.prt;*.stp;*.step;*.igs;*.iges;*.x_t;*.x_b)\0%s\0%s (*.*)\0*.*\0\0" % (
        T("imp.filter"), FILTER_PATTERN, T("imp.filter_all"))


# 导入选择窗：布局、控件、交互与导出侧的 SELECT_UI_PS 保持一致
# （左侧目录树 / 右侧文件列表 / 下方已选清单 / 右下 取消+确认），
# 差别只有：只收文件、清单上方带数量、下方说明改成导入行为说明。
IMPORT_UI_PS = r'''
param(
    [string]$Root = "",
    [Parameter(Mandatory=$true)][string]$OutFile
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "[[imp.ui.title]]"
$form.ClientSize = New-Object System.Drawing.Size(944, 726)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.TopMost = $true
$uiFont = [System.Drawing.SystemFonts]::MessageBoxFont
try { $uiFont = New-Object System.Drawing.Font("Microsoft YaHei UI", 9) } catch { }
$form.Font = $uiFont
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)

# ---- NX 2506 观感基准（实测值与来源见交付参数表）----
# 内容背景 #FFFFFF、分隔线 #DCDCDC、正文 #1E1E1E 取自 NX 2506 界面截图采样；
# 深色主题无本机 NX 截图可量，用 Windows 深色系统色近似；
# 环境变量 NX_PICKER_THEME=light|dark 可强制切换（仅用于渲染核对，功能无关）。
$script:isDark = $false
try {
    $themeKey = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -ErrorAction Stop
    if ($themeKey.AppsUseLightTheme -eq 0) { $script:isDark = $true }
} catch { }
if ($env:NX_PICKER_THEME -eq 'dark') { $script:isDark = $true }
elseif ($env:NX_PICKER_THEME -eq 'light') { $script:isDark = $false }
if ($script:isDark) {
    $uiBg = [System.Drawing.Color]::FromArgb(32, 32, 32)
    $uiField = [System.Drawing.Color]::FromArgb(24, 24, 24)
    $uiFg = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $uiLine = [System.Drawing.Color]::FromArgb(63, 63, 70)
    $uiBtn = [System.Drawing.Color]::FromArgb(51, 51, 51)
} else {
    $uiBg = [System.Drawing.Color]::FromArgb(255, 255, 255)
    $uiField = [System.Drawing.Color]::FromArgb(255, 255, 255)
    $uiFg = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $uiLine = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $uiBtn = [System.Drawing.Color]::FromArgb(240, 240, 240)
}
$form.BackColor = $uiBg
$form.ForeColor = $uiFg

function Add-Separator([int]$x, [int]$y, [int]$w) {
    $line = New-Object System.Windows.Forms.Panel
    $line.Location = New-Object System.Drawing.Point($x, $y)
    $line.Size = New-Object System.Drawing.Size($w, 1)
    $line.BackColor = $uiLine
    $form.Controls.Add($line)
}

function Apply-Theme([System.Windows.Forms.Control]$root) {
    foreach ($c in $root.Controls) {
        if ($c -is [System.Windows.Forms.GroupBox]) {
            $c.ForeColor = $uiFg
            $c.BackColor = $uiBg
        } elseif ($c -is [System.Windows.Forms.Label] -or $c -is [System.Windows.Forms.RadioButton]) {
            $c.ForeColor = $uiFg
            $c.BackColor = [System.Drawing.Color]::Transparent
        } elseif ($c -is [System.Windows.Forms.ComboBox]) {
            $c.BackColor = $uiField
            $c.ForeColor = $uiFg
            # 深色下标准 Combobox 会自绘白底，改 Flat 才吃 BackColor
            if ($script:isDark) { $c.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat }
        } elseif ($c -is [System.Windows.Forms.TextBox]) {
            $c.BackColor = $uiField
            $c.ForeColor = $uiFg
        } elseif ($c -is [System.Windows.Forms.TreeView] -or $c -is [System.Windows.Forms.ListView]) {
            $c.BackColor = $uiField
            $c.ForeColor = $uiFg
            # 列表网格线是系统色，深色下会变成刺眼白线，直接关掉
            if ($script:isDark -and $c -is [System.Windows.Forms.ListView]) { $c.GridLines = $false }
        } elseif ($c -is [System.Windows.Forms.Button]) {
            $c.BackColor = $uiBtn
            $c.ForeColor = $uiFg
            $c.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $c.FlatAppearance.BorderColor = $uiLine
        }
        if ($c.HasChildren) { Apply-Theme $c }
    }
}

$lblTree = New-Object System.Windows.Forms.Label
$lblTree.Text = "[[imp.ui.tree_label]]"
$lblTree.Location = New-Object System.Drawing.Point(12, 8)
$lblTree.Size = New-Object System.Drawing.Size(340, 20)
$form.Controls.Add($lblTree)

$tree = New-Object System.Windows.Forms.TreeView
$tree.Location = New-Object System.Drawing.Point(12, 30)
$tree.Size = New-Object System.Drawing.Size(340, 316)
$tree.HideSelection = $false
$tree.ItemHeight = 22
# NX 树形控件不用点状连线，展开箭头由系统绘制
$tree.ShowLines = $false
$tree.ShowRootLines = $false
$tree.Indent = 16
$tree.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($tree)

$lblList = New-Object System.Windows.Forms.Label
$lblList.Text = "[[imp.ui.list_label]]"
$lblList.Location = New-Object System.Drawing.Point(364, 8)
$lblList.Size = New-Object System.Drawing.Size(568, 20)
$form.Controls.Add($lblList)

$list = New-Object System.Windows.Forms.ListView
$list.Location = New-Object System.Drawing.Point(364, 30)
$list.Size = New-Object System.Drawing.Size(568, 316)
$list.View = "Details"
$list.FullRowSelect = $true
$list.MultiSelect = $true
$list.GridLines = $true
[void]$list.Columns.Add("[[imp.ui.col_name]]", 440)
[void]$list.Columns.Add("[[imp.ui.col_size]]", 100)
$list.ShowItemToolTips = $true
$form.Controls.Add($list)

$lblGrid = New-Object System.Windows.Forms.Label
$lblGrid.Location = New-Object System.Drawing.Point(12, 350)
$lblGrid.Size = New-Object System.Drawing.Size(920, 26)
$form.Controls.Add($lblGrid)

$grid = New-Object System.Windows.Forms.ListView
$grid.Location = New-Object System.Drawing.Point(12, 378)
$grid.Size = New-Object System.Drawing.Size(920, 160)
$grid.View = "Details"
$grid.FullRowSelect = $true
$grid.MultiSelect = $false
$grid.GridLines = $true
[void]$grid.Columns.Add("[[imp.ui.col_type]]", 60)
[void]$grid.Columns.Add("[[imp.ui.col_path]]", 780)
[void]$grid.Columns.Add("✕", 40)
$form.Controls.Add($grid)

# 行高按 NX 实测行距（22px@125% ≈ 17.6 逻辑 + 行间留白）统一为 22：
# ListView 只能借 ImageList 撑高行，图标列宽 1px 不占位
$script:rowHeight = 22
$script:rowImages = New-Object System.Windows.Forms.ImageList
$script:rowImages.ImageSize = New-Object System.Drawing.Size(1, $script:rowHeight)
$script:rowImages.Images.Add((New-Object System.Drawing.Bitmap(1, $script:rowHeight)))
$list.SmallImageList = $script:rowImages
$grid.SmallImageList = $script:rowImages

# 分隔线：位于已选清单与下方说明分组之间
Add-Separator 12 540 920

$grpNote = New-Object System.Windows.Forms.GroupBox
$grpNote.Text = "[[imp.ui.note_group]]"
$grpNote.Location = New-Object System.Drawing.Point(12, 546)
$grpNote.Size = New-Object System.Drawing.Size(920, 108)
$form.Controls.Add($grpNote)

$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Text = "[[imp.ui.note_text]]"
$lblNote.Location = New-Object System.Drawing.Point(14, 24)
$lblNote.Size = New-Object System.Drawing.Size(890, 72)
$lblNote.AutoSize = $false
$grpNote.Controls.Add($lblNote)

function Measure-Column([int]$index) {
    $font = $grid.Font
    $width = [System.Windows.Forms.TextRenderer]::MeasureText($grid.Columns[$index].Text, $font).Width + 28
    foreach ($row in $grid.Items) {
        $text = $row.SubItems[$index].Text
        $w = [System.Windows.Forms.TextRenderer]::MeasureText($text, $font).Width + 28
        if ($w -gt $width) { $width = $w }
    }
    return $width
}

function Fit-GridColumns {
    if ($grid.Columns.Count -lt 3) { return }
    $colType = Measure-Column 0
    $colMark = Measure-Column 2
    $colPath = Measure-Column 1
    $grid.Columns[0].Width = $colType
    $grid.Columns[2].Width = $colMark
    $rest = $grid.ClientSize.Width - $colType - $colMark - 4
    if ($rest -lt 200) { $rest = 200 }
    if ($colPath -lt $rest) { $colPath = $rest }
    $grid.Columns[1].Width = $colPath
}

function Update-Count {
    $lblGrid.Text = "[[imp.ui.grid_label]]".Replace("{0}", [string]$grid.Items.Count)
}

function Get-TypeText([string]$path) {
    # 第一列直接给“用户能认的”类型，和导出清单的「类型」列位置一致
    $ext = [System.IO.Path]::GetExtension($path).TrimStart('.').ToUpper()
    if ($ext -eq "STP" -or $ext -eq "STEP") { return "STEP" }
    if ($ext -eq "IGS" -or $ext -eq "IGES") { return "IGES" }
    if ($ext -eq "X_T" -or $ext -eq "X_B") { return "Parasolid" }
    return $ext
}

function Add-GridRow([string]$path) {
    foreach ($row in $grid.Items) {
        if ($row.SubItems[1].Text -eq $path) { return }
    }
    $item = New-Object System.Windows.Forms.ListViewItem((Get-TypeText $path))
    $item.ImageIndex = 0
    $item.Tag = $path
    $item.ToolTipText = $path
    [void]$item.SubItems.Add($path)
    [void]$item.SubItems.Add("✕")
    $grid.BeginUpdate()
    try { [void]$grid.Items.Add($item) } finally { $grid.EndUpdate() }
    Fit-GridColumns
    Update-Count
}

$script:PH = "__placeholder__"
# 同一目录只枚举一次：重复点击/折叠再展开不再走磁盘
$script:dirCache = @{}

function New-DirNode([System.Windows.Forms.TreeNode]$parent, [string]$path) {
    # 只加一层：子目录等到展开时才枚举，避免一上来递归整块硬盘
    $node = New-Object System.Windows.Forms.TreeNode([System.IO.Path]::GetFileName($path.TrimEnd('\')))
    if (-not $node.Text) { $node.Text = $path }
    $node.Tag = $path
    [void]$parent.Nodes.Add($node)
    $hasChild = $false
    try {
        foreach ($d in [System.IO.Directory]::GetDirectories($path)) { $hasChild = $true; break }
    } catch { }
    if ($hasChild) {
        $ph = New-Object System.Windows.Forms.TreeNode("...")
        $ph.Tag = $script:PH
        [void]$node.Nodes.Add($ph)
    }
    return $node
}

# 从「此电脑 → 盘符」沿路径逐级建节点并展开，返回目标目录节点（用于「上次位置」记忆）
function Expand-ToPath([string]$path) {
    if (-not $path) { return $null }
    $full = $path.TrimEnd('\')
    if (-not $full) { return $null }
    try {
        if (-not [System.IO.Directory]::Exists($full)) { return $null }
    } catch { return $null }
    $rootPart = [System.IO.Path]::GetPathRoot($full)
    if (-not $rootPart) { return $null }
    $node = $null
    # 盘符节点挂在「此电脑」下面，先查根节点本身，再查它的一级子节点
    foreach ($n in $tree.Nodes) {
        if (([string]$n.Tag).TrimEnd('\') -eq $rootPart.TrimEnd('\')) { $node = $n; break }
        foreach ($c in $n.Nodes) {
            if (([string]$c.Tag).TrimEnd('\') -eq $rootPart.TrimEnd('\')) { $node = $c; break }
        }
        if ($node) { break }
    }
    if (-not $node) { return $null }
    if ($node.Parent) { $node.Parent.Expand() }
    [void]$node.Expand()
    $rest = $full.Substring($rootPart.Length).Trim('\')
    if ($rest) {
        $walk = $rootPart
        foreach ($seg in $rest.Split('\')) {
            if (-not $seg) { continue }
            $walk = [System.IO.Path]::Combine($walk, $seg)
            $child = $null
            foreach ($c in $node.Nodes) {
                if (([string]$c.Tag).TrimEnd('\') -eq $walk.TrimEnd('\')) { $child = $c; break }
            }
            if (-not $child) { return $null }
            $node = $child
            [void]$node.Expand()
        }
    }
    return $node
}
# 顶层只有「此电脑」+ 所有盘符（含可移动盘）；上次位置靠展开定位，不再单独挂节点
$pcNode = New-Object System.Windows.Forms.TreeNode("[[imp.ui.this_pc]]")
$pcNode.Tag = ""
try {
    foreach ($drv in [System.IO.DriveInfo]::GetDrives()) {
        try {
            if (-not $drv.IsReady) { continue }
            $label = $drv.Name
            if ($drv.VolumeLabel) { $label = "$($drv.Name)  $($drv.VolumeLabel)" }
            $node = New-DirNode $pcNode $drv.Name
            $node.Text = $label
        } catch { continue }
    }
} catch { }
[void]$tree.Nodes.Add($pcNode)
$pcNode.Expand()

$tree.add_BeforeExpand({
    param($sender, $e)
    $node = $e.Node
    if ($node.Nodes.Count -eq 1 -and $node.Nodes[0].Tag -eq $script:PH) {
        $node.Nodes.Clear()
        $path = $node.Tag
        if ($path -and [System.IO.Directory]::Exists($path)) {
            try {
                $dirKey = $path + "|dirs"
                if ($script:dirCache.ContainsKey($dirKey)) {
                    $dirs = $script:dirCache[$dirKey]
                } else {
                    $dirs = [System.IO.Directory]::GetDirectories($path)
                    $script:dirCache[$dirKey] = $dirs
                }
                foreach ($d in $dirs) { [void](New-DirNode $node $d) }
            } catch {
                $lblList.Text = "[[imp.ui.denied]]"
            }
        }
    }
})

$tree.add_AfterSelect({
    param($sender, $e)
    $path = $e.Node.Tag
    $list.Items.Clear()
    $lblList.Text = "[[imp.ui.list_label]]"
    if (-not $path) { return }
    if (-not [System.IO.Directory]::Exists($path)) {
        $lblList.Text = "[[imp.ui.denied]]"
        try {
            $errLog = Join-Path $env:TEMP "nx_batch_import_ui_errors.log"
            Add-Content -LiteralPath $errLog -Encoding UTF8 -Value ("[{0}] picker: cannot access '{1}' (Directory.Exists=False; missing or permission denied)" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $path)
        } catch { }
        $hint = "[[imp.ui.denied_hint]]`n`n$path"
        $job = { [void][System.Windows.Forms.MessageBox]::Show($form, $hint, "[[imp.ui.title]]", "OK", "Warning") }.GetNewClosure()
        [void]$form.BeginInvoke([System.Windows.Forms.MethodInvoker]$job)
        return
    }
    try {
        $dirKey = $path + "|files"
        if ($script:dirCache.ContainsKey($dirKey)) {
            $files = $script:dirCache[$dirKey]
        } else {
            $files = [System.IO.Directory]::GetFiles($path, "*.*") | Where-Object {
                $l = $_.ToLower()
                $l.EndsWith(".prt") -or $l.EndsWith(".stp") -or $l.EndsWith(".step") -or
                $l.EndsWith(".igs") -or $l.EndsWith(".iges") -or $l.EndsWith(".x_t") -or
                $l.EndsWith(".x_b")
            }
            $script:dirCache[$dirKey] = $files
        }
        $arr = New-Object System.Collections.ArrayList
        foreach ($f in $files) {
            $fi = New-Object System.IO.FileInfo($f)
            $item = New-Object System.Windows.Forms.ListViewItem($fi.Name)
            $item.ImageIndex = 0
            [void]$item.SubItems.Add([string]([math]::Round($fi.Length / 1KB, 1)))
            $item.Tag = $f
            $item.ToolTipText = $f
            [void]$arr.Add($item)
        }
        if ($arr.Count -gt 0) {
            $list.BeginUpdate()
            $list.Items.AddRange([System.Windows.Forms.ListViewItem[]]$arr.ToArray())
            $list.EndUpdate()
        }
        if ($arr.Count -eq 0) { $lblList.Text = "[[imp.ui.empty_dir]]" }
    } catch {
        $lblList.Text = "[[imp.ui.denied]]"
        try {
            $errLog = Join-Path $env:TEMP "nx_batch_import_ui_errors.log"
            Add-Content -LiteralPath $errLog -Encoding UTF8 -Value ("[{0}] picker: failed to list files in '{1}'`r`n{2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $path, $_.Exception.ToString())
        } catch { }
        $hint = "[[imp.ui.denied_hint]]`n`n$path"
        $job = { [void][System.Windows.Forms.MessageBox]::Show($form, $hint, "[[imp.ui.title]]", "OK", "Warning") }.GetNewClosure()
        [void]$form.BeginInvoke([System.Windows.Forms.MethodInvoker]$job)
    }
})

$list.add_DoubleClick({
    param($sender, $e)
    if ($list.SelectedItems.Count -eq 0) { return }
    foreach ($sel in $list.SelectedItems) {
        Add-GridRow ([string]$sel.Tag)
    }
})

$grid.add_MouseClick({
    param($sender, $e)
    $hit = $grid.HitTest($e.X, $e.Y)
    if (-not $hit.Item) { return }
    if (-not $hit.SubItem) { return }
    if ($hit.Item.SubItems.IndexOf($hit.SubItem) -ne 2) { return }
    $path = $hit.Item.SubItems[1].Text
    $ans = [System.Windows.Forms.MessageBox]::Show("[[imp.ui.msg_remove_text]]`n`n$path", "[[imp.ui.msg_remove_title]]", "YesNo", "Question")
    if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
        $grid.Items.RemoveAt($hit.Item.Index)
        Fit-GridColumns
        Update-Count
    }
})

$btnAddFile = New-Object System.Windows.Forms.Button
$btnAddFile.Text = "[[imp.ui.btn_add_file]]"
$btnAddFile.Location = New-Object System.Drawing.Point(12, 669)
$btnAddFile.Size = New-Object System.Drawing.Size(110, 27)
$form.Controls.Add($btnAddFile)
$btnAddFile.add_Click({
    param($sender, $e)
    if ($list.SelectedItems.Count -eq 0) { return }
    foreach ($sel in $list.SelectedItems) {
        Add-GridRow ([string]$sel.Tag)
    }
})

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "[[imp.ui.btn_clear]]"
$btnClear.Location = New-Object System.Drawing.Point(132, 669)
$btnClear.Size = New-Object System.Drawing.Size(70, 27)
$form.Controls.Add($btnClear)
$btnClear.add_Click({
    param($sender, $e)
    $grid.Items.Clear()
    Fit-GridColumns
    Update-Count
})

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "[[imp.ui.btn_cancel]]"
$btnCancel.Location = New-Object System.Drawing.Point(742, 669)
$btnCancel.Size = New-Object System.Drawing.Size(80, 27)
$form.Controls.Add($btnCancel)
$btnCancel.add_Click({
    param($sender, $e)
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Close()
})

$btnOK = New-Object System.Windows.Forms.Button
$btnOK.Font = New-Object System.Drawing.Font($uiFont, [System.Drawing.FontStyle]::Bold)
$btnOK.Text = "[[imp.ui.btn_ok]]"
$btnOK.Location = New-Object System.Drawing.Point(832, 669)
$btnOK.Size = New-Object System.Drawing.Size(100, 27)
$form.Controls.Add($btnOK)
$btnOK.add_Click({
    param($sender, $e)
    if ($grid.Items.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("[[imp.ui.msg_empty_one]]`n[[imp.ui.msg_empty_two]]", "[[imp.ui.msg_empty_title]]", "OK", "Warning")
        return
    }
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})

# 先建树句柄再定位：展开/选中会触发 BeforeExpand 建子节点、AfterSelect 填充右侧列表，
# 保证窗口显示时列表已是上次目录的内容（而不是空列表）
[void]$tree.Handle

# 上次位置：沿「此电脑 → 盘符 → 逐级子目录」自动展开并选中，右侧列表随即显示该目录
$targetNode = Expand-ToPath $Root
if ($targetNode) {
    try {
        $tree.SelectedNode = $targetNode
        $targetNode.EnsureVisible()
    } catch { }
} else {
    try { $tree.SelectedNode = $pcNode } catch { }
}

Fit-GridColumns
Update-Count
Apply-Theme $form
[void]$form.ShowDialog()

$files = @()
$ok = ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK)
if ($ok) {
    foreach ($row in $grid.Items) { $files += $row.SubItems[1].Text }
}

$payload = [ordered]@{ ok = $ok; files = $files }
$json = $payload | ConvertTo-Json -Depth 3 -Compress
[System.IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
$form.Dispose()
'''


def _powershell_path():
    """返回 powershell.exe 的可执行路径；找不到时返回 None"""
    candidates = [
        os.path.join(os.environ.get("SystemRoot", r"C:\Windows"),
                     "System32", "WindowsPowerShell", "v1.0", "powershell.exe"),
        "powershell",
    ]
    for candidate in candidates:
        if os.path.exists(candidate) or candidate == "powershell":
            return candidate
    return None


def render_ps(template):
    """把 PowerShell 模板里的 [[key]] 占位符替换成当前语言文案（缺项回落英文）"""
    merged = dict(MSG["en"])
    merged.update(MSG.get(LANG) or {})
    for key, value in merged.items():
        template = template.replace("[[%s]]" % key, str(value))
    return template


def _ui_error_log(text):
    """选择窗内部异常写独立小日志，避免污染导入主日志"""
    import tempfile
    try:
        with open(os.path.join(tempfile.gettempdir(),
                               "nx_batch_import_ui_errors.log"), "a",
                  encoding="utf-8") as handle:
            handle.write("%s %s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"), text))
    except Exception:
        pass


def pick_files_ui(initial_dir):
    """弹出与导出侧同款的导入选择窗。

    返回 (文件列表, 是否走了新窗口)。窗口本身选完即确认，不再二次弹框。
    任何异常都抛给调用方，由调用方回落系统多选框。
    """
    import json
    import subprocess
    import tempfile

    if os.name != "nt":
        raise RuntimeError("import picker needs Windows")

    temp_dir = tempfile.gettempdir()
    ps_path = os.path.join(temp_dir, "nx_batch_import_picker.ps1")
    out_path = os.path.join(temp_dir, "nx_batch_import_picker.json")
    _rotate_log(os.path.join(temp_dir, "nx_batch_import_ui_errors.log"))

    with open(ps_path, "w", encoding="utf-8-sig") as handle:
        handle.write(render_ps(IMPORT_UI_PS))
    if os.path.exists(out_path):
        os.remove(out_path)

    powershell = _powershell_path()
    if powershell is None:
        raise RuntimeError("powershell.exe not found")

    arguments = ["-File", ps_path, "-OutFile", out_path]
    if initial_dir:
        arguments += ["-Root", initial_dir]
    command = [powershell, "-NoProfile", "-STA", "-ExecutionPolicy",
               "Bypass"] + arguments

    if hasattr(subprocess, "run"):
        process = subprocess.run(command, stdout=subprocess.PIPE,
                                 stderr=subprocess.PIPE)
        stderr = process.stderr
    else:
        process = subprocess.Popen(command, stdout=subprocess.PIPE,
                                   stderr=subprocess.PIPE)
        _, stderr = process.communicate()

    if not os.path.exists(out_path):
        detail = ""
        try:
            detail = (stderr or b"").decode("utf-8", "ignore")[:300]
        except Exception:
            pass
        raise RuntimeError("picker produced no result: %s" % detail)

    with open(out_path, "r", encoding="utf-8-sig") as handle:
        payload = json.load(handle)
    if not payload.get("ok"):
        return [], True
    raw = payload.get("files") or []
    if isinstance(raw, str):
        # 只有一个文件时 ConvertTo-Json 会退化成字符串
        raw = [raw]
    files = []
    for item in raw:
        text = str(item or "").strip()
        if text:
            files.append(text)
    trace("PICK_OK", "%d file(s)" % len(files))
    return files, True



def select_files(initial_dir):
    """弹系统多选框；取消返回 []，异常抛给调用方。"""
    if os.name != "nt":
        raise RuntimeError("只支持 Windows 系统文件对话框")

    buffer = ctypes.create_unicode_buffer(_FILE_BUFFER_CHARS)
    ofn = _OPENFILENAMEW()
    ofn.lStructSize = ctypes.sizeof(_OPENFILENAMEW)
    ofn.hwndOwner = None
    ofn.lpstrFilter = _file_filter()
    ofn.nFilterIndex = 1
    ofn.lpstrFile = ctypes.cast(buffer, ctypes.c_wchar_p)
    ofn.nMaxFile = _FILE_BUFFER_CHARS
    ofn.lpstrTitle = T("imp.title")
    ofn.lpstrInitialDir = initial_dir or None
    ofn.Flags = (OFN_ALLOWMULTISELECT | OFN_EXPLORER | OFN_FILEMUSTEXIST |
                 OFN_PATHMUSTEXIST | OFN_HIDEREADONLY | OFN_NOCHANGEDIR)

    try:
        comdlg32 = ctypes.WinDLL("comdlg32")
    except Exception:
        comdlg32 = ctypes.windll.comdlg32
    get_open = comdlg32.GetOpenFileNameW
    get_open.argtypes = [ctypes.POINTER(_OPENFILENAMEW)]
    get_open.restype = ctypes.c_int

    if not get_open(ctypes.byref(ofn)):
        # 0 = 用户取消；非 0 错误码说明对话框本身有问题
        try:
            code = comdlg32.CommDlgExtendedError()
        except Exception:
            code = 0
        if code:
            raise RuntimeError("CommDlgExtendedError=0x%08X" % code)
        return []

    try:
        raw = ctypes.string_at(ctypes.addressof(buffer),
                               _FILE_BUFFER_CHARS * 2).decode("utf-16-le")
    except Exception:
        raw = ctypes.wstring_at(ctypes.addressof(buffer), _FILE_BUFFER_CHARS)
    parts = [item for item in raw.split("\x00") if item]
    if not parts:
        return []
    if len(parts) == 1:
        return [parts[0]]
    folder = parts[0]
    return [os.path.join(folder, item) for item in parts[1:]]


# ================= NX 交互 =================

def message_box(kind, title, text):
    """NXMessageBox；kind = 'Information' / 'Question'。UI 不可用时返回 None。"""
    try:
        ui = NXOpen.UI.GetUI()
        if ui is None:
            return None
        kinds = getattr(NXOpen.NXMessageBox, "DialogType", None)
        value = getattr(kinds, kind, None) if kinds is not None else None
        if value is None:
            return None
        return ui.NXMessageBox.Show(title, value, text)
    except Exception:
        return None


def confirm_yes(title, text):
    response = message_box("Question", title, text)
    if response is None:
        log(T("imp.log.no_confirm"))
        return True
    try:
        return response == NXOpen.NXMessageBox.DialogResponse.Yes
    except Exception:
        return True


def get_listing(session):
    try:
        listing = session.ListingWindow
        if not listing.IsOpen:
            listing.Open()
        return listing
    except Exception:
        return None


def listing_write(listing, text):
    if listing is None:
        return
    try:
        listing.WriteLine(text)
    except Exception:
        pass


def find_open_part(session, path):
    """已打开的部件按 FullPath 匹配（NXOpen 没有 FindPartByPathName 这类接口）"""
    target = os.path.normcase(os.path.abspath(path))
    try:
        parts = list(session.Parts)
    except Exception:
        return None
    for part in parts:
        try:
            current = part.FullPath
        except Exception:
            continue
        if current and os.path.normcase(os.path.abspath(current)) == target:
            return part
    return None


def _split_open_result(result):
    """OpenDisplay / NewDisplay 可能返回 Part，也可能返回 (Part, PartLoadStatus)"""
    if isinstance(result, (tuple, list)):
        if len(result) >= 2:
            return result[0], result[1]
        if len(result) == 1:
            return result[0], None
    return result, None


def _describe_load_status(status):
    if status is None:
        return ""
    fields = []
    for attr in ("NumberUnloadedParts", "NumberParts", "NumberErrors"):
        try:
            value = getattr(status, attr)
        except Exception:
            continue
        if isinstance(value, int):
            fields.append("%s=%d" % (attr, value))
    detail = []
    try:
        getter = getattr(status, "GetPartLoadStatus", None)
        count = 0
        try:
            count = int(getattr(status, "NumberUnloadedParts", 0) or 0)
        except Exception:
            count = 0
        if getter is not None:
            for index in range(min(count, 5)):
                try:
                    info = getter(index)
                except Exception:
                    break
                text = []
                for name in ("PartName", "FileName", "Description", "Status"):
                    try:
                        value = getattr(info, name)
                    except Exception:
                        continue
                    if value:
                        text.append(str(value))
                if text:
                    detail.append(" / ".join(text))
    except Exception:
        pass
    if detail:
        return "; ".join(fields + detail)
    if fields:
        return ", ".join(fields)
    try:
        return str(status)[:200]
    except Exception:
        return "PartLoadStatus(unprintable)"


def display_option_keep_others():
    """取“在已有显示部件之外再显示一个”的枚举值。

    DisplayPartOption.ReplaceExisting 会把上一个显示部件替换掉：该部件
    未被修改时 NX 直接把它从会话里关掉，于是批量导入 N 个文件最后只剩
    最后一个（现场截图就是这个现象）。AllowAdditional 是追加显示。
    """
    options = getattr(NXOpen, "DisplayPartOption", None)
    return getattr(options, "AllowAdditional", None) if options is not None else None


def ensure_multi_display(session):
    """打开会话的“允许多个显示部件”选项。

    该选项被关掉时，任何切换显示部件的操作都会替换当前窗口内容，
    即使传 AllowAdditional 也只会剩一个部件。返回 (状态, 动作) 供日志。
    """
    parts = session.Parts
    try:
        status = parts.AllowMultipleDisplayedParts
    except Exception:
        return "unknown", "none"
    enum_cls = getattr(NXOpen.PartCollection, "MultipleDisplayedPartStatus", None)
    enabled = getattr(enum_cls, "Enabled", None)
    if enabled is not None:
        try:
            if status == enabled:
                return "Enabled", "none"
        except Exception:
            pass
    value = getattr(status, "value", None)
    setter = getattr(parts, "SetAllowMultipleDisplayedParts", None)
    if setter is not None:
        try:
            setter(True)
            return "value=%r enabled" % value, "SetAllowMultipleDisplayedParts(True)"
        except Exception as exc:
            return "value=%r failed %s" % (value, exc), "none"
    return "value=%r" % value, "none"


def activate_display(session, part):
    """把 part 设为显示部件，但不挤掉其它已加载部件。"""
    keep = display_option_keep_others()
    if keep is not None:
        session.Parts.SetActiveDisplay(
            part, keep, NXOpen.PartDisplayPartWorkPartOption.SameAsDisplay)
        return "SetActiveDisplay(AllowAdditional)"
    session.Parts.SetActiveDisplay(
        part,
        NXOpen.DisplayPartOption.ReplaceExisting,
        NXOpen.PartDisplayPartWorkPartOption.SameAsDisplay)
    return "SetActiveDisplay(ReplaceExisting)"


def set_work_part(session, part):
    """NewDisplay 只保证显示部件；这里显式把新部件设为工作部件。
    优先走 SetActiveDisplay(AllowAdditional)：新部件同时成为显示/工作部件，
    又不会关掉前面已经导入的部件。"""
    errors = []
    keep = display_option_keep_others()
    setter = getattr(session.Parts, "SetActiveDisplay", None)
    if setter is not None and keep is not None:
        try:
            setter(part, keep,
                   NXOpen.PartDisplayPartWorkPartOption.SameAsDisplay)
            return "SetActiveDisplay(AllowAdditional)"
        except Exception as exc:
            errors.append("SetActiveDisplay(AllowAdditional): %s" % exc)
    setter = getattr(session.Parts, "SetDisplay", None)
    if setter is not None:
        try:
            setter(part, False, True)
            return "SetDisplay"
        except Exception as exc:
            errors.append("SetDisplay: %s" % exc)
    setter = getattr(session.Parts, "SetWork", None)
    if setter is not None:
        try:
            setter(part)
            return "SetWork"
        except Exception as exc:
            errors.append("SetWork: %s" % exc)
    try:
        return activate_display(session, part)
    except Exception as exc:
        errors.append("SetActiveDisplay: %s" % exc)
    raise RuntimeError("; ".join(errors) or "no work-part API")


def close_part(session, part):
    """尽力关闭导入失败留下的新部件；返回是否成功。

    枚举是扁平的 NXOpen.BasePartCloseWholeTree / BasePartCloseModified。
    CloseModified 必须用 CloseModified（空部件是“已修改”状态，用
    DontCloseModified 会报 “Modified part not saved” 而关不掉）。
    """
    closer = getattr(part, "Close", None)
    whole = getattr(getattr(NXOpen, "BasePartCloseWholeTree", None),
                    "TrueValue", None)
    modified = getattr(getattr(NXOpen, "BasePartCloseModified", None),
                       "CloseModified", None)
    if closer is None or whole is None or modified is None:
        return False
    responses = None
    try:
        responses = session.Parts.NewPartCloseResponses()
    except Exception:
        responses = None
    try:
        closer(whole, modified, responses)
        return True
    except Exception:
        return False


def new_display_part(session, name):
    maker = getattr(session.Parts, "NewDisplay", None)
    if maker is None:
        raise RuntimeError("PartCollection.NewDisplay 不存在")
    units = getattr(NXOpen, "PartUnits", None)
    value = getattr(units, "Millimeters", None) if units is not None else None
    result = maker(name) if value is None else maker(name, value)
    part, status = _split_open_result(result)
    if part is None:
        raise RuntimeError("NewDisplay 没有返回部件")
    return part, status


def open_part_file(session, path):
    """打开 .prt，返回 (部件, loadStatus, 使用的方式)。

    OpenDisplay / OpenBaseDisplay 在 NX 里等价于 ReplaceExisting：上一个
    显示部件没被修改时会被直接关掉，批量打开 N 个 .prt 最后只剩最后一个。
    优先用 OpenActiveDisplay(AllowAdditional) 追加打开，全部留在会话里。
    """
    keep = display_option_keep_others()
    opener = getattr(session.Parts, "OpenActiveDisplay", None)
    if opener is not None and keep is not None:
        try:
            result = opener(path, keep)
        except Exception as exc:
            raise RuntimeError(T("imp.reason.open", exc))
        part, status = _split_open_result(result)
        if part is None:
            raise RuntimeError(T("imp.reason.open", "OpenActiveDisplay 没有返回部件"))
        return part, status, "OpenActiveDisplay(AllowAdditional)"
    opener = getattr(session.Parts, "OpenDisplay", None)
    if opener is None:
        opener = getattr(session.Parts, "OpenBaseDisplay", None)
    if opener is None:
        raise RuntimeError("PartCollection.OpenDisplay 不存在")
    try:
        result = opener(path)
    except Exception as exc:
        raise RuntimeError(T("imp.reason.open", exc))
    part, status = _split_open_result(result)
    if part is None:
        raise RuntimeError(T("imp.reason.open", "OpenDisplay 没有返回部件"))
    return part, status, "OpenDisplay(ReplaceExisting)"


def part_file_on_disk(part):
    """部件对应的磁盘文件；内存态部件返回空串"""
    try:
        full = part.FullPath or ""
    except Exception:
        return ""
    if full and os.path.isfile(full):
        return full
    return ""


def part_still_loaded(session, part, path):
    """部件是否仍在会话中：有磁盘路径按路径比对，否则按对象身份比对"""
    if path:
        return find_open_part(session, path) is not None
    try:
        for item in list(session.Parts):
            if item is part:
                return True
    except Exception:
        return True
    return False


def displayed_part_names(session):
    """当前显示集合里的部件名；老版本没有 MDP 时返回空列表"""
    try:
        shown = list(session.Parts.GetDisplayedParts())
    except Exception:
        return []
    names = []
    for item in shown:
        try:
            names.append(item.Leaf)
        except Exception:
            names.append(str(item))
    return names


def part_still_displayed(session, part, path):
    """部件是否还在显示集合里。

    STEP/IGS 分支用 Parts.NewDisplay 建新部件，NX 会把之前的显示集合
    换掉，部件仍留在会话里但不再显示。这里按对象身份或磁盘路径比对。
    """
    try:
        shown = list(session.Parts.GetDisplayedParts())
    except Exception:
        return True
    for item in shown:
        if item is part:
            return True
        try:
            if path and (item.FullPath or "") == path:
                return True
        except Exception:
            pass
    return False


def set_part_displayed(session, part):
    """把仍在会话里的部件重新加入显示集合（AllowAdditional，不挤掉别的部件）"""
    try:
        action = activate_display(session, part)
        trace("SETDISPLAY_OK", "%s %s" % (getattr(part, "Leaf", part), action))
        return True
    except Exception as exc:
        trace("SETDISPLAY_FAIL", "%s: %s" % (getattr(part, "Leaf", part), exc))
        return False


def reopen_missing_parts(session, records):
    """兜底：把被显示部件切换挤掉的部件补回会话，并补上显示。

    records = [(名称, 磁盘路径或空, 部件对象)]。
    返回 (补回数量, 仍缺失的名称列表)。补回数量含"只在会话里、没在显示集合里"
    被重新追加显示的部件。
    """
    restored = 0
    missing = []
    for name, path, part in records:
        if not part_still_loaded(session, part, path):
            if not path or not os.path.isfile(path):
                missing.append(name)
                continue
            try:
                open_part_file(session, path)
                restored += 1
                trace("REOPEN_OK", "%s <- %s" % (name, path))
            except Exception as exc:
                missing.append(name)
                trace("REOPEN_FAIL", "%s: %s" % (name, exc))
            continue
        if part_still_displayed(session, part, path):
            continue
        # 部件还在会话里，只是不显示：先试 SetDisplay，失败再退回重新打开
        if set_part_displayed(session, part):
            restored += 1
            trace("REDISPLAY_OK", "%s (SetDisplay)" % name)
            continue
        if not path or not os.path.isfile(path):
            continue
        try:
            open_part_file(session, path)
            restored += 1
            trace("REDISPLAY_OK", "%s <- %s" % (name, path))
        except Exception as exc:
            trace("REDISPLAY_FAIL", "%s: %s" % (name, exc))
    return restored, missing


def _dex_manager(session):
    """DexManager：优先会话级，其次当前工作部件"""
    manager = getattr(session, "DexManager", None)
    if manager is not None:
        return manager
    try:
        work = session.Parts.Work
    except Exception:
        work = None
    if work is not None:
        manager = getattr(work, "DexManager", None)
        if manager is not None:
            return manager
    raise RuntimeError("DexManager 不存在")


def create_importer(session, method_names):
    manager = _dex_manager(session)
    errors = []
    for method_name in method_names:
        factory = getattr(manager, method_name, None)
        if factory is None:
            errors.append("%s: missing" % method_name)
            continue
        try:
            return factory()
        except Exception as exc:
            errors.append("%s: %s" % (method_name, exc))
    raise RuntimeError("; ".join(errors))


def set_input_file(importer, path):
    for attr in ("InputFile", "FileName", "InputFileName"):
        if hasattr(importer, attr):
            try:
                setattr(importer, attr, path)
                return attr
            except Exception:
                continue
    setter = getattr(importer, "SetInputFile", None)
    if setter is not None:
        setter(path)
        return "SetInputFile"
    raise RuntimeError("导入器没有可用的输入文件属性")


def _import_to_enum(importer, member):
    """按嵌套枚举 / 扁平类名两种形式找 ImportTo 枚举成员。

    不同导入器的枚举类型不同（Step214Importer.ImportToOption、
    IgesImporter.ImportToEnum），先按导入器自身类名找；
    不能把 Step214 的枚举塞给 IGES，否则报 “Expecting ... type, found ...”。
    """
    cls = type(importer)
    owners = []
    for attr in ("ImportToOption", "ImportToEnum"):
        nested = getattr(cls, attr, None)
        if nested is not None:
            owners.append(nested)
    name = getattr(cls, "__name__", "")
    for suffix in ("ImportToEnum", "ImportToOption", "ImportTo"):
        flat = getattr(NXOpen, name + suffix, None)
        if flat is not None:
            owners.append(flat)
    if not owners:
        for attr in ("Step214ImporterImportToOption", "StepImporterImportToOption",
                     "IgesImporterImportToEnum", "IgesImporterImportToOption"):
            value = getattr(NXOpen, attr, None)
            if value is not None:
                owners.append(value)
    for owner in owners:
        value = getattr(owner, member, None)
        if value is not None:
            return value
    return None


def apply_import_to(importer):
    """三器统一口径：导入目标 = 工作部件。
    STEP / IGES 显式设 ImportTo=WorkPart；Parasolid 无 ImportTo（本就只导工作部件）。
    失败时抛错，由调用方记日志并继续。"""
    if not hasattr(importer, "ImportTo"):
        return "default(work part)"
    value = _import_to_enum(importer, "WorkPart")
    if value is None:
        raise RuntimeError("ImportTo 枚举里没有 WorkPart")
    importer.ImportTo = value
    return "ImportTo=WorkPart"


def translator_settings_file(label):
    r"""translator 默认设置文件：STEP 用 STEP214UG\step214ug.def，
    IGES 用 IGES\igesimport.def（与原生“文件→导入→STEP214 / IGES”同一份 def）。"""
    if label == "STEP":
        envs = (("STEP214UG_DIR", "step214ug.def"), ("UGII_STEP214_DIR", "step214ug.def"))
        relative = os.path.join("STEP214UG", "step214ug.def")
    elif label == "IGES":
        envs = (("IGESUG_DIR", "igesimport.def"), ("UGII_IGES_DIR", "igesimport.def"))
        relative = os.path.join("IGES", "igesimport.def")
    else:
        return ""
    for var, filename in envs:
        value = (os.environ.get(var) or "").strip()
        if value:
            path = os.path.join(value, filename)
            if os.path.isfile(path):
                return path
    base = (os.environ.get("UGII_BASE_DIR") or "").strip()
    if base:
        path = os.path.join(base, relative)
        if os.path.isfile(path):
            return path
    return ""


def translator_output_file(name):
    """translator 的输出 prt 路径：必须可写，且目标不能已存在。

    用 %TEMP%\\nx_batch_import 而不是 part.FullPath —— GUI 里 NX 的 cwd 可能是
    只读安装目录，写不了就没法翻译；文件名带 pid + 毫秒，避免目标已存在时
    translator 改写 _1.prt 而导入效果不确定。
    该 prt 只是 translator 的中间产物，NX 加载进工作部件后不留在磁盘上。
    仅 STEP / IGES 需要（两者都是“先翻成 prt 再加载”）。
    """
    folder = os.path.join(tempfile.gettempdir(), "nx_batch_import")
    os.makedirs(folder, exist_ok=True)
    stem = _ILLEGAL_CHARS.sub("_", name or "imported")[:MAX_NAME_LEN]
    return os.path.join(folder, "%s_%d_%d.prt" % (stem, os.getpid(),
                                                  int(time.time() * 1000)))


def import_temp_dir():
    """translator 中间产物目录：临时 prt 与 .log 都写在这里"""
    return os.path.join(tempfile.gettempdir(), "nx_batch_import")


def _logs_since(folder, since):
    """列出目录里修改时间 >= since 的 .log（即本次导入产生的转换器日志）

    用严格 >=（NTFS 的 mtime 精度足够），避免把同批次此前失败项保留的日志一起删掉。
    """
    result = []
    try:
        for name in os.listdir(folder):
            if not name.lower().endswith(".log"):
                continue
            path = os.path.join(folder, name)
            try:
                if os.path.getmtime(path) >= since:
                    result.append(path)
            except Exception:
                continue
    except Exception:
        pass
    return result


def cleanup_run_logs(since):
    """导入成功后删除本次产生的 translator .log；任何异常都不影响导入"""
    count = 0
    for path in _logs_since(import_temp_dir(), since):
        try:
            os.remove(path)
            count += 1
        except Exception:
            continue
    if count:
        log(T("imp.log.temp_clean", count, import_temp_dir()))
    return count


def cleanup_old_logs(days=TEMP_LOG_DAYS):
    """会话启动时清理 N 天前的 translator .log；失败静默，不影响导入"""
    cutoff = time.time() - days * 86400
    count = 0
    try:
        for name in os.listdir(import_temp_dir()):
            if not name.lower().endswith(".log"):
                continue
            path = os.path.join(import_temp_dir(), name)
            try:
                if os.path.getmtime(path) < cutoff:
                    os.remove(path)
                    count += 1
            except Exception:
                continue
    except Exception:
        pass
    return count


def prepare_importer(importer, label, part, name="", sew_surfaces=SEW_SURFACES):
    """补齐 Dex 导入器必填项：SettingsFile / OutputFile / FileOpenFlag。

    STEP / IGES 不是“直接读进内存”，而是 translator 先写 prt、NX 再加载；
    DexManager 的 Commit 走同一条链，所以必须告诉它输出到哪个 prt：
    OutputFile 必须指向可写目录且文件尚不存在（见 translator_output_file），
    否则 translator 报 “Unable to import selected file to work part”，
    部件既没有几何也没有转换器日志——表现为“卡住/失败但没有产物”。
    FileOpenFlag=False 表示这是导入而不是打开文件（Siemens 官方样例口径）。
    Parasolid 不经过 translator，也没有这些属性，直接导工作部件。
    属性按 hasattr 逐个补，缺的跳过（不同 NX 版本类定义有差异）。
    """
    if label == "Parasolid":
        return "n/a(parasolid)"
    applied = []
    settings = translator_settings_file(label)
    if settings:
        if hasattr(importer, "SettingsFile"):
            importer.SettingsFile = settings
            applied.append("SettingsFile")
    else:
        log(T("imp.log.settings_missing"))
    output = ""
    try:
        output = translator_output_file(name)
    except Exception:
        try:
            output = part.FullPath or ""
        except Exception:
            output = ""
    if output and hasattr(importer, "OutputFile"):
        importer.OutputFile = output
        applied.append("OutputFile")
    if hasattr(importer, "FileOpenFlag"):
        importer.FileOpenFlag = False
        applied.append("FileOpenFlag=False")
    # 自动缝合：默认关。只有显式传 True 才写属性，False 时仅探查一次当前默认值，
    # 保证默认路径下导入结果与改动前完全一致（属性不存在只记日志，不抛错）。
    if label in ("STEP", "IGES"):
        if hasattr(importer, "SewSurfaces"):
            if label not in _sew_probed:
                _sew_probed.add(label)
                try:
                    log(T("imp.log.sew_probe", label, bool(importer.SewSurfaces)))
                except Exception as exc:
                    log(T("imp.log.sew_probe_fail", label, exc))
            if sew_surfaces:
                try:
                    importer.SewSurfaces = True
                    applied.append("SewSurfaces=True")
                    log(T("imp.log.sew_set", label))
                except Exception as exc:
                    log(T("imp.log.sew_set_fail", label, exc))
        else:
            log(T("imp.log.sew_unavailable", label))
    return ",".join(applied) if applied else "none"


def fit_view(part):
    try:
        view = part.Views.WorkView
        if view is not None:
            view.Fit()
            return True
    except Exception:
        pass
    try:
        part.Views.FitAll()
        return True
    except Exception as exc:
        log(T("imp.log.fit_fail", exc))
        return False


# ================= 文件名 / 部件名 =================

_ILLEGAL_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def sanitize_part_name(stem, used):
    """文件名主干 -> 合法部件名：清非法字符、截断 80、同名追加 _2/_3"""
    name = _ILLEGAL_CHARS.sub("_", stem or "").strip().strip(".")
    if not name:
        name = "imported"
    name = name[:MAX_NAME_LEN]
    base = name
    index = 2
    while name.lower() in used:
        suffix = "_%d" % index
        name = base[:MAX_NAME_LEN - len(suffix)] + suffix
        index += 1
    used.add(name.lower())
    return name


def save_target(path):
    """转换文件的落盘目标：源文件同目录、同主干的 .prt；同名已存在则 _2/_3，不覆盖。

    返回 (目标路径, 是否因同名改过名)。与导出侧 build_plan 的改名策略一致。
    """
    folder = os.path.dirname(os.path.abspath(path))
    stem = os.path.splitext(os.path.basename(path))[0]
    first = os.path.join(folder, stem + ".prt")
    target = first
    serial = 2
    while os.path.exists(target):
        target = os.path.join(folder, "%s_%d.prt" % (stem, serial))
        serial += 1
    return target, target != first


def file_written(path):
    """落盘校验：文件存在且非 0 字节（SaveAs 偶有“不报错但没写盘”）。"""
    try:
        return os.path.isfile(path) and os.path.getsize(path) > 0
    except Exception:
        return False


def existing_part_names(session):
    used = set()
    try:
        for part in list(session.Parts):
            try:
                leaf = part.LeafName or ""
            except Exception:
                leaf = ""
            stem = os.path.splitext(leaf)[0]
            if stem:
                used.add(stem.lower())
    except Exception:
        pass
    return used


def dedupe(paths):
    """按 normcase(abspath) 去重，保留用户选择顺序；返回 (队列, 跳过项)"""
    queue = []
    skipped = []
    seen = set()
    for path in paths:
        full = os.path.abspath(path)
        key = os.path.normcase(full)
        if key in seen:
            skipped.append((full, "dup"))
            continue
        seen.add(key)
        if not os.path.isfile(full):
            skipped.append((full, "missing"))
            continue
        queue.append(full)
    return queue, skipped


# ================= 单个文件 =================

def import_one(session, path, used, save_as=True):
    """返回 (状态, 部件名, 说明, 落盘状态, 部件)。

    状态 = ok / fail；落盘状态 = saved（已写入 .prt）/ mem（几何只在内存）/
    n/a（.prt 输入，不做另存）。
    save_as=False：只导入几何、不另存（供导出侧链式转换复用，调用方负责关闭部件）。
    """
    ext = os.path.splitext(path)[1].lower()
    started = time.time()

    if ext in SUPPORTED_PRT:
        log(T("imp.log.prt_open", path))
        part = find_open_part(session, path)
        already = part is not None
        if already:
            log(T("imp.log.prt_already"))
            try:
                how = activate_display(session, part)
                log(T("imp.log.keep_open", how))
            except Exception as exc:
                raise RuntimeError(T("imp.reason.open", exc))
            status = None
        else:
            part, status, how = open_part_file(session, path)
            log(T("imp.log.open_mode", how))
            trace("OPEN_MODE", how)
        name = ""
        try:
            name = os.path.splitext(part.LeafName or "")[0]
        except Exception:
            name = os.path.basename(path)
        if status is not None:
            text = _describe_load_status(status)
            if text:
                log(T("imp.log.load_status", text))
        log(T("imp.log.prt_done", time.time() - started, name))
        return "ok", name, "already_open" if already else "opened", "n/a", part

    if ext in SUPPORTED_STEP:
        methods = ("CreateStep214Importer", "CreateStepImporter")
        label = "STEP"
    elif ext in SUPPORTED_IGES:
        methods = ("CreateIgesImporter", "CreateIGESImporter")
        label = "IGES"
    elif ext in SUPPORTED_PARASOLID:
        methods = ("CreateParasolidImporter", "CreateXTImporter")
        label = "Parasolid"
    else:
        raise RuntimeError(T("imp.reason.ext"))

    importer = None
    part = None
    previous = None
    done = False
    part_keep = False
    name = ""
    trace("IMPORT_START", "%s ext=%s" % (path, ext))
    try:
        # 先定落盘目标，再用目标主干命名新部件：
        # NX 的 NewDisplay 会拒绝与磁盘上已有 .prt 同名的部件（"File already exists"），
        # 必须在建部件之前就把 _2/_3 定下来，否则同名再导时 save_target 永远轮不到执行。
        if save_as:
            save_path, renamed = save_target(path)
            stem = os.path.splitext(os.path.basename(save_path))[0]
        else:
            save_path, renamed = "", False
            stem = os.path.splitext(os.path.basename(path))[0]
        name = sanitize_part_name(stem, used)
        try:
            previous = session.Parts.Display
        except Exception:
            previous = None
        trace("DISPLAY_BEFORE", repr(previous))
        try:
            part, _status = new_display_part(session, name)
        except Exception as exc:
            raise RuntimeError(T("imp.reason.no_part", exc))
        try:
            trace("NEW_DISPLAY_OK", "%s full=%r" % (name, part.FullPath))
        except Exception:
            trace("NEW_DISPLAY_OK", name)
        log(T("imp.log.new_part", name))
        try:
            how = set_work_part(session, part)
            log(T("imp.log.work_part", how))
            trace("SET_WORK_OK", how)
        except Exception as exc:
            raise RuntimeError(T("imp.reason.work_part", exc))
        try:
            importer = create_importer(session, methods)
        except Exception as exc:
            raise RuntimeError(T("imp.reason.import", exc))
        trace("IMPORTER_OK", type(importer).__name__)
        log(T("imp.log.import_begin", label, name))

        try:
            attr = set_input_file(importer, path)
        except Exception as exc:
            raise RuntimeError(T("imp.reason.import", exc))
        trace("INPUT_FILE_OK", attr)
        try:
            prep = prepare_importer(importer, label, part, name)
            log(T("imp.log.prep", prep, getattr(importer, "OutputFile", "")))
            trace("PREP_OK", prep)
        except Exception as exc:
            log(T("imp.log.prep_fail", exc))
            trace("PREP_FAIL", exc)
        try:
            target = apply_import_to(importer)
            trace("IMPORT_TO_OK", target)
        except Exception as exc:
            log(T("imp.log.import_to", exc))
            target = "default"
            trace("IMPORT_TO_FAIL", exc)
        trace("COMMIT_BEFORE", "")
        try:
            status = importer.Commit()
        except Exception as exc:
            trace("COMMIT_EXC", exc)
            raise RuntimeError(T("imp.reason.import", exc))
        trace("COMMIT_AFTER", "%.2fs" % (time.time() - started))
        # 几何已进内存：此后即使落盘失败也不再走“关闭空部件”路径
        part_keep = True
        text = _describe_load_status(status)
        if text:
            log(T("imp.log.load_status", text))
        try:
            trace("FIT_BEFORE", "")
            fit_view(part)
            trace("FIT_AFTER", "%d bodies" % len(list(part.Bodies)))
        except Exception as exc:
            trace("FIT_FAIL", exc)
        # 与 NX 打开 STEP 一致：产物直接落在源文件同级目录，同名加 _2/_3
        if not save_as:
            # 链式转换：几何留在内存，由调用方导出目标格式后关闭部件
            done = True
            log(T("imp.log.import_done", time.time() - started))
            return "ok", name, label + "/" + attr + "/" + target, "mem", part
        if renamed:
            log(T("imp.log.saveas_rename", save_path))
        log(T("imp.log.saveas_start", save_path))
        trace("SAVEAS_BEFORE", save_path)
        try:
            part.SaveAs(save_path)
        except Exception as exc:
            trace("SAVEAS_EXC", exc)
            reason = T("imp.err.saveas", exc, save_path).strip()
            log(reason)
            return "fail", name, reason, "mem", part
        if not file_written(save_path):
            trace("SAVEAS_MISSING", save_path)
            reason = T("imp.err.saveas_missing", save_path).strip()
            log(reason)
            return "fail", name, reason, "mem", part
        trace("SAVEAS_AFTER", save_path)
        done = True
        log(T("imp.log.import_done", time.time() - started))
        log(T("imp.log.saveas_ok", save_path))
        return "ok", name, label + "/" + attr + "/" + target, "saved", part
    finally:
        trace("IMPORT_END", "done=%r %.2fs" % (done, time.time() - started))
        if importer is not None:
            try:
                importer.Destroy()
            except Exception as exc:
                log(T("imp.log.destroy_fail", exc))
        # 失败清理：先切回原显示部件，再关掉空部件；关不掉就保留并记日志
        # part_keep=True 表示几何已导入但落盘失败，部件要留在内存并保持可见
        if not done and part is not None and not part_keep:
            if previous is not None and previous is not part:
                try:
                    session.Parts.SetDisplay(previous, False, True)
                except Exception as exc:
                    log(T("imp.log.fail", exc))
            if close_part(session, part):
                log(T("imp.log.close_empty", name))
            else:
                log(T("imp.log.close_fail", name))
        # 成功时清掉本次 translator 产生的 .log；失败时保留，供汇总里给出路径
        if done:
            cleanup_run_logs(started)


# ================= 主流程 =================

def main():
    trace("MAIN_START", "")
    session = NXOpen.Session.GetSession()
    log_path = open_log()
    trace("LOG_OPEN", log_path)
    multi_status, multi_action = ensure_multi_display(session)
    trace("MULTI_DISPLAY", "status=%s action=%s after=%s"
          % (multi_status, multi_action, getattr(session.Parts, "AllowMultipleDisplayedParts", "?")))
    if "nx_batch_import_" in os.path.basename(log_path):
        log(T("imp.log.log_fallback", log_path))

    # 轻量清理：7 天前的 translator 日志；失败静默，绝不影响导入
    removed = cleanup_old_logs()
    if removed:
        log(T("imp.log.temp_clean_old", removed, TEMP_LOG_DAYS))

    listing = get_listing(session)
    trace("LISTING_OK", repr(listing))

    used_custom = False
    try:
        trace("DIALOG_BEFORE", "")
        picked, used_custom = pick_files_ui(read_last_dir())
        trace("DIALOG_AFTER", "%d custom=%s" % (len(picked or []), used_custom))
    except Exception as exc:
        log(T("imp.err.dialog", exc))
        _ui_error_log("picker failed: %s" % exc)
        try:
            picked = select_files(read_last_dir())
            used_custom = False
            trace("DIALOG_FALLBACK", "%d" % len(picked or []))
        except Exception as exc2:
            log(T("imp.err.dialog", exc2))
            message_box("Information", T("imp.confirm.title"),
                        T("imp.err.dialog", exc2))
            return
    if not picked:
        log(T("imp.msg.cancelled"))
        return

    try:
        # 记忆「最后一个加入的文件」所在目录，下次打开直接定位到这里
        write_last_dir(os.path.dirname(os.path.abspath(picked[-1])))
        trace("LAST_DIR_OK", "")
    except Exception:
        pass

    queue, skipped_all = dedupe(picked)
    for path, kind in skipped_all:
        if kind == "dup":
            log(T("imp.log.skip_dup", path))
        else:
            log(T("imp.log.skip_missing", path))
    if not queue:
        log(T("imp.msg.no_files"))
        message_box("Information", T("imp.confirm.title"), T("imp.msg.no_files"))
        return

    log(T("imp.log.header", len(queue), log_path))

    if used_custom:
        # 新选择窗里已选清单就是确认界面，不再二次弹框
        listing_write(listing, T("imp.confirm.body", len(queue),
                                 "\n".join(os.path.basename(item) for item in queue)))
    else:
        preview = queue[:10]
        body = T("imp.confirm.body", len(queue),
                 "\n".join(os.path.basename(item) for item in preview))
        if len(queue) > len(preview):
            body += T("imp.confirm.more", len(queue) - len(preview))
        body += "\n\n" + T("imp.confirm.proceed_log")

        listing_write(listing, T("imp.confirm.body", len(queue),
                                "\n".join(preview)))
        if not confirm_yes(T("imp.confirm.title"), body):
            log(T("imp.msg.cancelled"))
            return

    used = existing_part_names(session)
    ok_count = 0
    fail_count = 0
    failures = []
    visible = []
    saved_count = 0
    mem_count = 0
    total = len(queue)
    converted_total = sum(1 for item in queue
                          if os.path.splitext(item)[1].lower() not in SUPPORTED_PRT)

    for index, path in enumerate(queue, 1):
        log("")
        log(T("imp.log.file", index, total, path))
        listing_write(listing, T("imp.log.file", index, total,
                                 os.path.basename(path)))
        try:
            state, name, detail, disk, part = import_one(session, path, used)
            if state == "ok":
                ok_count += 1
                if part is not None:
                    keep_path = (path if os.path.splitext(path)[1].lower() in SUPPORTED_PRT
                                 else part_file_on_disk(part))
                    visible.append((name, keep_path, part))
            else:
                fail_count += 1
                failures.append((path, detail))
            if disk == "saved":
                saved_count += 1
            elif disk == "mem":
                mem_count += 1
        except Exception as exc:
            fail_count += 1
            reason = str(exc) or traceback_text()
            failures.append((path, reason))
            log(T("imp.log.fail", reason))
            listing_write(listing, T("imp.log.fail", reason))

    # 显示部件切换在 NX 里可能把上一个未修改部件挤出会话，这里兜底核对/补开，
    # 保证“选了 N 个文件就有 N 个部件留在会话里”。
    if visible:
        restored, missing = reopen_missing_parts(session, visible)
        if restored:
            log(T("imp.log.parts_restored", restored))
        if missing:
            log(T("imp.log.parts_missing", ", ".join(missing)))
        loaded = sum(1 for name, keep_path, part in visible
                     if part_still_loaded(session, part, keep_path))
        shown = displayed_part_names(session)
        log(T("imp.log.parts_check", len(visible), loaded))
        trace("PARTS_CHECK", "expected=%d restored=%d loaded=%d displayed=%d shown=%s"
              % (len(visible), restored, loaded, len(shown), ",".join(shown)))

    skip_count = len(skipped_all)
    log("")
    log(T("imp.log.summary", ok_count, fail_count, skip_count))

    lines = []
    for path, reason in failures[:10]:
        lines.append(T("imp.sum.fail_item", os.path.basename(path), reason))
    if len(failures) > 10:
        lines.append(T("imp.sum.more_fail", len(failures) - 10))
    if skip_count:
        lines.append(T("imp.sum.skipped", skip_count))
    if converted_total:
        lines.append(T("imp.sum.disk", saved_count, mem_count))
    if not lines:
        lines.append(T("imp.sum.no_fail"))
    if fail_count:
        lines.append(T("imp.sum.temp_logs", import_temp_dir()))

    summary = T("imp.sum.body", ok_count, fail_count, skip_count,
                "\n".join(lines), log_path)
    log(summary)
    listing_write(listing, summary)
    message_box("Information", T("imp.sum.title"), summary)

    if _log_handle:
        try:
            _log_handle.close()
        except Exception:
            pass


def traceback_text():
    try:
        import traceback
        return traceback.format_exc().strip().splitlines()[-1]
    except Exception:
        return "unknown error"


if __name__ == "__main__":
    main()
