# -*- coding: utf-8 -*-
r"""
批量把文件夹下所有 .prt 导出为 STEP (AP214)

用法一（推荐，NX 界面内）：
    NX 菜单栏左上角「菜单(Menu)」→ STEP Tools → Batch STEP Export...
    （菜单位于本目录 startup\step_export.men，动作写的是裸文件名 batch_export_step.py；
      NX 会把带盘符的绝对路径当成未知操作类型，所以脚本必须放在菜单所在的 startup 目录）
    运行后弹出自定义选择窗口：左侧文件夹树、右侧 .prt 列表、下方已选清单，
    右下角「确认导出」开始导出。窗口基于 PowerShell WinForms，打不开时自动回退系统对话框。
    （也可用 NX 菜单 文件 → 执行 → NX Open 手动选中本脚本，效果相同）
用法二（命令行批处理，无界面）：
    "C:\\Program Files\\Siemens\\NX 12.0\\NXBIN\\run_journal.exe" startup\batch_export_step.py

注意：
    1. 需要 STEP 转换器许可（Translator 授权），否则创建导出器会失败。
    2. 为判断图层在当前视图里是否可见，脚本会临时打开并显示正在导出的部件；
       导出结束只关闭脚本自己打开的部件，不会关闭你原本打开、正在编辑的部件，
       并把显示部件恢复成你原来那个。
    3. 每个 .prt 导出的 .step 就放在该 .prt 所在的那个文件夹里，
       文件分散在不同子文件夹时，结果也分别留在各自文件夹。
    4. 导出后会校验产物：文件存在、大于 0 字节且确实含几何才算成功；
       为空或只有文件头会自动重试一次，仍无效则记为失败并删除该文件。
       已存在但只有文件头（没有几何）的旧 .step 也会被重新导出，不会当成已完成跳过。
    5. 只导出该部件当前工作视图中“可见”的图层（LayerManager.GetVisibilitiesInView），
       隐藏图层上的实体会被排除。查询可见性要求部件是显示部件，因此脚本会临时
       打开并显示该部件，导出后只关闭脚本自己打开的部件，并恢复你原本的显示部件。
    6. NX 2506 起必须显式指定导出对象类型（实体/曲面/曲线）与图层范围，
       否则 STEP 里只有文件头、没有几何，脚本已默认设置好。
"""

import os
import glob
import json
import time
import stat
import fnmatch
import traceback

import NXOpen

# ================= 多语言 =================
# UGII_LANG 的常见取值：simpl_chinese / trad_chinese / schinese / chinese / english
# NX 界面实测还会给出 schinese（简中）与 chinese（繁中），与 install.ps1 的映射保持一致。
# 未知或未设置时回落英文；三张表缺项时逐项回落英文。

_ZH_CN_ALIASES = frozenset((
    "simpl_chinese", "simplified_chinese", "zh_cn", "zh-cn", "chs",
    "schinese", "zh_hans",
))
_ZH_TW_ALIASES = frozenset((
    "trad_chinese", "traditional_chinese", "zh_tw", "zh-tw", "cht",
    "chinese", "zh_hant",
))


def _detect_lang():
    value = (os.environ.get("UGII_LANG") or "").strip().lower()
    # 精确成员判断：chinese 不能吞掉 simpl_chinese / trad_chinese
    if value in _ZH_CN_ALIASES:
        return "zh_cn"
    if value in _ZH_TW_ALIASES:
        return "zh_tw"
    return "en"


LANG = _detect_lang()

MSG = {
    "en": {
        "log.ignore_outside": "  Ignored (outside root): {0}",
        "log.already_included": "  Already included, skipped: {0}",
        "log.selected": "  Selected: {0}",
        "log.root_all": "<Root: all>",
        "log.rename_conflict": "Name conflict: {0} and {1} both map to {2}, using {3}",
        "log.export_from_fail": "      ExportFrom failed (ignored): {0}",
        "log.objecttypes_fail": "      Object types failed (ignored): {0}",
        "log.layermask_fail": "      Layer mask failed (ignored): {0}",
        "log.holdflag_fail": "      ProcessHoldFlag failed (ignored): {0}",
        "log.retry_deleted": "      Removed previous invalid output, retrying",
        "log.retry_delete_fail": "      Failed to remove invalid output (ignored): {0}",
        "log.retry2_ok": "      Attempt 2 succeeded, wrote {0} bytes",
        "log.invalid_attempt": "      Attempt {0} invalid: {1}",
        "log.cleanup_fail": "      Failed to clean invalid output (ignored): {0}",
        "log.folder_dialog_fail": "Folder dialog unavailable ({0})",
        "log.file_dialog_fail": "File dialog unavailable ({0})",
        "log.opening_picker": "Opening selection window...",
        "log.picker_fail": "Selection window unavailable ({0}), falling back to system dialogs",
        "log.picked_folders": "Selected folders: {0}",
        "log.picked_root": "Root folder selected: export everything.",
        "log.picked_files": "Selected {0} file(s):",
        "log.no_selection": "Nothing selected, cancelled, no file exported.",
        "log.no_prt_reread": "No .prt found in the selected scope, please choose again.",
        "log.preview_fail": "Preview window unavailable ({0}), exporting by the plan directly",
        "log.back": "Back to modify, reopening the selection window.",
        "log.preview_cancel": "Cancelled at the preview list, no file exported.",
        "log.src_dir": "Source : {0}",
        "log.src_any": "Source : any locations you selected (folders / files)",
        "log.output_line": "Output : {0} next to each .prt (in place)",
        "log.file_count": "Files  : {0}",
        "log.include_dirs": "Include folders: {0}",
        "log.exclude_dirs": "Exclude folders: {0}",
        "log.include_files": "Include files  : {0}",
        "log.exclude_files": "Exclude files  : {0}",
        "log.manual_files": "Manually selected: {0}",
        "log.no_files_end": "No matching .prt found, done.",
        "log.dryrun_hint": "(Hint: in interactive mode the preview list replaces DRY_RUN)",
        "log.dryrun_title": "Dry run: the following files would be exported (not executed)",
        "log.dryrun_skip_mark": "  [exists, would be skipped]",
        "log.modified_header": "{0} selected file(s) are open with unsaved changes:",
        "log.cancel_unsaved": "Export cancelled (nothing saved, no file exported).",
        "log.user_skip_mod": "  Skipped by user (unsaved changes): {0}",
        "log.saved_ok": "  Saved (user confirmed): {0}",
        "log.save_fail_skip": "  Save failed, skipping instead: {0} ({1})",
        "log.noninteractive_skip": "  Non-interactive mode, skipped by default: {0}",
        "reason.user_skip": "Skipped (unsaved changes, user chose to skip)",
        "reason.save_fail": "Skipped (unsaved changes, save failed)",
        "reason.noninteractive": "Skipped (unsaved changes, non-interactive default)",
        "log.skip_exists": "{0} Skipped (already exists): {1}",
        "log.no_visible": "      Failed: no visible layer in view {0}, not exported",
        "log.visible": "      Visible layers: {0} (view {1}, {2} layer(s))",
        "log.cover": "{0} Overwrite export (same name already exists): {1}",
        "log.export_start": "{0} Export: {1}",
        "log.progress": "Exporting {0}/{1}: {2}",
        "log.done": "      Done {0:.1f}s -> {1} ({2} bytes)",
        "log.failed": "      Failed: {0}",
        "log.close_fail": "  Failed to close a temporarily opened part (ignored): {0}",
        "log.keep_modified": "  Kept open (unsaved changes): {0}",
        "log.display_restored": "  Original display part restored: {0}",
        "log.display_restore_fail": "  Failed to restore display part: {0}",
        "log.work_restored": "  Original work part restored: {0}",
        "log.work_restore_fail": "  Failed to restore work part: {0}",
        "log.summary": "Succeeded {0}, skipped {1}, failed {2}",
        "log.failed_item": "  Failed: {0}",
        "log.err_no_powershell": "powershell.exe not found",
        "log.err_no_result": "The window returned no result: {0}",
        "log.err_no_dir": "Folder does not exist: {0}",
        "log.err_creator": "Cannot create the STEP exporter (make sure the STEP translator is installed and licensed): {0}",
        "log.err_method_missing": "{0} not available",
        "log.err_method_fail": "{0} failed: {1}",
        "log.err_open_part": "Cannot open the part (it may be locked by another process or unreadable): {0}",
        "log.err_display_none": "Failed to switch the display part (NX returned none): {0}",
        "log.err_out_missing": "output file does not exist",
        "log.err_out_empty": "output file is 0 bytes",
        "log.err_out_header": "output file has only the header and no geometry ({0} bytes)",
        "log.err_export_invalid": "STEP export produced no valid content ({0})",
        "log.delete_old_fail": "      Failed to delete the old file (ignored): {0}",
        "log.err_remove_target": "Cannot overwrite {0}: the {2} file is read-only or open in another program (Excel/Notepad). Close it or clear read-only and retry ({1})",
        "log.missing_part": "File no longer exists (deleted or moved after selection): {0}",
        "log.busy": "Another export is already running; this run was not started (to avoid two windows changing the same files).",
        "log.stale_lock": "Removed the stale lock left by PID {0} and retried",
        "log.lock_unknown": "Cannot verify whether lock PID {0} still exists; treating the lock as busy",
        "log.log_fallback": "The log file is in use; this run writes its log to: {0}",
        "log.unsupported_name": "File name contains characters NX cannot handle (e.g., emoji); rename the file and retry: {0}",
        "log.old_restored": "Overwrite failed; the original .step was restored: {0}",
        "log.backup_cleanup_fail": "Failed to remove the overwrite backup (ignored): {0}",
        "log.too_long_path": "Path exceeds the Windows 260-character limit; shorten the folder depth and retry: {0}",
        "log.walk_fail": "Cannot access a folder (path too long or permission denied); skipped: {0}",
        "ui.folder_pick_title": "Select the folder to export as STEP (Cancel = finish)",
        "ui.file_pick_title": "Select the .prt / .stp / .igs files to export or convert",
        "ui.filter_prt": "NX parts and STEP/IGES (*.prt;*.stp;*.step;*.igs;*.iges)",
        "ui.filter_all": "All files (*.*)",
        "ui.picker_title": "Select folders / files to export or convert",
        "ui.picker_this_pc": "This PC",
        "ui.picker_tree_label": "Folders (click to select, double-click to expand)",
        "ui.picker_list_label": "*.prt / *.stp / *.igs files in current folder (Ctrl/Shift multi-select, double-click to add)",
        "ui.picker_empty_dir": "(No .prt / .stp / .igs files in this folder)",
        "ui.picker_denied": "Cannot access this folder (permission denied or invalid path):",
        "ui.picker_denied_hint": "This folder cannot be accessed. Choose another folder, or contact your administrator.",
        "ui.picker_col_name": "File name",
        "ui.picker_col_size": "Size(KB)",
        "ui.picker_col_type": "Type",
        "ui.picker_col_path": "Path",
        "ui.type_dir": "Folder",
        "ui.type_file": "File",
        "ui.btn_add_folder": "+ Add folder",
        "ui.btn_add_file": "+ Add file",
        "ui.btn_clear": "Clear",
        "ui.btn_cancel": "Cancel",
        "ui.btn_ok": "Export",
        "ui.busy_title": "Export in progress",
        "ui.busy_text": "Another export is already running. Please wait for it to finish, then try again.",
        "ui.busy_ok": "OK",
        "ui.msg_remove_title": "Confirm removal",
        "ui.msg_remove_text": "Remove this item from the list?",
        "ui.msg_empty_title": "Notice",
        "ui.msg_empty_one": "No folder or file has been added yet.",
        "ui.msg_empty_two": "Please select one and click Add folder / Add file.",
        "ui.preview_title": "Confirm export list",
        "ui.preview_hint": "Files are exported as listed; rows whose target exists can be checked as Overwrite (all unchecked by default). Click Start export to proceed, or Back to modify.",
        "ui.col_cover": "Overwrite",
        "ui.col_no": "No.",
        "ui.col_src": "Source",
        "ui.col_dst": "Output",
        "ui.col_state": "Status",
        "ui.state_new": "New",
        "ui.state_skip": "Skipped (exists)",
        "ui.state_overwrite": "Overwrite",
        "ui.chk_all": "Check all overwrite (existing rows only, off by default)",
        "ui.btn_back": "Back",
        "ui.btn_start": "Start export",
        "ui.summary": "{0} file(s): new/overwrite {1}, skipped {2} (existing rows can be checked as Overwrite; unchecked rows follow the status above)",
        "ui.save_title": "Unsaved changes",
        "ui.save_hint": "These files are open in NX with unsaved changes. Checked = save then export; unchecked = skip this item (all unchecked by default).",
        "ui.save_item_prefix": "Save and export: ",
        "ui.save_note": "Unchecked files will not be saved or exported.",
        "ui.btn_cancel_export": "Cancel this export",
        "ui.btn_save_go": "Save checked and continue",
        "ui.out_group": "Output location",
        "ui.out_src": "Next to the source part (default)",
        "ui.out_dir": "Custom output folder",
        "ui.out_browse": "Browse...",
        "ui.name_prefix": "Prefix",
        "ui.name_suffix": "Suffix",
        "ui.name_group": "File name and format",
        "ui.on_exists": "If the target exists",
        "ui.on_exists_skip": "Skip",
        "ui.on_exists_rename": "Rename (_2, _3 ...)",
        "ui.on_exists_overwrite": "Overwrite (backup first)",
        "ui.format": "Format",
        "ui.format_step": "STEP (.step)",
        "ui.format_iges": "IGES (.igs)",
        "ui.state_rename": "Rename",
        "ui.state_self": "Same format (skipped)",
        "log.chain_same_fmt": "Source and target are the same format; skipped: {0}",
        "log.chain_begin": "Converting in memory: {0} -> {1}",
        "log.chain_import_fail": "Chain import failed: {0}",
        "log.chain_close_fail": "Could not close the temporary in-memory part: {0}",
        "log.chain_restore_fail": "Could not restore the previous display part: {0}",
        "log.err_step_from_part": "StepCreator does not support DisplayPart/WorkPart export; in-memory STEP export is unavailable",
        "log.err_iges_from_part": "IgesCreator does not support DisplayPart/WorkPart export; in-memory IGES export is unavailable",
        "log.out_dir_invalid": "Output folder is empty or does not exist, falling back to the source folder: {0}",
        "log.output_custom": "Output : {0} (flat, source folder structure not mirrored)",
        "log.format_line": "Format : {0}",
        "log.rename_policy": "Target exists, renamed: {0} -> {1}",
        "log.iges_exe_missing": "IGES translator iges.exe not found: {0}",
        "log.iges_def_fallback": "igesexport.def not found in the NX install, falling back to: {0}",
        "log.iges_timeout": "IGES translation timed out after {0}s and was terminated",
        "log.iges_exit": "IGES translation failed (exit code {0})",
        "log.iges_log_unread": "      (cannot read the IGES log: {0})",
        "log.iges_cmd_fail": "Cannot start the IGES translator: {0}",
        "log.err_iges_out": "IGES output is missing or 0 bytes: {0}",
    },
    "zh_cn": {
        "log.ignore_outside": "  已忽略（不在根目录内）: {0}",
        "log.already_included": "  已包含，跳过: {0}",
        "log.selected": "  已选择: {0}",
        "log.root_all": "<根目录：全部>",
        "log.rename_conflict": "重名冲突: {0} 与 {1} 转名后同为 {2}，改用 {3}",
        "log.export_from_fail": "      ExportFrom 设置失败（忽略）: {0}",
        "log.objecttypes_fail": "      对象类型设置失败（忽略）: {0}",
        "log.layermask_fail": "      图层范围设置失败（忽略）: {0}",
        "log.holdflag_fail": "      ProcessHoldFlag 设置失败（忽略）: {0}",
        "log.retry_deleted": "      已删除上次的无效文件，准备重试",
        "log.retry_delete_fail": "      删除无效文件失败（忽略）: {0}",
        "log.retry2_ok": "      第 2 次尝试成功，输出 {0} 字节",
        "log.invalid_attempt": "      第 {0} 次导出无效：{1}",
        "log.cleanup_fail": "      清理无效文件失败（忽略）: {0}",
        "log.folder_dialog_fail": "文件夹选择框不可用（{0}）",
        "log.file_dialog_fail": "文件选择框不可用（{0}）",
        "log.opening_picker": "正在打开选择窗口…",
        "log.picker_fail": "选择窗口不可用（{0}），回退到系统对话框",
        "log.picked_folders": "已选文件夹: {0}",
        "log.picked_root": "已选择根目录：导出全部。",
        "log.picked_files": "已选文件 {0} 个:",
        "log.no_selection": "未选择任何文件夹或文件，已取消，未导出任何文件。",
        "log.no_prt_reread": "所选范围里没有找到 .prt 文件，请重新选择。",
        "log.preview_fail": "预览窗口不可用（{0}），直接按清单导出",
        "log.back": "返回修改，重新打开选择窗口。",
        "log.preview_cancel": "已在预览清单处取消，未导出任何文件。",
        "log.src_dir": "源目录 : {0}",
        "log.src_any": "源目录 : 你选择的任意位置（文件夹 / 文件）",
        "log.output_line": "输出    : 与各 .prt 同目录（就地生成 {0}）",
        "log.file_count": "文件数  : {0}",
        "log.include_dirs": "仅包含文件夹: {0}",
        "log.exclude_dirs": "排除文件夹  : {0}",
        "log.include_files": "仅包含文件  : {0}",
        "log.exclude_files": "排除文件    : {0}",
        "log.manual_files": "手动指定文件: {0} 个",
        "log.no_files_end": "没有找到符合条件的 .prt 文件，结束。",
        "log.dryrun_hint": "（提示：交互模式下选完会先弹清单预览，不需要改 DRY_RUN）",
        "log.dryrun_title": "试运行模式：以下文件将被导出（未实际执行）",
        "log.dryrun_skip_mark": "  [已存在，会跳过]",
        "log.modified_header": "检测到 {0} 个已打开且有未保存修改的文件：",
        "log.cancel_unsaved": "已取消本次导出（未保存修改未保存，未导出任何文件）。",
        "log.user_skip_mod": "  用户选择跳过（未保存修改）: {0}",
        "log.saved_ok": "  已保存（用户确认）: {0}",
        "log.save_fail_skip": "  保存失败，改为跳过: {0} ({1})",
        "log.noninteractive_skip": "  非交互模式，按默认跳过（未保存修改）: {0}",
        "reason.user_skip": "跳过（未保存修改，用户选择跳过）",
        "reason.save_fail": "跳过（未保存修改，保存失败）",
        "reason.noninteractive": "跳过（未保存修改，非交互模式默认跳过）",
        "log.skip_exists": "{0} 跳过（已存在）: {1}",
        "log.no_visible": "      失败: 当前视图({0})没有可见图层，未导出",
        "log.visible": "      可见图层: {0}（视图 {1}，共 {2} 层）",
        "log.cover": "{0} 覆盖导出（已存在同名文件）: {1}",
        "log.export_start": "{0} 导出: {1}",
        "log.progress": "正在导出 {0}/{1}: {2}",
        "log.done": "      完成 {0:.1f}s -> {1} ({2} 字节)",
        "log.failed": "      失败: {0}",
        "log.close_fail": "  关闭临时打开的部件失败（忽略）: {0}",
        "log.keep_modified": "  保持打开（有未保存修改）: {0}",
        "log.display_restored": "  已切回原显示部件: {0}",
        "log.display_restore_fail": "  切回原显示部件失败: {0}",
        "log.work_restored": "  已恢复原工作部件: {0}",
        "log.work_restore_fail": "  恢复原工作部件失败: {0}",
        "log.summary": "成功 {0} 个，跳过 {1} 个，失败 {2} 个",
        "log.failed_item": "  失败: {0}",
        "log.err_no_powershell": "找不到 powershell.exe",
        "log.err_no_result": "窗口没有返回结果: {0}",
        "log.err_no_dir": "目录不存在: {0}",
        "log.err_creator": "无法创建 STEP 导出器（请确认已安装 STEP 转换器并具备许可）: {0}",
        "log.err_method_missing": "{0} 不存在",
        "log.err_method_fail": "{0} 失败: {1}",
        "log.err_open_part": "无法打开部件（可能被其他进程占用或文件不可读）: {0}",
        "log.err_display_none": "切换显示部件失败（NX 未返回显示部件）: {0}",
        "log.err_out_missing": "输出文件不存在",
        "log.err_out_empty": "输出文件为 0 字节",
        "log.err_out_header": "输出文件只有文件头、没有几何（{0} 字节）",
        "log.err_export_invalid": "STEP 导出没有产出有效内容（{0}）",
        "log.delete_old_fail": "      删除旧文件失败（忽略）: {0}",
        "log.err_remove_target": "无法覆盖 {0}：该 {2} 文件只读或正被其他程序占用（如 Excel/记事本），请关闭或取消只读后重试（{1}）",
        "log.missing_part": "文件不存在（可能在选择后被删除或移动）: {0}",
        "log.busy": "已有一次导出正在进行，本次未启动（避免两个窗口同时改同一批文件）。",
        "log.stale_lock": "已删除 PID {0} 遗留的陈旧锁，并重试本次导出",
        "log.lock_unknown": "无法确认锁内进程 {0} 是否仍存在，按占用处理",
        "log.log_fallback": "日志文件被占用，本次日志改写到: {0}",
        "log.unsupported_name": "文件名包含 NX 无法识别的特殊字符（如 emoji），请改名后重试: {0}",
        "log.old_restored": "覆盖失败，已恢复原 .step: {0}",
        "log.backup_cleanup_fail": "清理覆盖备份失败（忽略）: {0}",
        "log.too_long_path": "路径超过 Windows 260 字符限制，NX 无法访问；请缩短目录层级后重试: {0}",
        "log.walk_fail": "无法访问某个目录（路径过长或权限不足），已跳过: {0}",
        "ui.folder_pick_title": "选择要导出 STEP 的文件夹（取消 = 结束选择）",
        "ui.file_pick_title": "选择要导出/转换的 .prt / .stp / .igs 文件",
        "ui.filter_prt": "NX 部件与 STEP/IGES (*.prt;*.stp;*.step;*.igs;*.iges)",
        "ui.filter_all": "所有文件 (*.*)",
        "ui.picker_title": "选择要导出/转换的文件夹 / 文件",
        "ui.picker_this_pc": "此电脑",
        "ui.picker_tree_label": "文件夹（单击选中，双击展开）",
        "ui.picker_list_label": "当前文件夹中的 .prt / .stp / .igs（可 Ctrl/Shift 多选，双击加入）",
        "ui.picker_empty_dir": "（该目录下没有 .prt / .stp / .igs 文件）",
        "ui.picker_denied": "无法访问该目录（权限不足或路径无效）：",
        "ui.picker_denied_hint": "无法访问这个文件夹，请换一个目录，或联系管理员。",
        "ui.picker_col_name": "文件名",
        "ui.picker_col_size": "大小(KB)",
        "ui.picker_col_type": "类型",
        "ui.picker_col_path": "路径",
        "ui.type_dir": "文件夹",
        "ui.type_file": "文件",
        "ui.btn_add_folder": "＋ 加入文件夹",
        "ui.btn_add_file": "＋ 加入文件",
        "ui.btn_clear": "清空",
        "ui.btn_cancel": "取消",
        "ui.btn_ok": "确认导出",
        "ui.busy_title": "导出正在进行",
        "ui.busy_text": "已有一次导出正在进行，请等它结束后再试。",
        "ui.busy_ok": "确定",
        "ui.msg_remove_title": "确认移除",
        "ui.msg_remove_text": "确定要从已选清单中移除吗？",
        "ui.msg_empty_title": "提示",
        "ui.msg_empty_one": "还没有加入任何文件夹或文件。",
        "ui.msg_empty_two": "请先选中后点击「加入文件夹 / 加入文件」。",
        "ui.preview_title": "确认导出清单",
        "ui.preview_hint": "以下文件将按列导出；目标已存在的行可勾选「覆盖」（默认不勾），确认无误后点「开始导出」，如需调整请点「返回修改」。",
        "ui.col_cover": "覆盖",
        "ui.col_no": "序",
        "ui.col_src": "源文件",
        "ui.col_dst": "输出",
        "ui.col_state": "状态",
        "ui.state_new": "新建",
        "ui.state_skip": "跳过（已存在）",
        "ui.state_overwrite": "覆盖",
        "ui.chk_all": "全选覆盖（只对已存在的行生效，默认关闭）",
        "ui.btn_back": "返回修改",
        "ui.btn_start": "开始导出",
        "ui.summary": "共 {0} 个文件：新建/覆盖 {1} 个，跳过 {2} 个（已存在的行可勾选「覆盖」，未勾选按上表状态执行）",
        "ui.save_title": "有未保存的修改",
        "ui.save_hint": "以下文件已在本机 NX 中打开且有未保存修改。勾选 = 先保存再导出；不勾选 = 跳过该项（默认全部不勾）。",
        "ui.save_item_prefix": "保存并导出：",
        "ui.save_note": "未勾选的文件不会被保存，也不会导出。",
        "ui.btn_cancel_export": "取消本次导出",
        "ui.btn_save_go": "保存勾选项并继续导出",
        "ui.out_group": "输出位置",
        "ui.out_src": "原文档所在路径（默认）",
        "ui.out_dir": "指定导出目录",
        "ui.out_browse": "浏览…",
        "ui.name_prefix": "前缀",
        "ui.name_suffix": "后缀",
        "ui.name_group": "文件名与格式",
        "ui.on_exists": "同名文件已存在时",
        "ui.on_exists_skip": "跳过",
        "ui.on_exists_rename": "新建改名（_2、_3…）",
        "ui.on_exists_overwrite": "覆盖（先备份）",
        "ui.format": "格式",
        "ui.format_step": "STEP (.step)",
        "ui.format_iges": "IGES (.igs)",
        "ui.state_rename": "改名新建",
        "ui.state_self": "同格式（跳过）",
        "log.chain_same_fmt": "源与目标格式相同，已跳过：{0}",
        "log.chain_begin": "内存中转转换：{0} → {1}",
        "log.chain_import_fail": "链式导入失败：{0}",
        "log.chain_close_fail": "无法关闭临时内存部件：{0}",
        "log.chain_restore_fail": "无法恢复原来的显示部件：{0}",
        "log.err_step_from_part": "StepCreator 不支持“显示/工作部件”导出，无法做内存 STEP 导出",
        "log.err_iges_from_part": "IgesCreator 不支持“显示/工作部件”导出，无法做内存 IGES 导出",
        "log.out_dir_invalid": "指定导出目录为空或不存在，已回落到原文档所在路径: {0}",
        "log.output_custom": "输出    : {0}（扁平放置，不镜像原目录结构）",
        "log.format_line": "格式    : {0}",
        "log.rename_policy": "同名已存在，改名输出: {0} -> {1}",
        "log.iges_exe_missing": "找不到 IGES 转换器 iges.exe: {0}",
        "log.iges_def_fallback": "NX 安装目录下未找到 igesexport.def，回落到: {0}",
        "log.iges_timeout": "IGES 转换超时（{0} 秒），已终止",
        "log.iges_exit": "IGES 转换失败（exit code {0}）",
        "log.iges_log_unread": "      （无法读取 IGES 日志: {0}）",
        "log.iges_cmd_fail": "无法启动 IGES 转换器: {0}",
        "log.err_iges_out": "IGES 输出文件不存在或为 0 字节: {0}",
    },
    "zh_tw": {
        "log.ignore_outside": "  已忽略（不在根目錄內）: {0}",
        "log.already_included": "  已包含，略過: {0}",
        "log.selected": "  已選擇: {0}",
        "log.root_all": "<根目錄：全部>",
        "log.rename_conflict": "名稱衝突: {0} 與 {1} 轉名後同為 {2}，改用 {3}",
        "log.export_from_fail": "      ExportFrom 設定失敗（忽略）: {0}",
        "log.objecttypes_fail": "      物件類型設定失敗（忽略）: {0}",
        "log.layermask_fail": "      圖層範圍設定失敗（忽略）: {0}",
        "log.holdflag_fail": "      ProcessHoldFlag 設定失敗（忽略）: {0}",
        "log.retry_deleted": "      已刪除上次的無效檔案，準備重試",
        "log.retry_delete_fail": "      刪除無效檔案失敗（忽略）: {0}",
        "log.retry2_ok": "      第 2 次嘗試成功，輸出 {0} 位元組",
        "log.invalid_attempt": "      第 {0} 次匯出無效：{1}",
        "log.cleanup_fail": "      清理無效檔案失敗（忽略）: {0}",
        "log.folder_dialog_fail": "資料夾選擇框無法使用（{0}）",
        "log.file_dialog_fail": "檔案選擇框無法使用（{0}）",
        "log.opening_picker": "正在開啟選擇視窗…",
        "log.picker_fail": "選擇視窗無法使用（{0}），改回系統對話框",
        "log.picked_folders": "已選資料夾: {0}",
        "log.picked_root": "已選擇根目錄：全部匯出。",
        "log.picked_files": "已選檔案 {0} 個:",
        "log.no_selection": "未選擇任何資料夾或檔案，已取消，未匯出任何檔案。",
        "log.no_prt_reread": "所選範圍內找不到 .prt 檔案，請重新選擇。",
        "log.preview_fail": "預覽視窗無法使用（{0}），直接依清單匯出",
        "log.back": "返回修改，重新開啟選擇視窗。",
        "log.preview_cancel": "已在預覽清單處取消，未匯出任何檔案。",
        "log.src_dir": "來源目錄 : {0}",
        "log.src_any": "來源目錄 : 您選擇的任意位置（資料夾 / 檔案）",
        "log.output_line": "輸出    : 與各 .prt 同目錄（就地產生 {0}）",
        "log.file_count": "檔案數  : {0}",
        "log.include_dirs": "僅包含資料夾: {0}",
        "log.exclude_dirs": "排除資料夾  : {0}",
        "log.include_files": "僅包含檔案  : {0}",
        "log.exclude_files": "排除檔案    : {0}",
        "log.manual_files": "手動指定檔案: {0} 個",
        "log.no_files_end": "找不到符合條件的 .prt 檔案，結束。",
        "log.dryrun_hint": "（提示：互動模式下選完會先彈出清單預覽，不需要改 DRY_RUN）",
        "log.dryrun_title": "試執行模式：以下檔案將被匯出（未實際執行）",
        "log.dryrun_skip_mark": "  [已存在，會略過]",
        "log.modified_header": "偵測到 {0} 個已開啟且有未儲存修改的檔案：",
        "log.cancel_unsaved": "已取消本次匯出（未儲存修改未儲存，未匯出任何檔案）。",
        "log.user_skip_mod": "  使用者選擇略過（未儲存修改）: {0}",
        "log.saved_ok": "  已儲存（使用者確認）: {0}",
        "log.save_fail_skip": "  儲存失敗，改為略過: {0} ({1})",
        "log.noninteractive_skip": "  非互動模式，依預設略過（未儲存修改）: {0}",
        "reason.user_skip": "略過（未儲存修改，使用者選擇略過）",
        "reason.save_fail": "略過（未儲存修改，儲存失敗）",
        "reason.noninteractive": "略過（未儲存修改，非互動模式預設略過）",
        "log.skip_exists": "{0} 略過（已存在）: {1}",
        "log.no_visible": "      失敗: 目前視圖({0})沒有可見圖層，未匯出",
        "log.visible": "      可見圖層: {0}（視圖 {1}，共 {2} 層）",
        "log.cover": "{0} 覆蓋匯出（已存在同名檔案）: {1}",
        "log.export_start": "{0} 匯出: {1}",
        "log.progress": "正在匯出 {0}/{1}: {2}",
        "log.done": "      完成 {0:.1f}s -> {1} ({2} 位元組)",
        "log.failed": "      失敗: {0}",
        "log.close_fail": "  關閉臨時開啟的部件失敗（忽略）: {0}",
        "log.keep_modified": "  保持開啟（有未儲存修改）: {0}",
        "log.display_restored": "  已切回原顯示部件: {0}",
        "log.display_restore_fail": "  切回原顯示部件失敗: {0}",
        "log.work_restored": "  已還原原工作部件: {0}",
        "log.work_restore_fail": "  還原原工作部件失敗: {0}",
        "log.summary": "成功 {0} 個，略過 {1} 個，失敗 {2} 個",
        "log.failed_item": "  失敗: {0}",
        "log.err_no_powershell": "找不到 powershell.exe",
        "log.err_no_result": "視窗沒有傳回結果: {0}",
        "log.err_no_dir": "目錄不存在: {0}",
        "log.err_creator": "無法建立 STEP 匯出器（請確認已安裝 STEP 轉換器並具備授權）: {0}",
        "log.err_method_missing": "{0} 不存在",
        "log.err_method_fail": "{0} 失敗: {1}",
        "log.err_open_part": "無法開啟部件（可能被其他程序佔用或檔案無法讀取）: {0}",
        "log.err_display_none": "切換顯示部件失敗（NX 未傳回顯示部件）: {0}",
        "log.err_out_missing": "輸出檔案不存在",
        "log.err_out_empty": "輸出檔案為 0 位元組",
        "log.err_out_header": "輸出檔案只有檔頭、沒有幾何（{0} 位元組）",
        "log.err_export_invalid": "STEP 匯出沒有產出有效內容（{0}）",
        "log.delete_old_fail": "      刪除舊檔案失敗（忽略）: {0}",
        "log.err_remove_target": "無法覆蓋 {0}：該 {2} 檔案唯讀或正被其他程式佔用（如 Excel/記事本），請關閉或取消唯讀後重試（{1}）",
        "log.missing_part": "檔案不存在（可能在選擇後被刪除或移動）: {0}",
        "log.busy": "已有一次匯出正在進行，本次未啟動（避免兩個視窗同時修改同一批檔案）。",
        "log.stale_lock": "已刪除 PID {0} 遺留的陳舊鎖，並重試本次匯出",
        "log.lock_unknown": "無法確認鎖內進程 {0} 是否仍存在，按佔用處理",
        "log.log_fallback": "日誌檔案被佔用，本次日誌改寫到: {0}",
        "log.unsupported_name": "檔案名稱包含 NX 無法辨識的特殊字元（如 emoji），請改名後重試: {0}",
        "log.old_restored": "覆蓋失敗，已還原原 .step: {0}",
        "log.backup_cleanup_fail": "清理覆蓋備份失敗（忽略）: {0}",
        "log.too_long_path": "路徑超過 Windows 260 字元限制，NX 無法存取；請縮短目錄層級後重試: {0}",
        "log.walk_fail": "無法存取某個目錄（路徑過長或權限不足），已略過: {0}",
        "ui.folder_pick_title": "選擇要匯出 STEP 的資料夾（取消 = 結束選擇）",
        "ui.file_pick_title": "選擇要匯出/轉換的 .prt / .stp / .igs 檔案",
        "ui.filter_prt": "NX 部件與 STEP/IGES (*.prt;*.stp;*.step;*.igs;*.iges)",
        "ui.filter_all": "所有檔案 (*.*)",
        "ui.picker_title": "選擇要匯出/轉換的資料夾 / 檔案",
        "ui.picker_this_pc": "本機",
        "ui.picker_tree_label": "資料夾（單擊選取，雙擊展開）",
        "ui.picker_list_label": "目前資料夾中的 .prt / .stp / .igs（可 Ctrl/Shift 多選，雙擊加入）",
        "ui.picker_empty_dir": "（此目錄下沒有 .prt / .stp / .igs 檔案）",
        "ui.picker_denied": "無法存取此目錄（權限不足或路徑無效）：",
        "ui.picker_denied_hint": "無法存取這個資料夾，請換一個目錄，或聯絡管理員。",
        "ui.picker_col_name": "檔案名稱",
        "ui.picker_col_size": "大小(KB)",
        "ui.picker_col_type": "類型",
        "ui.picker_col_path": "路徑",
        "ui.type_dir": "資料夾",
        "ui.type_file": "檔案",
        "ui.btn_add_folder": "＋ 加入資料夾",
        "ui.btn_add_file": "＋ 加入檔案",
        "ui.btn_clear": "清除",
        "ui.btn_cancel": "取消",
        "ui.btn_ok": "確認匯出",
        "ui.busy_title": "匯出正在進行",
        "ui.busy_text": "已有一次匯出正在進行，請等它結束後再試。",
        "ui.busy_ok": "確定",
        "ui.msg_remove_title": "確認移除",
        "ui.msg_remove_text": "確定要從已選清單中移除嗎？",
        "ui.msg_empty_title": "提示",
        "ui.msg_empty_one": "尚未加入任何資料夾或檔案。",
        "ui.msg_empty_two": "請先選取後點擊「加入資料夾 / 加入檔案」。",
        "ui.preview_title": "確認匯出清單",
        "ui.preview_hint": "以下檔案將依欄位匯出；目標已存在的列可勾選「覆蓋」（預設不勾），確認無誤後點「開始匯出」，如需調整請點「返回修改」。",
        "ui.col_cover": "覆蓋",
        "ui.col_no": "序",
        "ui.col_src": "來源檔案",
        "ui.col_dst": "輸出",
        "ui.col_state": "狀態",
        "ui.state_new": "新建",
        "ui.state_skip": "略過（已存在）",
        "ui.state_overwrite": "覆蓋",
        "ui.chk_all": "全選覆蓋（僅對已存在的列生效，預設關閉）",
        "ui.btn_back": "返回修改",
        "ui.btn_start": "開始匯出",
        "ui.summary": "共 {0} 個檔案：新建/覆蓋 {1} 個，略過 {2} 個（已存在的列可勾選「覆蓋」，未勾選依上表狀態執行）",
        "ui.save_title": "有未儲存的修改",
        "ui.save_hint": "以下檔案已在本機 NX 中開啟且有未儲存修改。勾選 = 先儲存再匯出；不勾選 = 略過該項（預設全部不勾）。",
        "ui.save_item_prefix": "儲存並匯出：",
        "ui.save_note": "未勾選的檔案不會被儲存，也不會匯出。",
        "ui.btn_cancel_export": "取消本次匯出",
        "ui.btn_save_go": "儲存勾選項並繼續匯出",
        "ui.out_group": "輸出路徑",
        "ui.out_src": "原文件所在路徑（預設）",
        "ui.out_dir": "指定匯出目錄",
        "ui.out_browse": "瀏覽…",
        "ui.name_prefix": "前綴",
        "ui.name_suffix": "後綴",
        "ui.name_group": "檔案名稱與格式",
        "ui.on_exists": "同名檔案已存在時",
        "ui.on_exists_skip": "略過",
        "ui.on_exists_rename": "新建改名（_2、_3…）",
        "ui.on_exists_overwrite": "覆蓋（先備份）",
        "ui.format": "格式",
        "ui.format_step": "STEP (.step)",
        "ui.format_iges": "IGES (.igs)",
        "ui.state_rename": "改名新建",
        "ui.state_self": "同格式（略過）",
        "log.chain_same_fmt": "來源與目標格式相同，已略過：{0}",
        "log.chain_begin": "記憶體中轉轉換：{0} → {1}",
        "log.chain_import_fail": "鏈式匯入失敗：{0}",
        "log.chain_close_fail": "無法關閉暫存記憶體部件：{0}",
        "log.chain_restore_fail": "無法還原原本的顯示部件：{0}",
        "log.err_step_from_part": "StepCreator 不支援「顯示/工作部件」匯出，無法做記憶體 STEP 匯出",
        "log.err_iges_from_part": "IgesCreator 不支援「顯示/工作部件」匯出，無法做記憶體 IGES 匯出",
        "log.out_dir_invalid": "指定匯出目錄為空或不存在，已改回原文件所在路徑: {0}",
        "log.output_custom": "輸出    : {0}（扁平放置，不鏡像原目錄結構）",
        "log.format_line": "格式    : {0}",
        "log.rename_policy": "同名已存在，改名輸出: {0} -> {1}",
        "log.iges_exe_missing": "找不到 IGES 轉換器 iges.exe: {0}",
        "log.iges_def_fallback": "NX 安裝目錄下找不到 igesexport.def，改回: {0}",
        "log.iges_timeout": "IGES 轉換逾時（{0} 秒），已終止",
        "log.iges_exit": "IGES 轉換失敗（exit code {0}）",
        "log.iges_log_unread": "      （無法讀取 IGES 日誌: {0}）",
        "log.iges_cmd_fail": "無法啟動 IGES 轉換器: {0}",
        "log.err_iges_out": "IGES 輸出檔案不存在或為 0 位元組: {0}",
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


def render_ps(template):
    """把 PowerShell 模板里的 [[key]] 占位符替换成当前语言文案（缺项回落英文）"""
    merged = dict(MSG["en"])
    merged.update(MSG.get(LANG) or {})
    for key, value in merged.items():
        template = template.replace("[[%s]]" % key, str(value))
    return template
# ==========================================

# 非 Unicode 控制台（英文版 cp437、繁中版 cp950 等）打印中文会抛 UnicodeEncodeError，
# 且异常发生在日志函数里，主循环的 except 再调日志会二次抛错、整个导出中断。
# 这里把 stdout 的编码错误处理降级为替换字符，让控制台输出永远不致命；
# 日志文件仍按 UTF-8 完整写入，不受影响。
try:
    import sys
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

# 注意：不要在本文件里定义 startup(argc, args)。
# startup 目录下的 .py 会被 NX 当成“启动应用”导入；一旦它有 startup 函数，
# 导入就会成功并被注册，之后点菜单用 journal 播放同一个模块时，
# nxpythonldr 会直接崩（syslog 里是 signal 11，界面弹空白的 Journal execution error）。
# 没有 startup 时 NX 只会在 syslog 记一条“no attribute startup”的提示，功能正常。

# ================= 配置区 =================
# 已移除 BASE_DIR：不再有“项目根目录”概念。选择窗口从「此电脑」开始枚举所有盘符，
# 可以逐层展开任意目录，并在任意位置勾选文件夹或 .prt 加入清单。
RECURSIVE = True        # True = 递归处理所选文件夹的所有子文件夹（每个 .prt 各自导出到原处）
SKIP_EXISTING = True    # True = 已存在同名 .step 就跳过（只有文件头、没有几何的不算已存在）
VISIBLE_LAYERS_ONLY = True   # True = 只导出当前工作视图中可见的图层（隐藏层实体不进 STEP）
DRY_RUN = False         # 交互模式下选完会自动弹「清单预览」确认，无需改这里；
                        # 仅 INTERACTIVE=False（无界面批处理）时用来只看不导

# INTERACTIVE = True 时弹出自定义选择窗口：
#   左侧文件夹树（顶层为「此电脑」+ 所有盘符，单击选中、双击展开）、右侧该文件夹下的 .prt 列表
#   （可多选）、下方已选清单（点行尾 ✕ 确认后移除，其余行自动上移）、右下角「确认导出」。
# PICK_MODE 仅在窗口打不开、回退到系统对话框时决定弹哪些框。
PICK_MODE = "both"      # 仅影响回退方案："folder" / "file" / "both"
INTERACTIVE = True      # False = 不弹任何窗，完全按下面的配置规则跑

# ---- 筛选规则（不带 BASE_DIR，目录一律写绝对路径），支持 * 通配符，不区分大小写 ----
# 例：INCLUDE_DIRS = [r"D:\parts\box1", r"E:\proj"]   只处理这些文件夹（含其子文件夹）
#     EXCLUDE_DIRS = ["old*"]                         不处理名字匹配的文件夹（任意层级）
#     INCLUDE_FILES = ["顶盖.prt", "*上盖*"]           只处理文件名匹配这些的 prt
#     EXCLUDE_FILES = ["*draft*"]                     跳过文件名匹配的 prt
# 空列表 = 该条规则不生效；交互模式下以选择窗口勾选的结果为准
INCLUDE_DIRS = []
EXCLUDE_DIRS = []
INCLUDE_FILES = []
EXCLUDE_FILES = []
# =========================================

# 输出文件名里的全角/中文符号 -> 英文符号（只影响导出的 .step 文件名，不改原 .prt）
NAME_REPLACE = {
    u"\uff08": "(",   # （
    u"\uff09": ")",   # ）
    u"\u3010": "[",   # 【
    u"\u3011": "]",   # 】
}

# IGES 分支：直调 NX 自带的 translator（iges.exe），不走 NXOpen 导出器。
# 仅当 UI 里把“格式”选成 IGES 时才会用到这些常量；默认 STEP 流程完全不受影响。
IGES_DEFAULT_DIR = os.path.join(os.environ.get("UGII_BASE_DIR", ""), "IGES")   # 通过 NX 环境变量定位
IGES_TIMEOUT = 600                     # 单件 IGES 转换超时（秒）
IGES_LOG_TAIL_LINES = 10               # 失败时日志尾部打印行数

_log_handle = None
_RUN_LOCK = None


def default_options():
    """导出选项默认值：不选新选项时与改动前的行为完全一致"""
    return {
        "out_mode": "source",                       # source = 原文档所在路径
        "out_dir": "",                              # 指定导出目录（out_mode=dir 时有效）
        "prefix": "",
        "suffix": "",
        # SKIP_EXISTING 是旧配置项：关闭它时旧行为是“已存在则覆盖”，这里保持一致
        "on_exists": "skip" if SKIP_EXISTING else "overwrite",
        "format": "step",                           # step / iges
    }


def output_format(options):
    """规范化输出格式：只认 step / iges，其余一律回落 step"""
    value = str((options or {}).get("format") or "").strip().lower()
    return "iges" if value == "iges" else "step"


def output_ext(options):
    """输出扩展名（含点）：STEP -> .step，IGES -> .igs"""
    return ".igs" if output_format(options) == "iges" else ".step"


def output_is_valid(path, options=None):
    """目标文件是否算“已存在”：STEP 沿用含几何判定，IGES 只判存在且大小 > 0"""
    if output_format(options) == "iges":
        return step_size(path) > 0
    return step_is_valid(path)


def resolve_output_options(options):
    """校验“指定导出目录”：为空或不存在时回落到原文档所在路径，并记一条日志"""
    result = dict(options or default_options())
    if result.get("out_mode") != "dir":
        result["out_mode"] = "source"
        result["out_dir"] = ""
        return result
    raw = str(result.get("out_dir") or "").strip()
    path = os.path.abspath(raw) if raw else ""
    if not raw or not os.path.isdir(path):
        log(T("log.out_dir_invalid", path or raw or "-"))
        result["out_mode"] = "source"
        result["out_dir"] = ""
        return result
    result["out_dir"] = path
    return result


def normalize_name(name):
    """把文件名中的全角括号等字符替换为英文符号"""
    for old, new in NAME_REPLACE.items():
        name = name.replace(old, new)
    return name


def safe_name_part(text):
    """前缀/后缀里 Windows 文件名不允许的字符替换成下划线，避免拼出非法路径"""
    value = str(text or "")
    for ch in '\\/:*?"<>|':
        value = value.replace(ch, "_")
    return value.strip()


def force_remove(path):
    """删除文件；只读属性导致失败时先去掉只读再删（用于清理覆盖备份）"""
    try:
        os.remove(path)
    except PermissionError:
        os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
        os.remove(path)


def log(msg):
    """同时打印到 NX 信息窗口/控制台和日志文件"""
    text = str(msg)
    try:
        print(text)
    except Exception:
        # 控制台代码页无法表示这些字符（reconfigure 不可用时的兜底）：
        # 只放弃屏幕输出，日志文件仍照常写入
        pass
    if _log_handle:
        _log_handle.write(text + "\n")
        _log_handle.flush()


def plugin_dir():
    """插件根目录（脚本在 startup\ 下，取它的上一层；exe/被 exec 执行时回落当前目录）"""
    try:
        here = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()
    if os.path.basename(here).lower() == "startup":
        parent = os.path.dirname(here)
        if parent:
            return parent
    return here


# 选择窗「上次浏览目录」记忆：与导入侧 import_last_dir.txt 各存各的，互不影响
EXPORT_LAST_DIR_FILE = "export_last_dir.txt"


def _last_dir_path():
    return os.path.join(plugin_dir(), "logs", EXPORT_LAST_DIR_FILE)


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


# 选择窗「上次用过的选项」记忆：与目录记忆分开存，损坏或值非法时静默回落默认值
EXPORT_OPTIONS_FILE = "export_options.json"


def _options_path():
    return os.path.join(plugin_dir(), "logs", EXPORT_OPTIONS_FILE)


def read_options():
    try:
        with open(_options_path(), "r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except Exception:
        return {}
    return parse_picker_options(payload, {})


def write_options(options):
    try:
        path = _options_path()
        parent = os.path.dirname(path)
        if parent and not os.path.isdir(parent):
            os.makedirs(parent)
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(dict(options or {}), handle, ensure_ascii=False, indent=1, sort_keys=True)
    except Exception:
        pass


SELECT_UI_PS = r'''
param(
    [string]$Root = "",
    [string]$OutMode = "source",
    [string]$OutDir = "",
    [string]$Prefix = "",
    [string]$Suffix = "",
    [string]$OnExists = "skip",
    [string]$Format = "step",
    [Parameter(Mandatory=$true)][string]$OutFile
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "[[ui.picker_title]]"
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
$lblTree.Text = "[[ui.picker_tree_label]]"
$lblTree.Location = New-Object System.Drawing.Point(12, 8)
$lblTree.Size = New-Object System.Drawing.Size(340, 20)
$form.Controls.Add($lblTree)

$tree = New-Object System.Windows.Forms.TreeView
$tree.Location = New-Object System.Drawing.Point(12, 30)
$tree.Size = New-Object System.Drawing.Size(340, 336)
$tree.HideSelection = $false
$tree.ItemHeight = 22
# NX 树形控件不用点状连线，展开箭头由系统绘制
$tree.ShowLines = $false
$tree.ShowRootLines = $false
$tree.Indent = 16
$tree.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($tree)

$lblList = New-Object System.Windows.Forms.Label
$lblList.Text = "[[ui.picker_list_label]]"
$lblList.Location = New-Object System.Drawing.Point(364, 8)
$lblList.Size = New-Object System.Drawing.Size(568, 20)
$form.Controls.Add($lblList)

$list = New-Object System.Windows.Forms.ListView
$list.Location = New-Object System.Drawing.Point(364, 30)
$list.Size = New-Object System.Drawing.Size(568, 336)
$list.View = "Details"
$list.FullRowSelect = $true
$list.MultiSelect = $true
$list.GridLines = $true
[void]$list.Columns.Add("[[ui.picker_col_name]]", 440)
[void]$list.Columns.Add("[[ui.picker_col_size]]", 100)
$list.ShowItemToolTips = $true
$form.Controls.Add($list)

$grid = New-Object System.Windows.Forms.ListView
$grid.Location = New-Object System.Drawing.Point(12, 378)
$grid.Size = New-Object System.Drawing.Size(920, 160)
$grid.View = "Details"
$grid.FullRowSelect = $true
$grid.MultiSelect = $false
$grid.GridLines = $true
[void]$grid.Columns.Add("[[ui.picker_col_type]]", 60)
[void]$grid.Columns.Add("[[ui.picker_col_path]]", 780)
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

# 分隔线：位于已选清单与下方选项分组之间
Add-Separator 12 540 920

# ---- 导出选项：输出位置 / 文件名前后缀 / 同名策略 / 输出格式 ----
$grpOut = New-Object System.Windows.Forms.GroupBox
$grpOut.Text = "[[ui.out_group]]"
$grpOut.Location = New-Object System.Drawing.Point(12, 546)
$grpOut.Size = New-Object System.Drawing.Size(452, 108)
$form.Controls.Add($grpOut)

$rbOutSrc = New-Object System.Windows.Forms.RadioButton
$rbOutSrc.Text = "[[ui.out_src]]"
$rbOutSrc.Location = New-Object System.Drawing.Point(14, 22)
$rbOutSrc.Size = New-Object System.Drawing.Size(420, 22)
$grpOut.Controls.Add($rbOutSrc)

$rbOutDir = New-Object System.Windows.Forms.RadioButton
$rbOutDir.Text = "[[ui.out_dir]]"
$rbOutDir.Location = New-Object System.Drawing.Point(14, 48)
$rbOutDir.Size = New-Object System.Drawing.Size(150, 22)
$grpOut.Controls.Add($rbOutDir)

$txtOutDir = New-Object System.Windows.Forms.TextBox
$txtOutDir.Text = $OutDir
$txtOutDir.Location = New-Object System.Drawing.Point(168, 47)
$txtOutDir.Size = New-Object System.Drawing.Size(160, 24)
$grpOut.Controls.Add($txtOutDir)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "[[ui.out_browse]]"
$btnBrowse.Location = New-Object System.Drawing.Point(332, 46)
$btnBrowse.Size = New-Object System.Drawing.Size(104, 26)
$grpOut.Controls.Add($btnBrowse)

# 浏览按钮：FolderBrowserDialog，选中的目录直接回填到文本框
$btnBrowse.add_Click({
    param($sender, $e)
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "[[ui.out_browse]]"
    if ($txtOutDir.Text -and [System.IO.Directory]::Exists($txtOutDir.Text)) {
        $dlg.SelectedPath = $txtOutDir.Text
    }
    if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOutDir.Text = $dlg.SelectedPath
    }
})

if ($OutMode -eq "dir") { $rbOutDir.Checked = $true } else { $rbOutSrc.Checked = $true }

function Set-OutDirEnabled {
    $on = $rbOutDir.Checked
    $txtOutDir.Enabled = $on
    $btnBrowse.Enabled = $on
}
Set-OutDirEnabled
$rbOutSrc.add_CheckedChanged({ param($sender, $e) Set-OutDirEnabled })
$rbOutDir.add_CheckedChanged({ param($sender, $e) Set-OutDirEnabled })

$grpName = New-Object System.Windows.Forms.GroupBox
$grpName.Text = "[[ui.name_group]]"
$grpName.Location = New-Object System.Drawing.Point(472, 546)
$grpName.Size = New-Object System.Drawing.Size(460, 108)
$form.Controls.Add($grpName)

$lblPrefix = New-Object System.Windows.Forms.Label
$lblPrefix.Text = "[[ui.name_prefix]]"
$lblPrefix.Location = New-Object System.Drawing.Point(14, 26)
$lblPrefix.Size = New-Object System.Drawing.Size(46, 20)
$grpName.Controls.Add($lblPrefix)

$txtPrefix = New-Object System.Windows.Forms.TextBox
$txtPrefix.Text = $Prefix
$txtPrefix.Location = New-Object System.Drawing.Point(62, 24)
$txtPrefix.Size = New-Object System.Drawing.Size(130, 24)
$grpName.Controls.Add($txtPrefix)

$lblSuffix = New-Object System.Windows.Forms.Label
$lblSuffix.Text = "[[ui.name_suffix]]"
$lblSuffix.Location = New-Object System.Drawing.Point(200, 26)
$lblSuffix.Size = New-Object System.Drawing.Size(46, 20)
$grpName.Controls.Add($lblSuffix)

$txtSuffix = New-Object System.Windows.Forms.TextBox
$txtSuffix.Text = $Suffix
$txtSuffix.Location = New-Object System.Drawing.Point(248, 24)
$txtSuffix.Size = New-Object System.Drawing.Size(130, 24)
$grpName.Controls.Add($txtSuffix)

$lblExists = New-Object System.Windows.Forms.Label
$lblExists.Text = "[[ui.on_exists]]"
$lblExists.Location = New-Object System.Drawing.Point(14, 60)
$lblExists.Size = New-Object System.Drawing.Size(110, 20)
$grpName.Controls.Add($lblExists)

$cmbExists = New-Object System.Windows.Forms.ComboBox
$cmbExists.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbExists.Location = New-Object System.Drawing.Point(128, 58)
$cmbExists.Size = New-Object System.Drawing.Size(150, 24)
[void]$cmbExists.Items.Add("[[ui.on_exists_skip]]")
[void]$cmbExists.Items.Add("[[ui.on_exists_rename]]")
[void]$cmbExists.Items.Add("[[ui.on_exists_overwrite]]")
$cmbExists.SelectedIndex = 0
if ($OnExists -eq "rename") { $cmbExists.SelectedIndex = 1 }
elseif ($OnExists -eq "overwrite") { $cmbExists.SelectedIndex = 2 }
$grpName.Controls.Add($cmbExists)

$lblFormat = New-Object System.Windows.Forms.Label
$lblFormat.Text = "[[ui.format]]"
$lblFormat.Location = New-Object System.Drawing.Point(290, 60)
$lblFormat.Size = New-Object System.Drawing.Size(36, 20)
$grpName.Controls.Add($lblFormat)

$cmbFormat = New-Object System.Windows.Forms.ComboBox
$cmbFormat.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbFormat.Location = New-Object System.Drawing.Point(328, 58)
$cmbFormat.Size = New-Object System.Drawing.Size(110, 24)
[void]$cmbFormat.Items.Add("[[ui.format_step]]")
[void]$cmbFormat.Items.Add("[[ui.format_iges]]")
$cmbFormat.SelectedIndex = 0
if ($Format -eq "iges") { $cmbFormat.SelectedIndex = 1 }
$grpName.Controls.Add($cmbFormat)

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

function Add-GridRow([string]$typeCode, [string]$path) {
    foreach ($row in $grid.Items) {
        if ($row.SubItems[1].Text -eq $path) { return }
    }
    $typeText = "[[ui.type_file]]"
    if ($typeCode -eq "DIR") { $typeText = "[[ui.type_dir]]" }
    $item = New-Object System.Windows.Forms.ListViewItem($typeText)
    $item.ImageIndex = 0
    $item.Tag = $typeCode
    $item.ToolTipText = $path
    [void]$item.SubItems.Add($path)
    [void]$item.SubItems.Add("✕")
    $grid.BeginUpdate()
    try { [void]$grid.Items.Add($item) } finally { $grid.EndUpdate() }
    Fit-GridColumns
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
$pcNode = New-Object System.Windows.Forms.TreeNode("[[ui.picker_this_pc]]")
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
                $lblList.Text = "[[ui.picker_denied]]"
            }
        }
    }
})

$tree.add_AfterSelect({
    param($sender, $e)
    $path = $e.Node.Tag
    $list.Items.Clear()
    $lblList.Text = "[[ui.picker_list_label]]"
    if (-not $path) { return }
    # Directory.Exists 对“权限不足”的目录同样返回 False，这里必须显式提示，
    # 否则用户只会看到列表被清空，误以为该目录本来就没有 .prt。
    if (-not [System.IO.Directory]::Exists($path)) {
        $lblList.Text = "[[ui.picker_denied]]"
        # 界面上只给人话 + 路径；完整原因写日志，避免把 .NET 异常原文糊到用户脸上
        try {
            $errLog = Join-Path $env:TEMP "nx_step_export_ui_errors.log"
            Add-Content -LiteralPath $errLog -Encoding UTF8 -Value ("[{0}] picker: cannot access '{1}' (Directory.Exists=False; missing or permission denied)" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $path)
        } catch { }
        # 用 BeginInvoke 排到消息循环里再弹，并指定 owner：本窗体是 TopMost，
        # 无 owner 的弹窗会被压到窗体后面，看起来像没提示
        $hint = "[[ui.picker_denied_hint]]`n`n$path"
        $job = { [void][System.Windows.Forms.MessageBox]::Show($form, $hint, "[[ui.picker_title]]", "OK", "Warning") }.GetNewClosure()
        [void]$form.BeginInvoke([System.Windows.Forms.MethodInvoker]$job)
        return
    }
    try {
        $dirKey = $path + "|files"
        if ($script:dirCache.ContainsKey($dirKey)) {
            $files = $script:dirCache[$dirKey]
        } else {
            $files = [System.IO.Directory]::GetFiles($path, "*.*") | Where-Object {
                $_.ToLower().EndsWith(".prt") -or $_.ToLower().EndsWith(".stp") -or
                $_.ToLower().EndsWith(".step") -or $_.ToLower().EndsWith(".igs") -or
                $_.ToLower().EndsWith(".iges")
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
            # 批量添加，避免大目录逐行刷新
            $list.BeginUpdate()
            $list.Items.AddRange([System.Windows.Forms.ListViewItem[]]$arr.ToArray())
            $list.EndUpdate()
        }
        if ($arr.Count -eq 0) { $lblList.Text = "[[ui.picker_empty_dir]]" }
    } catch {
        $lblList.Text = "[[ui.picker_denied]]"
        # 异常原文只进日志（含类型与堆栈），弹窗只保留易懂的提示
        try {
            $errLog = Join-Path $env:TEMP "nx_step_export_ui_errors.log"
            Add-Content -LiteralPath $errLog -Encoding UTF8 -Value ("[{0}] picker: failed to list .prt in '{1}'`r`n{2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $path, $_.Exception.ToString())
        } catch { }
        $hint = "[[ui.picker_denied_hint]]`n`n$path"
        $job = { [void][System.Windows.Forms.MessageBox]::Show($form, $hint, "[[ui.picker_title]]", "OK", "Warning") }.GetNewClosure()
        [void]$form.BeginInvoke([System.Windows.Forms.MethodInvoker]$job)
    }
})

$list.add_DoubleClick({
    param($sender, $e)
    if ($list.SelectedItems.Count -eq 0) { return }
    foreach ($sel in $list.SelectedItems) {
        Add-GridRow "FILE" $sel.Tag
    }
})

$grid.add_MouseClick({
    param($sender, $e)
    $hit = $grid.HitTest($e.X, $e.Y)
    if (-not $hit.Item) { return }
    if (-not $hit.SubItem) { return }
    if ($hit.Item.SubItems.IndexOf($hit.SubItem) -ne 2) { return }
    $path = $hit.Item.SubItems[1].Text
    $ans = [System.Windows.Forms.MessageBox]::Show("[[ui.msg_remove_text]]`n`n$path", "[[ui.msg_remove_title]]", "YesNo", "Question")
    if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
        $grid.Items.RemoveAt($hit.Item.Index)
        Fit-GridColumns
    }
})

$btnAddFolder = New-Object System.Windows.Forms.Button
$btnAddFolder.Text = "[[ui.btn_add_folder]]"
$btnAddFolder.Location = New-Object System.Drawing.Point(12, 669)
$btnAddFolder.Size = New-Object System.Drawing.Size(120, 27)
$form.Controls.Add($btnAddFolder)
$btnAddFolder.add_Click({
    param($sender, $e)
    if ($tree.SelectedNode -and $tree.SelectedNode.Tag) {
        Add-GridRow "DIR" $tree.SelectedNode.Tag
    }
})

$btnAddFile = New-Object System.Windows.Forms.Button
$btnAddFile.Text = "[[ui.btn_add_file]]"
$btnAddFile.Location = New-Object System.Drawing.Point(142, 669)
$btnAddFile.Size = New-Object System.Drawing.Size(110, 27)
$form.Controls.Add($btnAddFile)
$btnAddFile.add_Click({
    param($sender, $e)
    if ($list.SelectedItems.Count -eq 0) { return }
    foreach ($sel in $list.SelectedItems) {
        Add-GridRow "FILE" $sel.Tag
    }
})

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "[[ui.btn_clear]]"
$btnClear.Location = New-Object System.Drawing.Point(262, 669)
$btnClear.Size = New-Object System.Drawing.Size(70, 27)
$form.Controls.Add($btnClear)
$btnClear.add_Click({
    param($sender, $e)
    $grid.Items.Clear()
    Fit-GridColumns
})

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "[[ui.btn_cancel]]"
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
$btnOK.Text = "[[ui.btn_ok]]"
$btnOK.Location = New-Object System.Drawing.Point(832, 669)
$btnOK.Size = New-Object System.Drawing.Size(100, 27)
$form.Controls.Add($btnOK)
$btnOK.add_Click({
    param($sender, $e)
    if ($grid.Items.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("[[ui.msg_empty_one]]`n[[ui.msg_empty_two]]", "[[ui.msg_empty_title]]", "OK", "Warning")
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
Apply-Theme $form
[void]$form.ShowDialog()

$folders = @()
$files = @()
$ok = ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK)
if ($ok) {
    foreach ($row in $grid.Items) {
        if ($row.Tag -eq "DIR") { $folders += $row.SubItems[1].Text }
        else { $files += $row.SubItems[1].Text }
    }
}

# 选项字段：取消时也照样回传（Python 侧只在 ok=True 时使用，便于“返回修改”保留上次选择）
$outModeValue = "source"
if ($rbOutDir.Checked) { $outModeValue = "dir" }
$onExistsValue = "skip"
if ($cmbExists.SelectedIndex -eq 1) { $onExistsValue = "rename" }
elseif ($cmbExists.SelectedIndex -eq 2) { $onExistsValue = "overwrite" }
$formatValue = "step"
if ($cmbFormat.SelectedIndex -eq 1) { $formatValue = "iges" }

$payload = [ordered]@{ ok = $ok; folders = $folders; files = $files;
    out_mode = $outModeValue; out_dir = $txtOutDir.Text.Trim();
    prefix = $txtPrefix.Text.Trim(); suffix = $txtSuffix.Text.Trim();
    on_exists = $onExistsValue; format = $formatValue }
$json = $payload | ConvertTo-Json -Depth 4 -Compress
[System.IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
$form.Dispose()
'''


PREVIEW_PS = r'''
param(
    [Parameter(Mandatory=$true)][string]$PlanFile,
    [Parameter(Mandatory=$true)][string]$OutFile
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$plan = Get-Content -Raw -Encoding UTF8 $PlanFile | ConvertFrom-Json

$form = New-Object System.Windows.Forms.Form
$form.Text = "[[ui.preview_title]]"
$form.ClientSize = New-Object System.Drawing.Size(960, 540)
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

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = "[[ui.preview_hint]]"
$lbl.Location = New-Object System.Drawing.Point(12, 10)
$lbl.Size = New-Object System.Drawing.Size(936, 22)
$form.Controls.Add($lbl)

$view = New-Object System.Windows.Forms.ListView
$view.Location = New-Object System.Drawing.Point(12, 36)
$view.Size = New-Object System.Drawing.Size(936, 396)
$view.View = "Details"
$view.FullRowSelect = $true
$view.MultiSelect = $false
$view.GridLines = $true
$view.CheckBoxes = $true
[void]$view.Columns.Add("[[ui.col_cover]]", 90)
[void]$view.Columns.Add("[[ui.col_no]]", 46)
[void]$view.Columns.Add("[[ui.col_src]]", 300)
[void]$view.Columns.Add("[[ui.col_dst]]", 300)
[void]$view.Columns.Add("[[ui.col_state]]", 150)
$view.ShowItemToolTips = $true
$form.Controls.Add($view)

# 只有“目标已存在”的行允许勾选；新建行点了也会被立刻取消，避免误勾
$view.add_ItemCheck({
    param($sender, $e)
    $tag = $view.Items[$e.Index].Tag
    if (-not $tag -or -not $tag.exists) {
        $e.NewValue = [System.Windows.Forms.CheckState]::Unchecked
    }
})

$items = New-Object System.Collections.ArrayList
foreach ($row in $plan.rows) {
    $stateText = "[[ui.state_new]]"
    if ($row.state -eq "SKIP") { $stateText = "[[ui.state_skip]]" }
    elseif ($row.state -eq "OVERWRITE") { $stateText = "[[ui.state_overwrite]]" }
    elseif ($row.state -eq "RENAME") { $stateText = "[[ui.state_rename]]" }
    elseif ($row.state -eq "SELF") { $stateText = "[[ui.state_self]]" }
    $item = New-Object System.Windows.Forms.ListViewItem("")
    [void]$item.SubItems.Add([string]$row.no)
    [void]$item.SubItems.Add($row.src)
    [void]$item.SubItems.Add($row.dst)
    [void]$item.SubItems.Add($stateText)
    $item.Tag = @{ exists = [bool]$row.exists; path = [string]$row.dst_abs }
    $item.ToolTipText = [string]$row.dst_abs
    if ($row.state -eq "NEW") { $item.ForeColor = [System.Drawing.Color]::Gray }
    if ($row.state -eq "SKIP") { $item.ForeColor = [System.Drawing.Color]::DimGray }
    if ($row.state -eq "SELF") { $item.ForeColor = [System.Drawing.Color]::DimGray }
    if ($row.state -eq "RENAME") { $item.ForeColor = [System.Drawing.Color]::DarkGreen }
    if ($row.exists -and $row.override) { $item.Checked = $true }
    [void]$items.Add($item)
}
$view.BeginUpdate()
$view.Items.AddRange([System.Windows.Forms.ListViewItem[]]$items.ToArray())
$view.EndUpdate()

$lblSum = New-Object System.Windows.Forms.Label
$lblSum.Text = $plan.summary
$lblSum.Location = New-Object System.Drawing.Point(12, 440)
$lblSum.Size = New-Object System.Drawing.Size(936, 22)
$form.Controls.Add($lblSum)

$chkAll = New-Object System.Windows.Forms.CheckBox
$chkAll.Text = "[[ui.chk_all]]"
$chkAll.Location = New-Object System.Drawing.Point(12, 466)
$chkAll.Size = New-Object System.Drawing.Size(360, 22)
$form.Controls.Add($chkAll)
$chkAll.add_CheckedChanged({
    param($sender, $e)
    foreach ($item in $view.Items) {
        if ($item.Tag -and $item.Tag.exists) { $item.Checked = $chkAll.Checked }
    }
})

$btnBack = New-Object System.Windows.Forms.Button
$btnBack.Text = "[[ui.btn_back]]"
$btnBack.Location = New-Object System.Drawing.Point(520, 494)
$btnBack.Size = New-Object System.Drawing.Size(100, 32)
$form.Controls.Add($btnBack)
$btnBack.add_Click({
    param($sender, $e)
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Retry
    $form.Close()
})

# 显式的「取消」按钮：与关窗等价（返回取消，不导出、不写盘）
$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "[[ui.btn_cancel]]"
$btnCancel.Location = New-Object System.Drawing.Point(630, 494)
$btnCancel.Size = New-Object System.Drawing.Size(100, 32)
$form.Controls.Add($btnCancel)
$btnCancel.add_Click({
    param($sender, $e)
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Close()
})

$btnOK = New-Object System.Windows.Forms.Button
$btnOK.Font = New-Object System.Drawing.Font($uiFont, [System.Drawing.FontStyle]::Bold)
$btnOK.Text = "[[ui.btn_start]]"
$btnOK.Location = New-Object System.Drawing.Point(740, 494)
$btnOK.Size = New-Object System.Drawing.Size(138, 32)
$form.Controls.Add($btnOK)
$btnOK.add_Click({
    param($sender, $e)
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})

[void]$form.ShowDialog()

$action = "cancel"
if ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK) { $action = "ok" }
elseif ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::Retry) { $action = "back" }

$overrides = @()
if ($action -eq "ok") {
    foreach ($item in $view.Items) {
        if (-not ($item.Tag -and $item.Tag.exists)) { continue }
        $overrides += [ordered]@{ path = $item.Tag.path; override = [bool]$item.Checked }
    }
}

$payload = [ordered]@{ action = $action; overrides = $overrides }
$json = $payload | ConvertTo-Json -Depth 4 -Compress
[System.IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
$form.Dispose()
'''


SAVE_PS = r'''
param(
    [Parameter(Mandatory=$true)][string]$ItemsFile,
    [Parameter(Mandatory=$true)][string]$OutFile
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$items = Get-Content -Raw -Encoding UTF8 $ItemsFile | ConvertFrom-Json

$form = New-Object System.Windows.Forms.Form
$form.Text = "[[ui.save_title]]"
$form.ClientSize = New-Object System.Drawing.Size(760, 480)
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

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = "[[ui.save_hint]]"
$lbl.Location = New-Object System.Drawing.Point(12, 10)
$lbl.Size = New-Object System.Drawing.Size(736, 40)
$form.Controls.Add($lbl)

$panel = New-Object System.Windows.Forms.FlowLayoutPanel
$panel.Location = New-Object System.Drawing.Point(12, 54)
$panel.Size = New-Object System.Drawing.Size(736, 356)
$panel.AutoScroll = $true
$panel.FlowDirection = "TopDown"
$panel.WrapContents = $false
$form.Controls.Add($panel)

foreach ($item in $items.items) {
    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = "[[ui.save_item_prefix]]" + $item.name
    $chk.Tag = $item.path
    $chk.Size = New-Object System.Drawing.Size(700, 26)
    $chk.Checked = $false
    $chk.ToolTipText = $item.path
    $panel.Controls.Add($chk)
}

$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Text = "[[ui.save_note]]"
$lblNote.Location = New-Object System.Drawing.Point(12, 416)
$lblNote.Size = New-Object System.Drawing.Size(736, 20)
$form.Controls.Add($lblNote)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "[[ui.btn_cancel_export]]"
$btnCancel.Location = New-Object System.Drawing.Point(500, 442)
$btnCancel.Size = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnCancel)
$btnCancel.add_Click({
    param($sender, $e)
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Close()
})

$btnOK = New-Object System.Windows.Forms.Button
$btnOK.Font = New-Object System.Drawing.Font($uiFont, [System.Drawing.FontStyle]::Bold)
$btnOK.Text = "[[ui.btn_save_go]]"
$btnOK.Location = New-Object System.Drawing.Point(628, 442)
$btnOK.Size = New-Object System.Drawing.Size(120, 30)
$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $btnOK
$form.Controls.Add($btnOK)

[void]$form.ShowDialog()

$ok = ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK)
$saved = @()
if ($ok) {
    foreach ($ctrl in $panel.Controls) {
        if ($ctrl -is [System.Windows.Forms.CheckBox] -and $ctrl.Checked) {
            $saved += $ctrl.Tag
        }
    }
}

$payload = [ordered]@{ ok = $ok; save = $saved }
$json = $payload | ConvertTo-Json -Depth 4 -Compress
[System.IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
$form.Dispose()
'''
# 并发提示：短提示窗，几秒后自动关闭，避免 journal 结束后窗口一直挂着
BUSY_PS = r'''
param(
    [string]$Title = "",
    [string]$Text = "",
    [string]$OkText = "OK",
    [int]$Seconds = 8
)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = $Title
$form.ClientSize = New-Object System.Drawing.Size(420, 150)
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

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = $Text
$lbl.Location = New-Object System.Drawing.Point(16, 18)
$lbl.Size = New-Object System.Drawing.Size(388, 68)
$form.Controls.Add($lbl)

$btn = New-Object System.Windows.Forms.Button
$btn.Font = New-Object System.Drawing.Font($uiFont, [System.Drawing.FontStyle]::Bold)
$btn.Text = $OkText
$btn.Location = New-Object System.Drawing.Point(150, 104)
$btn.Size = New-Object System.Drawing.Size(120, 30)
$btn.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $btn
$form.Controls.Add($btn)

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = [Math]::Max(1, $Seconds) * 1000
$timer.add_Tick({
    try { $timer.Stop() } catch { }
    $form.Close()
})
$timer.Start()

[void]$form.ShowDialog()
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


def run_powershell(script_text, arguments, tag):
    """把 PowerShell 脚本写到临时目录并执行，返回生成的结果 JSON（字典）"""
    import json
    import tempfile
    import subprocess

    temp_dir = tempfile.gettempdir()
    _rotate_log(os.path.join(temp_dir, "nx_step_export_ui_errors.log"))
    ps_path = os.path.join(temp_dir, "nx_step_export_%s.ps1" % tag)
    out_path = os.path.join(temp_dir, "nx_step_export_%s.json" % tag)

    with open(ps_path, "w", encoding="utf-8-sig") as handle:
        handle.write(render_ps(script_text))
    if os.path.exists(out_path):
        os.remove(out_path)

    powershell = _powershell_path()
    if powershell is None:
        raise RuntimeError(T("log.err_no_powershell"))

    command = [powershell, "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass",
               "-File", ps_path] + list(arguments) + ["-OutFile", out_path]

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
        raise RuntimeError(T("log.err_no_result", detail))

    with open(out_path, "r", encoding="utf-8-sig") as handle:
        return json.load(handle)


def _rotate_log(path, max_bytes=2 * 1024 * 1024, keep=3):
    """日志超过上限时轮转：当前 -> .1，旧的依次后移，最多保留 keep 份"""
    try:
        if (not os.path.exists(path)) or os.path.getsize(path) < max_bytes:
            return
        oldest = "%s.%d" % (path, keep)
        if os.path.exists(oldest):
            force_remove(oldest)
        for index in range(keep - 1, 0, -1):
            src = "%s.%d" % (path, index)
            if os.path.exists(src):
                os.replace(src, "%s.%d" % (path, index + 1))
        os.replace(path, path + ".1")
    except Exception:
        pass


def _prune_backups(step_path, keep=3):
    """只保留该目标文件最近 keep 份“时间戳”备份；不动旧格式 .bak 与用户文件"""
    import re
    try:
        folder = os.path.dirname(step_path)
        base = os.path.basename(step_path)
        pattern = re.compile(r"^%s\.\d{8}_\d{6}(_\d+)?\.bak$" % re.escape(base))
        items = []
        for name in os.listdir(folder):
            if not pattern.match(name):
                continue
            full = os.path.join(folder, name)
            try:
                items.append((os.path.getmtime(full), full))
            except Exception:
                pass
        items.sort(reverse=True)
        for _mtime, path in items[keep:]:
            try:
                force_remove(path)
            except Exception:
                pass
    except Exception:
        pass


def _make_backup(step_path):
    """把旧 .step 改名成带时间戳的备份（同秒冲突加 _2/_3），并只保留最近 3 份"""
    stamp = time.strftime("%Y%m%d_%H%M%S")
    backup = "%s.%s.bak" % (step_path, stamp)
    suffix = 2
    while os.path.exists(backup):
        backup = "%s.%s_%d.bak" % (step_path, stamp, suffix)
        suffix += 1
    os.replace(step_path, backup)
    _prune_backups(step_path, 3)
    return backup


def spawn_powershell(script_text, arguments, tag):
    """启动 PowerShell 但不等待退出（用于短提示窗：journal 结束窗口仍在，定时自动关闭）"""
    import tempfile
    import subprocess

    temp_dir = tempfile.gettempdir()
    _rotate_log(os.path.join(temp_dir, "nx_step_export_ui_errors.log"))
    ps_path = os.path.join(temp_dir, "nx_step_export_%s.ps1" % tag)
    with open(ps_path, "w", encoding="utf-8-sig") as handle:
        handle.write(render_ps(script_text))
    powershell = _powershell_path()
    if powershell is None:
        return False
    command = [powershell, "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass",
               "-File", ps_path] + list(arguments)
    try:
        process = subprocess.Popen(command, stdin=subprocess.DEVNULL,
                                   stdout=subprocess.DEVNULL,
                                   stderr=subprocess.DEVNULL,
                                   creationflags=0x08000000)  # CREATE_NO_WINDOW
    except TypeError:
        process = subprocess.Popen(command, stdin=subprocess.DEVNULL,
                                   stdout=subprocess.DEVNULL,
                                   stderr=subprocess.DEVNULL)
    return True


def notify_busy():
    """并发冲突时在 NX 里给一个短提示（8 秒自动关，可手动点确定）；失败静默"""
    try:
        spawn_powershell(BUSY_PS, [
            "-Title", T("ui.busy_title"),
            "-Text", T("ui.busy_text"),
            "-OkText", T("ui.busy_ok"),
            "-Seconds", "8",
        ], "busy")
    except Exception:
        pass


def preview_plan(plan, base_dir):
    """导出前弹出清单确认窗，返回 (action, overrides)

    plan 的每一项是 (源文件, 输出路径, 行状态)，状态在 build_plan 里按同名策略算好：
    NEW / SKIP / RENAME / OVERWRITE。

    action: "ok" / "back" / "cancel"
    overrides: {规范化绝对输出路径: True/False}，只包含窗口里可勾选的行（目标已存在）；
               没出现在字典里的行沿用该行状态。
    """
    import json
    import tempfile

    rows = []
    new_count = 0
    skip_count = 0
    for index, (prt_path, out_path, state) in enumerate(plan, 1):
        exists = os.path.exists(out_path)
        if state in ("SKIP", "SELF"):
            skip_count += 1
        else:
            new_count += 1
        rows.append({
            "no": index,
            "src": rel_display(prt_path, base_dir),
            # 预览里显示的是拼接前缀/后缀、并按同名策略算好的最终文件名
            "dst": rel_display(out_path, base_dir),
            "dst_abs": os.path.abspath(out_path),
            "state": state,
            "exists": exists,
            # 初始勾选状态：只有“覆盖”策略的行默认勾上，其余保持不勾
            "override": bool(exists and state == "OVERWRITE"),
        })

    summary = T("ui.summary", len(rows), new_count, skip_count)

    plan_path = os.path.join(tempfile.gettempdir(), "nx_step_export_plan.json")
    with open(plan_path, "w", encoding="utf-8-sig") as handle:
        json.dump({"rows": rows, "summary": summary}, handle, ensure_ascii=False)

    data = run_powershell(PREVIEW_PS, ["-PlanFile", plan_path], "preview") or {}
    overrides = {}
    for item in data.get("overrides") or []:
        try:
            key = os.path.normcase(os.path.abspath(item.get("path")))
            overrides[key] = bool(item.get("override"))
        except Exception:
            continue
    return data.get("action", "cancel"), overrides


def collect_modified_open(session, plan):
    """返回计划里“已在当前会话打开且 IsModified=True”的项

    判定只看内存中 Part.IsModified，不看文件时间戳：
    用户可能只改了图层、显示状态，时间戳没变，但导出前仍应让用户决定是否保存。
    """
    open_parts = {}
    try:
        for part in session.Parts:
            try:
                full_path = part.FullPath
            except Exception:
                continue
            if full_path:
                open_parts[os.path.normcase(os.path.abspath(full_path))] = part
    except Exception:
        return []

    items = []
    seen = set()
    for prt_path, _out_path, _state in plan:
        key = os.path.normcase(os.path.abspath(prt_path))
        if key in seen or key not in open_parts:
            continue
        part = open_parts[key]
        try:
            modified = bool(part.IsModified)
        except Exception:
            modified = False
        if modified:
            seen.add(key)
            items.append({"path": os.path.abspath(prt_path),
                          "name": os.path.basename(prt_path),
                          "part": part})
    return items


def ask_save_modified(items):
    """弹一次窗口列出所有“已打开且有未保存修改”的文件，返回 (是否继续, 保存集合)

    默认全部不勾选（= 跳过、不写盘）；用户勾选的文件才保存。
    关闭窗口按“取消本次导出”处理。
    """
    import json
    import tempfile

    payload = {"items": [{"path": item["path"], "name": item["name"]}
                         for item in items]}
    # 注意：这里必须是独立文件名。run_powershell 启动前会删除“输出文件”
    # （%TEMP%\nx_step_export_unsaved.json），若输入清单也叫这个名字，
    # 会在 PowerShell 读取前被删掉，对话框就一个勾选框都建不出来，
    # 用户勾了“保存”也只会返回空列表，进而被当成“跳过”。
    items_path = os.path.join(tempfile.gettempdir(),
                              "nx_step_export_unsaved_items.json")
    with open(items_path, "w", encoding="utf-8-sig") as handle:
        json.dump(payload, handle, ensure_ascii=False)

    data = run_powershell(SAVE_PS, ["-ItemsFile", items_path], "unsaved") or {}
    if not data.get("ok"):
        return False, set()
    save_keys = set()
    for path in data.get("save") or []:
        try:
            save_keys.add(os.path.normcase(os.path.abspath(path)))
        except Exception:
            continue
    return True, save_keys


def parse_picker_options(payload, previous=None):
    """把选择窗口回传的导出选项规范化；缺字段或非法值沿用上次值/默认值"""
    result = dict(previous or default_options())
    if not isinstance(payload, dict):
        return result
    mode = str(payload.get("out_mode") or "").strip().lower()
    if mode in ("source", "dir"):
        result["out_mode"] = mode
    result["out_dir"] = str(payload.get("out_dir") or "").strip()
    result["prefix"] = str(payload.get("prefix") or "")
    result["suffix"] = str(payload.get("suffix") or "")
    policy = str(payload.get("on_exists") or "").strip().lower()
    if policy in ("skip", "rename", "overwrite"):
        result["on_exists"] = policy
    fmt = str(payload.get("format") or "").strip().lower()
    if fmt in ("step", "iges"):
        result["format"] = fmt
    return result


def choose_targets(base_dir=None, options=None):
    """弹出 PowerShell 选择窗口（默认从「此电脑」+ 所有盘符开始）

    返回 (folders, files, options)

    folders: None = 没选文件夹；[绝对路径...] = 勾选的文件夹
    files:   绝对路径列表
    options: 导出选项字典（out_mode / out_dir / prefix / suffix / on_exists / format）
    """
    opts = dict(default_options())
    opts.update(read_options())
    opts.update(options or {})
    args = []
    if base_dir:
        args += ["-Root", base_dir]
    # 把上次的选项回传，用户点「返回修改」时窗口里保留上次填写的内容
    # 注意：PowerShell 会丢弃空字符串参数（会报 “Missing an argument”），
    # 所以值为空时干脆不传，让脚本里的默认值（空串）生效。
    args += ["-OutMode", opts["out_mode"],
             "-OnExists", opts["on_exists"],
             "-Format", opts["format"]]
    for flag, key in (("-OutDir", "out_dir"),
                      ("-Prefix", "prefix"),
                      ("-Suffix", "suffix")):
        value = str(opts.get(key) or "")
        if value:
            args += [flag, value]
    data = run_powershell(SELECT_UI_PS, args, "picker")

    if not data or not data.get("ok"):
        return None, [], opts

    folders = data.get("folders") or []
    files = data.get("files") or []
    if isinstance(folders, str):
        folders = [folders] if folders else []
    if isinstance(files, str):
        files = [files] if files else []

    chosen_dirs = []
    for folder in folders:
        path = os.path.abspath(folder)
        if os.path.isdir(path):
            chosen_dirs.append(path)
        else:
            log(T("log.ignore_outside", folder))

    options = parse_picker_options(data, opts)
    # 确认后记住本次选项，并把「最后一个目标文件所在目录」作为下次展开位置
    # （没有文件时退回最后一个文件夹）；取消时上面已提前返回，不会写
    write_options(options)
    picked_files = [os.path.abspath(p) for p in files if p]
    if picked_files:
        last_dir = os.path.dirname(picked_files[-1])
    elif chosen_dirs:
        last_dir = chosen_dirs[-1]
    else:
        last_dir = ""
    if last_dir:
        write_last_dir(last_dir)
    if chosen_dirs:
        return chosen_dirs, files, options
    return None, files, options


def rel_display(path, base_dir=None):
    """日志里显示的路径：有根目录时相对显示，否则显示绝对路径"""
    if base_dir:
        try:
            return os.path.relpath(path, base_dir)
        except ValueError:
            return path
    return os.path.abspath(path) if path else path


def pick_folder(title):
    """弹出 Windows 文件夹选择框，返回所选路径；用户取消返回 None"""
    import ctypes
    from ctypes import wintypes

    class BROWSEINFOW(ctypes.Structure):
        _fields_ = [
            ("hwndOwner", wintypes.HWND),
            ("pidlRoot", ctypes.c_void_p),
            ("pszDisplayName", wintypes.LPWSTR),
            ("lpszTitle", wintypes.LPCWSTR),
            ("ulFlags", wintypes.UINT),
            ("lpfn", ctypes.c_void_p),
            ("lParam", ctypes.c_void_p),
            ("iImage", ctypes.c_int),
        ]

    shell32 = ctypes.windll.shell32
    user32 = ctypes.windll.user32
    shell32.SHBrowseForFolderW.restype = ctypes.c_void_p
    shell32.SHBrowseForFolderW.argtypes = [ctypes.POINTER(BROWSEINFOW)]
    shell32.SHGetPathFromIDListW.argtypes = [ctypes.c_void_p, wintypes.LPWSTR]

    display_name = ctypes.create_unicode_buffer(260)
    info = BROWSEINFOW()
    info.hwndOwner = user32.GetForegroundWindow()
    info.pszDisplayName = ctypes.cast(display_name, wintypes.LPWSTR)
    info.lpszTitle = title
    info.ulFlags = 0x00000001 | 0x00000040   # 只列文件系统目录 + 新式可调整对话框

    pidl = shell32.SHBrowseForFolderW(ctypes.byref(info))
    if not pidl:
        return None
    try:
        path_buffer = ctypes.create_unicode_buffer(260)
        if not shell32.SHGetPathFromIDListW(pidl, path_buffer):
            return None
        return path_buffer.value
    finally:
        ctypes.windll.ole32.CoTaskMemFree(ctypes.c_void_p(pidl))


def pick_files(title, initial_dir):
    """弹出 Windows 多选文件对话框，返回选中的文件绝对路径列表；取消返回 []"""
    import ctypes
    from ctypes import wintypes

    class OPENFILENAMEW(ctypes.Structure):
        _fields_ = [
            ("lStructSize", wintypes.DWORD),
            ("hwndOwner", wintypes.HWND),
            ("hInstance", wintypes.HINSTANCE),
            ("lpstrFilter", wintypes.LPCWSTR),
            ("lpstrCustomFilter", wintypes.LPWSTR),
            ("nMaxCustFilter", wintypes.DWORD),
            ("nFilterIndex", wintypes.DWORD),
            ("lpstrFile", wintypes.LPWSTR),
            ("nMaxFile", wintypes.DWORD),
            ("lpstrFileTitle", wintypes.LPWSTR),
            ("nMaxFileTitle", wintypes.DWORD),
            ("lpstrInitialDir", wintypes.LPCWSTR),
            ("lpstrTitle", wintypes.LPCWSTR),
            ("Flags", wintypes.DWORD),
            ("nFileOffset", wintypes.WORD),
            ("nFileExtension", wintypes.WORD),
            ("lpstrDefExt", wintypes.LPCWSTR),
            ("lCustData", wintypes.LPARAM),
            ("lpfnHook", ctypes.c_void_p),
            ("lpTemplateName", wintypes.LPCWSTR),
            ("pvReserved", ctypes.c_void_p),
            ("dwReserved", wintypes.DWORD),
            ("FlagsEx", wintypes.DWORD),
        ]

    buffer_size = 65536
    file_buffer = ctypes.create_unicode_buffer(buffer_size)
    dialog = OPENFILENAMEW()
    dialog.lStructSize = ctypes.sizeof(OPENFILENAMEW)
    dialog.hwndOwner = ctypes.windll.user32.GetForegroundWindow()
    dialog.lpstrFilter = u"%s\0*.prt;*.stp;*.step;*.igs;*.iges\0%s\0*.*\0\0" % (
        T("ui.filter_prt"), T("ui.filter_all"))
    dialog.lpstrFile = ctypes.cast(file_buffer, wintypes.LPWSTR)
    dialog.nMaxFile = buffer_size
    dialog.lpstrTitle = title
    dialog.lpstrInitialDir = initial_dir
    dialog.lpstrDefExt = u"prt"
    # 允许多选 + 资源管理器风格 + 文件必须存在 + 隐藏只读勾选
    dialog.Flags = 0x00000200 | 0x00080000 | 0x00001000 | 0x00000800 | 0x00000004

    if not ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(dialog)):
        return []

    # 单选时返回完整路径；多选时返回 "目录\0文件1\0文件2\0\0"
    raw = ctypes.wstring_at(ctypes.addressof(file_buffer), buffer_size)
    parts = [p for p in raw.split(u"\0") if p]
    if not parts:
        return []
    if len(parts) == 1:
        return [parts[0]]
    folder = parts[0]
    return [os.path.join(folder, name) for name in parts[1:]]


def to_rel_dir(folder, base_dir=None):
    """无根目录时返回绝对路径；有根目录时返回相对路径，超出根目录返回 None"""
    folder = os.path.abspath(folder)
    if not base_dir:
        return folder
    try:
        rel = os.path.relpath(folder, base_dir)
    except ValueError:
        return None
    if rel == ".." or rel.startswith(".." + os.sep):
        return None
    return rel


def select_folders(base_dir=None):
    """弹窗逐个选择文件夹（系统对话框，仅在选择窗口打不开时用），取消即结束

    返回绝对路径列表；一个都没选则返回 None
    """
    import ctypes
    ole32 = ctypes.windll.ole32
    com_ready = ole32.CoInitialize(None) in (0, 1)   # S_OK / S_FALSE

    chosen = []
    try:
        while True:
            folder = pick_folder(T("ui.folder_pick_title"))
            if not folder:
                break

            rel = to_rel_dir(folder, base_dir)
            if rel is None:
                log(T("log.ignore_outside", folder))
                continue

            if rel != ".":
                # 已选目录的子目录不再重复添加
                if any(rel == c or rel.startswith(c + os.sep) for c in chosen):
                    log(T("log.already_included", rel))
                    continue
                # 选了上级目录，把此前选的子目录收掉
                chosen = [c for c in chosen if not c.startswith(rel + os.sep)]

            if rel in chosen:
                continue
            chosen.append(rel)
            log(T("log.selected", T("log.root_all") if rel == "." else rel))
    finally:
        if com_ready:
            ole32.CoUninitialize()

    if not chosen:
        return None
    return chosen


def match_any(text, patterns):
    """text 命中任意一个通配符规则即返回 True"""
    text = text.lower().replace("/", os.sep)
    for pattern in patterns:
        if fnmatch.fnmatch(text, pattern.lower().replace("/", os.sep)):
            return True
    return False


def dir_match(rel_dir, patterns):
    """rel_dir 自身、它的任一路径层级、或其中任一文件夹名命中规则即算匹配

    例如 INCLUDE_DIRS=["box1"] 能覆盖 box1、box1/sub；
    EXCLUDE_DIRS=["old*"] 能排除根下的 old 和 box2/old_data。
    """
    if not patterns:
        return False
    candidates = ["."]
    if rel_dir not in ("", "."):
        parts = rel_dir.split(os.sep)
        for i in range(len(parts)):
            candidates.append(os.sep.join(parts[:i + 1]))
        candidates.extend(parts)   # 单独文件夹名也算，便于匹配任意层级的同名目录
    return any(match_any(c, patterns) for c in candidates)


# 链式转换的源扩展名：非 .prt 源在“批量导出”里自动走“内存导入 → 直接导出目标格式”
CHAIN_SOURCE_EXTS = (".stp", ".step", ".igs", ".iges")
# 文件对话框/目录列表可选的源类型（.prt 走原有导出，其余走链式转换）
SELECTABLE_EXTS = (".prt",) + CHAIN_SOURCE_EXTS


def find_prt_files(base_dir=None, include_dirs=None, picked_files=None, scan_dirs=True):
    """收集 .prt 并按 EXCLUDE_DIRS / INCLUDE_FILES / EXCLUDE_FILES 筛选

    base_dir     : 已废弃，保留参数只为兼容旧调用；现在为 None
    include_dirs : None = 用配置里的 INCLUDE_DIRS；[绝对路径] = 只收这些目录（含子目录）
    picked_files : 手动点选的文件（绝对路径），不受文件夹/文件名规则限制
    scan_dirs    : False = 不遍历目录，只导出 picked_files
    """
    include_dirs = list(INCLUDE_DIRS if include_dirs is None else include_dirs)
    picked_files = [os.path.abspath(p) for p in (picked_files or [])]

    def _walk_error(err):
        # 路径过长/权限不足时 os.walk 会跳过该目录，这里显式记一条，避免静默漏导
        try:
            log(T("log.walk_fail", getattr(err, "filename", "") or ""))
        except Exception:
            pass

    found = []
    if not scan_dirs:
        pass   # 只导手动点选的文件，不扫目录
    else:
        roots = include_dirs if include_dirs else ([base_dir] if base_dir else [])
        for root_dir in roots:
            if RECURSIVE:
                for root, _dirs, names in os.walk(root_dir, onerror=_walk_error):
                    for name in names:
                        if name.lower().endswith(".prt"):
                            found.append(os.path.join(root, name))
            else:
                for path in glob.glob(os.path.join(root_dir, "*.prt")):
                    if os.path.isfile(path):
                        found.append(path)

    selected = []
    for path in found:
        rel_path = os.path.relpath(path, base_dir) if base_dir else os.path.abspath(path)
        rel_dir = os.path.dirname(rel_path)
        file_name = os.path.basename(path)

        # 文件夹过滤：排除优先，其次必须落在包含范围内（空 = 不限）
        if dir_match(rel_dir, EXCLUDE_DIRS):
            continue
        if include_dirs and not dir_match(rel_dir, include_dirs):
            continue

        # 文件过滤：文件名或相对路径命中即可
        if EXCLUDE_FILES and (match_any(file_name, EXCLUDE_FILES)
                              or match_any(rel_path, EXCLUDE_FILES)):
            continue
        if INCLUDE_FILES and not (match_any(file_name, INCLUDE_FILES)
                                  or match_any(rel_path, INCLUDE_FILES)):
            continue

        selected.append(path)

    # 手动点选的具体文件优先，不受文件夹/文件名规则限制
    # 放开 .stp/.igs 是为了链式转换；目录扫描仍然只收 .prt
    for path in picked_files:
        if path.lower().endswith(SELECTABLE_EXTS) and os.path.isfile(path) \
                and path not in selected:
            selected.append(path)

    selected.sort(key=lambda p: p.lower())
    return selected


def build_plan(files, options=None):
    """生成 (源文件, 输出路径, 行状态) 列表

    - 输出目录：options["out_mode"]=="dir" 时统一落到 options["out_dir"]（扁平放置，
      不镜像原目录结构）；否则仍是各 .prt 所在目录（默认，与改动前一致）
    - 文件名：前缀 + 原名 + 后缀 + 扩展名（STEP -> .step，IGES -> .igs）
    - 同名策略 options["on_exists"]：
        skip      = 已存在有效产物则标记 SKIP（只有文件头、没有几何的 STEP 不算已存在）
        rename    = 已存在就追加 _2/_3…直到不冲突，标记 RENAME
        overwrite = 照原名输出，标记 OVERWRITE（导出时仍先备份旧文件）
    - 同一批次内两件原始文件撞到同一输出名时，仍按原逻辑追加 _2/_3（标记 RENAME）
    """
    options = options or {}
    ext = output_ext(options)
    out_mode = options.get("out_mode") or "source"
    out_dir = options.get("out_dir") or ""
    prefix = safe_name_part(options.get("prefix"))
    suffix = safe_name_part(options.get("suffix"))
    on_exists = options.get("on_exists") or "skip"

    plan = []
    used = {}   # (目录小写, 输出名小写) -> 来源文件名
    for prt_path in files:
        src_dir = os.path.dirname(prt_path)
        target_dir = out_dir if (out_mode == "dir" and out_dir) else src_dir
        src_name = os.path.basename(prt_path)
        src_ext = os.path.splitext(src_name)[1].lower()
        if src_ext == ext:
            # 源与目标同格式（如 stp→stp）：跳过，否则“源目录 + overwrite”会把源文件覆盖掉
            log(T("log.chain_same_fmt", src_name))
            plan.append((prt_path, prt_path, "SELF"))
            continue
        stem = normalize_name(os.path.splitext(src_name)[0])
        base_name = prefix + stem + suffix

        out_name = base_name + ext
        state = "NEW"
        key = (target_dir.lower(), out_name.lower())
        if key in used:
            # 同一批次内的撞名：沿用原来的“追加 _2/_3”，避免同目录内互相覆盖
            serial = 2
            while (target_dir.lower(),
                   ("%s_%d%s" % (base_name, serial, ext)).lower()) in used:
                serial += 1
            new_name = "%s_%d%s" % (base_name, serial, ext)
            log(T("log.rename_conflict", src_name, used[key], out_name, new_name))
            out_name = new_name
            state = "RENAME"
        elif os.path.exists(os.path.join(target_dir, out_name)):
            target_path = os.path.join(target_dir, out_name)
            if on_exists == "rename":
                serial = 2
                while True:
                    candidate = "%s_%d%s" % (base_name, serial, ext)
                    if (target_dir.lower(), candidate.lower()) not in used and \
                            not os.path.exists(os.path.join(target_dir, candidate)):
                        break
                    serial += 1
                log(T("log.rename_policy", out_name, candidate))
                out_name = candidate
                state = "RENAME"
            elif on_exists == "overwrite":
                state = "OVERWRITE"
            elif output_is_valid(target_path, options):
                state = "SKIP"

        used[(target_dir.lower(), out_name.lower())] = src_name
        plan.append((prt_path, os.path.join(target_dir, out_name), state))
    return plan


def create_step_creator(session):
    """按 NX 版本依次尝试创建 STEP 导出器"""
    dm = session.DexManager
    errors = []
    for method_name in ("CreateStep214Creator", "CreateStepCreator"):
        factory = getattr(dm, method_name, None)
        if factory is None:
            errors.append(T("log.err_method_missing", method_name))
            continue
        try:
            return factory()
        except Exception as e:
            errors.append(T("log.err_method_fail", method_name, e))
    raise RuntimeError(T("log.err_creator", "; ".join(errors)))


def configure_step_creator(creator, prt_path, step_path, layer_mask, source_part=None):
    """把导出器配置成“从磁盘现有部件导出当前可见图层的几何”

    NX 2506 的导出器默认既不指定对象类型、也不指定图层范围，
    若沿用默认值，生成的 STEP 只有文件头、没有几何，所以这里全部显式设置。
    source_part 非空时改为“从显示/工作部件导出”，供链式转换使用内存部件（无磁盘路径）。
    """
    creator.OutputFile = step_path

    # 从磁盘读取，而不是从“显示部件”读取：导出内容不依赖当前显示状态
    option = getattr(type(creator), "ExportFromOption", None)
    if source_part is None:
        creator.InputFile = prt_path
        if option is not None and hasattr(option, "ExistingPart"):
            try:
                creator.ExportFrom = option.ExistingPart
            except Exception as e:
                log(T("log.export_from_fail", e))
    else:
        applied = False
        for flag in ("DisplayPart", "WorkPart"):
            value = getattr(option, flag, None) if option is not None else None
            if value is None:
                continue
            try:
                creator.ExportFrom = value
                applied = True
                break
            except Exception as e:
                log(T("log.export_from_fail", e))
        if not applied:
            raise RuntimeError(T("log.err_step_from_part"))

    # 对象类型：实体 + 曲面 + 曲线（与 STEP 转换器模块掩码一致）
    types = getattr(creator, "ObjectTypes", None)
    if types is not None:
        try:
            for flag in ("Solids", "Surfaces", "Curves"):
                if hasattr(type(types), flag):
                    setattr(types, flag, True)
        except Exception as e:
            log(T("log.objecttypes_fail", e))

    # 图层范围：只写当前视图里可见的层；空掩码等于不导任何对象，调用方会先拦截
    if hasattr(type(creator), "LayerMask"):
        try:
            creator.LayerMask = layer_mask
        except Exception as e:
            log(T("log.layermask_fail", e))

    # 等待转换真正结束再返回，否则可能读到还没写完的文件
    if hasattr(type(creator), "ProcessHoldFlag"):
        try:
            creator.ProcessHoldFlag = True
        except Exception as e:
            log(T("log.holdflag_fail", e))


def step_size(step_path):
    """返回 STEP 文件大小；不存在返回 -1"""
    try:
        return os.path.getsize(step_path) if os.path.exists(step_path) else -1
    except Exception:
        return -1


def step_has_geometry(step_path):
    """判断 STEP 是否真的含有几何（而不只是文件头 + 产品结构）

    只输出文件头的 STEP 也有 2KB 左右、大小非 0，单看文件大小会误判成功。
    这里统计几类几何实体；只有 1 个 CARTESIAN_POINT 的情况属于“空壳”，判为无效。
    """
    try:
        faces = edges = points = 0
        with open(step_path, "rb") as handle:
            while True:
                chunk = handle.read(1048576)
                if not chunk:
                    break
                faces += chunk.count(b"ADVANCED_FACE")
                edges += chunk.count(b"EDGE_CURVE")
                points += chunk.count(b"CARTESIAN_POINT")
        return faces > 0 or edges > 0 or points > 1
    except Exception:
        return False


def step_is_valid(step_path):
    """已存在的 STEP 是否可以直接跳过：必须非空且确实含几何"""
    return step_size(step_path) > 0 and step_has_geometry(step_path)


def format_layer_mask(layers):
    """把层号列表压缩成 1-5,8,10-12 形式；无可见层返回空串"""
    if not layers:
        return ""
    layers = sorted(set(layers))
    result = []
    start = prev = layers[0]
    for number in layers[1:]:
        if number == prev + 1:
            prev = number
            continue
        result.append("%d" % start if start == prev else "%d-%d" % (start, prev))
        start = prev = number
    result.append("%d" % start if start == prev else "%d-%d" % (start, prev))
    return ",".join(result)


def find_open_part(session, path):
    """在会话里按完整路径查找已打开的部件；找不到返回 None"""
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


def _part_key(part):
    try:
        return os.path.normcase(os.path.abspath(part.FullPath))
    except Exception:
        return None


def _session_parts(session):
    try:
        return list(session.Parts)
    except Exception:
        return []


def ensure_displayed(session, prt_path, opened_parts):
    """确保部件已加载并成为显示部件（图层可见性查询要求显示部件）

    返回显示部件；只有脚本自己新打开的部件（含装配加载出来的组件）才记入 opened_parts，
    调用方在每件导出后立即关闭这些部件。
    """
    before = set(k for k in (_part_key(p) for p in _session_parts(session)) if k)
    part = find_open_part(session, prt_path)
    try:
        display = session.Parts.Display
        display_path = display.FullPath if display else None
    except Exception:
        display_path = None

    if part is None:
        if not os.path.isfile(prt_path):
            raise RuntimeError(T("log.missing_part", prt_path))
        part, _status = session.Parts.OpenBaseDisplay(prt_path)
        if part is None:
            # 文件被其他进程独占、不可读等情况下 NX 可能返回 None 而不是抛异常
            raise RuntimeError(T("log.err_open_part", prt_path))
    elif not display_path or os.path.normcase(os.path.abspath(display_path)) != \
            os.path.normcase(os.path.abspath(prt_path)):
        # 已打开但不是显示部件：用“切换活动显示部件”而不是再次打开文件，
        # 后者对已打开部件会被 NX 拒绝（文件已存在）
        session.Parts.SetActiveDisplay(
            part,
            NXOpen.DisplayPartOption.ReplaceExisting,
            NXOpen.PartDisplayPartWorkPartOption.SameAsDisplay)

    # 记录本次新加载的部件（装配的组件也会被 OpenBaseDisplay 一起加载）
    for candidate in _session_parts(session):
        key = _part_key(candidate)
        if key and key not in before and candidate not in opened_parts:
            opened_parts.append(candidate)

    display = session.Parts.Display
    if display is None:
        raise RuntimeError(T("log.err_display_none", prt_path))
    return display


def _is_assembly(part):
    try:
        return part.ComponentAssembly.RootComponent is not None
    except Exception:
        return False


def _restore_state(session, display_part, work_part):
    """把显示/工作部件切回导出前的状态（用于关闭临时部件前）"""
    if display_part is not None:
        try:
            session.Parts.SetActiveDisplay(
                display_part,
                NXOpen.DisplayPartOption.ReplaceExisting,
                NXOpen.PartDisplayPartWorkPartOption.SameAsDisplay)
        except Exception:
            pass
    if work_part is not None:
        try:
            session.Parts.SetWork(work_part)
        except Exception:
            pass


def close_opened_parts(session, parts, original_keys):
    """关闭脚本自己打开、且无未保存修改的部件：先组件后装配，逐个释放"""
    ordered = sorted(parts, key=lambda p: 1 if _is_assembly(p) else 0)
    for part in ordered:
        key = _part_key(part)
        if key and key in original_keys:
            continue
        try:
            if bool(part.IsModified):
                log(T("log.keep_modified", key or part.FullPath))
                continue
        except Exception:
            pass
        try:
            part.Close(NXOpen.BasePart.CloseWholeTree.FalseValue,
                       NXOpen.BasePart.CloseModified.CloseModified, None)
        except Exception as e:
            log(T("log.close_fail", e))


def visible_layer_mask(part):
    """按当前工作视图的图层可见性生成 LayerMask

    采用 NX 的“视图内可见”判定（LayerManager.GetVisibilitiesInView），
    不是“图层可显示/可隐藏”的全局设置。工作层永远可见，用它的状态值作为
    “可见”的参照，避免依赖枚举名字（不同语言/版本命名可能不同）。
    """
    view = part.Views.WorkView
    infos = part.Layers.GetVisibilitiesInView(view)

    visible_value = None
    try:
        work_layer = int(part.Layers.WorkLayer)
    except Exception:
        work_layer = None
    if work_layer and 1 <= work_layer <= len(infos):
        visible_value = getattr(infos[work_layer - 1].State, "value", None)

    layers = []
    for index, info in enumerate(infos):
        state = getattr(info.State, "value", None)
        if visible_value is None or state == visible_value:
            layers.append(index + 1)
    return format_layer_mask(layers), view.Name, len(layers)


OUTPUT_NAME_EXTS = (".step", ".stp", ".igs", ".iges")


def output_candidates(out_path):
    """同一主干的其它扩展名写法：导出器可能把 .stp 规范化成 .step、把 .igs 规范化成 .iges"""
    stem = os.path.splitext(out_path)[0]
    names = []
    for ext in OUTPUT_NAME_EXTS:
        candidate = stem + ext
        if os.path.normcase(candidate) != os.path.normcase(out_path):
            names.append(candidate)
    return names


def snapshot_candidates(out_path):
    """导出前记录已存在的同名其它扩展名文件，避免把用户原有文件当成本次产物"""
    return set(os.path.normcase(p) for p in output_candidates(out_path) if os.path.exists(p))


def reclaim_output(out_path, before):
    """保证预览名 == 落盘名：导出器若换了扩展名写法，就把本次新产物改回计划里的名字"""
    if os.path.exists(out_path):
        return out_path
    for candidate in output_candidates(out_path):
        if os.path.normcase(candidate) in before or not os.path.exists(candidate):
            continue
        try:
            os.replace(candidate, out_path)
            log("[name] %s -> %s" % (os.path.basename(candidate), os.path.basename(out_path)))
        except Exception as e:
            log("[name] %s: %s" % (candidate, e))
        return out_path
    return out_path


def export_step(session, prt_path, step_path, layer_mask, source_part=None):
    """导出单个文件；产出为空或无几何时重试一次，仍无效则报错

    source_part 非空时从该内存部件导出（链式转换），此时 prt_path 传空串。
    """
    last_error = ""
    before = snapshot_candidates(step_path)
    for attempt in (1, 2):
        if attempt == 2:
            # 先清掉上次留下的无效文件（含导出器换了扩展名写出的），避免转换器因“目标已存在”而拒绝覆盖
            stale = [step_path] + [p for p in output_candidates(step_path)
                                   if os.path.normcase(p) not in before]
            for path in stale:
                try:
                    if os.path.exists(path):
                        os.remove(path)
                        log(T("log.retry_deleted"))
                except Exception as e:
                    log(T("log.retry_delete_fail", e))
        creator = create_step_creator(session)
        try:
            configure_step_creator(creator, prt_path, step_path, layer_mask, source_part)
            creator.Commit()
        finally:
            try:
                creator.Destroy()
            except Exception:
                pass

        reclaim_output(step_path, before)
        size = step_size(step_path)
        if size > 0 and step_has_geometry(step_path):
            if attempt == 2:
                log(T("log.retry2_ok", size))
            return size

        if size < 0:
            last_error = T("log.err_out_missing")
        elif size == 0:
            last_error = T("log.err_out_empty")
        else:
            last_error = T("log.err_out_header", size)
        log(T("log.invalid_attempt", attempt, last_error))
        time.sleep(0.4)

    # 无效产物不留在原目录，否则下次会被“已存在”规则当成已完成
    try:
        if os.path.exists(step_path):
            os.remove(step_path)
    except Exception as e:
        log(T("log.cleanup_fail", e))

    raise RuntimeError(T("log.err_export_invalid", last_error))


def _iges_tool_path(filename):
    """在 IGES 安装目录里找 iges.exe / igesexport.def；找不到返回 None"""
    for env_name in ("IGESUG_DIR", "UGII_IGES_DIR"):
        base = (os.environ.get(env_name) or "").strip()
        if not base:
            continue
        candidate = os.path.join(base, filename)
        if os.path.isfile(candidate):
            return candidate
    base = (os.environ.get("UGII_BASE_DIR") or "").strip()
    if base:
        candidate = os.path.join(base, "IGES", filename)
        if os.path.isfile(candidate):
            return candidate
    return None


def iges_exe_path():
    """IGES 转换器路径；优先环境变量，其次 NX 安装目录，最后回落到默认目录"""
    path = _iges_tool_path("iges.exe")
    if path:
        return path
    fallback = os.path.join(IGES_DEFAULT_DIR, "iges.exe")
    if os.path.isfile(fallback):
        return fallback
    raise RuntimeError(T("log.iges_exe_missing", fallback))


def iges_def_path():
    """IGES 转换配置 igesexport.def；回落到默认目录时记一条警告"""
    path = _iges_tool_path("igesexport.def")
    if path:
        return path
    fallback = os.path.join(IGES_DEFAULT_DIR, "igesexport.def")
    log(T("log.iges_def_fallback", fallback))
    return fallback


def _iges_log_tail(log_path, lines=None):
    """失败时返回日志尾部若干行；读不到就说明读不到"""
    import io
    lines = IGES_LOG_TAIL_LINES if lines is None else lines
    try:
        with io.open(log_path, "r", encoding="utf-8", errors="replace") as handle:
            text = handle.read()
    except Exception as e:
        return "\n" + T("log.iges_log_unread", e)
    tail = [line for line in text.splitlines() if line.strip()][-lines:]
    if not tail:
        return "\n" + T("log.iges_log_unread", log_path)
    return "\n" + "\n".join("      " + line for line in tail)


def export_iges(prt_path, igs_path):
    """直调 NX 的 IGES translator 导出单个文件，返回产物字节数

    命令形如：iges.exe <prt> "out=<igs>" "log=<log>" "def=<def>"
    日志写到 %TEMP%\\nx_batch_export_iges\\，不落在用户文档目录。
    IGES 不做“是否含几何”校验，只判存在且大小 > 0。
    """
    import subprocess
    import tempfile

    exe = iges_exe_path()
    def_file = iges_def_path()

    log_dir = os.path.join(tempfile.gettempdir(), "nx_batch_export_iges")
    try:
        if not os.path.isdir(log_dir):
            os.makedirs(log_dir)
    except Exception:
        pass
    stem = os.path.splitext(os.path.basename(prt_path))[0]
    log_path = os.path.join(log_dir, "%s_%d_%d.log" % (
        stem, os.getpid(), int(time.time() * 1000)))

    command = [exe, prt_path,
               "out=%s" % igs_path,
               "log=%s" % log_path,
               "def=%s" % def_file]
    before = snapshot_candidates(igs_path)
    try:
        process = subprocess.run(
            command,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=IGES_TIMEOUT,
            creationflags=0x08000000)   # CREATE_NO_WINDOW
    except TypeError:
        # 老版本 Python 不认 creationflags 关键字
        process = subprocess.run(
            command,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=IGES_TIMEOUT)
    except Exception as e:
        # TimeoutExpired 在不同 Python 版本里可能是不同名字，按类型名判断更稳
        if type(e).__name__ == "TimeoutExpired":
            raise RuntimeError(T("log.iges_timeout", IGES_TIMEOUT))
        raise RuntimeError(T("log.iges_cmd_fail", e))

    if process.returncode != 0:
        raise RuntimeError(T("log.iges_exit", process.returncode)
                           + _iges_log_tail(log_path))

    reclaim_output(igs_path, before)
    size = step_size(igs_path)
    if size <= 0:
        raise RuntimeError(T("log.err_iges_out", igs_path)
                           + _iges_log_tail(log_path))
    return size


def create_iges_creator(session):
    """按 NX 版本依次尝试创建 IGES 导出器（内存部件导出用）"""
    dm = session.DexManager
    errors = []
    for method_name in ("CreateIgesCreator", "CreateIges214Creator"):
        factory = getattr(dm, method_name, None)
        if factory is None:
            errors.append(T("log.err_method_missing", method_name))
            continue
        try:
            return factory()
        except Exception as e:
            errors.append(T("log.err_method_fail", method_name, e))
    raise RuntimeError(T("log.err_creator", "; ".join(errors)))


def export_iges_mem(session, part, igs_path, layer_mask):
    """把内存部件直接导成 IGES（不落中间 .prt，不调用 iges.exe）

    与 StepCreator 同族的“输出文件 + 从显示/工作部件导出”契约；
    该契约需实机验证，失败时抛异常，由 chain_convert 关闭部件。
    """
    last_error = ""
    before = snapshot_candidates(igs_path)
    for attempt in (1, 2):
        if attempt == 2:
            try:
                if os.path.exists(igs_path):
                    os.remove(igs_path)
                    log(T("log.retry_deleted"))
            except Exception as e:
                log(T("log.retry_delete_fail", e))
        creator = create_iges_creator(session)
        try:
            creator.OutputFile = igs_path
            option = getattr(type(creator), "ExportFromOption", None)
            applied = False
            for flag in ("DisplayPart", "WorkPart"):
                value = getattr(option, flag, None) if option is not None else None
                if value is None:
                    continue
                try:
                    creator.ExportFrom = value
                    applied = True
                    break
                except Exception as e:
                    log(T("log.export_from_fail", e))
            if not applied:
                raise RuntimeError(T("log.err_iges_from_part"))
            if hasattr(type(creator), "LayerMask"):
                try:
                    creator.LayerMask = layer_mask
                except Exception as e:
                    log(T("log.layermask_fail", e))
            if hasattr(type(creator), "ProcessHoldFlag"):
                try:
                    creator.ProcessHoldFlag = True
                except Exception as e:
                    log(T("log.holdflag_fail", e))
            creator.Commit()
        finally:
            try:
                creator.Destroy()
            except Exception:
                pass

        reclaim_output(igs_path, before)
        size = step_size(igs_path)
        if size > 0:
            if attempt == 2:
                log(T("log.retry2_ok", size))
            return size
        last_error = T("log.err_iges_out", igs_path)
        log(T("log.invalid_attempt", attempt, last_error))
        time.sleep(0.4)

    try:
        if os.path.exists(igs_path):
            os.remove(igs_path)
    except Exception as e:
        log(T("log.cleanup_fail", e))
    raise RuntimeError(T("log.err_export_invalid", last_error))


def batch_import_module():
    """按需加载同目录的 batch_import（不改动脚本启动行为）"""
    import sys
    here = os.path.dirname(os.path.abspath(__file__))
    if here not in sys.path:
        sys.path.insert(0, here)
    import batch_import
    return batch_import


def chain_convert(session, src_path, out_path, fmt, used, layer_mask):
    """stp/igs -> stp/igs：内存建件导入后直接导出，不落中间 .prt、不删源文件

    返回产物字节数；失败抛出异常。finally 关闭本次新建的部件，
    避免导入成功但导出失败时留下内存垃圾。源文件在整个过程中不被修改。
    """
    BI = batch_import_module()
    if not used:
        used.update(BI.existing_part_names(session))
    try:
        previous = session.Parts.Display
    except Exception:
        previous = None
    state, name, detail, _disk, part = BI.import_one(session, src_path, used, save_as=False)
    if state != "ok" or part is None:
        raise RuntimeError(T("log.chain_import_fail", detail or state))
    try:
        log(T("log.chain_begin", name, fmt.upper()))
        if fmt == "iges":
            return export_iges_mem(session, part, out_path, layer_mask)
        return export_step(session, "", out_path, layer_mask, source_part=part)
    finally:
        try:
            BI.close_part(session, part)
        except Exception as e:
            log(T("log.chain_close_fail", e))
        if previous is not None:
            try:
                session.Parts.SetDisplay(previous, False, True)
            except Exception as e:
                log(T("log.chain_restore_fail", e))


def legacy_pick(base_dir=None):
    """选择窗口不可用时的后备方案：分步弹出系统文件夹/文件对话框"""
    folders = None
    files = []
    if PICK_MODE in ("folder", "both"):
        try:
            folders = select_folders(base_dir)
        except Exception as e:
            log(T("log.folder_dialog_fail", e))
    if PICK_MODE in ("file", "both"):
        try:
            files = pick_files(T("ui.file_pick_title"), base_dir or "")
        except Exception as e:
            files = []
            log(T("log.file_dialog_fail", e))
    return folders, files


def _lock_path():
    import tempfile
    return os.path.join(tempfile.gettempdir(), "nx_step_export.lock")


def _pid_alive(pid):
    """判断锁文件里的 PID 是否仍在运行。

    返回 True=进程存在，False=进程已不存在（陈旧锁），None=无法判断（按占用处理）。
    """
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.OpenProcess(0x1000, False, int(pid))  # QUERY_LIMITED_INFORMATION
        if handle:
            kernel32.CloseHandle(handle)
            return True
        # 打不开句柄时区分“PID 不存在”(ERROR_INVALID_PARAMETER=87) 与“权限不足”
        last = kernel32.GetLastError()
        if last == 87:
            return False
        return None
    except Exception:
        return None


def _lock_event(message):
    """锁事件落盘：此时主日志可能尚未打开，先写控制台 + %TEMP% 下的锁日志"""
    log(message)
    try:
        import tempfile
        path = os.path.join(tempfile.gettempdir(), "nx_step_export_busy.log")
        _rotate_log(path)
        with open(path, "a", encoding="utf-8") as handle:
            handle.write(message + "\n")
    except Exception:
        pass


def _acquire_run_lock():
    """单实例锁：防止连点两次菜单/重复触发时两个导出同时改同一批文件"""
    global _RUN_LOCK
    path = _lock_path()
    for attempt in (1, 2):
        try:
            fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode("ascii"))
            os.close(fd)
            _RUN_LOCK = path
            return True
        except FileExistsError:
            pid = None
            try:
                with open(path, "r", encoding="ascii") as handle:
                    pid = int((handle.read() or "0").strip() or 0)
            except Exception:
                pid = None
            # 先验证锁里的进程是否真的存在：已退出 → 删陈旧锁并重试；
            # 存在 → busy；无法判断（权限/异常）→ 记日志后按 busy 处理（安全优先）
            state = _pid_alive(pid) if pid else None
            if attempt == 1 and state is False:
                try:
                    os.remove(path)
                    _lock_event(T("log.stale_lock", pid))
                except Exception:
                    pass
                continue
            if attempt == 1 and state is None:
                _lock_event(T("log.lock_unknown", pid if pid else "?"))
            message = T("log.busy")
            try:
                print(message)
            except Exception:
                pass
            try:
                import tempfile
                busy_log = os.path.join(tempfile.gettempdir(), "nx_step_export_busy.log")
                _rotate_log(busy_log)
                with open(busy_log, "a", encoding="utf-8") as handle:
                    handle.write(message + "\n")
            except Exception:
                pass
            notify_busy()
            return False
        except Exception:
            # 锁机制本身不可用时，不要阻止正常导出
            return True
    return False


def _release_run_lock():
    global _RUN_LOCK
    if _RUN_LOCK:
        try:
            os.remove(_RUN_LOCK)
        except Exception:
            pass
        _RUN_LOCK = None


def main():
    if not _acquire_run_lock():
        return
    try:
        _run_locked()
    finally:
        _release_run_lock()


def _run_locked():
    global _log_handle

    session = NXOpen.Session.GetSession()
    # 不再有“项目根目录”；base_dir 恒为 None，仅作为兼容参数传给旧函数
    base_dir = None

    # 日志放到插件目录下的 logs\；插件目录不可写时回落 %TEMP%
    # utf-8-sig：带 BOM 的 UTF-8，英文版/繁中版 Windows 下用 Excel 直接打开不乱码；
    # 本脚本自身从不回读该日志，所以 BOM 不会影响任何逻辑
    log_path = None
    try:
        log_dir = os.path.join(plugin_dir(), "logs")
        if not os.path.isdir(log_dir):
            os.makedirs(log_dir)
        log_path = os.path.join(log_dir, "batch_export_step.log")
        _rotate_log(log_path)
        _log_handle = open(log_path, "w", encoding="utf-8-sig")
    except Exception:
        # 插件目录只读/被占用时退到 %TEMP%，保证导出仍可继续
        import tempfile
        fallback = os.path.join(tempfile.gettempdir(),
                                "nx_step_export_%d.log" % os.getpid())
        _log_handle = open(fallback, "w", encoding="utf-8-sig")
        log(T("log.log_fallback", fallback))

    # 交互选择结果：include_dirs=None 表示沿用配置里的 INCLUDE_DIRS
    include_dirs = None
    picked_files = []
    scan_dirs = True

    plan = None
    overrides = {}
    # 交互模式首次不预置选项：由 choose_targets 从记忆文件读取上次选择，
    # 避免默认值把记忆值覆盖掉；非交互模式在下面统一回落到默认值
    options = None if INTERACTIVE else default_options()
    if INTERACTIVE:
        while True:
            folders = None
            picked = []
            try:
                log(T("log.opening_picker"))
                folders, picked, picked_options = choose_targets(
                    base_dir=read_last_dir(), options=options)
                options = picked_options
            except Exception as e:
                log(T("log.picker_fail", e))
                folders, picked = legacy_pick()
            # 指定导出目录为空/不存在时回落到原文档所在路径（交互模式下当次即生效）
            options = resolve_output_options(options)

            include_dirs = None
            picked_files = []
            scan_dirs = True

            if folders is not None:
                include_dirs = folders
                if folders:
                    log(T("log.picked_folders", ", ".join(folders)))
                else:
                    log(T("log.picked_root"))
            if picked:
                picked_files = picked
                log(T("log.picked_files", len(picked_files)))
                for path in picked_files:
                    log("  %s" % rel_display(path, base_dir))

            # 只选了文件、没选文件夹时，不再扫描目录
            scan_dirs = not (picked_files and folders is None)

            if folders is None and not picked_files:
                log(T("log.no_selection"))
                _log_handle.close()
                _log_handle = None
                return

            plan = build_plan(find_prt_files(base_dir, include_dirs,
                                             picked_files, scan_dirs), options)
            if not plan:
                log(T("log.no_prt_reread"))
                continue

            try:
                action, overrides = preview_plan(plan, base_dir)
            except Exception as e:
                log(T("log.preview_fail", e))
                action = "ok"
                overrides = {}

            if action == "back":
                log(T("log.back"))
                continue
            if action != "ok":
                log(T("log.preview_cancel"))
                _log_handle.close()
                _log_handle = None
                return
            break

    if plan is None:
        options = resolve_output_options(options)
        plan = build_plan(find_prt_files(base_dir, include_dirs,
                                         picked_files, scan_dirs), options)

    total = len(plan)
    log("=" * 60)
    log(T("log.src_any") if base_dir is None else T("log.src_dir", base_dir))
    if options.get("out_mode") == "dir":
        log(T("log.output_custom", options.get("out_dir") or ""))
    else:
        log(T("log.output_line", output_ext(options)))
    log(T("log.format_line", output_format(options).upper()))
    log(T("log.file_count", total))
    if scan_dirs:
        effective_dirs = list(INCLUDE_DIRS if include_dirs is None else include_dirs)
        if effective_dirs:
            log(T("log.include_dirs", ", ".join(effective_dirs)))
    if EXCLUDE_DIRS:
        log(T("log.exclude_dirs", ", ".join(EXCLUDE_DIRS)))
    if INCLUDE_FILES:
        log(T("log.include_files", ", ".join(INCLUDE_FILES)))
    if EXCLUDE_FILES:
        log(T("log.exclude_files", ", ".join(EXCLUDE_FILES)))
    if picked_files:
        log(T("log.manual_files", len(picked_files)))
    log("=" * 60)

    if total == 0:
        log(T("log.no_files_end"))
        log(T("log.dryrun_hint"))
        _log_handle.close()
        _log_handle = None
        return

    if DRY_RUN:
        log(T("log.dryrun_title"))
        for index, (prt_path, out_path, state) in enumerate(plan, 1):
            mark = T("log.dryrun_skip_mark") if state == "SKIP" else ""
            log("  %3d. %s  ->  %s%s" % (
                index,
                rel_display(prt_path, base_dir),
                rel_display(out_path, base_dir),
                mark))
        log("=" * 60)
        _log_handle.close()
        _log_handle = None
        return

    # 导出前检查：计划里哪些文件已经打开且有未保存修改（只按内存 IsModified 判断）
    modified_items = collect_modified_open(session, plan)
    skip_unsaved = {}
    if modified_items:
        log(T("log.modified_header", len(modified_items)))
        for item in modified_items:
            log("  %s" % rel_display(item["path"], base_dir))
        if INTERACTIVE:
            confirmed, save_keys = ask_save_modified(modified_items)
            if not confirmed:
                log(T("log.cancel_unsaved"))
                _log_handle.close()
                _log_handle = None
                return
            for item in modified_items:
                item_key = os.path.normcase(os.path.abspath(item["path"]))
                if item_key not in save_keys:
                    log(T("log.user_skip_mod", item["name"]))
                    skip_unsaved[item_key] = "reason.user_skip"
                    continue
                try:
                    item["part"].Save(NXOpen.BasePartSaveComponents.TrueValue,
                                      NXOpen.BasePartCloseAfterSave.FalseValue)
                    log(T("log.saved_ok", rel_display(item["path"], base_dir)))
                except Exception as e:
                    log(T("log.save_fail_skip", item["name"], e))
                    skip_unsaved[item_key] = "reason.save_fail"
        else:
            for item in modified_items:
                log(T("log.noninteractive_skip", item["name"]))
                skip_unsaved[os.path.normcase(os.path.abspath(item["path"]))] = \
                    "reason.noninteractive"

    ok, skipped, failed = 0, 0, []
    original_display_part = None
    original_display_path = None
    original_work_part = None
    try:
        original_display_part = session.Parts.Display
        if original_display_part is not None:
            original_display_path = original_display_part.FullPath
        original_work_part = session.Parts.Work
    except Exception:
        original_display_part = None
        original_work_part = None
    original_keys = set(k for k in (_part_key(p) for p in _session_parts(session)) if k)

    fmt = output_format(options)
    chain_used = set()   # 链式转换新建的部件名，跨文件复用避免重名
    for index, (prt_path, out_path, state) in enumerate(plan, 1):
        item_parts = []
        relative = rel_display(prt_path, base_dir)
        prefix = "[%d/%d]" % (index, total)
        src_ext = os.path.splitext(prt_path)[1].lower()

        # 选择/预览与真正导出之间有时间差，文件可能已被删除或移走
        if len(prt_path) > 259:
            log(T("log.too_long_path", relative))
            failed.append(relative)
            continue
        if not os.path.isfile(prt_path):
            log(T("log.missing_part", relative))
            failed.append(relative)
            continue

        # NX 无法打开含 emoji（非 BMP 字符）的文件路径，提前给出明确提示而不是底层异常
        if any(ord(ch) > 0xFFFF for ch in relative):
            log(T("log.unsupported_name", relative))
            failed.append(relative)
            continue

        # 未保存修改且用户选择跳过（或非交互模式）的项，先于覆盖/跳过规则处理
        prt_key = os.path.normcase(os.path.abspath(prt_path))
        if prt_key in skip_unsaved:
            log("%s %s: %s" % (prefix, T(skip_unsaved[prt_key]), relative))
            skipped += 1
            continue

        # 行级勾选优先；未勾选的行按 build_plan 算好的同名策略状态执行
        if state == "SELF":
            # 源与目标同格式，不转换（避免把源文件当产物覆盖）
            log(T("log.chain_same_fmt", relative))
            skipped += 1
            continue
        key = os.path.normcase(os.path.abspath(out_path))
        override = overrides.get(key)
        existing_valid = output_is_valid(out_path, options)
        if state == "SKIP" and existing_valid and override is not True:
            log(T("log.skip_exists", prefix, rel_display(out_path, base_dir)))
            skipped += 1
            continue
        if state == "OVERWRITE" and override is False:
            log(T("log.skip_exists", prefix, rel_display(out_path, base_dir)))
            skipped += 1
            continue

        start = time.time()
        log(T("log.progress", index, total, relative))
        layer_mask = "1-256"
        try:
            if fmt == "step" and VISIBLE_LAYERS_ONLY and src_ext == ".prt":
                # 图层可见性只能在显示部件上查询；脚本自己打开的部件会在结束时关闭
                part = ensure_displayed(session, prt_path, item_parts)
                layer_mask, view_name, visible_count = visible_layer_mask(part)
                if not layer_mask:
                    log(T("log.no_visible", view_name))
                    failed.append(relative)
                    continue
                log(T("log.visible", layer_mask, view_name, visible_count))

            # 覆盖前先把旧产物改名备份：新产物无效时自动还原，避免丢旧文件
            backup = None
            if os.path.exists(out_path):
                if existing_valid:
                    log(T("log.cover", prefix, relative))
                    backup = _make_backup(out_path)
                else:
                    # 目标存在但不是有效产物（只有文件头等）：先移除，避免转换器拒绝写入
                    try:
                        os.remove(out_path)
                    except Exception as e:
                        raise RuntimeError(T("log.err_remove_target", out_path, e,
                                              os.path.splitext(out_path)[1] or "PART"))

            try:
                if src_ext in CHAIN_SOURCE_EXTS:
                    # 链式转换：内存建件 → 导入几何 → 直接导出目标格式（不落中间 .prt）
                    size = chain_convert(session, prt_path, out_path, fmt,
                                         chain_used, layer_mask)
                elif fmt == "iges":
                    size = export_iges(prt_path, out_path)
                else:
                    size = export_step(session, prt_path, out_path, layer_mask)
            except Exception:
                if backup:
                    try:
                        if os.path.exists(out_path):
                            os.remove(out_path)
                        os.replace(backup, out_path)
                        backup = None
                        log(T("log.old_restored", rel_display(out_path, base_dir)))
                    except Exception as e:
                        log(T("log.backup_cleanup_fail", e))
                raise
            if backup:
                _prune_backups(out_path, 3)
            log(T("log.done", time.time() - start,
                  rel_display(out_path, base_dir), size))
            ok += 1
        except Exception as e:
            log(T("log.failed", e))
            for line in traceback.format_exc().splitlines():
                log("      " + line)
            failed.append(relative)
        finally:
            if item_parts:
                # 先切回用户原状态，再逐件关闭“本件由脚本新打开”的部件
                _restore_state(session, original_display_part, original_work_part)
                close_opened_parts(session, item_parts, original_keys)
                del item_parts[:]

    # 恢复用户原本的显示部件
    if original_display_part is not None:
        try:
            # 用“切回活动显示部件”的方式恢复：对已打开的部件再调 OpenBaseDisplay 会被 NX 拒绝，
            # 导致显示区变空；SetActiveDisplay 不会重复打开文件。
            _status, _load = session.Parts.SetActiveDisplay(
                original_display_part,
                NXOpen.DisplayPartOption.ReplaceExisting,
                NXOpen.PartDisplayPartWorkPartOption.SameAsDisplay)
            log(T("log.display_restored", original_display_path))
        except Exception as e:
            log(T("log.display_restore_fail", e))
    # 工作部件也还原成执行前那个（切显示部件会把工作部件跟着改掉）
    if original_work_part is not None:
        try:
            current_work = session.Parts.Work
            if current_work is None or \
                    os.path.normcase(os.path.abspath(current_work.FullPath)) != \
                    os.path.normcase(os.path.abspath(original_work_part.FullPath)):
                session.Parts.SetWork(original_work_part)
                log(T("log.work_restored", original_work_part.FullPath))
        except Exception as e:
            log(T("log.work_restore_fail", e))

    log("=" * 60)
    log(T("log.summary", ok, skipped, len(failed)))
    for name in failed:
        log(T("log.failed_item", name))
    log("=" * 60)

    _log_handle.close()
    _log_handle = None


if __name__ == "__main__":
    main()
