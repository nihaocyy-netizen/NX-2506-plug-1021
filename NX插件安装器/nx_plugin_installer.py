"""NX / UG 插件安装工具（图形界面版）

用法：
  1. 将本工具（或打包后的 exe）放在插件文件夹内，双击运行；
  2. 点击"浏览…"选择 NX(UG) 的安装根目录
     （即包含 UGII 文件夹的目录，例如 C:\\Program Files\\Siemens\\NX 12.0）；
  3. 点击"安装插件"，工具会把插件所在目录写入
     UGII\\menus\\custom_dirs.dat，重启 NX 后生效。

功能：
  - 启动时自动识别 NX/UG 安装目录（也可点击“自动识别”重新扫描）
  - 根据 Windows 语言自动选择简体或繁体 NX 菜单
  - 浏览选择目录 + 自动校验（所选目录下没有 UGII 时给出提示）
  - 只读的 custom_dirs.dat 自动处理：临时取消只读 → 写入 → 恢复原状（无需手动改属性）
  - 安装时自动去重，避免重复写入
  - 支持一键卸载（从 custom_dirs.dat 中移除记录）
  - 安装状态检测、日志输出、权限不足提醒

打包为单个 exe（无控制台窗口）：
  pip install pyinstaller
  pyinstaller -F -w nx_plugin_installer.py
"""

import locale
import os
import stat
import sys
from contextlib import contextmanager
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


def _valid_nx_root(path):
    """Return a normalized NX root when it contains the UGII directory."""
    if not path:
        return ""
    path = os.path.normpath(os.path.expandvars(os.path.expanduser(str(path))))
    if os.path.basename(path).lower() in ("ugii", "nxbin"):
        path = os.path.dirname(path)
    if os.path.isdir(os.path.join(path, "UGII")):
        return path
    return ""


def detect_nx_installations():
    """Find installed NX/UG roots without requiring a manual browse.

    The order is deliberate: active NX environment, Siemens registry entries,
    then common Siemens/NX folders. Duplicate paths are removed while keeping
    the first (usually the currently running NX version).
    """
    candidates = []
    for name in ("UGII_BASE_DIR", "UGII_ROOT_DIR", "NX2512_ROOT", "NX_ROOT"):
        value = os.environ.get(name)
        if value:
            candidates.append(value)
    if os.name == "nt":
        try:
            import winreg
            keys = (
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Siemens\NX"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Siemens\NX"),
                (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Siemens\NX"),
            )
            for hive, subkey in keys:
                try:
                    with winreg.OpenKey(hive, subkey) as key:
                        for i in range(winreg.QueryInfoKey(key)[0]):
                            version = winreg.EnumKey(key, i)
                            with winreg.OpenKey(key, version) as item:
                                for name in ("InstallDir", "InstallationDirectory", "BaseDir"):
                                    try:
                                        candidates.append(winreg.QueryValueEx(item, name)[0])
                                    except OSError:
                                        pass
                except OSError:
                    pass
        except ImportError:
            pass
    roots = [
        os.path.expandvars(r"%ProgramFiles%\Siemens"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Siemens"),
        r"C:\Siemens", r"D:\Siemens", r"E:\Siemens",
        r"C:\NX", r"D:\NX", r"E:\NX",
    ]
    for root in roots:
        if os.path.isdir(root):
            candidates.append(root)
            try:
                candidates.extend(os.path.join(root, name) for name in os.listdir(root))
            except OSError:
                pass
    found = []
    seen = set()
    for candidate in candidates:
        valid = _valid_nx_root(candidate)
        if valid:
            key = os.path.normcase(valid)
            if key not in seen:
                seen.add(key)
                found.append(valid)
    return found


def get_windows_locale():
    """Return the Windows user locale, for example zh-CN or zh-TW."""
    if os.name == "nt":
        try:
            import ctypes
            buffer = ctypes.create_unicode_buffer(85)
            if ctypes.windll.kernel32.GetUserDefaultLocaleName(buffer, 85):
                return buffer.value
        except Exception:
            pass
    return (locale.getlocale()[0] or "").replace("_", "-")


def sync_menu_locale(plugin_dir):
    """Select the locale-specific Chinese menu before NX is restarted."""
    startup = ""
    cursor = os.path.abspath(plugin_dir)
    for _ in range(5):
        candidate = os.path.join(cursor, "startup")
        if os.path.isdir(candidate):
            startup = candidate
            break
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    if not startup:
        raise FileNotFoundError("找不到 startup 菜单目录")
    active = os.path.join(startup, "myDemo.men")
    locale_name = get_windows_locale().lower().replace("_", "-")
    traditional = locale_name.startswith(("zh-tw", "zh-hk", "zh-mo"))
    template = os.path.join(startup, "myDemo_zhTW.men" if traditional else "myDemo_zhCN.men")
    encoding = "big5" if traditional else "gb18030"
    with open(template, "rb") as stream:
        content = stream.read().decode(encoding)
    with open(active, "w", encoding=encoding, newline="") as stream:
        stream.write(content)
    return active, ("繁体中文" if traditional else "简体中文"), locale_name

# ---------------- 核心逻辑（与界面无关） ----------------


def decode_line(raw):
    """把单行字节按 utf-8/gbk/latin-1 依次尝试解码（同一文件各行编码可能不同）"""
    for enc in ("utf-8-sig", "gbk", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return ""


def is_readonly(path):
    """判断文件是否带只读属性（Windows: 写权限位被清除 == FILE_ATTRIBUTE_READONLY）"""
    try:
        return not os.stat(path).st_mode & stat.S_IWRITE
    except OSError:
        return False


@contextmanager
def open_writable(path, mode="ab"):
    """打开 custom_dirs.dat 写入；若文件带只读属性则临时取消，写完后自动恢复原状。

    NX 出厂的 custom_dirs.dat 默认带只读属性，直接写入会报“权限不足”。
    取消/恢复只读属性属于“写属性”操作，普通权限即可完成，无需管理员。
    """
    was_readonly = is_readonly(path) if os.path.isfile(path) else False
    if was_readonly:
        os.chmod(path, stat.S_IWRITE)  # 临时取消只读
    try:
        with open(path, mode) as f:
            yield f
    finally:
        if was_readonly:
            try:
                os.chmod(path, stat.S_IREAD)  # 恢复只读原状
            except OSError:
                pass


def is_installed_line(line, plugin_dir):
    """判断 custom_dirs.dat 中某一行是否就是插件目录（忽略大小写）"""
    t = line.strip()
    if not t:
        return False
    try:
        return os.path.normpath(t).lower() == plugin_dir.lower()
    except (ValueError, OSError):
        return False


def is_installed(custom_dirs_path, plugin_dir):
    """检测插件路径是否已存在于 custom_dirs.dat"""
    if not os.path.isfile(custom_dirs_path):
        return False
    try:
        with open(custom_dirs_path, "rb") as f:
            raw = f.read()
    except OSError:
        return False
    return any(is_installed_line(decode_line(p.strip()), plugin_dir) for p in raw.splitlines())


def install_plugin(custom_dirs_path, plugin_dir, default_encoding=None):
    """把插件目录追加写入 custom_dirs.dat；已存在时返回 False"""
    if os.path.isfile(custom_dirs_path):
        if is_installed(custom_dirs_path, plugin_dir):
            return False
        with open(custom_dirs_path, "rb") as f:
            raw = f.read()
        if raw:
            nl = b"\r\n" if b"\r\n" in raw else b"\n"
            # 文件已以换行结尾时不补前置换行，避免产生多余空行
            prefix = b"" if raw.endswith((b"\n", b"\r")) else nl
        else:
            nl, prefix = b"\r\n", b""
    else:
        os.makedirs(os.path.dirname(custom_dirs_path), exist_ok=True)
        nl, prefix = b"\r\n", b""
    # 追加行按系统区域编码写入（NX 以系统编码读取 custom_dirs.dat），无法编码时退回 UTF-8
    enc = default_encoding or locale.getpreferredencoding(False) or "utf-8"
    try:
        line = plugin_dir.encode(enc)
    except UnicodeEncodeError:
        line = plugin_dir.encode("utf-8")
    with open_writable(custom_dirs_path, "ab") as f:
        f.write(prefix + line + nl)
    return True


def uninstall_plugin(custom_dirs_path, plugin_dir):
    """从 custom_dirs.dat 移除插件目录；逐行按原始字节处理，其他记录保持原样不受影响"""
    if not os.path.isfile(custom_dirs_path):
        return False
    with open(custom_dirs_path, "rb") as f:
        raw = f.read()
    kept, removed = [], False
    for part in raw.splitlines(keepends=True):
        if is_installed_line(decode_line(part.strip()), plugin_dir):
            removed = True  # 丢弃该行（连同换行符）
        else:
            kept.append(part)
    if not removed:
        return False
    with open_writable(custom_dirs_path, "wb") as f:
        f.write(b"".join(kept))
    return True


def get_plugin_dir():
    """获取插件所在目录（兼容 PyInstaller 打包后的 exe）"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def is_admin():
    try:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return True


# ---------------- 图形界面 ----------------


class InstallerApp:
    def __init__(self, root):
        self.root = root
        self.plugin_dir = os.path.normpath(get_plugin_dir())
        self._build_ui()
        self.log(f"插件目录：{self.plugin_dir}")
        try:
            menu_path, language, locale_name = sync_menu_locale(self.plugin_dir)
            self.log(f"已根据系统语言 {locale_name} 加载{language}菜单：{menu_path}", "ok")
        except Exception as error:
            self.log(f"菜单语言自动选择失败：{error}", "warn")
        detected = detect_nx_installations()
        if detected:
            self.path_var.set(detected[0])
            self.refresh_status()
            self.log(f"已自动识别 NX 安装目录：{detected[0]}", "ok")
            if len(detected) > 1:
                self.log("检测到多个 NX 安装：" + "；".join(detected), "warn")
        else:
            self.log("未自动识别到 NX 安装目录，请点击“浏览…”手动选择。", "warn")
        if os.name == "nt" and not is_admin():
            self.log("提示：当前以普通权限运行。若 NX 安装在 Program Files 下，写入可能被拒绝，届时请以管理员身份重新运行本工具。", "warn")

    # ---------- 界面 ----------
    def _build_ui(self):
        self.root.title("NX / UG 插件安装工具")
        self.root.geometry("680x540")
        self.root.minsize(600, 460)
        try:
            ttk.Style().theme_use("vista")
        except tk.TclError:
            pass

        main = ttk.Frame(self.root, padding=14)
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            main,
            text="使用说明：请将本工具放在插件文件夹内运行；\n"
                 "选择 NX(UG) 的安装根目录后，点击“安装插件”。",
            justify=tk.LEFT,
        ).pack(anchor=tk.W)

        # ① NX 安装路径
        box1 = ttk.LabelFrame(main, text="① NX (UG) 安装路径", padding=10)
        box1.pack(fill=tk.X, pady=(10, 6))
        row = ttk.Frame(box1)
        row.pack(fill=tk.X)
        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(row, textvariable=self.path_var)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(row, text="浏览…", command=self.browse_path, width=8).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(row, text="自动识别", command=self.auto_detect, width=10).pack(side=tk.LEFT, padx=(6, 0))
        self.path_entry.bind("<FocusOut>", lambda e: self.refresh_status())
        self.path_entry.bind("<Return>", lambda e: self.refresh_status())
        self.path_status = ttk.Label(box1, text="尚未选择路径", foreground="gray")
        self.path_status.pack(anchor=tk.W, pady=(6, 0))

        # ② 插件路径
        box2 = ttk.LabelFrame(main, text="② 插件路径（将注册到 custom_dirs.dat 的目录）", padding=10)
        box2.pack(fill=tk.X, pady=6)
        ttk.Label(box2, text=self.plugin_dir, wraplength=560, foreground="#0a5aa0").pack(anchor=tk.W)
        self.install_status = ttk.Label(box2, text="插件状态：未知", foreground="gray")
        self.install_status.pack(anchor=tk.W, pady=(6, 0))

        # ③ 操作按钮
        box3 = ttk.Frame(main)
        box3.pack(pady=10)
        self.install_btn = ttk.Button(box3, text="安装插件", command=self.install)
        self.install_btn.pack(side=tk.LEFT, padx=4)
        self.uninstall_btn = ttk.Button(box3, text="卸载插件", command=self.uninstall)
        self.uninstall_btn.pack(side=tk.LEFT, padx=4)
        self.check_btn = ttk.Button(box3, text="检测状态", command=self.check_status)
        self.check_btn.pack(side=tk.LEFT, padx=4)

        # ④ 日志
        box4 = ttk.LabelFrame(main, text="日志", padding=6)
        box4.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        self.log_text = scrolledtext.ScrolledText(box4, height=9, state=tk.DISABLED, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        for tag, color in (("ok", "#0a7d28"), ("err", "#c62828"), ("warn", "#b26a00")):
            self.log_text.tag_configure(tag, foreground=color)

    def log(self, msg, tag=None):
        self.log_text.configure(state=tk.NORMAL)
        if tag:
            self.log_text.insert(tk.END, msg + "\n", tag)
        else:
            self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    # ---------- 路径 ----------
    def browse_path(self):
        folder = filedialog.askdirectory(title="请选择 NX (UG) 安装根目录")
        if folder:
            self.path_var.set(os.path.normpath(folder))
            self.refresh_status()

    def auto_detect(self):
        detected = detect_nx_installations()
        if not detected:
            self.path_status.config(text="未找到 NX 安装目录", foreground="#c62828")
            self.log("自动识别失败，请点击“浏览…”选择包含 UGII 文件夹的 NX 根目录。", "warn")
            return
        self.path_var.set(detected[0])
        self.refresh_status()
        self.log(f"已选择 NX 安装目录：{detected[0]}", "ok")
        if len(detected) > 1:
            self.log("其他候选目录：" + "；".join(detected[1:]), "warn")

    def _nx_path(self):
        """取输入框中的 NX 根目录（自动纠正误选 UGII 目录的情况）"""
        p = self.path_var.get().strip().strip('"').strip("'")
        if not p:
            return ""
        p = os.path.normpath(p)
        if os.path.basename(p).lower() == "ugii":
            p = os.path.dirname(p)
        return p

    def _custom_dirs(self, nx_path):
        return os.path.join(nx_path, "UGII", "menus", "custom_dirs.dat")

    def refresh_status(self):
        """校验路径并刷新状态标签；路径有效时返回 custom_dirs.dat 路径，否则 None"""
        nx_path = self._nx_path()
        if not nx_path:
            self.path_status.config(text="尚未选择路径", foreground="gray")
            self.install_status.config(text="插件状态：未知", foreground="gray")
            return None
        if not os.path.isdir(nx_path):
            self.path_status.config(text="✗ 目录不存在", foreground="#c62828")
            self.install_status.config(text="插件状态：未知", foreground="gray")
            return None
        if not os.path.isdir(os.path.join(nx_path, "UGII")):
            self.path_status.config(
                text="✗ 该目录下未找到 UGII 文件夹，请确认选择的是 NX 安装根目录",
                foreground="#c62828",
            )
            self.install_status.config(text="插件状态：未知", foreground="gray")
            return None
        cd = self._custom_dirs(nx_path)
        if os.path.isfile(cd):
            self.path_status.config(text="✓ 路径有效（custom_dirs.dat 已存在）", foreground="#0a7d28")
            if is_installed(cd, self.plugin_dir):
                self.install_status.config(text="插件状态：已安装 ✓", foreground="#0a7d28")
            else:
                self.install_status.config(text="插件状态：未安装", foreground="#b26a00")
        else:
            self.path_status.config(
                text="路径有效（custom_dirs.dat 不存在，安装时将自动创建）",
                foreground="#b26a00",
            )
            self.install_status.config(text="插件状态：未安装", foreground="#b26a00")
        return cd

    # ---------- 操作 ----------
    def install(self):
        nx_path = self._nx_path()
        if not nx_path:
            messagebox.showwarning("提示", "请先选择 NX (UG) 安装根目录。")
            return
        cd = self.refresh_status()
        if cd is None:
            messagebox.showerror("错误", "所选路径无效，请确认选择的是 NX 安装根目录。")
            return
        readonly = os.path.isfile(cd) and is_readonly(cd)
        if readonly:
            self.log("检测到 custom_dirs.dat 带只读属性，将临时取消后写入，完成后自动恢复原状。", "warn")
        try:
            if install_plugin(cd, self.plugin_dir):
                if readonly:
                    self.log("custom_dirs.dat 只读属性已恢复原状。")
                self.log(f"已写入 custom_dirs.dat：{cd}")
                self.log("恭喜：插件安装完成！重启 NX 后生效。", "ok")
                messagebox.showinfo("安装成功", "恭喜：插件安装完成！\n重启 NX 后即可使用。")
            else:
                self.log("插件路径已存在于 custom_dirs.dat，无需重复安装。", "warn")
                messagebox.showinfo("提示", "插件已安装过，无需重复安装。")
        except PermissionError:
            msg = "写入被拒绝：权限不足。\n请右键“以管理员身份运行”本工具后重试。"
            self.log(msg, "err")
            messagebox.showerror("权限不足", msg)
        except OSError as e:
            self.log(f"安装失败：{e}", "err")
            messagebox.showerror("错误", f"安装失败：\n{e}")
        self.refresh_status()

    def uninstall(self):
        nx_path = self._nx_path()
        if not nx_path:
            messagebox.showwarning("提示", "请先选择 NX (UG) 安装根目录。")
            return
        cd = self.refresh_status()
        if cd is None:
            messagebox.showerror("错误", "所选路径无效，请确认选择的是 NX 安装根目录。")
            return
        if not os.path.isfile(cd):
            self.log("未找到 custom_dirs.dat，无需卸载。", "warn")
            messagebox.showinfo("提示", "未找到 custom_dirs.dat，无需卸载。")
            return
        readonly = is_readonly(cd)
        if readonly:
            self.log("检测到 custom_dirs.dat 带只读属性，将临时取消后写入，完成后自动恢复原状。", "warn")
        try:
            if uninstall_plugin(cd, self.plugin_dir):
                if readonly:
                    self.log("custom_dirs.dat 只读属性已恢复原状。")
                self.log(f"已从 custom_dirs.dat 移除插件路径：{cd}")
                messagebox.showinfo("卸载成功", "插件已卸载。\n重启 NX 后生效。")
            else:
                self.log("custom_dirs.dat 中没有本插件的记录，无需卸载。", "warn")
                messagebox.showinfo("提示", "未发现本插件的安装记录。")
        except PermissionError:
            msg = "写入被拒绝：权限不足。\n请右键“以管理员身份运行”本工具后重试。"
            self.log(msg, "err")
            messagebox.showerror("权限不足", msg)
        except OSError as e:
            self.log(f"卸载失败：{e}", "err")
            messagebox.showerror("错误", f"卸载失败：\n{e}")
        self.refresh_status()

    def check_status(self):
        cd = self.refresh_status()
        if cd is None:
            return  # refresh_status 已在界面给出提示
        installed = is_installed(cd, self.plugin_dir)
        state = "已安装 ✓" if installed else "未安装"
        self.log(f"检测结果：插件状态 → {state}")
        messagebox.showinfo("安装状态", f"插件状态：{state}")


def main():
    root = tk.Tk()
    InstallerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
