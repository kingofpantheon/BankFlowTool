# -*- coding: utf-8 -*-
r"""
银行流水梳理工具 V1.7.0

本版相对 V1.6.9 只做一件事：把"出问题时的证据"留下来，业务逻辑没有改动。

日志规则：
  * 日志就在软件（exe）同目录下，文件名带软件名和版本号，例如：
        银行流水梳理工具 V1.7.0 运行日志.txt
    双击就能打开看，不用满硬盘去找；
  * 平时永远只有这一个日志文件，每次启动覆盖重写，不会越攒越多；
  * 但如果上次是闪退/被强制结束（日志里没有"正常退出"标记），
    本次启动会把上次日志的末尾（最多 300 行）原样搬到新日志开头留档，
    所以"崩了马上重启"也不会把证据冲掉；
  * 界面上还有「日志存到桌面」按钮，一键复制到桌面，直接发出来即可。

本版还做到：
  1. 每写一行立即落盘，卡死、闪退、被任务管理器结束时已写内容都还在。
  2. 未捕获异常、子线程异常、界面回调异常，全部连完整堆栈写进日志。
     这些在无控制台的 exe 里以前是完全静默的。
  3. 界面超过 90 秒没有响应时，自动把当时所有线程的调用堆栈写进同一个日志，
     可以直接看出卡在哪一行、正在处理哪个工作表。
  4. 上次异常退出后，下次启动会弹窗告知，并可一键打开日志。

日志位置（自动选择第一个可写的）：
  软件(exe)所在目录\银行流水梳理工具 V1.7.0 运行日志.txt
  %LOCALAPPDATA%\银行流水梳理工具\logs\（同名文件）
  %APPDATA%\银行流水梳理工具\logs\（同名文件）
  %TEMP%\银行流水梳理工具_logs\（同名文件）
也可以用环境变量 BANKTOOL_LOG_DIR 指定目录。

注意：日志记录的是文件名、工作表名、行列数、选项开关和错误信息，不写单元格内容；
但文件名本身可能含客户姓名，转发出去之前请自行确认。
"""

import os
import sys
import time
import atexit
import platform
import datetime
import logging
import faulthandler
import threading
import traceback

VERSION = "V1.7.0"
APP_NAME = "银行流水梳理工具"

# 卡死判定：主线程超过这么多秒没有心跳，就把所有线程堆栈写进日志
HANG_WARN_SECONDS = 90
# 正常退出时写的标记行；下次启动靠它判断上次是否正常结束
_EXIT_SENTINEL = "===== 程序正常退出 ====="
# 打包时会把这个文件名一起塞进 exe；软件里的“导出源码备份”就是把它写出来
SOURCE_BACKUP_NAME = "_source_backup.py"
ICON_BACKUP_NAME = "_icon_backup.ico"
# 上次异常退出时，最多把旧日志末尾这么多字符搬到新日志开头留档
CARRY_OVER_BYTES = 200 * 1024
CARRY_OVER_LINES = 300

logger = logging.getLogger("banktool")
logger.setLevel(logging.DEBUG)
logger.propagate = False

_log_dir = None
_log_file_path = None
_log_fp = None          # 唯一的日志文件句柄
_crash_fp = None        # 兼容旧代码的别名，实际就是 _log_fp
_prev_was_abnormal = False   # 上次运行是否异常退出
_notice_done = False
_log_lines = []              # 日志的内存镜像：文件被占用时用它导出
_log_memory_bytes = 0
LOG_MEMORY_LIMIT = 4 * 1024 * 1024
_start_time_text = ""
_last_beat = time.time()
_watchdog_started = False
_watchdog_stop = threading.Event()


# ==================== 一、日志文件的位置 ====================
def _program_dir():
    """程序所在目录（打包成 exe 后取 exe 所在文件夹）。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _try_make_dir(path):
    """能建目录、并且真能写进去，才算可用。"""
    try:
        os.makedirs(path, exist_ok=True)
        probe = os.path.join(path, ".write_test")
        with open(probe, "wb") as f:
            f.write(b"1")
        os.remove(probe)
        return True
    except Exception:
        return False


def _log_file_name():
    """日志文件名：软件名 + 版本号 + 运行日志.txt。放在软件旁边，双击就能看。"""
    return "%s %s 运行日志.txt" % (APP_NAME, VERSION)


def _candidate_log_dirs():
    """按优先级列出候选目录，第一个就是软件（exe）所在目录。

    用户要求"就在 exe 同目录下"，这里照办；只有那里写不进去
    （装在 Program Files、只读共享盘、被杀软受控文件夹拦下）才逐级退到
    用户目录 / 临时目录，并且界面底部会把实际路径显示出来。
    """
    dirs = []
    env_dir = os.environ.get("BANKTOOL_LOG_DIR")
    if env_dir:
        dirs.append(env_dir)
    dirs.append(_program_dir())
    for env_name in ("LOCALAPPDATA", "APPDATA"):
        base = os.environ.get(env_name)
        if base:
            dirs.append(os.path.join(base, APP_NAME, "logs"))
    tmp = os.environ.get("TEMP") or os.environ.get("TMP")
    if tmp:
        dirs.append(os.path.join(tmp, APP_NAME + "_logs"))
    return dirs


# ==================== 二、单一文件日志 ====================
class _SingleFileHandler(logging.Handler):
    """只写一个文件、每行都 flush 的极简 handler。
    不用 RotatingFileHandler：本工具要求日志始终只有一份，启动即覆盖。"""

    def __init__(self, fp):
        logging.Handler.__init__(self)
        self.fp = fp
        self.setFormatter(logging.Formatter(
            "%(asctime)s.%(msecs)03d [%(levelname)-7s] [%(threadName)s] %(message)s",
            "%Y-%m-%d %H:%M:%S",
        ))

    def emit(self, record):
        try:
            line = self.format(record)
            self.fp.write(line + "\n")
            self.fp.flush()
            _remember_line(line)
        except Exception:
            pass


def _read_tail(path, max_bytes=12288):
    """读文件末尾一段（旧日志可能很大，不整份读）。"""
    try:
        if not os.path.exists(path):
            return ""
        size = os.path.getsize(path)
        with open(path, "rb") as f:
            if size > max_bytes:
                f.seek(size - max_bytes)
            return f.read().decode("utf-8", "replace")
    except Exception:
        return ""


def _remember_line(line):
    """把日志同时留在内存里一份：Windows 上文件被本进程占着时，
    其它程序（记事本、复制）读不到，导出就走这份镜像。"""
    global _log_memory_bytes
    try:
        _log_lines.append(line)
        _log_memory_bytes += len(line) + 1
        while _log_memory_bytes > LOG_MEMORY_LIMIT and len(_log_lines) > 1:
            _log_memory_bytes -= len(_log_lines.pop(0)) + 1
    except Exception:
        pass


def _log_snapshot_text():
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    head = [
        "%s %s 运行日志（快照）" % (APP_NAME, VERSION),
        "生成时间: %s" % now,
        "原始日志: %s" % (_log_file_path or "（未能创建）"),
        "=" * 80,
    ]
    return "\n".join(head + list(_log_lines)) + "\n"


def _log_file_readable():
    try:
        with open(_log_file_path, "rb") as f:
            f.read(1)
        return True
    except Exception:
        return False


def _copy_log_to(target):
    """优先复制真实日志（含 faulthandler 写的原始堆栈）；
    文件读不到（被本进程独占）就退回内存镜像写一份。"""
    if _log_file_path:
        try:
            with open(_log_file_path, "rb") as src, open(target, "wb") as dst:
                dst.write(src.read())
            return True
        except Exception:
            pass
    try:
        with open(target, "w", encoding="utf-8") as dst:
            dst.write(_log_snapshot_text())
        return True
    except Exception:
        return False


def _open_log_file():
    """挑目录并打开日志；返回 (路径, 文件对象, 旧日志内容, 上次是否正常退出)。"""
    global _prev_was_abnormal
    name = _log_file_name()
    for d in _candidate_log_dirs():
        if not d or not _try_make_dir(d):
            continue
        path = os.path.join(d, name)
        prev_tail = _read_tail(path, CARRY_OVER_BYTES)
        prev_clean = True if not prev_tail.strip() else (_EXIT_SENTINEL in prev_tail)
        try:
            # "w" = 每次启动覆盖重写；buffering=1 = 行缓冲，每写一行就落盘
            fp = open(path, "w", encoding="utf-8", buffering=1)
        except Exception:
            continue
        _prev_was_abnormal = not prev_clean
        return path, fp, prev_tail, prev_clean
    return None, None, "", True


def _build_handlers():
    global _log_dir, _log_file_path, _log_fp, _crash_fp

    path, fp, prev_tail, prev_clean = _open_log_file()
    if fp is None:
        return False, "", True

    _log_file_path = path
    _log_dir = os.path.dirname(path)
    _log_fp = fp
    # faulthandler 和看门狗直接写同一个句柄：整个程序只有这一份日志
    _crash_fp = fp

    logger.addHandler(_SingleFileHandler(fp))

    # 有控制台时顺带打印（从源码运行有用；打包成无控制台的 exe 时这里是 None）
    if sys.stdout is not None:
        try:
            sh = logging.StreamHandler(sys.stdout)
            sh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S"))
            logger.addHandler(sh)
        except Exception:
            pass

    return True, prev_tail, prev_clean


def _flush_all():
    try:
        for h in logger.handlers:
            try:
                h.flush()
            except Exception:
                pass
        if _log_fp is not None:
            _log_fp.flush()
    except Exception:
        pass


class _NullStream(object):
    """无控制台的 exe 里 sys.stdout/stderr 是 None，第三方库一 print 就会崩。
    这里给个黑洞流顶上，保证程序不会因为"打印"而挂掉。"""

    encoding = "utf-8"
    errors = "replace"

    def write(self, s):
        return len(s) if s else 0

    def flush(self):
        pass

    def isatty(self):
        return False


def _make_console_safe():
    try:
        if sys.stdout is None:
            sys.stdout = _NullStream()
        if sys.stderr is None:
            sys.stderr = _NullStream()
    except Exception:
        pass


def _install_faulthandler():
    if _crash_fp is None:
        return
    try:
        # 段错误/解释器级崩溃时，faulthandler 会把所有线程堆栈直接写进日志
        faulthandler.enable(file=_crash_fp, all_threads=True)
    except Exception:
        pass


def _install_excepthooks():
    def _hook(exc_type, exc, tb):
        try:
            logger.critical("未捕获异常：%s: %s", getattr(exc_type, "__name__", exc_type), exc,
                            exc_info=(exc_type, exc, tb))
            _flush_all()
        except Exception:
            pass
        try:
            if sys.__excepthook__ is not None:
                sys.__excepthook__(exc_type, exc, tb)
        except Exception:
            pass

    try:
        sys.excepthook = _hook
    except Exception:
        pass

    def _thread_hook(args):
        try:
            name = getattr(getattr(args, "thread", None), "name", "?")
            logger.critical("子线程「%s」未捕获异常：%s: %s", name,
                            getattr(args.exc_type, "__name__", args.exc_type), args.exc_value,
                            exc_info=(args.exc_type, args.exc_value, args.exc_traceback))
            _flush_all()
        except Exception:
            pass

    try:
        threading.excepthook = _thread_hook
    except Exception:
        pass


def _install_tk_hook():
    """tkinter 会自己吞掉按钮回调里的异常，只往 stderr 打（无控制台=看不见）。
    接管 report_callback_exception，把这些异常也记进日志。"""
    try:
        import tkinter
    except Exception:
        return
    try:
        original = tkinter.Tk.report_callback_exception

        def _report(self, exc, val, tb):
            try:
                logger.error("界面回调异常（按钮/定时器里的代码出错了）：%s: %s",
                             getattr(exc, "__name__", exc), val, exc_info=(exc, val, tb))
                _flush_all()
            except Exception:
                pass
            try:
                original(self, exc, val, tb)
            except Exception:
                pass

        tkinter.Tk.report_callback_exception = _report
    except Exception:
        pass


# ==================== 三、卡死看门狗 ====================
def _heartbeat():
    global _last_beat
    _last_beat = time.time()


def _start_watchdog():
    global _watchdog_started
    if _watchdog_started:
        return
    _watchdog_started = True
    _heartbeat()

    def _run():
        last_dump = 0.0
        while not _watchdog_stop.wait(5.0):
            try:
                gap = time.time() - _last_beat
                if gap >= HANG_WARN_SECONDS and (time.time() - last_dump) >= HANG_WARN_SECONDS:
                    last_dump = time.time()
                    logger.error("界面已经 %.0f 秒没有响应，下面把所有线程的调用堆栈写进日志"
                                 "（主线程停在哪一行，就是卡在哪）", gap)
                    _flush_all()
                    if _crash_fp is not None:
                        _crash_fp.write("\n===== 卡死检测 %s：无响应 %.0f 秒 pid=%s =====\n"
                                        % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                           gap, os.getpid()))
                        _crash_fp.flush()
                        faulthandler.dump_traceback(file=_crash_fp, all_threads=True)
                        _crash_fp.flush()
            except Exception:
                pass

    t = threading.Thread(target=_run, name="watchdog", daemon=True)
    t.start()


# ==================== 四、给用户取日志的入口 ====================
def _reveal(path):
    """在资源管理器里定位并选中这个文件，比只打开文件夹更省事。"""
    try:
        import subprocess
        subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
        return True
    except Exception:
        return False


def open_log_folder():
    """打开日志看内容。

    程序运行期间日志文件是被本进程独占的（Windows 下记事本会提示“文件被占用”），
    这时先导一份副本到临时目录再打开，保证任何情况下都点得开。
    """
    if not _log_file_path:
        messagebox.showwarning("提示", "这次没能创建日志文件，请检查磁盘空间或权限。")
        return
    if _log_file_readable():
        try:
            os.startfile(_log_file_path)
            return
        except Exception:
            pass
    import tempfile
    tmp = os.path.join(tempfile.gettempdir(), "%s 日志副本.txt" % APP_NAME)
    if _copy_log_to(tmp):
        try:
            os.startfile(tmp)
            return
        except Exception as e:
            messagebox.showwarning("提示", "打开日志失败：%s" % e)
            return
    messagebox.showwarning("提示", "日志暂时打不开，请改用“日志存到桌面”。")


def _desktop_dir():
    for cand in (os.path.join(os.path.expanduser("~"), "Desktop"),
                 os.path.join(os.environ.get("USERPROFILE", ""), "Desktop")):
        if cand and os.path.isdir(cand):
            return cand
    return os.path.expanduser("~")


def save_log_to_desktop():
    """一键把日志复制到桌面，同事直接把这个文件发出来就行，不用进目录找。"""
    if not _log_file_path:
        messagebox.showwarning("提示", "还没有生成日志文件。")
        return
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    target = os.path.join(_desktop_dir(), "%s %s 日志_%s.txt" % (APP_NAME, VERSION, stamp))
    ok = False
    for _ in range(6):
        if _copy_log_to(target):
            ok = True
            break
        time.sleep(0.25)
    if not ok:
        messagebox.showerror("错误", "复制日志失败（文件可能正被占用）。\n"
                                     "可以稍后再点一次，或直接到程序目录复制。")
        return
    messagebox.showinfo("已存到桌面",
                        "日志已复制到桌面：\n%s\n\n把这个文件直接发出去就行。" % target)
    _reveal(target)


# ---------- 源码/图标备份：打包时一起塞进 exe，这里负责把它们写出来 ----------
def _find_source_text():
    """找打包时一起带进来的源码，返回 (文本, 来源路径)；找不到返回 (None, None)。"""
    candidates = []
    mei = getattr(sys, "_MEIPASS", None)          # onefile / onedir 的解包目录
    if mei:
        candidates.append(os.path.join(mei, SOURCE_BACKUP_NAME))
    try:
        candidates.append(os.path.join(_program_dir(), SOURCE_BACKUP_NAME))
        candidates.append(os.path.join(_program_dir(), "_internal", SOURCE_BACKUP_NAME))
    except Exception:
        pass
    if not getattr(sys, "frozen", False):
        try:
            candidates.append(os.path.abspath(__file__))   # 从源码直接运行
        except Exception:
            pass

    for p in candidates:
        try:
            if p and os.path.exists(p):
                # newline="" 表示不做换行转换，保证导出的源码和打进 exe 的那份逐字节一致
                with open(p, "r", encoding="utf-8", errors="replace", newline="") as f:
                    txt = f.read()
                # 粗校验：确实是本工具的源码，别把别的东西当源码导出去
                if len(txt) > 1000 and "VERSION" in txt and "def main(" in txt:
                    return txt, p
        except Exception:
            continue
    return None, None


def _find_icon_bytes():
    """找打包时一起带进来的图标，返回 (字节, 来源)；找不到返回 (None, None)。"""
    candidates = []
    mei = getattr(sys, "_MEIPASS", None)
    if mei:
        candidates.append(os.path.join(mei, ICON_BACKUP_NAME))
    try:
        candidates.append(os.path.join(_program_dir(), ICON_BACKUP_NAME))
        candidates.append(os.path.join(_program_dir(), "_internal", ICON_BACKUP_NAME))
        if not getattr(sys, "frozen", False):
            # 从源码运行时，工程目录里那个 .ico 就是图标
            for n in os.listdir(_program_dir()):
                if n.lower().endswith(".ico"):
                    candidates.append(os.path.join(_program_dir(), n))
    except Exception:
        pass

    for p in candidates:
        try:
            if p and os.path.exists(p):
                with open(p, "rb") as f:
                    data = f.read()
                if len(data) > 100 and data[:4] == b"\x00\x00\x01\x00":   # ICO 文件头
                    return data, p
        except Exception:
            continue
    return None, None


def _export_source_to(target):
    """把内置源码写成单个 .py；成功返回 True。"""
    txt, _ = _find_source_text()
    if not txt:
        return False
    try:
        with open(target, "w", encoding="utf-8", newline="") as f:
            f.write(txt)
        return True
    except Exception:
        return False


def _export_bundle(folder):
    """导出完整备份目录（源码 + 图标），返回 (是否成功, 说明)。"""
    txt, srcpath = _find_source_text()
    if not txt:
        return False, "这个 exe 里没有内置源码"
    try:
        os.makedirs(folder, exist_ok=True)
        py_path = os.path.join(folder, "源代码 %s 增加运行日志.py" % VERSION)
        with open(py_path, "w", encoding="utf-8", newline="") as f:
            f.write(txt)
    except Exception as e:
        return False, "写源码失败：%s" % e

    icon_data, _ = _find_icon_bytes()
    if icon_data:
        try:
            with open(os.path.join(folder, "软件图标.ico"), "wb") as f:
                f.write(icon_data)
        except Exception:
            pass
    return True, folder


def export_source():
    txt, srcpath = _find_source_text()
    if not txt:
        messagebox.showwarning(
            "提示",
            "这个 exe 里没有内置源码。\n\n"
            "多半是用旧脚本或手动 PyInstaller 打的包。\n"
            "请用工程目录里的「打包成EXE就双击我 V1.7.0.bat」重新打包，"
            "它会把源码和图标一起打进去。")
        return
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    folder = os.path.join(_desktop_dir(), "%s %s 源码备份_%s" % (APP_NAME, VERSION, stamp))
    ok, info = _export_bundle(folder)
    if not ok:
        messagebox.showerror("错误", "导出失败：%s" % info)
        return
    icon_data, _ = _find_icon_bytes()
    logger.info("已导出源码备份 -> %s（源码来源 %s）", folder, srcpath)
    messagebox.showinfo(
        "源码已导出",
        "已保存到桌面的文件夹：\n%s\n\n里面有两样东西：\n"
        "  · 源代码 %s 增加运行日志.py（当前版本完整源码）\n"
        "  · %s\n\n拷回工程目录就能继续改、继续打包。"
        % (folder, VERSION, "软件图标.ico" if icon_data else "（这个包没带图标，可忽略）"))
    _reveal(folder)


def _cli_export_source():
    """命令行导出源码：界面打不开时也能把源码救出来。
    用法：程序.exe --export-source [目标目录 或 目标.py]"""
    args = sys.argv[1:]
    if not args or args[0] != "--export-source":
        return False
    target = args[1] if len(args) > 1 else None

    if target and target.lower().endswith((".py", ".txt")):
        ok = _export_source_to(target)
        note = target
    else:
        folder = target or os.path.join(
            _desktop_dir(),
            "%s %s 源码备份_%s" % (APP_NAME, VERSION,
                                   datetime.datetime.now().strftime("%Y%m%d_%H%M%S")))
        ok, note = _export_bundle(folder)
    logger.info("命令行导出源码 -> %s 成功=%s", note, ok)
    return True


def maybe_notice_previous_crash(root):
    """上次异常退出的话，启动后提示一次，并可直接打开日志。"""
    global _notice_done
    if _notice_done or not _prev_was_abnormal:
        return
    _notice_done = True
    try:
        yes = messagebox.askyesno(
            APP_NAME + " 提示",
            "上次运行没有正常结束（闪退或被强制关闭）。\n\n"
            "上次的日志末尾已经自动保留在本次日志的开头，没有丢失：\n"
            "%s\n\n现在打开这份日志吗？" % (_log_file_path or ""),
            parent=root)
        if yes:
            _reveal(_log_file_path)
    except Exception:
        pass


def _show_startup_error():
    """连 tkinter 都可能没起来时的最后一道提示。"""
    text = ("程序依赖加载失败，无法启动。\n\n"
            "日志文件：%s\n" % (_log_file_path or "（未能创建）"))
    try:
        import tkinter.messagebox as _mb
        _mb.showerror(APP_NAME + " 启动失败", text)
    except Exception:
        pass


# ==================== 五、启动 ====================
def _bootstrap():
    global _start_time_text
    _start_time_text = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    ok, prev_tail, prev_clean = False, "", True
    try:
        ok, prev_tail, prev_clean = _build_handlers()
    except Exception:
        ok = False

    _make_console_safe()
    _install_faulthandler()
    _install_excepthooks()
    atexit.register(_on_exit)

    if not ok:
        return
    try:
        logger.info("=" * 78)
        logger.info("%s %s 启动（本次日志已按规则清空重写）", APP_NAME, VERSION)
        logger.info("进程号=%s  冻结(exe)=%s", os.getpid(), getattr(sys, "frozen", False))
        logger.info("程序路径=%s", sys.executable)
        logger.info("工作目录=%s", os.getcwd())
        logger.info("Python=%s %s  %s", platform.python_version(),
                    platform.architecture()[0], platform.platform())
        logger.info("日志文件=%s", _log_file_path)
        logger.info("TEMP=%s", os.environ.get("TEMP"))
        logger.info("文件编码=%s", sys.getfilesystemencoding())

        if not prev_clean:
            logger.warning("上一次运行没有正常结束（闪退 / 被强制结束 / 断电）。"
                           "本次日志虽然重写了，但下面先把上次的末尾原样搬过来留档，"
                           "证据不会丢：")
            lines = [ln for ln in prev_tail.splitlines() if ln.strip()][-CARRY_OVER_LINES:]
            if lines:
                logger.warning("---- 上一次运行日志的末尾（共 %d 行，原文照搬）----", len(lines))
                for ln in lines:
                    logger.warning("历史| %s", ln)
                logger.warning("---- 上一次运行日志末尾结束，以下是本次运行 ----")
        elif prev_tail.strip():
            logger.info("上一次运行正常结束。")
        else:
            logger.info("这是第一次运行（没有旧日志）。")

        _flush_all()
    except Exception:
        pass


def _on_exit():
    try:
        logger.info("进程退出（atexit 触发）")
    except Exception:
        pass
    _flush_all()


_bootstrap()

# 第三方依赖放在日志初始化之后导入：万一环境缺库，日志里能看到原因
try:
    import pandas as pd
    import openpyxl
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import PatternFill, Alignment, Font
    from openpyxl.utils import get_column_letter
    import tkinter as tk
    from tkinter import messagebox, filedialog, BooleanVar, Entry, Label, Frame, Checkbutton, Button
    from pathlib import Path
    import re
    import io
    import csv
    import hashlib
    import random
    import string
except BaseException:
    logger.critical("依赖加载失败，程序无法启动", exc_info=True)
    _flush_all()
    _show_startup_error()
    raise

try:
    logger.info("依赖版本：pandas=%s openpyxl=%s Tk=%s",
                getattr(pd, "__version__", "?"), getattr(openpyxl, "__version__", "?"),
                getattr(tk, "TkVersion", "?"))
    logger.info("依赖加载完成")
except Exception:
    pass


# ==================== 六、运行日志（沿用 V1.6.9 的用法） ====================
log_entries = []


def log(msg):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = "[%s] %s" % (timestamp, msg)
    try:
        logger.info("%s", msg)
    except Exception:
        pass
    log_entries.append(line)
    # 有控制台就同时打印；打包成 exe（无控制台）时这里什么都不做，但绝不允许报错
    if sys.stdout is not None:
        try:
            print(line, flush=True)
        except Exception:
            pass


def log_exc(msg):
    """记录一条错误，并把完整堆栈一起写进日志（V1.6.9 只记了错误信息，定位不了位置）。"""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    tb_text = traceback.format_exc()
    try:
        logger.error("%s\n%s", msg, tb_text)
    except Exception:
        pass
    log_entries.append("[%s] %s" % (timestamp, msg))
    for ln in tb_text.rstrip().splitlines():
        log_entries.append("    " + ln)
    if sys.stdout is not None:
        try:
            print(tb_text, file=sys.stderr)
        except Exception:
            pass


# ---------- 诊断信息（会显示在“处理完成”提示框里，避免静默跳过） ----------
diagnostics = []

# 当前工作表的处理上下文：表头都识别不出来时（多半不是流水表），
# 只给一条总结性提示，不把每个区块的细节都塞进提示框。
_sheet_ctx = {"sheet": "", "quiet": False, "header_score": 0}


def warn(msg):
    """记录一条用户可见的警告：同时写入运行日志和本次处理的诊断列表。"""
    if _sheet_ctx.get("quiet"):
        log("（细节）" + msg)
        return
    log("警告：" + msg)
    diagnostics.append(msg)


def export_log():
    if not log_entries:
        messagebox.showinfo("提示", "暂无日志记录。")
        return
    file_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
        title="保存日志文件"
    )
    if not file_path:
        return
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write("%s %s 运行日志\n" % (APP_NAME, VERSION))
            f.write("导出时间: %s\n" % datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            f.write("本次日志文件: %s\n" % (_log_file_path or "（未能创建）"))
            f.write("=" * 80 + "\n")
            f.write("\n".join(log_entries))
        messagebox.showinfo("完成", "日志已保存至：%s" % file_path)
    except Exception as e:
        messagebox.showerror("错误", "保存日志失败：%s" % e)


# ---------- 脱敏映射管理 ----------
class DesensitizeMapper:
    def __init__(self):
        self.orig_to_mask = {}
        self.mask_to_orig = {}
        self.charset_digit = list(string.digits)
        self.charset_lower = list(string.ascii_lowercase)
        self.charset_upper = list(string.ascii_uppercase)
        self.charset_ch = [chr(i) for i in range(0x4e00, 0x9fff) if chr(i).isalpha()]

    def _random_char(self, ch_type, seed_str):
        # V1.6.9：以前用内置 hash()，它每次启动进程都会随机加盐，
        # 所以同一份数据两次运行会得到不同的脱敏结果（与“确定性脱敏”的说明不符）。
        # 改用 md5 取种子，跨进程/跨机器都稳定；同时不再污染全局 random 状态。
        h = int(hashlib.md5(seed_str.encode('utf-8')).hexdigest()[:8], 16)
        rnd = random.Random(h)
        if ch_type == 'digit':
            result = rnd.choice(self.charset_digit)
        elif ch_type == 'lower':
            result = rnd.choice(self.charset_lower)
        elif ch_type == 'upper':
            result = rnd.choice(self.charset_upper)
        elif ch_type == 'ch':
            result = rnd.choice(self.charset_ch)
        else:
            result = seed_str
        return result

    def _generate_masked(self, s, salt=0):
        masked_chars = []
        for i, ch in enumerate(s):
            seed = f"{s}_{i}" if not salt else f"{s}_{i}_{salt}"
            if ch.isdigit():
                masked_chars.append(self._random_char('digit', seed))
            elif ch.islower():
                masked_chars.append(self._random_char('lower', seed))
            elif ch.isupper():
                masked_chars.append(self._random_char('upper', seed))
            elif '\u4e00' <= ch <= '\u9fff':
                masked_chars.append(self._random_char('ch', seed))
            else:
                masked_chars.append(ch)
        return ''.join(masked_chars)

    def mask(self, original):
        if original is None:
            return None
        s = str(original)
        if s not in self.orig_to_mask:
            masked = self._generate_masked(s)
            salt = 0
            # 两个不同原文万一撞到同一个脱敏串，还原就会串号；撞了换个盐重来
            while masked in self.mask_to_orig and self.mask_to_orig[masked] != s and salt < 50:
                salt += 1
                masked = self._generate_masked(s, salt)
            self.orig_to_mask[s] = masked
            self.mask_to_orig[masked] = s
        return self.orig_to_mask[s]

    def unmask(self, masked):
        if masked is None:
            return None
        return self.mask_to_orig.get(str(masked), masked)

    def save_mapping(self, file_path):
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump({
                'orig_to_mask': self.orig_to_mask,
                'mask_to_orig': self.mask_to_orig
            }, f, ensure_ascii=False, indent=2)

    def load_mapping(self, file_path):
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        self.orig_to_mask = data['orig_to_mask']
        self.mask_to_orig = data['mask_to_orig']

desensitize_mapper = DesensitizeMapper()

# ---------- 辅助函数 ----------
_FULLWIDTH_TABLE = str.maketrans('０１２３４５６７８９．，－＋％', '0123456789.,-+%')
_CURRENCY_RE = re.compile(r'[¥￥$€£]')
_AMOUNT_UNIT_RE = re.compile(r'(元整|元|圆|人民币|RMB|CNY|USD)$', re.IGNORECASE)

def _norm_key(p):
    try:
        return os.path.normcase(os.path.abspath(str(p)))
    except Exception:
        return str(p)


def parse_amount(cell_value):
    """把单元格值解析成 float；不像金额就返回 None。

    只接受“数字 + 常见金额格式”（货币符号、千分位、空格、全角数字、括号负数）。
    V1.6.9：以前 is_numeric_cell 会把“导出说明 第1行”“订单A123”这类含数字的文本
    也算成金额，导致金额列可能识别到文本列上，进而把文本改成数字。
    """
    if cell_value is None or isinstance(cell_value, bool):
        return None
    if isinstance(cell_value, (int, float)):
        val = float(cell_value)
        return val if abs(val) <= 1e12 else None
    s = str(cell_value).strip().translate(_FULLWIDTH_TABLE)
    if not s:
        return None
    negative = False
    if s.startswith('(') and s.endswith(')'):
        negative = True
        s = s[1:-1].strip()
    s = _CURRENCY_RE.sub('', s)
    s = s.replace('\u00a0', '').replace('\u3000', '').replace(' ', '')
    s = _AMOUNT_UNIT_RE.sub('', s)
    if s.endswith('%'):
        s = s[:-1]
    if s in ('', '-', '+', '.', '--'):
        return None
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            # 欧式写法 1.234,56
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        parts = s.split(',')
        if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
            s = ''.join(parts)
        else:
            s = s.replace(',', '.')
    if not re.fullmatch(r'[+-]?\d+(\.\d+)?', s):
        return None
    val = float(s)
    if negative:
        val = -val
    return val if abs(val) <= 1e12 else None

def is_numeric_cell(cell_value):
    return parse_amount(cell_value) is not None

def clean_money_cell(cell_value):
    """清理金额列：能解析成数字就返回 float，否则原样返回，避免改坏非金额内容。"""
    if cell_value is None:
        return None
    if isinstance(cell_value, str) and cell_value.strip() == '':
        return None
    val = parse_amount(cell_value)
    return val if val is not None else cell_value

def cell_text(value):
    """把单元格值转成干净的文本（整数型浮点不带 .0），用于对手方名称/ID。"""
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()

def col_fill_ratio(ws, col, start_row, end_row):
    """某列在指定行区间内的非空占比，用于在同名/多个候选列里挑真正有数据的那个。"""
    if end_row < start_row:
        return 0.0
    total = end_row - start_row + 1
    filled = 0
    for row in range(start_row, end_row + 1):
        val = ws.cell(row, col).value
        if val is not None and str(val).strip() != '':
            filled += 1
    return filled / total if total else 0.0

INVALID_SHEET_TITLE_CHARS = re.compile(r'[\\/*?:\[\]]')
STAT_SHEET_SUFFIX = "对手方统计汇总"

def _trim_dangling_bracket(head):
    """截断后如果留下落单的左括号，把它连同后面的一起去掉，避免出现
    “微信支付账单(20220101-20221231” 这种看起来像出错的表名。"""
    while head and head.count('(') > head.count(')'):
        head = head[:head.rfind('(')]
    return head.rstrip(' -_·、')

def make_sheet_title(base, suffix, existing_titles):
    """生成 Excel 合法的工作表名：去掉非法字符、总长 ≤31、不与已有表重名。

    V1.6.9：以前直接把“原工作表名 + 对手方统计汇总”当新表名。名称超过 31 字符时
    （例如 微信支付账单(20220101-20221231) 就是 25 字符），Excel 打开会提示修复并
    丢掉这张表——这正是“偶发看不到统计汇总工作表”的原因之一。
    """
    base = INVALID_SHEET_TITLE_CHARS.sub('_', str(base or '')).strip().strip("'")
    suffix = INVALID_SHEET_TITLE_CHARS.sub('_', str(suffix or '')).strip("'")
    existing = set(existing_titles or [])
    limit = 31
    room = limit - len(suffix)
    head_base = _trim_dangling_bracket(base[:room]) if room > 0 else ''

    title = f"{head_base}{suffix}"
    idx = 1
    while not title or title in existing:
        tail = f"{suffix}_{idx}"
        head = head_base[:max(limit - len(tail), 0)]
        title = f"{head}{tail}"
        idx += 1
    return title

def is_stat_sheet_name(sheet_name):
    """判断是不是本工具生成的统计汇总表（重复处理时不能再当成数据表统计）。"""
    return bool(re.search(r'对手方统计汇总(_\d+)?$', str(sheet_name or '')))

def _claimed_stat_titles(wb):
    """本轮已经生成的统计表名（同一个工作簿里，防止误删别的数据表刚生成的统计表）。"""
    claimed = getattr(wb, "_dsh_stat_titles", None)
    if claimed is None:
        claimed = set()
        try:
            setattr(wb, "_dsh_stat_titles", claimed)
        except Exception:
            pass
    return claimed

def remove_stale_stat_sheet(wb, title):
    """如果这个数据表上次处理生成的统计表还在，先删掉它。

    V1.6.9：重复处理已处理文件时，旧统计表会和新统计表同时留在文档里
    （新表还被迫改名成 _1、_2…），这里先把同名旧表删掉再生成。
    """
    if not title or title not in wb.sheetnames:
        return None
    if title in _claimed_stat_titles(wb) or not is_stat_sheet_name(title):
        return None
    wb.remove(wb[title])
    return title

def sheet_has_enough_data(ws, min_rows=2):
    """粗略判断工作表里有没有值得处理的数据，避免对空白/装饰性工作表发一堆告警。"""
    max_row = min(ws.max_row, 300)
    max_col = min(ws.max_column, 120)
    non_empty_rows = 0
    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        if any(cell.value is not None and str(cell.value).strip() != '' for cell in row):
            non_empty_rows += 1
            if non_empty_rows >= min_rows:
                return True
    return False

def _empty_stats():
    return {'交易次数': 0, '流入金额': 0.0, '流出金额': 0.0, '金额合计': 0.0}

def aggregate_counterparty(rows):
    """把对手方明细汇总成 名称统计 / 名称→ID统计 / 总计。

    V1.6.9：名称与ID分开累计。对手方ID为空时不再造一个空的“对手方账户ID”分块，
    也不再因为 ID 为空把名称的合计算成 0。
    """
    name_stats = {}
    name_id_stats = {}
    totals = _empty_stats()
    for d in rows:
        name = d.get('对手方名称') or ''
        uid = cell_text(d.get('对手方ID'))
        inflow = float(d.get('inflow') or 0.0)
        outflow = float(d.get('outflow') or 0.0)
        amount = float(d.get('abs_amount') or 0.0)

        stat = name_stats.setdefault(name, _empty_stats())
        stat['交易次数'] += 1
        stat['流入金额'] += inflow
        stat['流出金额'] += outflow
        stat['金额合计'] += amount

        totals['交易次数'] += 1
        totals['流入金额'] += inflow
        totals['流出金额'] += outflow
        totals['金额合计'] += amount

        if uid:
            id_map = name_id_stats.setdefault(name, {})
            id_stat = id_map.setdefault(uid, _empty_stats())
            id_stat['交易次数'] += 1
            id_stat['流入金额'] += inflow
            id_stat['流出金额'] += outflow
            id_stat['金额合计'] += amount
    return name_stats, name_id_stats, totals

# ---------- 自动识别表头行 ----------
HEADER_KEYWORDS = [
    '用户ID', '交易单号', '大单号', '账号名称', '借贷类型', '交易业务类型',
    '交易用途类型', '交易时间', '金额', '交易额', '发生额', '交易金额',
    '账户余额', '银行卡号', '网银联单号', '对手', '对方', '对面',
    '第三方账户', '银行名称', '基金公司', '间联', '备注'
]

def _header_score(ws, row, scan_cols):
    score = 0
    for col in range(1, scan_cols + 1):
        cell_val = ws.cell(row, col).value
        if cell_val is None:
            continue
        cell_str = str(cell_val).strip()
        if not cell_str:
            continue
        for kw in HEADER_KEYWORDS:
            if kw in cell_str:
                score += 1
                break
    return score

def detect_header_row(ws, max_row, max_col, max_scan_rows=80):
    """识别表头行。

    V1.6.9：以前只扫前 30 行，表头在 30 行之后的导出文件会被静默按第 1 行处理，
    整张表随之被跳过（输出文档里就没有统计汇总表）。现在先扫前 80 行，
    仍找不到明显表头时继续向下深挖到 500 行，并明确写告警。
    """
    scan_limit = min(max_row, max_scan_rows)
    scan_cols = min(max_col, 120)
    best_row = 1
    best_score = 0
    for row in range(1, scan_limit + 1):
        score = _header_score(ws, row, scan_cols)
        if score > best_score:
            best_score = score
            best_row = row
        if best_score >= 8:
            break
    if best_score < 3 and max_row > scan_limit:
        deep_limit = min(max_row, 500)
        for row in range(scan_limit + 1, deep_limit + 1):
            score = _header_score(ws, row, scan_cols)
            if score > best_score:
                best_score = score
                best_row = row
                if best_score >= 3:
                    break
    if best_score < 2:
        _sheet_ctx["header_score"] = best_score
        log(f"未识别到明显的表头行（前 {min(max_row, 500)} 行关键词命中数仅 {best_score}），"
            f"已按第 1 行处理；该工作表很可能整体被跳过")
        return 1
    _sheet_ctx["header_score"] = best_score
    if best_score < 4:
        log(f"表头行把握不大：第 {best_row} 行（关键词命中数 {best_score}），识别结果可能有偏差")
    log(f"自动识别表头行: 第 {best_row} 行 (关键词命中数: {best_score})")
    return best_row

# ---------- 识别列（基于表头行） ----------
def identify_money_columns_auto(ws, max_row, max_col, header_row, start_row=None):
    start_row = header_row + 1 if start_row is None else start_row
    candidate_cols = []
    for col in range(1, max_col + 1):
        total = 0
        numeric = 0
        for row in range(start_row, max_row + 1):
            cell = ws.cell(row, col)
            if cell.value is None:
                continue
            total += 1
            if is_numeric_cell(cell.value):
                numeric += 1
        if total > 0 and numeric / total >= 0.3:
            candidate_cols.append(col)
    if not candidate_cols:
        return None
    col_names = {}
    for col in candidate_cols:
        val = ws.cell(header_row, col).value
        if val:
            col_names[col] = str(val).strip()
    money_keywords = ['金额', '交易额', '发生额', '交易金额', '发生金额']
    bad_keywords = ['余额', '卡号', '单号', '账号', 'ID', '编号', '时间', '日期']
    best_col = None
    best_score = None
    for col in candidate_cols:
        name = col_names.get(col, '')
        score = 0
        for kw in money_keywords:
            if kw in name:
                score += 10
        for kw in bad_keywords:
            if kw in name:
                score -= 5
        if best_score is None or score > best_score:
            best_score = score
            best_col = col
    if best_score is not None and best_score <= 0:
        warn(f"未找到标题带“金额/交易额/发生额”的列，暂用第 {best_col} 列作为金额列，请核对")
    return [best_col]

def identify_counterparty_column_auto(ws, max_col, header_row, start_row=None, end_row=None):
    """自动识别对手方名称列。

    V1.6.9：
      - 关键词扩充：对手/对方/对面/交易相对人/第三方账户/商户/收款/付款/交易方；
      - 排除 ID、卡号、账号、单号、时间、金额、余额、银行名称 等明显不是“名称”的列
        （以前会把“对手方ID”“对手侧银行名称”当成名称列，导致统计表不生成或名称全是ID）；
      - 多个候选时按“实际有数据的比例”优先，避免选中一列全空的名称列。
    """
    col_names = {}
    for col in range(1, max_col + 1):
        val = ws.cell(header_row, col).value
        if val:
            col_names[col] = str(val).strip()
    if start_row is None:
        start_row = header_row + 1
    if end_row is None:
        end_row = ws.max_row

    relation_keywords = ['对手', '对方', '对面', '交易相对人', '相对人', '第三方账户',
                         '商户', '收款', '付款', '交易方', '收付方']
    name_hints = ['名称', '户名', '姓名', '全称']
    exclude_keywords = ['ID', 'id', '卡号', '账号', '账户号', '单号', '编号', '流水号',
                        '时间', '日期', '金额', '余额', '银行名称', '支行', '机构',
                        '地址', '备注', '手机号', '证件']

    candidates = []
    for col, name in col_names.items():
        if not any(kw in name for kw in relation_keywords):
            continue
        if any(kw in name for kw in exclude_keywords):
            continue
        is_name_like = any(h in name for h in name_hints)
        fill = col_fill_ratio(ws, col, start_row, end_row)
        chinese = sum(1 for ch in name if '\u4e00' <= ch <= '\u9fff')
        candidates.append((col, name, is_name_like, fill, chinese))

    if not candidates:
        log("未找到包含“对手/对方/对面/交易相对人/商户”等关键词的对手方名称列")
        return None

    name_like = [c for c in candidates if c[2]] or candidates
    name_like.sort(key=lambda c: (c[3], c[4], -len(c[1])), reverse=True)
    best = name_like[0]
    if best[3] == 0:
        log(f"注意：识别到的对手方名称列「{best[1]}」在本区块内没有任何数据")
    log(f"识别到对手方名称列: {best[1]} (列索引 {best[0]}，填充率 {best[3]:.0%})")
    return best[0]

def identify_counterparty_id_column_auto(ws, max_col, header_row):
    """自动识别对手方ID列"""
    col_names = {}
    for col in range(1, max_col + 1):
        val = ws.cell(header_row, col).value
        if val:
            col_names[col] = str(val).strip()
    id_keywords = ['对手方ID', '对手ID', '对方ID', '交易方ID', '账户ID',
                   '对手方账号', '对手账号', '对方账号', '对手方账户号', '交易对方账号']
    for kw in id_keywords:
        for col, name in col_names.items():
            if name == kw or (kw in name and len(name) - len(kw) <= 2):
                log(f"识别到对手方ID列: {name} (列索引 {col})")
                return col
    log("未找到对手方ID列，将不使用ID分项统计")
    return None

def detect_direction_column_auto(ws, max_col, header_row):
    direction_keywords = ['借贷类型', '交易方向', '借贷', '方向', '借/贷', '出入']
    for col in range(1, max_col + 1):
        val = ws.cell(header_row, col).value
        if val:
            name = str(val).strip()
            for kw in direction_keywords:
                if kw in name:
                    log(f"识别到方向列: {name} (列索引 {col})")
                    return col
    log("未找到方向列，将根据金额正负判断流向")
    return None

def get_direction_from_cell(cell_value):
    if cell_value is None:
        return None
    s = str(cell_value).strip()
    if s in ['入', '贷', '收入', '增加', '+']:
        return 'inflow'
    elif s in ['出', '借', '支出', '减少', '-']:
        return 'outflow'
    else:
        if any(k in s for k in ['入', '贷', '收入']):
            return 'inflow'
        elif any(k in s for k in ['出', '借', '支出']):
            return 'outflow'
    return None

def find_column_by_title(ws, max_col, header_row, title_list, start_row=None, end_row=None):
    """按用户给定的标题找列。

    V1.6.9：银行导出里同名重复列很常见（例如两列都叫“第三方账户名称”），
    以前盲取最左边那列，如果那列是空的，就会一条明细都收集不到、连统计表都不生成。
    现在：精确匹配优先，其次取实际有数据的列。
    """
    col_names = {}
    for col in range(1, max_col + 1):
        val = ws.cell(header_row, col).value
        if val:
            col_names[col] = str(val).strip()
    if start_row is None:
        start_row = header_row + 1
    if end_row is None:
        end_row = ws.max_row
    matches = []
    seen = set()
    for user_title in title_list:
        for col, name in col_names.items():
            if col in seen:
                continue
            if name == user_title or user_title in name:
                seen.add(col)
                matches.append((col, name == user_title))
    if not matches:
        return None
    matches.sort(key=lambda m: (m[1], col_fill_ratio(ws, m[0], start_row, end_row)), reverse=True)
    return matches[0][0]

# ---------- 区块处理 ----------
def find_blocks(ws, max_row, max_col, header_row):
    key_cols = []
    for col in range(1, min(max_col, 10) + 1):
        if ws.cell(header_row, col).value is not None:
            key_cols.append(col)
    if not key_cols:
        return [(header_row + 1, max_row, 1)]

    header_values = {}
    for col in key_cols:
        val = ws.cell(header_row, col).value
        header_values[col] = str(val).strip() if val is not None else ""

    candidate_rows = []
    for row in range(header_row + 1, max_row + 1):
        has_content = any(ws.cell(row, col).value is not None for col in range(1, max_col+1))
        if not has_content:
            continue
        match = True
        for col in key_cols:
            cell_val = ws.cell(row, col).value
            cur_val = str(cell_val).strip() if cell_val is not None else ""
            if cur_val != header_values[col]:
                match = False
                break
        if match:
            candidate_rows.append(row)

    if not candidate_rows:
        return [(header_row + 1, max_row, 1)]

    blocks = []
    prev = header_row
    for idx, row in enumerate(candidate_rows):
        start = prev + 1
        end = row - 1
        if start <= end:
            blocks.append((start, end, idx + 1))
        prev = row
    if prev < max_row:
        blocks.append((prev + 1, max_row, len(candidate_rows) + 1))
    return blocks

def process_block(ws, wb, block_start, block_end, block_idx, header_row, enable_sum, enable_internal_stats, collect_data,
                  manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                  enable_desensitize, desensitize_cols, enable_money_clean):
    max_col = ws.max_column
    if block_start > block_end:
        return []
    block_desc = f"第 {block_idx} 区块（第 {block_start}-{block_end} 行）"

    # 确定金额列
    if manual_mode and user_amount_titles:
        amount_col = find_column_by_title(ws, max_col, header_row, user_amount_titles, block_start, block_end)
        if amount_col is None:
            warn(f"{block_desc}：未找到手动指定的金额列标题 {user_amount_titles}，已跳过该区块")
            return []
        money_cols = [amount_col]
        log(f"  使用手动金额列: {user_amount_titles}")
    else:
        money_cols = identify_money_columns_auto(ws, block_end, max_col, header_row, block_start)
        if not money_cols:
            warn(f"{block_desc}：未找到金额列，已跳过该区块（不会进入统计表）")
            return []
        log(f"  自动识别金额列: {ws.cell(header_row, money_cols[0]).value}")

    # 确定对手方ID列（先定 ID，名称列整列为空时可以拿它兜底）
    counterparty_id_col = None
    if manual_mode and user_counterparty_id_titles:
        counterparty_id_col = find_column_by_title(ws, max_col, header_row, user_counterparty_id_titles,
                                                   block_start, block_end)
        if counterparty_id_col is not None:
            log(f"  使用手动对手方ID列: {user_counterparty_id_titles}")
        else:
            log(f"  注意：未找到手动指定的对手方ID列标题 {user_counterparty_id_titles}，将不使用ID分项统计")
    else:
        counterparty_id_col = identify_counterparty_id_column_auto(ws, max_col, header_row)

    # 确定对手方名称列
    if manual_mode and user_counterparty_titles:
        counterparty_col = find_column_by_title(ws, max_col, header_row, user_counterparty_titles,
                                                block_start, block_end)
        if counterparty_col is None:
            warn(f"{block_desc}：未找到手动指定的对手方名称列标题 {user_counterparty_titles}，已跳过该区块")
            return []
        log(f"  使用手动对手方名称列: {user_counterparty_titles}")
    else:
        counterparty_col = identify_counterparty_column_auto(ws, max_col, header_row, block_start, block_end)
        if counterparty_col is None:
            if counterparty_id_col is not None and col_fill_ratio(ws, counterparty_id_col, block_start, block_end) > 0:
                counterparty_col = counterparty_id_col
                warn(f"{block_desc}：未识别到对手方名称列，改用「"
                     f"{ws.cell(header_row, counterparty_id_col).value}」作为对手方名称")
            else:
                warn(f"{block_desc}：未找到对手方名称列，已跳过该区块（不会生成统计汇总表）")
                return []
        else:
            log(f"  自动识别对手方名称列: {ws.cell(header_row, counterparty_col).value}")
            if (counterparty_id_col is not None and counterparty_id_col != counterparty_col
                    and col_fill_ratio(ws, counterparty_col, block_start, block_end) == 0
                    and col_fill_ratio(ws, counterparty_id_col, block_start, block_end) > 0):
                warn(f"{block_desc}：对手方名称列「{ws.cell(header_row, counterparty_col).value}」整列为空，"
                     f"改用「{ws.cell(header_row, counterparty_id_col).value}」显示对手方")

    direction_col = detect_direction_column_auto(ws, max_col, header_row)

    # 清理金额
    if enable_money_clean:
        for col in money_cols:
            for row in range(block_start, block_end + 1):
                cell = ws.cell(row, col)
                if cell.value is None:
                    continue
                cleaned = clean_money_cell(cell.value)
                if cleaned is not None:
                    cell.value = cleaned

    # 脱敏
    if enable_desensitize and desensitize_cols:
        for col in desensitize_cols:
            if col is None:
                continue
            for row in range(block_start, block_end + 1):
                cell = ws.cell(row, col)
                if cell.value is None:
                    continue
                cell.value = desensitize_mapper.mask(cell.value)

    # 收集对手方数据（含ID）—— 必须放在插入求和行之前，
    # 否则 insert_rows 会把行整体下移，收集区间错位。
    collected = []
    if enable_internal_stats or collect_data:
        amount_col = money_cols[0]
        for row in range(block_start, block_end + 1):
            counterparty = ws.cell(row, counterparty_col).value
            name = cell_text(counterparty)
            counterparty_id = ws.cell(row, counterparty_id_col).value if counterparty_id_col is not None else None
            uid = cell_text(counterparty_id)
            if not name:
                # V1.6.9：名称列为空但对手方ID有值时不丢弃该行，用ID兜底，
                # 否则整列名称都空时一条明细都收集不到，统计表直接消失。
                if not uid:
                    continue
                name = uid
            amount_cell = ws.cell(row, amount_col)
            amount = amount_cell.value
            if amount is None:
                continue
            val = parse_amount(amount)
            if val is None:
                continue
            inflow = outflow = 0.0
            if direction_col:
                dir_val = ws.cell(row, direction_col).value
                direction = get_direction_from_cell(dir_val)
                if direction == 'inflow':
                    inflow = abs(val)
                elif direction == 'outflow':
                    outflow = abs(val)
                else:
                    if val > 0:
                        inflow = val
                    elif val < 0:
                        outflow = abs(val)
            else:
                if val > 0:
                    inflow = val
                elif val < 0:
                    outflow = abs(val)
            if inflow == 0 and outflow == 0:
                continue
            collected.append({
                '对手方名称': name,
                '对手方ID': uid,
                'inflow': inflow,
                'outflow': outflow,
                'abs_amount': abs(val),
                'block_index': block_idx
            })

    # 求和行（黄色）—— 查找最后一个有效数据行
    last_data_row = None
    if enable_sum:
        sums = {}
        amount_col = money_cols[0]
        for row in range(block_start, block_end + 1):
            cell = ws.cell(row, amount_col)
            if cell.value is not None:
                val = parse_amount(cell.value)
                if val is not None:
                    last_data_row = row
        for col in money_cols:
            total = 0.0
            for row in range(block_start, block_end + 1):
                val = parse_amount(ws.cell(row, col).value)
                if val is not None:
                    total += abs(val)
            if total != 0:
                sums[col] = total
        if sums and last_data_row is not None:
            ws.insert_rows(last_data_row + 1)
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            for col, val in sums.items():
                cell = ws.cell(row=last_data_row + 1, column=col, value=val)
                cell.fill = yellow_fill

    return collected

# ---------- 统计区块写入函数（符合打样样式，支持独立ID排序和颜色） ----------
def write_stat_block(ws, start_row, title, name_stats, name_id_stats, totals=None, is_total_section=False):
    """写出一段统计区（总计区 / 各区块区）。

    name_stats    : {对手方名称: {'交易次数','流入金额','流出金额','金额合计'}}   ← 左侧 A~E 列
    name_id_stats : {对手方名称: {对手方ID: 同结构}}                               ← 右侧 ID 分块
    totals        : 合计行的数值（dict）；为 None 时不写合计行

    V1.6.9：名称汇总与 ID 汇总分开计算。没有对手方ID时不再画一排空的
    “对手方账户ID”分块（与打样文件一致），也不会因为 ID 为空把合计算成 0。
    """
    current_row = start_row
    name_id_stats = name_id_stats or {}
    max_id_count = 0
    for id_dict in name_id_stats.values():
        max_id_count = max(max_id_count, len(id_dict))
    id_col_groups = max_id_count

    # 第一行：大标题行
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
    title_cell = ws.cell(row=current_row, column=1, value=title)
    title_cell.font = Font(bold=True)
    title_cell.alignment = Alignment(horizontal='center')
    col_offset = 6
    for idx in range(1, id_col_groups + 1):
        start_col = col_offset
        end_col = col_offset + 4
        ws.merge_cells(start_row=current_row, start_column=start_col, end_row=current_row, end_column=end_col)
        cell = ws.cell(row=current_row, column=start_col, value=f"对手方账户ID{idx}")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
        col_offset += 5
    current_row += 1

    # 第二行：子标题行
    left_headers = ['对手方名称', '交易次数', '流入金额', '流出金额', '金额合计']
    for col, h in enumerate(left_headers, start=1):
        cell = ws.cell(row=current_row, column=col, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    col_offset = 6
    sub_headers = ['对手方账户ID', '交易次数', '流入金额', '流出金额', '金额合计']
    for _ in range(id_col_groups):
        for sub_col, sub_h in enumerate(sub_headers, start=col_offset):
            cell = ws.cell(row=current_row, column=sub_col, value=sub_h)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
        col_offset += 5
    current_row += 1

    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    green_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    blue_fill = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
    id_color_sequence = [green_fill, blue_fill]

    if totals is not None:
        ws.cell(row=current_row, column=1, value="合计")
        ws.cell(row=current_row, column=2, value=totals['交易次数'])
        ws.cell(row=current_row, column=3, value=totals['流入金额'])
        ws.cell(row=current_row, column=4, value=totals['流出金额'])
        ws.cell(row=current_row, column=5, value=totals['金额合计'])
        if is_total_section:
            for col in range(1, 6):
                ws.cell(row=current_row, column=col).fill = yellow_fill
        current_row += 1

    sorted_names = sorted(name_stats.keys(),
                          key=lambda x: name_stats[x]['金额合计'], reverse=True)
    for name in sorted_names:
        stats = name_stats[name]
        ws.cell(row=current_row, column=1, value=name)
        ws.cell(row=current_row, column=2, value=stats['交易次数'])
        ws.cell(row=current_row, column=3, value=stats['流入金额'])
        ws.cell(row=current_row, column=4, value=stats['流出金额'])
        ws.cell(row=current_row, column=5, value=stats['金额合计'])
        id_dict = name_id_stats.get(name, {})
        sorted_ids = sorted(id_dict.keys(), key=lambda x: id_dict[x]['金额合计'], reverse=True)
        col_offset = 6
        for idx in range(max_id_count):
            start_col = col_offset
            if idx < len(sorted_ids):
                uid = sorted_ids[idx]
                stat = id_dict[uid]
                ws.cell(row=current_row, column=start_col, value=uid)
                ws.cell(row=current_row, column=start_col+1, value=stat['交易次数'])
                ws.cell(row=current_row, column=start_col+2, value=stat['流入金额'])
                ws.cell(row=current_row, column=start_col+3, value=stat['流出金额'])
                ws.cell(row=current_row, column=start_col+4, value=stat['金额合计'])
            col_offset += 5
        if is_total_section:
            for col in range(1, 6):
                ws.cell(row=current_row, column=col).fill = yellow_fill
        col_offset = 6
        for idx in range(max_id_count):
            color = id_color_sequence[idx % 2]
            for c in range(5):
                ws.cell(row=current_row, column=col_offset + c).fill = color
            col_offset += 5
        current_row += 1

    current_row += 1
    return current_row

def process_worksheet_multi(ws, wb, sheet_name, enable_sum, enable_internal_stats, collect_data,
                           manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                           enable_desensitize, desensitize_cols, enable_money_clean, header_row=None):
    max_row = ws.max_row
    max_col = ws.max_column
    log(f"工作表「{sheet_name}」开始处理（{max_row} 行 × {max_col} 列）")
    if max_row <= 1 or max_col == 0:
        return []
    if not sheet_has_enough_data(ws):
        log(f"工作表「{sheet_name}」没有有效数据，跳过")
        return []

    if header_row is None:
        header_row = detect_header_row(ws, max_row, max_col)
    # 表头都没识别出来，说明这张表多半不是流水表：细节只写日志，不刷屏
    _sheet_ctx["sheet"] = sheet_name
    _sheet_ctx["quiet"] = _sheet_ctx["header_score"] < 2
    blocks = find_blocks(ws, max_row, max_col, header_row)
    if not blocks:
        return []

    all_collected = []
    for block_start, block_end, block_idx in reversed(blocks):
        data = process_block(ws, wb, block_start, block_end, block_idx, header_row, enable_sum, enable_internal_stats, collect_data,
                            manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                            enable_desensitize, desensitize_cols, enable_money_clean)
        all_collected.extend(data)

    if enable_internal_stats:
        if all_collected:
            preferred = make_sheet_title(sheet_name, STAT_SHEET_SUFFIX, [])
            removed = remove_stale_stat_sheet(wb, preferred)
            if removed:
                log(f"  已覆盖上一次生成的统计表：{removed}")
            stats_sheet_name = make_sheet_title(sheet_name, STAT_SHEET_SUFFIX, wb.sheetnames)
            stats_ws = wb.create_sheet(stats_sheet_name)
            _claimed_stat_titles(wb).add(stats_sheet_name)
            current_row = 1

            block_dict = {}
            for item in all_collected:
                block_dict.setdefault(item.get('block_index', 1), []).append(item)

            name_stats, name_id_stats, totals = aggregate_counterparty(all_collected)
            current_row = write_stat_block(stats_ws, current_row, "总计", name_stats, name_id_stats,
                                           totals=totals, is_total_section=True)

            for blk_idx in sorted(block_dict.keys()):
                b_name_stats, b_name_id_stats, b_totals = aggregate_counterparty(block_dict[blk_idx])
                current_row = write_stat_block(stats_ws, current_row, f"区块{blk_idx}",
                                               b_name_stats, b_name_id_stats,
                                               totals=b_totals, is_total_section=False)

            for col in stats_ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                stats_ws.column_dimensions[col_letter].width = min(max_len+2, 20)

            log(f"已生成对手方统计汇总表：{stats_sheet_name}（{len(all_collected)} 条明细）")
            _sheet_ctx["quiet"] = False
        else:
            quiet = _sheet_ctx.get("quiet", False)
            _sheet_ctx["quiet"] = False
            if quiet:
                warn(f"工作表「{sheet_name}」未识别到表头、也没有收集到任何明细，已跳过（看起来不是流水表）")
            else:
                warn(f"工作表「{sheet_name}」没有收集到任何对手方明细，未生成对手方统计汇总表"
                     f"（常见原因：对手方名称列整列为空、金额列识别错误、表头行识别错误；可勾选“手动指定列”重试）")

    if collect_data:
        return [{'对手方名称': d['对手方名称'], '对手方ID': d['对手方ID'], 'inflow': d['inflow'],
                 'outflow': d['outflow'], 'abs_amount': d['abs_amount']} for d in all_collected]
    else:
        return []

# ---------- 文件处理 ----------
def read_csv_grid(file_path):
    """用标准库把 CSV 读成二维列表（不做任何类型推断）。

    V1.6.9：银行/支付平台导出的 CSV 常有“说明行字段数比数据行少”的情况，
    pandas 的 C 解析器会直接报 “Expected 1 fields in line 4, saw 6”，
    整份文件处理失败（也就看不到统计表）。改用标准库 csv 逐行读取，容错更好。
    """
    last_err = None
    text = None
    for enc in ('utf-8-sig', 'gb18030', 'utf-16'):
        try:
            with open(file_path, 'r', encoding=enc, newline='') as f:
                text = f.read()
            break
        except (UnicodeDecodeError, UnicodeError) as e:
            last_err = e
            text = None
    if text is None:
        raise Exception(f"无法识别 CSV 编码，请另存为 UTF-8 或 GBK 后重试（{last_err}）")

    sample = text[:4096]
    delimiter = ','
    if sample.count('\t') > sample.count(','):
        delimiter = '\t'
    elif sample.count(';') > sample.count(','):
        delimiter = ';'
    return list(csv.reader(io.StringIO(text), delimiter=delimiter))

def grid_to_workbook(rows):
    """把二维列表转成工作表（空字符串写成空单元格）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row, start=1):
            if val is None:
                continue
            text = str(val)
            if text.strip() == '':
                continue
            ws.cell(row=r, column=c, value=text)
    return wb

def resolve_desensitize_cols(ws, max_col, header_row, enable_desensitize, desensitize_titles, sheet_label="工作表"):
    desensitize_cols = []
    if enable_desensitize and desensitize_titles:
        for title in desensitize_titles:
            col = find_column_by_title(ws, max_col, header_row, [title])
            if col is not None:
                desensitize_cols.append(col)
            else:
                warn(f"{sheet_label}未找到脱敏列标题 '{title}'，跳过该列")
    return desensitize_cols

def process_csv(file_path, enable_sum, enable_internal_stats, collect_data,
                manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                enable_desensitize, desensitize_titles, enable_money_clean):
    file_path = Path(file_path)
    rows = read_csv_grid(file_path)
    wb = grid_to_workbook(rows)
    ws = wb.active
    header_row = detect_header_row(ws, ws.max_row, ws.max_column)
    desensitize_cols = resolve_desensitize_cols(ws, ws.max_column, header_row,
                                                enable_desensitize, desensitize_titles)
    data = process_worksheet_multi(ws, wb, "Sheet1", enable_sum, enable_internal_stats, collect_data,
                                   manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                                   enable_desensitize, desensitize_cols, enable_money_clean, header_row)
    output_path = file_path.parent / f"{file_path.stem}_已处理.xlsx"
    wb.save(output_path)
    return output_path, data

def process_old_xls(file_path, enable_sum, enable_internal_stats, collect_data,
                    manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                    enable_desensitize, desensitize_titles, enable_money_clean):
    file_path = Path(file_path)
    try:
        df = pd.read_excel(file_path, header=None, dtype=str, engine='xlrd')
    except Exception as e:
        raise Exception(f"无法读取旧版 .xls 文件（需要 xlrd 库，或文件已损坏）：{e}")
    df = df.astype(object).where(pd.notna(df), None)
    wb = grid_to_workbook(df.values.tolist())
    ws = wb.active
    header_row = detect_header_row(ws, ws.max_row, ws.max_column)
    desensitize_cols = resolve_desensitize_cols(ws, ws.max_column, header_row,
                                                enable_desensitize, desensitize_titles)
    data = process_worksheet_multi(ws, wb, "Sheet1", enable_sum, enable_internal_stats, collect_data,
                                   manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                                   enable_desensitize, desensitize_cols, enable_money_clean, header_row)
    output_path = file_path.parent / f"{file_path.stem}_已处理.xlsx"
    wb.save(output_path)
    return output_path, data

def process_xlsx(file_path, enable_sum, enable_internal_stats, collect_data,
                 manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                 enable_desensitize, desensitize_titles, enable_money_clean):
    file_path = Path(file_path)
    wb = load_workbook(file_path, data_only=True)
    all_data = []
    for sheet_name in list(wb.sheetnames):
        # V1.6.9：重复处理“已处理”文件时，不能把上次生成的统计表再当数据表统计
        if is_stat_sheet_name(sheet_name):
            log(f"跳过本工具生成的统计汇总表：{sheet_name}")
            continue
        ws = wb[sheet_name]
        try:
            header_row = detect_header_row(ws, ws.max_row, ws.max_column)
            desensitize_cols = []
            if enable_desensitize and desensitize_titles:
                for title in desensitize_titles:
                    col = find_column_by_title(ws, ws.max_column, header_row, [title])
                    if col is not None:
                        desensitize_cols.append(col)
                    else:
                        warn(f"工作表「{sheet_name}」未找到脱敏列标题 '{title}'，跳过该列")
            data = process_worksheet_multi(ws, wb, sheet_name, enable_sum, enable_internal_stats, collect_data,
                                           manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                                           enable_desensitize, desensitize_cols, enable_money_clean, header_row)
            all_data.extend(data)
        except Exception as e:
            # 单个工作表出错不再拖垮整批文件
            warn(f"工作表「{sheet_name}」处理失败：{e}（已跳过该工作表，其余工作表继续）")
    output_path = file_path.parent / f"{file_path.stem}_已处理.xlsx"
    wb.save(output_path)
    return output_path, all_data

def fix_files(file_paths, enable_sum, enable_internal_stats, enable_cross_summary,
              manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
              enable_desensitize, desensitize_titles, enable_money_clean,
              progress_callback, finish_callback, overrides=None):
    # 每轮处理都从干净的映射表开始，避免上一轮的映射混进来
    desensitize_mapper = DesensitizeMapper()
    diagnostics.clear()
    log("=" * 25 + " 新一轮处理开始 " + "=" * 25)
    log(f"开始处理，文件数量: {len(file_paths)}")
    log(f"选项状态: 求和={enable_sum}, 内部统计={enable_internal_stats}, 跨文件汇总={enable_cross_summary}")
    log(f"手动模式: {manual_mode}, 金额标题={user_amount_titles}, 对手方名称标题={user_counterparty_titles}, 对手方ID标题={user_counterparty_id_titles}")
    log(f"脱敏开关: {enable_desensitize}, 脱敏列标题={desensitize_titles}")
    log(f"金额清理开关: {enable_money_clean}")
    processed_paths = []
    failed = []
    all_collected = []
    summary_path = None
    mapping_path = None
    total = len(file_paths)

    try:
        for idx, file_path_str in enumerate(file_paths):
            file_path = Path(file_path_str)
            log(f"处理文件 {idx+1}/{total}: {file_path.name}")
            progress_callback(idx + 1, total, file_path.name)
            if overrides and _norm_key(file_path) in overrides:
                _ov = overrides[_norm_key(file_path)]
                _amt, _cp, _cid, _des = _ov
            else:
                _amt, _cp, _cid, _des = (user_amount_titles, user_counterparty_titles, user_counterparty_id_titles, desensitize_titles)
            _t_file = time.time()
            try:
                ext = file_path.suffix.lower()
                if ext == '.csv':
                    out_path, data = process_csv(file_path, enable_sum, enable_internal_stats, enable_cross_summary,
                                                 manual_mode, _amt, _cp, _cid,
                                                 enable_desensitize, _des, enable_money_clean)
                elif ext == '.xls':
                    out_path, data = process_old_xls(file_path, enable_sum, enable_internal_stats, enable_cross_summary,
                                                     manual_mode, _amt, _cp, _cid,
                                                     enable_desensitize, _des, enable_money_clean)
                elif ext == '.xlsx':
                    out_path, data = process_xlsx(file_path, enable_sum, enable_internal_stats, enable_cross_summary,
                                                  manual_mode, _amt, _cp, _cid,
                                                  enable_desensitize, _des, enable_money_clean)
                else:
                    raise Exception(f"不支持的文件类型：{ext}")
            except Exception as e:
                # V1.6.9：单个文件失败不再中断整批任务
                log_exc(f"文件处理失败：{file_path.name} -> {e}")
                failed.append(f"{file_path.name}：{e}")
                continue
            log(f"文件处理完成：{file_path.name}（用时 {time.time() - _t_file:.1f} 秒）")
            processed_paths.append(out_path)
            if enable_cross_summary:
                all_collected.extend(data)
                log(f"  收集到 {len(data)} 条对手方明细")

        if enable_desensitize and desensitize_mapper.orig_to_mask:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            mapping_path = Path(file_paths[0]).parent / f"脱敏映射_{timestamp}.json"
            try:
                desensitize_mapper.save_mapping(mapping_path)
                log(f"脱敏映射已保存至: {mapping_path}")
            except Exception as e:
                mapping_path = None
                warn(f"脱敏映射保存失败：{e}")

        log(f"汇总数据总量: {len(all_collected)} 条")
        if enable_cross_summary:
            if len(file_paths) <= 1:
                log("仅选择单个文件，跨文件汇总已启用但未满足条件（需要至少2个文件）")
            elif not all_collected:
                warn("跨文件汇总已启用，但没有收集到任何对手方明细，未生成汇总文件")
            else:
                name_stats, name_id_stats, totals = aggregate_counterparty(all_collected)
                first_dir = Path(file_paths[0]).parent
                summary_path = first_dir / "对手方汇总统计.xlsx"
                wb_summary = Workbook()
                ws_summary = wb_summary.active
                ws_summary.title = "对手方汇总"
                write_stat_block(ws_summary, 1, "总计", name_stats, name_id_stats,
                                 totals=totals, is_total_section=True)
                for col in ws_summary.columns:
                    max_len = 0
                    col_letter = get_column_letter(col[0].column)
                    for cell in col:
                        if cell.value:
                            max_len = max(max_len, len(str(cell.value)))
                    ws_summary.column_dimensions[col_letter].width = min(max_len+2, 25)
                try:
                    wb_summary.save(summary_path)
                    log(f"成功生成跨文件汇总表: {summary_path}")
                except Exception as e:
                    summary_path = None
                    warn(f"跨文件汇总表保存失败（是否被 Excel 打开占用？）：{e}")

        msg_parts = []
        if failed:
            msg_parts.append(f"处理完成，但有 {len(failed)} 个文件失败。")
        else:
            msg_parts.append("所有文件处理完成！")
        msg_parts.append(f"共处理 {len(processed_paths)} 个文件。")
        if enable_cross_summary:
            if summary_path:
                msg_parts.append(f"汇总统计表已生成：{summary_path}")
            elif len(file_paths) > 1:
                msg_parts.append("未生成跨文件汇总表（原因见下方诊断）。")
        if mapping_path:
            msg_parts.append(f"脱敏映射已保存：{mapping_path}")
        if enable_internal_stats:
            msg_parts.append("已在处理后的文档里按工作表生成“工作表名+对手方统计汇总”（无有效数据的工作表会跳过，原因见下）。")
        if diagnostics:
            msg_parts.append("")
            msg_parts.append("—— 本次诊断 ——")
            for d in diagnostics[:10]:
                msg_parts.append("· " + d)
            if len(diagnostics) > 10:
                msg_parts.append(f"· ……另有 {len(diagnostics) - 10} 条，请点“导出运行日志”查看")
        if failed:
            msg_parts.append("")
            msg_parts.append("—— 失败文件 ——")
            for f in failed[:8]:
                msg_parts.append("· " + f)
        finish_callback((not failed) or bool(processed_paths), "\n".join(msg_parts), processed_paths)
    except Exception as e:
        log_exc(f"处理出错：{e}")
        finish_callback(False, f"处理出错：{e}", None)

# ---------- 还原功能 ----------
def restore_file():
    file_path = filedialog.askopenfilename(
        title="请选择已脱敏的 Excel 或 CSV 文件",
        filetypes=[("Excel/CSV 文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")]
    )
    if not file_path:
        return
    mapping_path = filedialog.askopenfilename(
        title="请选择对应的脱敏映射文件（.json）",
        filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")]
    )
    if not mapping_path:
        return

    try:
        mapper = DesensitizeMapper()
        mapper.load_mapping(mapping_path)

        file_ext = Path(file_path).suffix.lower()
        if file_ext == '.csv':
            wb = grid_to_workbook(read_csv_grid(file_path))
            output_path = Path(file_path).parent / f"{Path(file_path).stem}_还原.xlsx"
        elif file_ext == '.xls':
            # V1.6.9：openpyxl 读不了旧版 .xls，先用 pandas+xlrd 转成新工作簿再还原
            df = pd.read_excel(file_path, header=None, dtype=str, engine='xlrd')
            df = df.astype(object).where(pd.notna(df), None)
            wb = grid_to_workbook(df.values.tolist())
            output_path = Path(file_path).parent / f"{Path(file_path).stem}_还原.xlsx"
        elif file_ext == '.xlsx':
            wb = load_workbook(file_path, data_only=True)
            output_path = Path(file_path).parent / f"{Path(file_path).stem}_还原.xlsx"
        else:
            messagebox.showerror("错误", "不支持的文件格式")
            return

        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None and isinstance(cell.value, str):
                        original = mapper.unmask(cell.value)
                        if original != cell.value:
                            cell.value = original
        wb.save(output_path)
        messagebox.showinfo("完成", f"还原完成！\n已保存至：{output_path}")
    except Exception as e:
        messagebox.showerror("错误", f"还原失败：{e}")

# ---------- GUI ----------
class App:
    def __init__(self, root):
        self.root = root
        self.root.title(f"银行流水梳理工具 {VERSION}")
        self.root.geometry("680x850")
        self.root.configure(bg="#f0f0f0")

        self.enable_money_clean = BooleanVar(value=True)
        self.enable_sum = BooleanVar(value=False)
        self.enable_internal_stats = BooleanVar(value=False)
        self.enable_cross_summary = BooleanVar(value=False)
        self.manual_mode = BooleanVar(value=False)
        self.enable_desensitize = BooleanVar(value=False)
        self.user_amount_titles = tk.StringVar()
        self.user_counterparty_titles = tk.StringVar()
        self.user_counterparty_id_titles = tk.StringVar()
        self.desensitize_titles = tk.StringVar()
        self.processing = False
        self.progress_win = None
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        info_frame = tk.Frame(root, bg="#f0f0f0")
        info_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        info_text = tk.Text(info_frame, height=8, width=65, wrap=tk.WORD, bg="#fafafa", font=("微软雅黑", 9), relief=tk.FLAT)
        info_text.insert(tk.END, "本软件不会去动原始表格的数据，而是生成一个处理后的表格。\n", "red_bold")
        info_text.insert(tk.END, "银行拉取的流水文件中，金额单元格常包含不可见格式（如货币符号、千分位、空格等），"
                                 "导致Excel公式无法正常求和。本软件可自动识别金额列，清理不可见字符，并转换为数值，"
                                 "确保统计准确。同时可生成对手方流水统计表（含流入/流出/合计）以及按对手方账户ID的细分统计。\n\n"
                                 "【信息脱敏】\n可对指定列进行确定性脱敏（数字→数字、中文→中文、英文→英文，长度不变），"
                                 "并保存映射文件，方便后续还原。\n\n"
                                 "【还原脱敏】\n点击下方按钮，分两步操作：先选择脱敏后的文件，再选择对应的映射文件，恢复原始数据。\n\n"
                                 "【导出运行日志】\n当遇到异常时，可点击“导出运行日志”按钮，将详细运行信息导出供分析。\n\n"
                                 "【处理完成提示】\n如果某个工作表没有生成统计汇总表（例如对手方名称列整列为空、"
                                 "金额列识别不到），完成提示框里会列出具体原因，可据此勾选“手动指定列”重跑。")
        info_text.insert(tk.END, "\n\n【自动运行日志】\n本软件一启动，就会在本软件同目录下写一个日志文件（文件名带版本号，如「银行流水梳理工具 V1.7.0 运行日志.txt」），只有一个，每次启动覆盖重写；上次若是闪退，会自动把上次的末尾留在新日志开头，证据不会丢。遇到卡死或闪退，点“日志存到桌面”，把桌面上那个 txt 发回即可。\n")
        info_text.insert(tk.END, "\n\n【源码备份】\n打包时源码和图标已经一并打进了 exe。点“导出源码备份”，桌面会多出一个文件夹，里面是当前版本的完整源码（.py）和图标（.ico）——换电脑、重装系统都不会再丢。万一界面打不开，还可以在命令行执行：程序.exe --export-source\n")
        info_text.tag_config("red_bold", foreground="red", font=("微软雅黑", 9, "bold"))
        info_text.config(state=tk.DISABLED)
        info_text.pack(fill=tk.BOTH, expand=True)

        options_frame = tk.Frame(root, bg="#f0f0f0")
        options_frame.pack(pady=5)

        row = 0
        cb_money_clean = tk.Checkbutton(options_frame, text="清理不可见字符，使金额列数据可被正常统计", variable=self.enable_money_clean, bg="#f0f0f0", font=("微软雅黑", 10))
        cb_money_clean.grid(row=row, column=0, sticky="w", padx=10, pady=2)
        row += 1

        cb_sum = tk.Checkbutton(options_frame, text="在每个区块末尾插入绝对值求和行（标黄）", variable=self.enable_sum, bg="#f0f0f0", font=("微软雅黑", 10))
        cb_sum.grid(row=row, column=0, sticky="w", padx=10, pady=2)
        row += 1

        cb_stats = tk.Checkbutton(options_frame, text="在生成表格内添加对手方流水统计表", variable=self.enable_internal_stats, bg="#f0f0f0", font=("微软雅黑", 10))
        cb_stats.grid(row=row, column=0, sticky="w", padx=10, pady=2)
        row += 1

        cb_cross = tk.Checkbutton(options_frame, text="生成对手方流水统计表汇总文件（仅在勾选并选择多个文件时可用）", variable=self.enable_cross_summary, bg="#f0f0f0", font=("微软雅黑", 10))
        cb_cross.grid(row=row, column=0, sticky="w", padx=10, pady=2)
        row += 1

        row += 1
        self.manual_frame = Frame(options_frame, bg="#f0f0f0")
        self.manual_frame.grid(row=row, column=0, sticky="w", padx=10, pady=10)
        manual_cb = Checkbutton(self.manual_frame, text="手动勾选（对手方流水自动识别失败时使用）",
                                variable=self.manual_mode, bg="#f0f0f0", font=("微软雅黑", 10),
                                command=self.toggle_manual_input)
        manual_cb.pack(anchor="w")

        self.manual_row = Frame(self.manual_frame, bg="#f0f0f0")
        self.manual_pick = Label(self.manual_row, text="", bg="#f0f0f0", font=("微软雅黑", 9),
                                 fg="#006400", wraplength=620, justify="left")
        self.manual_pick.pack(anchor="w")
        self.manual_row.pack_forget()

        row += 1
        self.desensitize_frame = Frame(options_frame, bg="#f0f0f0")
        self.desensitize_frame.grid(row=row, column=0, sticky="w", padx=10, pady=10)
        desensitize_cb = Checkbutton(self.desensitize_frame, text="启用信息脱敏", variable=self.enable_desensitize, bg="#f0f0f0", font=("微软雅黑", 10), command=self.toggle_desensitize_input)
        desensitize_cb.pack(anchor="w")

        self.desensitize_label = Label(self.desensitize_frame,
                                       text="勾选后：选完文件会先弹出「选择要脱敏的列」，从该文件的真实表头里勾选",
                                       bg="#f0f0f0", font=("微软雅黑", 9), fg="#444444",
                                       wraplength=620, justify="left")
        self.desensitize_pick = Label(self.desensitize_frame, text="", bg="#f0f0f0", font=("微软雅黑", 9),
                                      fg="#006400", wraplength=620, justify="left")
        self.desensitize_label.pack_forget()
        self.desensitize_pick.pack_forget()

        note_label = tk.Label(root, text="提示：手动勾选模式下，金额列 / 对手方名称列 / 对手方ID列都用「点按钮从表头里勾选」的方式指定；找不到时会跳过该区块并写日志。脱敏列同样从表头勾选。",
                              bg="#f0f0f0", font=("微软雅黑", 9), fg="blue", wraplength=600)
        note_label.pack(pady=5)

        btn_frame = tk.Frame(root, bg="#f0f0f0")
        btn_frame.pack(pady=10)
        self.btn = tk.Button(btn_frame, text="选择文件（可多选）", command=self.select_files, font=("微软雅黑", 10))
        self.btn.pack()

        restore_btn = Button(btn_frame, text="还原脱敏（分两步：先选文件，再选映射文件）", command=restore_file, font=("微软雅黑", 10), bg="#e0e0e0")
        restore_btn.pack(pady=5)

        restore_hint = tk.Label(btn_frame, text="还原操作：依次选择脱敏后的Excel/CSV文件，然后选择对应的脱敏映射文件（.json）",
                                bg="#f0f0f0", font=("微软雅黑", 8), fg="gray")
        restore_hint.pack(pady=2)

        log_btn_row = tk.Frame(btn_frame, bg="#f0f0f0")
        log_btn_row.pack(pady=5)
        send_log_btn = tk.Button(log_btn_row, text="日志存到桌面", command=save_log_to_desktop,
                                 font=("微软雅黑", 9), bg="#ffe2a8")
        send_log_btn.pack(side=tk.LEFT, padx=4)
        open_log_btn = tk.Button(log_btn_row, text="打开日志文件", command=open_log_folder,
                                 font=("微软雅黑", 9), bg="#e0e0e0")
        open_log_btn.pack(side=tk.LEFT, padx=4)
        export_btn = tk.Button(log_btn_row, text="导出运行日志", command=export_log,
                               font=("微软雅黑", 9), bg="#e0e0e0")
        export_btn.pack(side=tk.LEFT, padx=4)
        src_btn_row = tk.Frame(btn_frame, bg="#f0f0f0")
        src_btn_row.pack(pady=(0, 4))
        src_btn = tk.Button(src_btn_row, text="导出源码备份（源码+图标已打进 exe，不怕丢）",
                            command=export_source, font=("微软雅黑", 9), bg="#dce9ff")
        src_btn.pack()
        log_path_label = tk.Label(btn_frame, text=f"日志就在本软件旁边：{_log_file_path or '（未能创建日志文件）'}",
                                  bg="#f0f0f0", font=("微软雅黑", 8), fg="gray", wraplength=620,
                                  justify="left")
        log_path_label.pack(pady=2)

    def toggle_manual_input(self):
        if self.manual_mode.get():
            self.manual_row.pack(anchor="w", pady=(4, 0))
            self._refresh_manual_picks()
        else:
            self.manual_row.pack_forget()

    def _refresh_manual_picks(self):
        def fmt(var, label):
            v = (var.get() or "").strip()
            return "%s：%s" % (label, v) if v else "%s：未选" % label
        text = "  ｜  ".join([
            fmt(self.user_amount_titles, "金额"),
            fmt(self.user_counterparty_titles, "名称"),
            fmt(self.user_counterparty_id_titles, "ID"),
        ])
        done = all((var.get() or "").strip() for var in (
            self.user_amount_titles, self.user_counterparty_titles, self.user_counterparty_id_titles))
        try:
            self.manual_pick.config(text=text, fg="#006400" if done else "#b00020")
        except Exception:
            pass

    def toggle_desensitize_input(self):
        if self.enable_desensitize.get():
            self.desensitize_label.pack(anchor="w", pady=(5, 0))
            self._refresh_desensitize_pick()
            self.desensitize_pick.pack(anchor="w")
        else:
            self.desensitize_label.pack_forget()
            self.desensitize_pick.pack_forget()

    def _refresh_desensitize_pick(self):
        """把已选列显示在界面上；还没选时说明下一步会做什么。"""
        cur = [p.strip() for p in re.split(r'[，,、]+', self.desensitize_titles.get()) if p.strip()]
        try:
            if cur:
                self.desensitize_pick.config(
                    text="已选 %d 列：%s（想换可直接重新选文件）" % (len(cur), "、".join(cur)), fg="#006400")
            else:
                self.desensitize_pick.config(
                    text="尚未选择列 —— 接下来点「选择文件」时会先弹窗让你从表头里勾选", fg="#b00020")
        except Exception:
            pass

    def _collect_header_titles(self, file_path):
        """读取文件，返回所有非统计工作表的表头列名（去重、保序）。"""
        titles = []
        wb = None
        try:
            ext = str(file_path).lower()
            if ext.endswith('.csv'):
                wb = grid_to_workbook(read_csv_grid(file_path))
            else:
                wb = load_workbook(file_path, data_only=True)
        except Exception as exc:
            warn("读取表头失败：%s" % exc)
            return titles
        try:
            for ws in wb.worksheets:
                if is_stat_sheet_name(ws.title):
                    continue
                try:
                    hr = detect_header_row(ws, ws.max_row, ws.max_column)
                except Exception:
                    continue
                if not hr:
                    continue
                for c in range(1, ws.max_column + 1):
                    v = ws.cell(row=hr, column=c).value
                    if v is None:
                        continue
                    t = cell_text(v)
                    if t and t not in titles:
                        titles.append(t)
        finally:
            try:
                wb.close()
            except Exception:
                pass
        return titles
    def _split_var(self, var):
        return [p.strip() for p in re.split(r"[，,、]+", var.get()) if p.strip()]

    def _pick_columns(self, titles, need, cur, title_prefix=""):
        """统一勾选窗口：每个表头一行，列为 金额 / 名称 / ID / 脱敏（按 need 显示）。

        need: 可含 "amount" / "cp" / "id" / "des" 的子集；cur: 同键→已选列表。
        单击勾选；滚轮可滚动；确定返回 dict，取消返回 None。
        """
        if not titles:
            messagebox.showwarning("列不出表头", "这份文件里没识别到表头行，无法列出可选的列。")
            return None
        cols = []
        if "amount" in need:
            cols.append(("amount", "金额"))
        if "cp" in need:
            cols.append(("cp", "名称"))
        if "id" in need:
            cols.append(("id", "ID"))
        if "des" in need:
            cols.append(("des", "脱敏"))
        win = tk.Toplevel(self.root)
        win.title(title_prefix + "选择列（" + " / ".join(c[1] for c in cols) + "）")
        win.geometry("600x560")
        win.transient(self.root)
        tk.Label(win, text="每个表头一行；勾上它属于哪种列（可同时勾多种）：",
                 font=("微软雅黑", 10)).pack(anchor="w", padx=10, pady=(10, 0))
        tk.Label(win, text="金额=交易金额列，名称=对手方名称列，ID=对手方账户ID列，脱敏=该列做脱敏；取消=本次不处理。",
                 font=("微软雅黑", 9), fg="#666666", wraplength=560, justify="left").pack(anchor="w", padx=10, pady=(0, 6))
        holder = tk.Frame(win)
        holder.pack(fill="both", expand=True, padx=10)
        canvas = tk.Canvas(holder, borderwidth=0, highlightthickness=0, bg="white")
        sb = tk.Scrollbar(holder, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg="white")
        canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        tk.Label(inner, text="表头", bg="white", font=("微软雅黑", 9, "bold")).grid(row=0, column=0, sticky="w", padx=6, pady=2)
        for i, (key, label) in enumerate(cols, start=1):
            tk.Label(inner, text=label, bg="white", font=("微软雅黑", 9, "bold")).grid(row=0, column=i, padx=8)
        state = {}
        for r, t in enumerate(titles, start=1):
            tk.Label(inner, text=t, bg="white", font=("微软雅黑", 9), anchor="w", justify="left").grid(row=r, column=0, sticky="w", padx=6)
            per = {}
            for i, (key, label) in enumerate(cols, start=1):
                v = tk.BooleanVar(value=(t in (cur.get(key) or [])))
                tk.Checkbutton(inner, variable=v, bg="white", activebackground="white").grid(row=r, column=i, padx=8)
                per[key] = v
            state[t] = per
        result = {"ok": False}

        def _ok():
            result["ok"] = True
            win.destroy()

        btns = tk.Frame(win)
        btns.pack(pady=8)
        tk.Button(btns, text="确定", width=10, command=_ok).pack(side=tk.LEFT, padx=6)
        tk.Button(btns, text="取消", width=10, command=win.destroy).pack(side=tk.LEFT, padx=6)

        def _on_wheel(event, _c=canvas):
            try:
                _c.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except Exception:
                pass

        win.bind("<MouseWheel>", _on_wheel)
        win.grab_set()
        win.lift()
        win.focus_force()
        self.root.wait_window(win)
        if not result["ok"]:
            return None
        out = {}
        for key, _label in cols:
            out[key] = [t for t in titles if state[t][key].get()]
        for key in ("amount", "cp", "id", "des"):
            out.setdefault(key, [])
        return out


    def select_files(self):
        if self.processing:
            return
        file_paths = filedialog.askopenfilenames(
            title="选择要修复的文件（可多选）",
            filetypes=[("Excel/CSV 文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")]
        )
        if not file_paths:
            return
        overrides = {}
        need_manual = self.manual_mode.get()
        need_des = self.enable_desensitize.get()
        if need_manual or need_des:
            need = []
            if need_manual:
                need += ["amount", "cp", "id"]
            if need_des:
                need += ["des"]
            groups = {}
            for fp in file_paths:
                sig = tuple(self._collect_header_titles(fp))
                groups.setdefault(sig, []).append(fp)
            if len(groups) > 1:
                messagebox.showinfo(
                    "检测到 %d 种不同表头" % len(groups),
                    "选中的 %d 个文件里有 %d 种不同的表头，将逐个表头组分别设置列；\\n同一组文件只弹一次窗口。" % (
                        len(file_paths), len(groups)))
            first_group = True
            group_list = list(groups.items())
            for gi, (sig, files) in enumerate(group_list, start=1):
                rep = files[0]
                prefix = "第 %d/%d 组（%s）：" % (gi, len(group_list), Path(rep).name) if len(group_list) > 1 else ""
                try:
                    self.root.update()
                except Exception:
                    pass
                cur = {
                    "amount": self._split_var(self.user_amount_titles),
                    "cp": self._split_var(self.user_counterparty_titles),
                    "id": self._split_var(self.user_counterparty_id_titles),
                    "des": self._split_var(self.desensitize_titles),
                }
                picked = self._pick_columns(self._collect_header_titles(rep), need, cur, title_prefix=prefix)
                if picked is None:
                    messagebox.showinfo("已取消", "没有确认要用的列，本次不处理任何文件。")
                    return
                amt = picked["amount"] if need_manual else []
                cp = picked["cp"] if need_manual else []
                cid = picked["id"] if need_manual else []
                des = picked["des"] if need_des else []
                if need_des and not des:
                    if not messagebox.askyesno("一列都没选", "没有勾选任何脱敏列，继续处理等于「不做脱敏」。\\n确定要继续吗？"):
                        return
                for fp in files:
                    overrides[_norm_key(fp)] = (list(amt), list(cp), list(cid), list(des))
                if first_group:
                    self.user_amount_titles.set("，".join(amt))
                    self.user_counterparty_titles.set("，".join(cp))
                    self.user_counterparty_id_titles.set("，".join(cid))
                    self.desensitize_titles.set("，".join(des))
                    self._refresh_manual_picks()
                    self._refresh_desensitize_pick()
                    first_group = False
        self.start_processing(file_paths, overrides)


    def start_processing(self, file_paths, overrides=None):
        self.processing = True
        self.btn.config(state=tk.DISABLED)
        self.progress_win = tk.Toplevel(self.root)
        self.progress_win.title("处理中")
        self.progress_win.geometry("400x150")
        self.progress_win.transient(self.root)
        self.progress_win.grab_set()
        self.progress_win.protocol("WM_DELETE_WINDOW", lambda: None)

        self.progress_label = tk.Label(self.progress_win, text="正在处理文件，请稍候...\n", font=("微软雅黑", 10))
        self.progress_label.pack(pady=10)
        self.progress_bar = tk.Label(self.progress_win, text="", font=("微软雅黑", 9), fg="blue")
        self.progress_bar.pack()

        def parse_titles(txt):
            if not txt:
                return []
            parts = re.split(r'[，,、\s]+', txt)
            return [p.strip() for p in parts if p.strip()]

        user_amount = parse_titles(self.user_amount_titles.get())
        user_counterparty = parse_titles(self.user_counterparty_titles.get())
        user_counterparty_id = parse_titles(self.user_counterparty_id_titles.get())
        desensitize_titles = parse_titles(self.desensitize_titles.get()) if self.enable_desensitize.get() else []

        def update_progress(current, total, filename):
            # V1.6.9：tkinter 控件只能在主线程操作；工作线程里直接改控件
            # 在部分电脑上会偶发卡死/闪退，这里统一丢回主线程执行。
            self.root.after(0, lambda: self._update_progress(current, total, filename))

        def finish_callback(success, message, extra):
            self.root.after(0, lambda: self._finish_processing(success, message))

        thread = threading.Thread(target=fix_files, args=(
            file_paths, self.enable_sum.get(), self.enable_internal_stats.get(), self.enable_cross_summary.get(),
            self.manual_mode.get(), user_amount, user_counterparty, user_counterparty_id,
            self.enable_desensitize.get(), desensitize_titles,
            self.enable_money_clean.get(),
            update_progress, finish_callback, overrides
        ))
        thread.daemon = True
        thread.start()

    # ---------- 主线程里的界面更新 ----------
    def _update_progress(self, current, total, filename):
        try:
            if self.progress_win is not None and self.progress_win.winfo_exists():
                self.progress_label.config(text=f"正在处理: {filename}\n({current}/{total})")
                self.progress_bar.config(text=f"进度: {current}/{total}")
        except tk.TclError:
            pass

    def _finish_processing(self, success, message):
        global desensitize_mapper
        desensitize_mapper = DesensitizeMapper()
        try:
            if self.progress_win is not None and self.progress_win.winfo_exists():
                self.progress_win.grab_release()
                self.progress_win.destroy()
        except tk.TclError:
            pass
        self.progress_win = None
        self.btn.config(state=tk.NORMAL)
        self.processing = False
        if success:
            messagebox.showinfo("完成", message)
        else:
            messagebox.showerror("错误", message)

    def on_close(self):
        if self.processing:
            if not messagebox.askyesno("提示", "正在处理文件，现在退出可能留下不完整的输出文件。\n确定要退出吗？"):
                return
        logger.info("用户点击关闭窗口")
        self.root.destroy()

def _tk_heartbeat_loop(root):
    """每秒给看门狗打一次心跳；主线程卡住时心跳就断了。"""
    def _beat():
        try:
            _heartbeat()
            root.after(1000, _beat)
        except Exception:
            # 窗口已经销毁时 after 会失败，这里静默结束
            pass
    _beat()


def main():
    _install_tk_hook()
    logger.info("正在创建主窗口…")
    root = tk.Tk()
    app = App(root)          # noqa: F841  保持原写法
    _start_watchdog()
    _tk_heartbeat_loop(root)
    # 上次是闪退的话，窗口出来后提示一次（不挡住启动）
    root.after(400, lambda: maybe_notice_previous_crash(root))
    logger.info("主窗口已显示，进入事件循环")
    try:
        root.mainloop()
    except BaseException:
        logger.critical("事件循环异常退出", exc_info=True)
        raise
    else:
        logger.info("事件循环正常结束")
    finally:
        _watchdog_stop.set()
        _flush_all()

    logger.info(_EXIT_SENTINEL)
    logger.info("程序正常退出 %s", VERSION)
    _flush_all()


if __name__ == "__main__":
    # 命令行抢救源码：程序.exe --export-source [目标路径]
    # 不弹界面、直接写文件退出，界面起不来时也能把源码取回来
    try:
        if _cli_export_source():
            sys.exit(0)
    except Exception:
        logger.exception("命令行处理失败")
        sys.exit(2)

    try:
        main()
    except BaseException:
        try:
            logger.critical("程序异常退出", exc_info=True)
            _flush_all()
        except Exception:
            pass
        try:
            messagebox.showerror(
                "程序异常退出",
                "程序遇到未处理的错误，已经退出。\n\n"
                "日志文件：%s\n\n"
                "请把这个文件发给开发者。"
                % (_log_file_path or "（未能创建）"))
        except Exception:
            pass
        raise