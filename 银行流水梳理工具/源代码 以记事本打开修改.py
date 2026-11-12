# -*- coding: utf-8 -*-
r"""
银行流水梳理工具 V1.7.1

本版相对 V1.6.9 只做一件事：把"出问题时的证据"留下来，业务逻辑没有改动。

日志规则：
  * 日志就在软件（exe）同目录下，文件名带软件名和版本号，例如：
        银行流水梳理工具 V1.7.3 运行日志.txt
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
  软件(exe)所在目录\银行流水梳理工具 V1.7.3 运行日志.txt
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
import json
import logging
import faulthandler
import threading
import traceback

VERSION = "V2.0.7"
APP_NAME = "银行流水梳理工具"

# 卡死判定：主线程超过这么多秒没有心跳，就把所有线程堆栈写进日志
HANG_WARN_SECONDS = 90
# 正常退出时写的标记行；下次启动靠它判断上次是否正常结束
_EXIT_SENTINEL = "===== 程序正常退出 ====="
# V1.8.1：搬运历史时的分隔标记与"历史|"前缀剥离规则
_CARRY_MARK = "---- 上一次运行日志末尾结束，以下是本次运行 ----"
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


LOG_HISTORY_KEEP = 4   # 除当前日志外，再保留最近 4 次（合计 5 次）
_CARRY_PREFIX_RE = None   # 真正编译放在 re 导入之后（见 _init_log_regex）


def _init_log_regex():
    global _CARRY_PREFIX_RE
    try:
        _CARRY_PREFIX_RE = re.compile(r"^(?:\s*历史\|\s*)+")
    except Exception:
        _CARRY_PREFIX_RE = None


def _history_path(path, idx):
    base, ext = os.path.splitext(path)
    return "%s_历史%d%s" % (base, idx, ext)


def _rotate_logs(path):
    """把当前日志挪成_历史1，旧的依次后移，只保留最近 LOG_HISTORY_KEEP 份历史。

    V1.8.1：返回是否成功。以前整个函数包在 except: pass 里，
    另一个实例占着日志文件时 os.replace 会失败，但循环前半段
    （删历史4、历史3->4、历史2->3、历史1->2）已经执行完了，
    于是"历史1"被挪走却没人补上 —— 现象就是每多开一个实例就少一份历史。
    现在失败要回报，调用方改用自己的 PID 日志，绝不碰别人的历史。
    """
    try:
        last = _history_path(path, LOG_HISTORY_KEEP)
        if os.path.exists(last):
            os.remove(last)
        for i in range(LOG_HISTORY_KEEP - 1, 0, -1):
            src = _history_path(path, i)
            if os.path.exists(src):
                os.replace(src, _history_path(path, i + 1))
        if os.path.exists(path):
            os.replace(path, _history_path(path, 1))
        return True
    except Exception:
        return False


_SINGLETON = {"checked": False, "second": False, "fp": None, "mutex": None}


def _is_cmdline_only():
    """--export-source 是"界面起不来时救源码"的只读通道：不写业务文件，也不该抢日志。"""
    return len(sys.argv) > 1 and sys.argv[1] == "--export-source"


def _acquire_single_instance():
    """V1.8.2：全局单实例互斥。

    用 Windows 命名互斥量（Local\\ 前缀 = 当前登录会话），进程一退出内核自动释放，
    不会留下"死锁文件"让人误以为程序还在跑。
    返回 True = 本进程是唯一实例；False = 已经有一个在跑。
    拿不到句柄（非 Windows / 异常）一律返回 True：宁可多开，也不把用户挡在门外。
    """
    if _SINGLETON.get("mutex") is not None:
        return True
    try:
        import ctypes
        from ctypes import wintypes
        kernel32 = ctypes.windll.kernel32
        try:
            kernel32.SetLastError(0)
        except Exception:
            pass
        handle = kernel32.CreateMutexW(None, wintypes.BOOL(True), "Local\\BankFlowTool_single")
        if not handle:
            return True
        already = (kernel32.GetLastError() == 183)   # ERROR_ALREADY_EXISTS
        _SINGLETON["mutex"] = handle                  # 持有到进程结束，别关
        return not already
    except Exception:
        return True


def _activate_existing_window():
    """把已经在跑的那个窗口调到前台；找不到就算了，不影响退出。"""
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        found = []

        @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        def _enum(hwnd, _lparam):
            try:
                if not user32.IsWindowVisible(hwnd):
                    return True
                n = user32.GetWindowTextLengthW(hwnd)
                if n <= 0:
                    return True
                buf = ctypes.create_unicode_buffer(n + 1)
                user32.GetWindowTextW(hwnd, buf, n + 1)
                if APP_NAME in buf.value:
                    found.append(hwnd)
                    return False
            except Exception:
                pass
            return True

        user32.EnumWindows(_enum, 0)
        if found:
            h = found[0]
            if user32.IsIconic(h):
                user32.ShowWindow(h, 9)          # SW_RESTORE
            user32.SetForegroundWindow(h)
            return True
    except Exception:
        pass
    return False


def _notify_already_running():
    msg = (APP_NAME + " 已经在运行了。\n\n"
           "同一时间只允许开一个窗口：两个窗口同时处理会互相覆盖输出——\n"
           "  · xxx_已处理.xlsx 会被后来的那份覆盖；\n"
           "  · 两份进程还会抢同一份运行日志。\n\n"
           "请先关掉已经打开的那个窗口，再重新启动。")
    if os.environ.get("BFT_NO_MUTEX_NOTICE"):
        try:
            print(msg)
        except Exception:
            pass
        return
    try:
        import ctypes
        # MB_OK | MB_ICONINFORMATION | MB_SETFOREGROUND | MB_TOPMOST
        ctypes.windll.user32.MessageBoxW(
            None, msg, "%s %s" % (APP_NAME, VERSION), 0x40 | 0x10000 | 0x40000)
        return
    except Exception:
        pass
    try:
        print(msg)
    except Exception:
        pass


def _i_am_second_instance():
    """（V1.8.2 起这是兜底）万一互斥量失效、真跑出第二个实例，
    日志就改名带 PID，至少不覆盖别人的历史。正常路径走不到这里。

    用 TEMP 下一个固定名字的锁文件 + Windows 文件锁判断：
    第一个实例锁住它并持有到进程退出；后面的实例拿不到锁就知道自己是第二个。
    不禁止多开（有些用法就是要开两个窗口），只是各写各的日志。
    """
    if _SINGLETON["checked"]:
        return _SINGLETON["second"]
    _SINGLETON["checked"] = True
    try:
        import msvcrt
        lock_dir = os.environ.get("TEMP") or os.path.expanduser("~")
        lock_path = os.path.join(lock_dir, ".bankflow_tool.instance.lock")
        fh = open(lock_path, "a+")
        try:
            fh.seek(0)
            msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
        except Exception:
            _SINGLETON["second"] = True
            try:
                fh.close()
            except Exception:
                pass
            return True
        _SINGLETON["fp"] = fh      # 持有到进程结束，别关
    except Exception:
        pass
    return _SINGLETON["second"]


def _open_log_file():
    """挑目录并打开日志；返回 (路径, 文件对象, 旧日志内容, 上次是否正常退出)。"""
    global _prev_was_abnormal
    name = _log_file_name()
    for d in _candidate_log_dirs():
        if not d or not _try_make_dir(d):
            continue
        path = os.path.join(d, name)
        base, ext = os.path.splitext(name)
        if _is_cmdline_only():
            # 命令行导出：固定文件名、覆盖写、不轮转。
            # 以前它会走"第二实例"分支，凭空留下一个 _PIDxxxx.txt，让人以为是垃圾文件。
            path = os.path.join(d, "%s_命令行%s" % (base, ext))
            prev_tail, prev_clean, second = "", True, True
        elif _i_am_second_instance():
            # 兜底：真跑出第二个实例时，自己的 PID 日志，不轮转、不搬运、不动别人
            path = os.path.join(d, "%s_PID%d%s" % (base, os.getpid(), ext))
            prev_tail, prev_clean, second = "", True, True
        else:
            second = False
            prev_tail = _read_tail(path, CARRY_OVER_BYTES)
            prev_clean = True if not prev_tail.strip() else (_EXIT_SENTINEL in prev_tail)
        try:
            if not second:
                if not _rotate_logs(path):
                    # 日志被占着（多半是另一个实例）：退回 PID 日志，历史保持原样
                    path = os.path.join(d, "%s_PID%d%s" % (base, os.getpid(), ext))
                    prev_tail, prev_clean = "", True
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
            "请用工程目录里的「打包成EXE就双击我 V1.7.3.bat」重新打包，"
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

    # V1.8.2：单实例互斥，必须放在建日志之前 —— 否则第二个实例会去轮转别人的日志。
    # --export-source 是"界面起不来时救源码"的只读通道，不写任何业务文件，放行。
    _cmdline_only = _is_cmdline_only()
    if not _cmdline_only and not _acquire_single_instance():
        _activate_existing_window()
        _notify_already_running()
        try:
            _flush_all()
        except Exception:
            pass
        sys.exit(0)

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
    _init_log_regex()
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
            # V1.8.1：只搬"上一次自己写的那段"，并且剥掉已有的"历史|"前缀，
            # 否则每崩一次就在上一代的搬运结果上再套一层，变成"历史| 历史| 历史| ..."
            tail = prev_tail
            if _CARRY_MARK and _CARRY_MARK in tail:
                tail = tail.rsplit(_CARRY_MARK, 1)[1]
            lines = [ln for ln in tail.splitlines() if ln.strip()][-CARRY_OVER_LINES:]
            _strip = _CARRY_PREFIX_RE.sub if _CARRY_PREFIX_RE else (lambda pat, s: s)
            lines = [_strip('', ln) for ln in lines]
            if lines:
                logger.warning("---- 上一次运行日志的末尾（共 %d 行，原文照搬）----", len(lines))
                for ln in lines:
                    logger.warning("历史| %s", ln)
                logger.warning("%s", _CARRY_MARK)
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
    from tkinter import messagebox, filedialog, simpledialog, BooleanVar, Entry, Label, Frame, Checkbutton, Button
    from tkinter import ttk
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


# ---------- 脱敏映射管理（V1.8.0：可解密的格式保留加密） ----------
# 旧实现的三个致命缺陷：
#   1) mask 被 process_block 与 _mask_sheet_columns 各调一次 -> 同一格脱敏两代，映射表只记一代
#   2) "可解密"靠明文 JSON 旁路，表一丢就永久还原不了
#   3) 种子是 md5(f"{s}_{i}")，无密钥 -> 拿到 exe 就能枚举候选明文比对出来
# V1.8.0 做法：口令 -> PBKDF2 -> 格式保留加密(FPE)，密文同形同长，有口令才可逆。
# 不落明文映射表；口令丢了才真还原不了（这是"可解密"的代价，已在界面提示）。
import hmac
import secrets
import base64

_ALPHA = {
    "digit": list(string.digits),
    "lower": list(string.ascii_lowercase),
    "upper": list(string.ascii_uppercase),
    "cjk": [chr(i) for i in range(0x4E00, 0x9FFF) if chr(i).isalpha()],
}
_FPE_ROUNDS = 12          # 默认值
DESENS_ROUNDS = 12        # 实际使用（界面上可改，1~999）
_VAULT_MAGIC = b"BFTV2"
_PAD_CHAR = "\u9fa5"

DESENS_PASSWORD = ""
DESENS_BATCH = ""
DESENS_PER_FILE_SALT = False
DESENS_PUBLISH = False            # True = 脱敏发布模式（不可逆，外网随便改）
DESENS_RANDOMIZE = False          # True = 同一值每次不同密文（带 nonce 前缀）
DESENS_INCLUDE_HEADER = False     # True = 表头标题行也一起加密
DESENS_COLS_MAP = {}              # {sheet_index: set(列号)} 这一轮实际加密了哪些列
DESENS_ROUNDS_DEFAULT = 12


def set_desens_password(pwd):
    global DESENS_PASSWORD
    DESENS_PASSWORD = pwd or ""


def reset_masked_cols():
    global DESENS_COLS_MAP
    DESENS_COLS_MAP = {}


def record_masked_cols(wb, sheet_name, cols):
    """记下这个工作表里真正被加密的列号，还原时只解这些列。"""
    try:
        idx = wb.sheetnames.index(sheet_name)
    except Exception:
        return
    if not cols:
        return
    DESENS_COLS_MAP.setdefault(idx, set()).update(cols)


def set_desens_include_header(flag):
    """True = 表头那一行也加密；False = 只加密数据行（默认）。"""
    global DESENS_INCLUDE_HEADER
    DESENS_INCLUDE_HEADER = bool(flag)


def set_desens_rounds(n):
    """加密轮数：默认 12，用户爱填几轮填几轮（1~999）。越高越慢、越难反推。"""
    global DESENS_ROUNDS
    try:
        n = int(float(str(n).strip()))
    except Exception:
        n = DESENS_ROUNDS_DEFAULT
    DESENS_ROUNDS = max(1, min(999, n))


def set_desens_publish(flag):
    """True = 脱敏发布：单向假名化，不可还原、不签名、随便改。

    False = 加密保管：可逆、要口令、改过会报警。
    """
    global DESENS_PUBLISH
    DESENS_PUBLISH = bool(flag)


def set_desens_randomize(flag):
    """True = 随机化加密（同值不同密文，抗频率分析）；False = 确定性（同值同密文，可聚类）。"""
    global DESENS_RANDOMIZE
    DESENS_RANDOMIZE = bool(flag)


def set_desens_per_file_salt(flag):
    """True = 一份文件一把盐（文件之间对不上）；False = 同一次处理共用一把盐。"""
    global DESENS_PER_FILE_SALT
    DESENS_PER_FILE_SALT = bool(flag)


def set_desens_batch(batch_id):
    """同一次处理的一批文件共用一个批次号，写进文件属性，替代原来的金库清单。"""
    global DESENS_BATCH
    DESENS_BATCH = batch_id or ""


def _b64(s):
    """提示里可能有 ';' '=' 之类，塞进文件属性前先编码。"""
    try:
        return base64.urlsafe_b64encode((s or "").encode("utf-8")).decode("ascii").rstrip("=")
    except Exception:
        return ""


def _unb64(s):
    try:
        return base64.urlsafe_b64decode((s or "").encode("ascii") + b"=" * (-len(s or "") % 4)).decode("utf-8")
    except Exception:
        return ""


def _char_class(ch):
    if "0" <= ch <= "9":
        return "digit"
    if "a" <= ch <= "z":
        return "lower"
    if "A" <= ch <= "Z":
        return "upper"
    if "\u4e00" <= ch <= "\u9fff" and ch.isalpha():
        return "cjk"
    return None


def _hkdf(prk, info, length=32):
    out, t, i = b"", b"", 1
    while len(out) < length:
        t = hmac.new(prk, t + info + bytes([i]), hashlib.sha256).digest()
        out += t
        i += 1
    return out[:length]


def _prf_mod(k, tweak, rnd, n):
    return int.from_bytes(
        hmac.new(k, tweak + b"|K|" + rnd.to_bytes(4, "big"), hashlib.sha256).digest(), "big") % n


def _prf_bit(k, tweak, rnd, lo, n):
    msg = (tweak + b"|B|" + str(rnd).encode() + b"|" + hex(n).encode() + b"|" + hex(lo).encode())
    return hmac.new(k, msg, hashlib.sha256).digest()[0] & 1


def _fpe_round(k, tweak, i, x, n):
    """一轮 swap-or-not：对合函数，加解密共用。"""
    ki = _prf_mod(k, tweak, i, n)
    y = (ki - x) % n
    lo = x if x < y else y
    return y if _prf_bit(k, tweak, i, lo, n) else x


def _fpe_forward(k, tweak, x, n):
    for i in range(DESENS_ROUNDS):
        x = _fpe_round(k, tweak, i, x, n)
    return x


def _fpe_backward(k, tweak, x, n):
    for i in reversed(range(DESENS_ROUNDS)):
        x = _fpe_round(k, tweak, i, x, n)
    return x


def _keystream(k, nonce, length):
    out, ctr = b"", 0
    while len(out) < length:
        out += hmac.new(k, b"|S|" + nonce + ctr.to_bytes(8, "big"), hashlib.sha256).digest()
        ctr += 1
    return out[:length]


def _table_signature(mapper, wb):
    """对整张工作簿的内容做 HMAC，用来发现"事后被改过"。

    用无序多重集：排序、筛选、拖行都不会误报；改一个值、增删一行一定对不上。
    """
    try:
        vals = []
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is None:
                        continue
                    vals.append(str(cell.value))
        vals.sort()
        msg = "\x1f".join(vals).encode("utf-8")
        return hmac.new(mapper._k_mac, msg, hashlib.sha256).hexdigest()
    except Exception:
        return ""


def _read_sig(wb):
    try:
        d = wb.properties.description or ""
    except Exception:
        return ""
    m = re.search(r"BFTSIG=([0-9a-f]{64})", d)
    return m.group(1) if m else ""


def _scrub_metadata(wb, masked_gen=0, keyfp="", salt="", hint="", src="", batch="", mode="", rounds=0, sig=""):
    """清掉作者/修改人/时间/custom.xml，只留一个不可逆的脱敏标记。"""
    try:
        p = wb.properties
        p.creator = ""
        p.lastModifiedBy = ""
        p.lastPrinted = None
        p.created = datetime.datetime(2000, 1, 1)
        p.modified = datetime.datetime(2000, 1, 1)
        p.title = None
        p.subject = None
        p.description = None
        p.company = None
        p.manager = None
        p.category = None
        if masked_gen:
            # V1.8.3：salt 必须一起写进文件。密钥 = PBKDF2(口令, salt)，
            # salt 不落盘的话"只凭口令还原"就是假的（口令对也解成乱码）。
            # V1.8.4：HINT 也写进文件 —— 换一台电脑 / 发给别人时，
            # 对方看不到本机的提示列表，只能从文件里读。
            _cols_str = ""
            try:
                if DESENS_COLS_MAP:
                    _parts = []
                    for _i in sorted(DESENS_COLS_MAP):
                        _cs = ",".join(str(c) for c in sorted(DESENS_COLS_MAP[_i]))
                        if _cs:
                            _parts.append("%d:%s" % (_i, _cs))
                    _cols_str = ";".join(_parts)
                    if len(_cols_str) > 160:      # 属性长度有限，超了就退化为"全解"
                        _cols_str = ""
            except Exception:
                _cols_str = ""
            _mode = mode or ("publish" if DESENS_PUBLISH else ("rnd" if DESENS_RANDOMIZE else "det"))
            _pub = (_mode == "publish")
            p.keywords = ("BFT_MASK_GEN=%d;SALT=%s;KEYFP=%s;HINT=%s;SRC=%s;BATCH=%s;"
                          "MODE=%s;ROUNDS=%d;HDR=%d;COLS=%s;ALGO=%s") % (
                masked_gen, salt or "-", keyfp or "-",
                _b64(hint)[:80] or "-", _b64(src)[:80] or "-", batch or "-",
                _mode,
                0 if _pub else (rounds or DESENS_ROUNDS),
                1 if DESENS_INCLUDE_HEADER else 0,
                _cols_str or "-",
                "alias-v1" if _pub else "fpe-v1")
            # 完整性印记单独放 description：keywords 已经很长，再塞 64 位会超字段长度
            p.description = ("BFTSIG=%s" % sig) if sig else None
        else:
            p.keywords = None
    except Exception as e:
        log(f"元数据清理失败（不影响数据）：{e}")


def _read_hint(wb):
    """从文件属性里取回密码提示（base64）。"""
    try:
        kw = wb.properties.keywords or ""
    except Exception:
        return ""
    m = re.search(r"HINT=([A-Za-z0-9_\-]+)", kw)
    if not m or m.group(1) == "-":
        return ""
    return _unb64(m.group(1))


def _read_cols_map(wb):
    """返回 {sheet_index: [列号]}；老文件没有这项时返回 {}。"""
    try:
        kw = wb.properties.keywords or ""
        m = re.search(r"COLS=([0-9:;,\-]+)", kw)
        if not m or m.group(1) == "-":
            return {}
        out = {}
        for part in m.group(1).split(";"):
            if ":" not in part:
                continue
            i, cs = part.split(":", 1)
            out[int(i)] = [int(x) for x in cs.split(",") if x.strip().isdigit()]
        return out
    except Exception:
        return {}


def _read_hdr_masked(wb):
    try:
        kw = wb.properties.keywords or ""
        m = re.search(r"HDR=(\d)", kw)
        return bool(int(m.group(1))) if m else False
    except Exception:
        return False


def _read_rounds(wb):
    try:
        kw = wb.properties.keywords or ""
        m = re.search(r"ROUNDS=(\d+)", kw)
        return max(1, min(999, int(m.group(1)))) if m else DESENS_ROUNDS_DEFAULT
    except Exception:
        return DESENS_ROUNDS_DEFAULT


def _read_mode(wb):
    """返回 'det'（确定性）或 'rnd'（随机化）。"""
    try:
        kw = wb.properties.keywords or ""
    except Exception:
        return "det"
    m = re.search(r"MODE=([a-z]+)", kw)
    return (m.group(1) if m else "det")


def _read_mask_marker(wb):
    """返回 (代数, keyfp, salt_hex)；没标记返回 (0, '', '')。"""
    try:
        kw = wb.properties.keywords or ""
    except Exception:
        return 0, "", ""
    m = re.search(r"BFT_MASK_GEN=(\d+)", kw)
    if not m:
        return 0, "", ""
    g = int(m.group(1))
    k = re.search(r"KEYFP=([0-9a-f]+)", kw)
    s = re.search(r"SALT=([0-9a-fA-F]+)", kw)
    return g, (k.group(1) if k else ""), (s.group(1) if s else "")


class DesensitizeMapper:
    """V1.8.0 内核。接口与旧版兼容：mask / unmask / save_mapping / load_mapping。"""

    VERSION = "fpe-v1"

    def __init__(self, password=None, salt=None):
        self.salt = salt or secrets.token_bytes(16)
        self.orig_to_mask = {}
        self.mask_to_orig = {}
        self.legacy_to_orig = {}
        self.manifest = {}
        self.randomized = False
        self.publish = False
        self.password = None
        self.keyfp = ""
        self._k_fpe = self._k_mac = self._k_wrap = b""
        if password:
            self.set_password(password)

    # --- 密钥
    def set_password(self, password):
        self.password = password
        master = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), self.salt, 200000, 32)
        self._k_fpe = _hkdf(master, b"fpe")
        self._k_mac = _hkdf(master, b"mac")
        self._k_wrap = _hkdf(master, b"wrap")
        self.keyfp = hmac.new(self._k_fpe, b"keyfp", hashlib.sha256).hexdigest()[:8]

    @property
    def ready(self):
        return bool(self._k_fpe)

    def _need_key(self):
        if not self._k_fpe:
            raise RuntimeError("未设置脱敏口令")

    # --- 单段
    def _seg(self, text, cls, tweak, forward=True):
        alpha = _ALPHA[cls]
        radix = len(alpha)
        idx = {c: i for i, c in enumerate(alpha)}
        x = 0
        for ch in text:
            x = x * radix + idx[ch]
        n = radix ** len(text)
        y = _fpe_forward(self._k_fpe, tweak, x, n) if forward else _fpe_backward(self._k_fpe, tweak, x, n)
        out = []
        for _ in range(len(text)):
            y, r = divmod(y, radix)
            out.append(alpha[r])
        return "".join(reversed(out))

    def _walk(self, s, tweak, forward):
        tb = ("v1|" + self.VERSION + "|" + (tweak or "")).encode("utf-8")
        out, buf, cur = [], [], None
        for ch in s:
            cls = _char_class(ch)
            if cls is None:
                if buf:
                    out.append(self._seg("".join(buf), cur, tb, forward))
                    buf, cur = [], None
                out.append(ch)
            else:
                if cur is not None and cls != cur:
                    out.append(self._seg("".join(buf), cur, tb, forward))
                    buf = []
                cur = cls
                buf.append(ch)
        if buf:
            out.append(self._seg("".join(buf), cur, tb, forward))
        return "".join(out)

    # --- 对外
    def _alias(self, s):
        """单向假名：同值同结果，但没有密钥算不回去。

        发布模式专用。前缀按值的字符类型给，方便肉眼分辨是哪类数据。
        """
        key = self._k_mac or self.salt
        h = hmac.new(key, s.encode("utf-8"), hashlib.sha256).digest()
        code = base64.b32encode(h).decode("ascii").rstrip("=")[:6]
        c0 = s[:1]
        if "\u4e00" <= c0 <= "\u9fff":
            pre = "名"
        elif c0.isdigit():
            pre = "号"
        elif c0.isalpha():
            pre = "户"
        else:
            pre = "值"
        return "%s-%s" % (pre, code)

    def mask(self, original, tweak=""):
        if self.publish:
            if original is None:
                return None
            s = str(original)
            res = self._alias(s)
            self.orig_to_mask.setdefault(s, res)
            self.mask_to_orig.setdefault(res, s)
            return res
        self._need_key()
        if original is None:
            return None
        s = str(original)
        if DESENS_RANDOMIZE:
            # 随机化：每格一个 2 字随机 nonce 拼在最前面，作为这一格的 tweak 种子。
            # 同值不同格 -> nonce 不同 -> 密文不同；解密时把 nonce 剥出来重建 tweak。
            _alpha = _ALPHA["cjk"]
            nonce = _alpha[secrets.randbelow(len(_alpha))] + _alpha[secrets.randbelow(len(_alpha))]
            res = nonce + self._walk(s, tweak + "|rnd|" + nonce, True)
        else:
            res = self._walk(s, tweak, True)
        self.orig_to_mask.setdefault(s, res)
        self.mask_to_orig.setdefault(res, s)
        return res

    def unmask(self, masked, tweak=""):
        if masked is None:
            return None
        if getattr(self, "publish", False):
            return masked          # 发布模式不可逆
        s = str(masked)
        if s in self.legacy_to_orig:
            return self.legacy_to_orig[s]
        if not self._k_fpe:
            return s
        if getattr(self, "randomized", False) and len(s) >= 2:
            _c0, _c1 = s[0], s[1]
            if _char_class(_c0) == "cjk" and _char_class(_c1) == "cjk":
                nonce, body = _c0 + _c1, s[2:]
                if body:
                    return self._walk(body, tweak + "|rnd|" + nonce, False)
        return self._walk(s, tweak, False)

    # --- 金库：映射表落盘即密文
    def vault_bytes(self, manifest=None):
        self._need_key()
        _payload = {
            "version": self.VERSION,
            "keyfp": self.keyfp,
            "rounds": DESENS_ROUNDS,
            "orig_to_mask": self.orig_to_mask,
        }
        if manifest:
            _payload["manifest"] = manifest
        payload = json.dumps(_payload, ensure_ascii=False).encode("utf-8")
        nonce = secrets.token_bytes(16)
        ks = _keystream(self._k_wrap, nonce, len(payload))
        ct = bytes(a ^ b for a, b in zip(payload, ks))
        tag = hmac.new(self._k_mac, self.salt + nonce + ct, hashlib.sha256).digest()
        return _VAULT_MAGIC + self.salt + nonce + tag + ct

    def save_mapping(self, file_path, manifest=None):
        with open(file_path, "wb") as f:
            f.write(self.vault_bytes(manifest))

    def load_mapping(self, file_path, password=None):
        """兼容三样：.vault（口令解）、旧 .json（明文，只做只读兜底）。"""
        path = str(file_path)
        raw = open(path, "rb").read()
        if raw.startswith(_VAULT_MAGIC):
            salt = raw[len(_VAULT_MAGIC):len(_VAULT_MAGIC) + 16]
            off = len(_VAULT_MAGIC) + 16
            nonce, tag, ct = raw[off:off + 16], raw[off + 16:off + 48], raw[off + 48:]
            pwd = password or self.password
            if not pwd:
                raise RuntimeError("金库需要口令")
            master = hashlib.pbkdf2_hmac("sha256", pwd.encode("utf-8"), salt, 200000, 32)
            k_mac = _hkdf(master, b"mac")
            k_wrap = _hkdf(master, b"wrap")
            if not hmac.compare_digest(tag, hmac.new(k_mac, salt + nonce + ct, hashlib.sha256).digest()):
                raise RuntimeError("口令错误（金库校验失败）")
            ks = _keystream(k_wrap, nonce, len(ct))
            data = json.loads(bytes(a ^ b for a, b in zip(ct, ks)).decode("utf-8"))
            self.salt = salt
            self.set_password(pwd)
            self.orig_to_mask = data.get("orig_to_mask", {})
            self.mask_to_orig = {v: k for k, v in self.orig_to_mask.items()}
            self.manifest = data.get("manifest", {})
            return True
        # 旧明文 JSON：只当"只读反查表"用，不参与新脱敏
        try:
            data = json.loads(raw.decode("utf-8-sig"))
            self.legacy_to_orig = data.get("mask_to_orig", {})
            self.mask_to_orig.update(self.legacy_to_orig)
            return True
        except Exception as e:
            raise RuntimeError("映射文件读不出来：%s" % e)


desensitize_mapper = DesensitizeMapper()

# ---------- 辅助函数 ----------
_FULLWIDTH_TABLE = str.maketrans('０１２３４５６７８９．，－＋％', '0123456789.,-+%')
_CURRENCY_RE = re.compile(r'[¥￥$€£]')
_AMOUNT_UNIT_RE = re.compile(r'(元整|元|圆|人民币|RMB|CNY|USD)$', re.IGNORECASE)

_LAST_DIR_FILE = os.path.join(os.path.expanduser("~"), ".bankflow_lastdir.txt")
_last_dir_cache = {"dir": ""}


def remember_dir(path):
    """记住最近的输出/映射目录，供还原时自动定位。"""
    try:
        d = os.path.dirname(os.path.abspath(str(path)))
        _last_dir_cache["dir"] = d
        with open(_LAST_DIR_FILE, "w", encoding="utf-8") as fh:
            fh.write(d)
    except Exception:
        pass


def last_dir(default=""):
    """取上次记住的目录；没有就返回 default（或当前目录）。"""
    try:
        if _last_dir_cache["dir"] and os.path.isdir(_last_dir_cache["dir"]):
            return _last_dir_cache["dir"]
        if os.path.exists(_LAST_DIR_FILE):
            with open(_LAST_DIR_FILE, encoding="utf-8") as fh:
                d = fh.read().strip()
            if d and os.path.isdir(d):
                _last_dir_cache["dir"] = d
                return d
    except Exception:
        pass
    return default or os.getcwd()

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

def _generic_header_score(ws, row, scan_cols):
    """通用（非银行流水）表头打分：短文本、非数字、非日期、非空。"""
    score = 0
    for col in range(1, scan_cols + 1):
        v = ws.cell(row, col).value
        if v is None:
            continue
        if isinstance(v, (int, float, bool, datetime.date, datetime.datetime)):
            continue
        s = str(v).strip()
        if not s or len(s) > 20:
            continue
        if parse_amount(s) is not None:
            continue
        score += 1
    return score


def _detect_generic_header_row(ws, max_row, scan_cols, max_scan=20):
    """关键词全不命中时，找最像表头的一行。返回 (row, score)。"""
    best_row, best = 1, 0
    for row in range(1, min(max_row, max_scan) + 1):
        s = _generic_header_score(ws, row, scan_cols)
        if s > best:
            best_row, best = row, s
    return best_row, best


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
        # 银行关键词一个都没命中：多半是普通表格（单独用脱敏时很常见）。
        # 退到通用规则找"最像表头的一行"，并明确告诉用户可以改。
        gen_row, gen_score = _detect_generic_header_row(ws, max_row, scan_cols)
        _sheet_ctx["header_score"] = best_score
        if gen_score >= 2:
            log(f"未命中银行表头关键词，按通用规则推测表头为第 {gen_row} 行"
                f"（{gen_score} 个文本列）；不对的话在选列窗口里改表头行")
            return gen_row
        log(f"未识别到明显的表头行（关键词命中数 {best_score}，文本列也少于 2 个），已按第 1 行处理")
        return 1
    if best_score < 4:
        # 关键词证据不足（2~3 个）时不能直接采信：数据行里出现"备注/账号名称"这类词很常见，
        # 会把它误判成表头（真实案例：表头在第 1 行，却被判成第 28 行）。
        # 这时跟"通用规则认定的那一行"比一比：谁更像表头（短文本列更多）用谁，同分取靠上的。
        gen_row, gen_score = _detect_generic_header_row(ws, max_row, scan_cols)
        kw_gen = _generic_header_score(ws, best_row, scan_cols)
        if gen_score >= 2 and (gen_score > kw_gen or (gen_score == kw_gen and gen_row < best_row)):
            _sheet_ctx["header_score"] = best_score
            log(f"关键词命中仅 {best_score}（第 {best_row} 行），改用更像表头的第 {gen_row} 行"
                f"（{gen_score} 个文本列）；不对的话在选列窗口里改表头行")
            return gen_row
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

    # 脱敏：V1.8.0 起只由 process_worksheet_multi 的 _mask_sheet_columns 统一执行一次。
    # 原来这里还有一段按列索引的 mask 循环，没有 _mask_already 守卫，
    # 导致同一格被连续脱敏两代（文件里是 gen2/gen3，映射表只到 gen1 -> 还原不出来）。
    _ = (enable_desensitize, desensitize_cols)

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



def _sheet_segments(ws, max_col):
    """按整行空白把工作表切成若干段，返回 [(start_row, end_row), ...]。"""
    segs, cur = [], None
    for r in range(1, ws.max_row + 1):
        empty = True
        for c in range(1, max_col + 1):
            v = ws.cell(r, c).value
            if v is not None and not (isinstance(v, str) and not v.strip()):
                empty = False
                break
        if empty:
            if cur:
                segs.append(cur)
                cur = None
        else:
            cur = (cur[0], r) if cur else (r, r)
    if cur:
        segs.append(cur)
    return segs


def _segment_header_row(ws, start, end, max_col):
    """在一段内找表头行：先银行关键词，再通用短文本规则；取前 3 行里最像的。"""
    best_row, best_score = start, -1
    for row in range(start, min(end, start + 3) + 1):
        kw = _header_score(ws, row, min(max_col, 120))
        gen = _generic_header_score(ws, row, min(max_col, 120))
        score = kw * 10 + gen
        if score > best_score:
            best_row, best_score = row, score
    return best_row



def _pick_header_by_titles(ws, start, end, max_col, wanted):
    """在某一段里挑"能匹配到最多目标列"的那一行当表头；一个都匹配不到返回 None。"""
    best_row, best_hit = None, 0
    last = min(end, start + 80)
    for row in range(start, last + 1):
        hit = 0
        for col in range(1, max_col + 1):
            title = cell_text(ws.cell(row=row, column=col).value).strip()
            if title and title in wanted:
                hit += 1
        if hit > best_hit:
            best_row, best_hit = row, hit
    return best_row


def _mask_sheet_columns(ws, wanted_titles, fallback_header=None):
    """按列标题在整张表逐段脱敏：每段用自己那一行表头，互不影响。

    这样一张表里堆了多张不同表头的表也能各自命中。
    """
    total = 0
    _hit_cols = set()
    wanted = set(t.strip() for t in (wanted_titles or []) if t and t.strip())
    if not wanted:
        return total, _hit_cols
    segs = _sheet_segments(ws, ws.max_column)
    if not segs:
        segs = [(fallback_header or 1, ws.max_row)]
    for (s, e) in segs:
        hdr = _pick_header_by_titles(ws, s, e, ws.max_column, wanted) or _segment_header_row(ws, s, e, ws.max_column)
        if not (s <= hdr <= e):
            hdr = s
        for col in range(1, ws.max_column + 1):
            title = cell_text(ws.cell(row=hdr, column=col).value).strip()
            if not title or title not in wanted:
                continue
            _start = hdr if DESENS_INCLUDE_HEADER else hdr + 1
            for r in range(_start, e + 1):
                cell = ws.cell(row=r, column=col)
                v = cell.value
                if v is None:
                    continue
                if isinstance(v, str) and not v.strip():
                    continue
                if _mask_already(v):
                    continue
                cell.value = desensitize_mapper.mask(v)
                total += 1
                _hit_cols.add(col)
    return total, _hit_cols


def _parse_row_input(text, fallback=1):
    """把"表头在第几行"输入框的文本解析成行号；解析不了就用 fallback。"""
    try:
        v = int(float(str(text).strip()))
    except Exception:
        return fallback
    return v if v >= 1 else fallback


def _mask_already(value):
    """单元格是不是已经被本轮的脱敏表替换过（防止二次脱敏）。"""
    try:
        return str(value) in desensitize_mapper.mask_to_orig
    except Exception:
        return False


def _mask_columns_whole(ws, header_row, cols):
    """把指定列在表头之后的所有非空单元格脱敏，不依赖流水区块。

    通用表格（没有金额列、切不出区块）单独用脱敏时走这条路。
    """
    n = 0
    _hit = set()
    if not cols:
        return n, _hit
    _hr = int(header_row or 1)
    start = _hr if DESENS_INCLUDE_HEADER else _hr + 1
    max_row = ws.max_row
    for col in cols:
        if col is None:
            continue
        for row in range(start, max_row + 1):
            cell = ws.cell(row=row, column=col)
            v = cell.value
            if v is None:
                continue
            if isinstance(v, str) and not v.strip():
                continue
            if _mask_already(v):
                continue
            cell.value = desensitize_mapper.mask(v)
            n += 1
            _hit.add(col)
    return n, _hit


def process_worksheet_multi(ws, wb, sheet_name, enable_sum, enable_internal_stats, collect_data,
                           manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                           enable_desensitize, desensitize_cols, enable_money_clean, header_row=None,
                           desensitize_titles=None):
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

    # 脱敏独立于流水区块：整列脱敏先做，通用表格也能用。
    # V1.8.7：这是【唯一】脱敏入口。金额列不再特殊对待 ——
    # 脱敏和金额梳理已经互斥（fix_files 会把求和/统计/汇总/金额清理全关掉），
    # 所以金额列加密不会污染任何统计，勾了就该加密。
    if enable_desensitize and (desensitize_titles or desensitize_cols):
        if not desensitize_mapper.ready and not getattr(desensitize_mapper, "publish", False):
            warn("已勾选脱敏但没有口令，本工作表未脱敏")
        else:
            if desensitize_titles:
                _n_masked = _mask_sheet_columns(ws, desensitize_titles, header_row)
            else:
                _n_masked = _mask_columns_whole(ws, header_row, desensitize_cols)
            if _n_masked:
                log(f"脱敏：工作表「{sheet_name}」共处理 {_n_masked[0]} 个单元格（单次）")
                record_masked_cols(wb, sheet_name, _n_masked[1])
    # 表头都没识别出来，说明这张表多半不是流水表：细节只写日志，不刷屏
    _sheet_ctx["sheet"] = sheet_name
    _sheet_ctx["quiet"] = _sheet_ctx["header_score"] < 2
    if enable_desensitize and not (enable_sum or enable_internal_stats or collect_data or enable_money_clean):
        # 纯脱敏模式：不做区块切分、不找金额列、不出统计，
        # 也就不会有"未找到金额列，已跳过该区块"这种跟脱敏无关的诊断
        log(f"工作表「{sheet_name}」脱敏完成，跳过区块识别（脱敏模式不做统计）")
        return []

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
                enable_desensitize, desensitize_titles, enable_money_clean, header_rows=None,
                desensitize_by_sheet=None):
    file_path = Path(file_path)
    rows = read_csv_grid(file_path)
    wb = grid_to_workbook(rows)
    ws = wb.active
    header_row = (header_rows or {}).get('Sheet1') or detect_header_row(ws, ws.max_row, ws.max_column)
    _titles = (desensitize_by_sheet or {}).get("Sheet1", desensitize_titles)
    desensitize_cols = resolve_desensitize_cols(ws, ws.max_column, header_row,
                                                enable_desensitize, _titles)
    data = process_worksheet_multi(ws, wb, "Sheet1", enable_sum, enable_internal_stats, collect_data,
                                   manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                                   enable_desensitize, desensitize_cols, enable_money_clean, header_row,
                                   desensitize_titles=_titles)
    output_path = file_path.parent / f"{file_path.stem}_已处理.xlsx"
    if enable_desensitize and DESENS_PUBLISH:
        # 发布模式：外网要改数据、复制副本，签名只会添乱 —— 不生成
        _scrub_metadata(wb, 1, "-", desensitize_mapper.salt.hex(),
                        DESENS_HINT, file_path.name, DESENS_BATCH, "publish", DESENS_ROUNDS, "")
    elif enable_desensitize and desensitize_mapper.ready:
        _sig = _table_signature(desensitize_mapper, wb)
        _scrub_metadata(wb, 1, desensitize_mapper.keyfp, desensitize_mapper.salt.hex(),
                        DESENS_HINT, file_path.name, DESENS_BATCH, "", DESENS_ROUNDS, _sig)
    else:
        _scrub_metadata(wb, 0, "", "")
    wb.save(output_path)
    return output_path, data

def process_old_xls(file_path, enable_sum, enable_internal_stats, collect_data,
                    manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                    enable_desensitize, desensitize_titles, enable_money_clean, header_rows=None,
                    desensitize_by_sheet=None):
    file_path = Path(file_path)
    try:
        df = pd.read_excel(file_path, header=None, dtype=str, engine='xlrd')
    except Exception as e:
        raise Exception(f"无法读取旧版 .xls 文件（需要 xlrd 库，或文件已损坏）：{e}")
    df = df.astype(object).where(pd.notna(df), None)
    wb = grid_to_workbook(df.values.tolist())
    ws = wb.active
    header_row = (header_rows or {}).get('Sheet1') or detect_header_row(ws, ws.max_row, ws.max_column)
    _titles = (desensitize_by_sheet or {}).get("Sheet1", desensitize_titles)
    desensitize_cols = resolve_desensitize_cols(ws, ws.max_column, header_row,
                                                enable_desensitize, _titles)
    data = process_worksheet_multi(ws, wb, "Sheet1", enable_sum, enable_internal_stats, collect_data,
                                   manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                                   enable_desensitize, desensitize_cols, enable_money_clean, header_row,
                                   desensitize_titles=_titles)
    output_path = file_path.parent / f"{file_path.stem}_已处理.xlsx"
    if enable_desensitize and DESENS_PUBLISH:
        # 发布模式：外网要改数据、复制副本，签名只会添乱 —— 不生成
        _scrub_metadata(wb, 1, "-", desensitize_mapper.salt.hex(),
                        DESENS_HINT, file_path.name, DESENS_BATCH, "publish", DESENS_ROUNDS, "")
    elif enable_desensitize and desensitize_mapper.ready:
        _sig = _table_signature(desensitize_mapper, wb)
        _scrub_metadata(wb, 1, desensitize_mapper.keyfp, desensitize_mapper.salt.hex(),
                        DESENS_HINT, file_path.name, DESENS_BATCH, "", DESENS_ROUNDS, _sig)
    else:
        _scrub_metadata(wb, 0, "", "")
    wb.save(output_path)
    return output_path, data

def process_xlsx(file_path, enable_sum, enable_internal_stats, collect_data,
                 manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                 enable_desensitize, desensitize_titles, enable_money_clean, header_rows=None,
                 desensitize_by_sheet=None):
    file_path = Path(file_path)
    wb = load_workbook(file_path, data_only=True)
    _gen, _kfp, _salt = _read_mask_marker(wb)
    if _gen and enable_desensitize:
        raise Exception("该文件已是第 %d 代脱敏结果，拒绝再次脱敏（请先用「还原脱敏」解开）" % _gen)
    all_data = []
    for sheet_name in list(wb.sheetnames):
        # V1.6.9：重复处理“已处理”文件时，不能把上次生成的统计表再当数据表统计
        if is_stat_sheet_name(sheet_name):
            log(f"跳过本工具生成的统计汇总表：{sheet_name}")
            continue
        ws = wb[sheet_name]
        try:
            header_row = (header_rows or {}).get(sheet_name) or detect_header_row(ws, ws.max_row, ws.max_column)
            if desensitize_by_sheet is not None:
                sheet_titles = desensitize_by_sheet.get(sheet_name, [])
            else:
                sheet_titles = desensitize_titles
            desensitize_cols = []
            if enable_desensitize and sheet_titles:
                for title in sheet_titles:
                    col = find_column_by_title(ws, ws.max_column, header_row, [title])
                    if col is not None:
                        desensitize_cols.append(col)
                    else:
                        warn(f"工作表「{sheet_name}」未找到脱敏列标题 '{title}'，跳过该列")
            data = process_worksheet_multi(ws, wb, sheet_name, enable_sum, enable_internal_stats, collect_data,
                                           manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
                                           enable_desensitize, desensitize_cols, enable_money_clean, header_row,
                                           desensitize_titles=sheet_titles)
            all_data.extend(data)
        except Exception as e:
            # 单个工作表出错不再拖垮整批文件
            warn(f"工作表「{sheet_name}」处理失败：{e}（已跳过该工作表，其余工作表继续）")
    output_path = file_path.parent / f"{file_path.stem}_已处理.xlsx"
    if enable_desensitize and DESENS_PUBLISH:
        # 发布模式：外网要改数据、复制副本，签名只会添乱 —— 不生成
        _scrub_metadata(wb, 1, "-", desensitize_mapper.salt.hex(),
                        DESENS_HINT, file_path.name, DESENS_BATCH, "publish", DESENS_ROUNDS, "")
    elif enable_desensitize and desensitize_mapper.ready:
        _sig = _table_signature(desensitize_mapper, wb)
        _scrub_metadata(wb, 1, desensitize_mapper.keyfp, desensitize_mapper.salt.hex(),
                        DESENS_HINT, file_path.name, DESENS_BATCH, "", DESENS_ROUNDS, _sig)
    else:
        _scrub_metadata(wb, 0, "", "")
    wb.save(output_path)
    return output_path, all_data

def fix_files(file_paths, enable_sum, enable_internal_stats, enable_cross_summary,
              manual_mode, user_amount_titles, user_counterparty_titles, user_counterparty_id_titles,
              enable_desensitize, desensitize_titles, enable_money_clean,
              progress_callback, finish_callback, overrides=None):
    global desensitize_mapper
    # V1.8.7：互斥不只做在界面上，后端也强制一遍 ——
    # 走代码/命令行调用时不会有 GUI 帮你清零，金额列加密后要是还在算求和，数就成了密文的和。
    if enable_desensitize and (enable_sum or enable_internal_stats or enable_cross_summary
                               or enable_money_clean or manual_mode):
        log("脱敏模式：已忽略 求和/内部统计/跨文件汇总/金额清理/手动选列（与脱敏互斥）")
        enable_sum = False
        enable_internal_stats = False
        enable_cross_summary = False
        enable_money_clean = False
        manual_mode = False
    # V1.8.8：给这一批文件一个共同批次号（写进每个文件属性，代替原来的金库清单）
    set_desens_batch(datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
    # V2.0.0：两种模式共用一套内核
    #   发布模式：单向假名，盐口令可有可无（没有就用本批随机盐），不可还原
    #   保管模式：可逆，必须有口令
    if DESENS_PUBLISH:
        desensitize_mapper = DesensitizeMapper()
        desensitize_mapper.publish = True
        if DESENS_PASSWORD:
            desensitize_mapper.set_password(DESENS_PASSWORD)
        log("脱敏方式 = 发布（单向不可逆，外网可随意修改/复制）")
    else:
        desensitize_mapper = DesensitizeMapper(password=DESENS_PASSWORD)
    if enable_desensitize and not DESENS_PUBLISH and not DESENS_PASSWORD:
        warn("已勾选“启用信息脱敏”，但没有设置口令，本次不会脱敏")
    elif enable_desensitize and DESENS_PUBLISH:
        log("发布模式：不需要口令，也不生成完整性签名")
    elif enable_desensitize:
        log(f"脱敏内核 = {DesensitizeMapper.VERSION}，口令指纹 = {desensitize_mapper.keyfp}")
        try:
            save_password_hint(DESENS_HINT, desensitize_mapper.keyfp, desensitize_mapper.salt.hex())
        except Exception:
            pass
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
            reset_masked_cols()
            if enable_desensitize and DESENS_PER_FILE_SALT and DESENS_PASSWORD:
                # 一份文件一把盐：每个文件重新派生子密钥，密文互不相通
                desensitize_mapper = DesensitizeMapper(password=DESENS_PASSWORD)
                log("  一份文件一把盐：本文件独立盐 = %s" % desensitize_mapper.salt.hex()[:12])
            progress_callback(idx + 1, total, file_path.name)
            _hr = None
            _dbs = None
            if overrides and _norm_key(file_path) in overrides:
                _ov = overrides[_norm_key(file_path)]
                _amt, _cp, _cid, _des = _ov[0], _ov[1], _ov[2], _ov[3]
                if len(_ov) >= 5:
                    _hr = _ov[4]
                if len(_ov) >= 6:
                    _dbs = _ov[5]
            else:
                _amt, _cp, _cid, _des = (user_amount_titles, user_counterparty_titles, user_counterparty_id_titles, desensitize_titles)
            _t_file = time.time()
            try:
                ext = file_path.suffix.lower()
                if ext == '.csv':
                    out_path, data = process_csv(file_path, enable_sum, enable_internal_stats, enable_cross_summary,
                                                 manual_mode, _amt, _cp, _cid,
                                                 enable_desensitize, _des, enable_money_clean, header_rows=_hr,
                                                 desensitize_by_sheet=_dbs)
                elif ext == '.xls':
                    out_path, data = process_old_xls(file_path, enable_sum, enable_internal_stats, enable_cross_summary,
                                                     manual_mode, _amt, _cp, _cid,
                                                     enable_desensitize, _des, enable_money_clean, header_rows=_hr,
                                                     desensitize_by_sheet=_dbs)
                elif ext == '.xlsx':
                    out_path, data = process_xlsx(file_path, enable_sum, enable_internal_stats, enable_cross_summary,
                                                  manual_mode, _amt, _cp, _cid,
                                                  enable_desensitize, _des, enable_money_clean, header_rows=_hr,
                                                  desensitize_by_sheet=_dbs)
                else:
                    raise Exception(f"不支持的文件类型：{ext}")
            except Exception as e:
                # V1.6.9：单个文件失败不再中断整批任务
                log_exc(f"文件处理失败：{file_path.name} -> {e}")
                failed.append(f"{file_path.name}：{e}")
                continue
            log(f"文件处理完成：{file_path.name}（用时 {time.time() - _t_file:.1f} 秒）")
            processed_paths.append(out_path)
            remember_dir(out_path)
            if enable_cross_summary:
                all_collected.extend(data)
                log(f"  收集到 {len(data)} 条对手方明细")

        if enable_desensitize and getattr(desensitize_mapper, "publish", False):
            # V2.0.3：发布模式产出「映射表」——代号 ↔ 真实值，留在内网。
            # 外网只有脱敏表，反推不出代号；回内网拿这张表一换就回来。
            _mp = getattr(desensitize_mapper, "mask_to_orig", {}) or {}
            if _mp:
                _ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                _stem = Path(file_paths[0]).stem if len(file_paths) == 1 else "批量%d个" % len(file_paths)
                _map_path = Path(file_paths[0]).parent / ("密码本_%s_%s.json" % (_stem, _ts))
                try:
                    with open(_map_path, "w", encoding="utf-8") as f:
                        json.dump({
                            "version": "alias-v1",
                            "created": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "salt": desensitize_mapper.salt.hex(),
                            "batch": DESENS_BATCH,
                            "files": [os.path.basename(str(p)) for p in file_paths],
                            "map": _mp,
                        }, f, ensure_ascii=False, indent=1)
                    mapping_path = _map_path
                    remember_dir(_map_path)
                    log("密码本已保存：%s（%d 条代号）" % (_map_path, len(_mp)))
                    log("  密码本留在内网；外网弄完，用「还原脱敏」查它换回真身。")
                except Exception as e:
                    warn("脱敏映射表保存失败：%s" % e)

        if enable_desensitize and desensitize_mapper.ready:
            # V1.8.8：不再生成金库。
            # 密钥 = PBKDF2(口令, salt)，salt 写在每个 xlsx 的自定义属性里，
            # 所以"换一台电脑、发给别人，只要输同一个口令"就能还原，
            # 不需要任何附带的映射/金库文件。
            log("脱敏完成：未生成金库文件（密钥由口令+文件内 salt 决定，凭口令即可还原）"
                "；本批批次号 = %s" % (DESENS_BATCH or "-"))

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

# ---------- 口令对话框（参照 Windows 账户设置） ----------
_HINT_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "BankFlowTool")
_HINT_FILE = os.path.join(_HINT_DIR, "password_hints.json")
DESENS_HINT = ""


def set_desens_hint(h):
    global DESENS_HINT
    DESENS_HINT = (h or "").strip()


def load_password_hints():
    try:
        with open(_HINT_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return [d for d in data if isinstance(d, dict)]
    except Exception:
        return []


def save_password_hint(hint, keyfp="", salt=""):
    """只存"提示 + 指纹 + 盐"，绝不存口令本身。指纹只够校验对错，反推不出口令。"""
    hint = (hint or "").strip()
    if not hint and not (keyfp and salt):
        return
    try:
        os.makedirs(_HINT_DIR, exist_ok=True)
        items = [r for r in load_password_hints() if not (keyfp and r.get("keyfp") == keyfp)]
        items.insert(0, {
            "hint": hint,
            "keyfp": keyfp or "",
            "salt": salt or "",
            "used_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        })
        items = items[:20]
        tmp = _HINT_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(items, f, ensure_ascii=False, indent=2)
        os.replace(tmp, _HINT_FILE)
    except Exception as e:
        log(f"密码提示保存失败（不影响使用）：{e}")


def _eye_widget(parent, entry, bg):
    """画一只小眼睛：点一下在明文/密文之间切换。"""
    cv = tk.Canvas(parent, width=28, height=20, highlightthickness=0, bg=bg, cursor="hand2")
    cv.create_oval(3, 5, 25, 15, outline="#555", width=1)
    pupil = cv.create_oval(10, 7, 18, 13, fill="#333", outline="")
    slash = cv.create_line(4, 17, 24, 3, fill="#c0392b", width=1.6, state="hidden")

    def toggle(_e=None):
        try:
            hidden = bool(entry.cget("show"))
            entry.config(show="" if hidden else "*")
            cv.itemconfig(pupil, state="hidden" if hidden else "normal")
            cv.itemconfig(slash, state="normal" if hidden else "hidden")
            entry.focus_set()
        except Exception:
            pass

    cv.bind("<Button-1>", toggle)
    try:
        entry._toggle_show = toggle      # 也给键盘/测试留个入口
    except Exception:
        pass
    return cv


def _pwd_strength(p):
    """返回 (等级, 颜色, 一句人话)。只评估，不拦。"""
    n = len(p)
    if n == 0:
        return ("", "#888888", "输入口令后这里显示强度")
    has_d = any(c.isdigit() for c in p)
    has_l = any("a" <= c <= "z" for c in p)
    has_u = any("A" <= c <= "Z" for c in p)
    has_s = any(not c.isalnum() for c in p)
    kinds = sum([has_d, has_l, has_u, has_s])
    if n < 6:
        return ("太短", "#c0392b", "至少 6 位才能真正加密")
    if n <= 7 and kinds <= 1:
        return ("很弱", "#c0392b", "%d 位全是同一类字符，单卡几分钟到几小时能撞开" % n)
    score = n + kinds * 3
    if score < 14:
        return ("偏弱", "#b8860b", "长度够，但字类单一，经不住针对性字典")
    if score < 20:
        return ("一般", "#b8860b", "能用。再加长一点或掺个符号会更稳")
    if score < 30:
        return ("不错", "#0a5", "这个长度，撞库基本没戏")
    return ("很强", "#0a5", "够用了")

class PasswordDialog(tk.Toplevel):
    """mode='set' 设置口令；mode='restore' 输入口令并即时校验。"""

    def __init__(self, parent, mode="set", keyfp="", salt="", note=""):
        tk.Toplevel.__init__(self, parent)
        self.result = None
        self.mode = mode
        self.keyfp = (keyfp or "").strip().lower()
        self.salt = (salt or "").strip()
        self._hint_map = {}
        self.title("设置脱敏口令" if mode == "set" else "输入脱敏口令")
        self.resizable(False, False)
        bg = "#f7f7f7"
        self.configure(bg=bg, padx=18, pady=14)
        try:
            self.transient(parent)
            self.grab_set()
            self.attributes("-topmost", True)
        except Exception:
            pass

        if mode == "set":
            head = "设置本次脱敏的口令"
            sub = "口令是唯一的还原钥匙：不落盘、不写日志。忘了就再也解不开。"
        else:
            head = "输入脱敏时用的口令"
            sub = note or "凭口令即可还原，不再依赖映射文件。"
        tk.Label(self, text=head, bg=bg, font=("微软雅黑", 11, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w")
        tk.Label(self, text=sub, bg=bg, fg="#666", font=("微软雅黑", 9),
                 wraplength=400, justify="left").grid(row=1, column=0, columnspan=3, sticky="w", pady=(2, 10))

        r = 2
        if mode == "set":
            tk.Label(self, text="密码提示", bg=bg, font=("微软雅黑", 9)).grid(
                row=r, column=0, sticky="e", padx=(0, 8), pady=3)
            self.hint_var = tk.StringVar(value=DESENS_HINT)
            tk.Entry(self, textvariable=self.hint_var, width=36, font=("微软雅黑", 10)).grid(
                row=r, column=1, columnspan=2, sticky="we", pady=3)
            tk.Label(self, text="（可留空。还原时它会列出来，帮你想起口令）", bg=bg, fg="#888",
                     font=("微软雅黑", 8)).grid(row=r + 1, column=1, columnspan=2, sticky="w")
            r += 2
        else:
            hints = load_password_hints()
            labels = []
            for it in hints:
                h = (it.get("hint") or "").strip() or "（未填提示）"
                lab = "%s    %s" % (it.get("used_at", ""), h)
                labels.append(lab)
                self._hint_map[lab] = it
            if labels:
                tk.Label(self, text="已记录的提示", bg=bg, font=("微软雅黑", 9)).grid(
                    row=r, column=0, sticky="e", padx=(0, 8), pady=3)
                self.hint_cb = ttk.Combobox(self, values=labels, state="readonly",
                                            width=34, font=("微软雅黑", 9))
                self.hint_cb.grid(row=r, column=1, columnspan=2, sticky="we", pady=3)
                self.hint_cb.bind("<<ComboboxSelected>>", self._on_hint_pick)
                r += 1
                self.hint_label = tk.Label(self, text="（选一条就知道当时写的是什么提示）", bg=bg,
                                           fg="#0a5", font=("微软雅黑", 10, "bold"),
                                           wraplength=400, justify="left")
                self.hint_label.grid(row=r, column=0, columnspan=3, sticky="w", pady=(0, 6))
                r += 1

        tk.Label(self, text="口令", bg=bg, font=("微软雅黑", 9)).grid(
            row=r, column=0, sticky="e", padx=(0, 8), pady=3)
        self.pwd_entry = tk.Entry(self, show="*", width=36, font=("微软雅黑", 10))
        self.pwd_entry.grid(row=r, column=1, sticky="we", pady=3)
        _eye_widget(self, self.pwd_entry, bg).grid(row=r, column=2, padx=(4, 0))
        self.pwd_entry.bind("<KeyRelease>", self._on_change)
        r += 1

        if mode == "set":
            tk.Label(self, text="确认口令", bg=bg, font=("微软雅黑", 9)).grid(
                row=r, column=0, sticky="e", padx=(0, 8), pady=3)
            self.pwd2_entry = tk.Entry(self, show="*", width=36, font=("微软雅黑", 10))
            self.pwd2_entry.grid(row=r, column=1, sticky="we", pady=3)
            _eye_widget(self, self.pwd2_entry, bg).grid(row=r, column=2, padx=(4, 0))
            self.pwd2_entry.bind("<KeyRelease>", self._on_change)
            r += 1

        if mode == "set":
            self.strength_label = tk.Label(self, text="", bg=bg, font=("微软雅黑", 9),
                                           wraplength=400, justify="left")
            self.strength_label.grid(row=r, column=0, columnspan=3, sticky="w", pady=(8, 0))
            r += 1

        self.state_label = tk.Label(self, text="", bg=bg, font=("微软雅黑", 9),
                                    wraplength=400, justify="left")
        self.state_label.grid(row=r, column=0, columnspan=3, sticky="w", pady=(2, 8))
        r += 1

        btns = tk.Frame(self, bg=bg)
        btns.grid(row=r, column=0, columnspan=3, sticky="e")
        self.ok_btn = tk.Button(btns, text="确定", width=10, font=("微软雅黑", 9), command=self._ok)
        self.ok_btn.pack(side=tk.LEFT, padx=4)
        tk.Button(btns, text="取消", width=10, font=("微软雅黑", 9), command=self._cancel).pack(side=tk.LEFT)

        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self._cancel())
        self.pwd_entry.focus_set()
        self._on_change()
        self.update_idletasks()
        try:
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw = max(parent.winfo_width(), 400)
            ph = max(parent.winfo_height(), 300)
            w, h = self.winfo_width(), self.winfo_height()
            self.geometry("+%d+%d" % (px + (pw - w) // 2, py + (ph - h) // 3))
        except Exception:
            pass

    # --- 交互
    def _on_change(self, _e=None):
        if self.mode != "set":
            return
        try:
            p1 = self.pwd_entry.get()
            p2 = self.pwd2_entry.get()
        except Exception:
            return
        # 强度：随口令实时变，跟确认框填没填无关
        try:
            lvl, color, tip = _pwd_strength(p1)
            self.strength_label.config(
                text=("强度：%s —— %s" % (lvl, tip)) if lvl else tip, fg=color)
        except Exception:
            pass

        # 一致性：只在"该判的时候"才判
        _tip = "口令至少 6 位，建议 8 位以上，含数字、大小写英文字符等"
        if len(p1) < 6:
            self.state_label.config(text=_tip, fg="#c0392b")
            self.ok_btn.config(state=tk.DISABLED)
        elif not p2:
            self.state_label.config(text=_tip, fg="#888888")
            self.ok_btn.config(state=tk.DISABLED)
        elif p2 != p1:
            self.state_label.config(text="两次输入不一致", fg="#c0392b")
            self.ok_btn.config(state=tk.DISABLED)
        else:
            self.state_label.config(text="✓ 两次一致，可以点确定", fg="#0a5")
            self.ok_btn.config(state=tk.NORMAL)

    def _on_hint_pick(self, _e=None):
        it = self._hint_map.get(self.hint_cb.get()) or {}
        h = (it.get("hint") or "").strip()
        self.hint_label.config(text="提示：" + (h if h else "（这条没填提示）"))

    def _cancel(self):
        self.result = None
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def _ok(self):
        pwd = self.pwd_entry.get()
        if self.mode == "set":
            if len(pwd) < 6 or pwd != self.pwd2_entry.get():
                return
            self.result = {"password": pwd, "hint": self.hint_var.get().strip()}
            try:
                self.grab_release()
            except Exception:
                pass
            self.destroy()
            return

        if not pwd:
            self.state_label.config(text="请输入口令", fg="#c0392b")
            return
        if self.keyfp and self.salt:
            self.state_label.config(text="正在校验口令…", fg="#666")
            try:
                self.update_idletasks()
            except Exception:
                pass
            try:
                probe = DesensitizeMapper(password=pwd, salt=bytes.fromhex(self.salt))
                if probe.keyfp != self.keyfp:
                    self.state_label.config(
                        text="✗ 口令不对（和文件里记录的指纹 %s 不符），再试试" % self.keyfp, fg="#c0392b")
                    self.pwd_entry.select_range(0, tk.END)
                    self.pwd_entry.focus_set()
                    return
                self.state_label.config(text="✓ 口令正确", fg="#0a5")
            except Exception:
                pass
        self.result = {"password": pwd, "hint": ""}
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()


def _ask_password(mode, keyfp="", salt="", note=""):
    parent = tk._default_root
    if parent is None:
        return None
    dlg = PasswordDialog(parent, mode=mode, keyfp=keyfp, salt=salt, note=note)
    parent.wait_window(dlg)
    return dlg.result


# ---------- 还原功能（V1.8.3：口令直解 + 提示列表 + 指纹即时校验） ----------
_ALIAS_RE = None


def _alias_re():
    global _ALIAS_RE
    if _ALIAS_RE is None:
        # 名-ZYPPZ7 这种：1 个中文前缀 + 连字符 + 6 位 base32（A-Z 2-7）
        _ALIAS_RE = re.compile(r"^[名号户值址]-[A-Za-z2-7]{6}$")
    return _ALIAS_RE


def _count_aliases(wb, limit=5):
    """数一数工作簿里有多少个「代号」。

    用途：文件可能被复制粘贴进新工作簿，MODE 标记早就没了 ——
    这时候只能靠内容判断这是不是密码本加密过的数据。
    """
    pat = _alias_re()
    n, samples = 0, []
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                if row is None:
                    continue
                for v in row:
                    if isinstance(v, str):
                        s = v.strip()
                        if pat.match(s):
                            n += 1
                            if len(samples) < limit and s not in samples:
                                samples.append(s)
    except Exception:
        pass
    return n, samples


def _restore_by_alias_map(file_path):
    """发布模式专用：拿内网的映射表，把代号换回真实值。

    跟"解密"不是一回事 —— 不看文件有没有被改过、被复制过，
    只要单元格里还留着代号，就换成真实值。
    """
    map_path = filedialog.askopenfilename(
        title="选内网生成的那份密码本（.json）",
        initialdir=os.path.dirname(os.path.abspath(file_path)) or last_dir(),
        filetypes=[("密码本", "*.json"), ("所有文件", "*.*")]
    )
    if not map_path:
        messagebox.showinfo("已取消", "没有密码本就换不回真身。")
        return
    try:
        with open(map_path, "r", encoding="utf-8") as f:
            raw = json.load(f)
        mp = raw.get("map") or {}
        if not mp:
            messagebox.showerror("密码本是空的", "这个文件里没有 map 字段。")
            return
    except Exception as e:
        messagebox.showerror("读不了密码本", str(e))
        return

    try:
        ext = Path(file_path).suffix.lower()
        if ext == '.csv':
            wb = grid_to_workbook(read_csv_grid(file_path))
        elif ext == '.xls':
            df = pd.read_excel(file_path, header=None, dtype=str, engine='xlrd')
            df = df.astype(object).where(pd.notna(df), None)
            wb = grid_to_workbook(df.values.tolist())
        elif ext == '.xlsx':
            wb = load_workbook(file_path, data_only=True)
        else:
            messagebox.showerror("错误", "不支持的文件格式")
            return

        # 宽容匹配：外面可能被手工改过大小写、前后多了空格
        _mp_norm = {}
        for k, v in mp.items():
            _mp_norm[k.strip().upper()] = v

        changed, miss = 0, set()
        sample = []
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if not isinstance(v, str) or not v.strip():
                        continue
                    key = v.strip().upper()
                    if key in _mp_norm:
                        if len(sample) < 5:
                            sample.append("%s: %s -> %s" % (cell.coordinate, v, _mp_norm[key]))
                        cell.value = _mp_norm[key]
                        changed += 1
                    elif _alias_re().match(key):
                        miss.add(key)
        out_path = Path(file_path).parent / f"{Path(file_path).stem}_换回真身.xlsx"
        _scrub_metadata(wb, 0, "", "")
        wb.save(out_path)
        log(f"查表还原完成：{out_path}，替换 {changed} 个单元格，缺失代号 {len(miss)} 个")
        tip = "换回 %d 个单元格。\n已保存至：\n%s" % (changed, out_path)
        if sample:
            tip += "\n\n样例：\n" + "\n".join(sample)
        if miss:
            tip += "\n\n有 %d 个代号在密码本里找不到（可能是外网新写的），已保持原样。" % len(miss)
        messagebox.showinfo("完成", tip)
    except Exception as e:
        log_exc("映射还原失败")
        messagebox.showerror("错误", f"还原失败：{e}")


def restore_file():
    file_path = filedialog.askopenfilename(
        title="选择已脱敏的 Excel 或 CSV 文件（凭口令还原，不需要任何附带文件）",
        initialdir=last_dir(),
        filetypes=[("Excel/CSV 文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")]
    )
    if not file_path:
        return

    # 先从文件里读脱敏标记：代数 / salt / 指纹
    gen, kfp, salt_hex = 0, "", ""
    try:
        if Path(file_path).suffix.lower() == ".xlsx":
            _wb = load_workbook(file_path, read_only=True, data_only=True)
            try:
                gen, kfp, salt_hex = _read_mask_marker(_wb)
                globals()["_file_hint"] = _read_hint(_wb)
                globals()["_file_mode"] = _read_mode(_wb)
                globals()["_file_rounds"] = _read_rounds(_wb)
                globals()["_file_cols"] = _read_cols_map(_wb)
                globals()["_file_hdr"] = _read_hdr_masked(_wb)
            finally:
                _wb.close()
    except Exception as e:
        log(f"读取脱敏标记失败（继续尝试还原）：{e}")
    log(f"待还原文件: {file_path} | 代数={gen} 指纹={kfp or '-'} salt={'有' if salt_hex else '无'}")

    # V2.0.7：判断依据改成「内容里有没有代号」。
    # 文件被复制粘贴进新工作簿后，MODE 标记就没了，只看标记会漏判。
    _alias_n, _alias_samples = 0, []
    try:
        if Path(file_path).suffix.lower() == ".xlsx":
            _wb2 = load_workbook(file_path, read_only=True, data_only=True)
            try:
                _alias_n, _alias_samples = _count_aliases(_wb2)
            finally:
                _wb2.close()
    except Exception:
        pass

    if _alias_n:
        log(f"内容里检测到 {_alias_n} 个代号，例如 {_alias_samples[:3]}")
        if messagebox.askyesno(
                "这是密码本加密的内容",
                "这个文件里有 %d 个代号（例如 %s）。\n\n"
                "用密码本查表换回真身吗？\n"
                "  是 —— 下一步选内网生成的那份密码本（.json）\n"
                "  否 —— 按口令加密来解" % (_alias_n, "、".join(_alias_samples[:3]))):
            return _restore_by_alias_map(file_path)
    elif gen and globals().get("_file_mode") == "publish":
        messagebox.showinfo(
            "这是密码本加密的文件",
            "文件标记显示它用密码本加密过，但内容里没扫到代号。\n"
            "可能代号已经被替换或删掉了。\n\n"
            "接下来仍然可以选密码本试试查表。")
        return _restore_by_alias_map(file_path)

    if gen >= 2:
        warn("这个文件是第 %d 代脱敏（历史遗留的多代叠加），只能恢复出第 %d 代的值，"
             "要真正还原本请用 restore_multigen.py 走多代反查。" % (gen, gen - 1))

    note = ""
    _fh = globals().pop("_file_hint", "") or ""
    if _fh:
        note = "文件里记着提示：%s\n" % _fh
    if kfp:
        note += "已从文件读到校验指纹 %s，输错会当场提示。" % kfp

    if gen and not salt_hex:
        messagebox.showerror(
            "这个文件得先用老办法解",
            "它里面没记录 salt —— 那份文件是 V1.8.2 及更早版本生成的，\n"
            "当年随机盐只存在金库文件里，光凭口令算不出同一个密钥。\n\n"
            "两条路：\n"
            "  1) 找到同批次的 脱敏金库_*.vault + 老版本软件还原；\n"
            "  2) 直接拿原始数据用 V1.8.8 重跑一遍，以后就只认口令，不再有金库。")
        return

    got = _ask_password("restore", keyfp=kfp, salt=salt_hex, note=note)
    if not got:
        messagebox.showinfo("已取消", "还原需要口令。")
        return
    pwd = got["password"]

    try:
        if salt_hex:
            mapper = DesensitizeMapper(password=pwd, salt=bytes.fromhex(salt_hex))
        else:
            mapper = DesensitizeMapper(password=pwd)
        mapper.randomized = (globals().pop("_file_mode", "det") == "rnd")
        set_desens_rounds(globals().pop("_file_rounds", DESENS_ROUNDS_DEFAULT))
        if kfp and mapper.keyfp != kfp:
            messagebox.showerror(
                "口令不对",
                "按这个口令算出来的指纹是 %s，文件里记录的是 %s。\n"
                "继续下去只会得到乱码。" % (mapper.keyfp, kfp))
            return
    except Exception as e:
        messagebox.showerror("错误", "准备密钥失败：%s" % e)
        return

    try:
        file_ext = Path(file_path).suffix.lower()
        if file_ext == '.csv':
            wb = grid_to_workbook(read_csv_grid(file_path))
        elif file_ext == '.xls':
            df = pd.read_excel(file_path, header=None, dtype=str, engine='xlrd')
            df = df.astype(object).where(pd.notna(df), None)
            wb = grid_to_workbook(df.values.tolist())
        elif file_ext == '.xlsx':
            wb = load_workbook(file_path, data_only=True)
        else:
            messagebox.showerror("错误", "不支持的文件格式")
            return

        _sig_saved = _read_sig(wb)
        if _sig_saved:
            _sig_now = _table_signature(mapper, wb)
            if _sig_now == _sig_saved:
                log("完整性校验通过：文件内容和加密时一致")
            else:
                log("完整性校验失败：内容与加密时不一致")
                if not messagebox.askyesno(
                        "这份数据被改过",
                        "文件内容和当初加密的那份不一致 —— 有人动过它。\n\n"
                        "改过之后解出来的名字可能是错的，而你还会以为那是真的。\n\n"
                        "仍然要继续还原吗？（建议先找发件人重新要一份）"):
                    return
        else:
            log("这个文件没有完整性印记（V1.9.8 及更早生成），跳过校验")

        out_path = Path(file_path).parent / f"{Path(file_path).stem}_还原.xlsx"
        changed, sample = 0, []
        _cmap = globals().pop("_file_cols", {}) or {}
        _hdr_masked = bool(globals().pop("_file_hdr", False))
        for _si, sheet in enumerate(wb.worksheets):
            _cols = set(_cmap.get(_si) or [])
            if _hdr_masked:
                # 表头也加密了：从第 1 行开始解，但仍然只解真正加密过的列，
                # 否则用户没勾的列（比如金额列）会被误解成乱码
                _start_row, _only = 1, (set(_cols) if _cols else None)
            elif _cols:
                # 只解真正加密过的列，从表头下一行开始 —— 否则明文表头会被解成乱码
                try:
                    _h = detect_header_row(sheet, sheet.max_row, sheet.max_column)
                except Exception:
                    _h = 1
                _start_row, _only = _h + 1, _cols
            else:
                # 老文件没记录列：保持原来的全表解
                _start_row, _only = 1, None
            for row in sheet.iter_rows(min_row=_start_row):
                for cell in row:
                    if cell.value is None or not isinstance(cell.value, str):
                        continue
                    if _only is not None and cell.column not in _only:
                        continue
                    before = cell.value
                    after = mapper.unmask(before)
                    if after != before:
                        cell.value = after
                        changed += 1
                        if len(sample) < 5:
                            sample.append("%s: %s -> %s" % (cell.coordinate, before[:14], after[:14]))

        _scrub_metadata(wb, 0, "", "")
        wb.save(out_path)
        log(f"还原完成：{out_path}，替换 {changed} 个单元格")
        messagebox.showinfo(
            "完成",
            "还原完成！共替换 %d 个单元格。\n已保存至：\n%s\n\n样例：\n%s"
            % (changed, out_path, "\n".join(sample) if sample else "（没有可还原的内容）"))
    except Exception as e:
        log_exc("还原失败")
        messagebox.showerror("错误", f"还原失败：{e}")


# ---------- GUI ----------
class App:
    def __init__(self, root):
        self.root = root
        self.root.title(f"银行流水梳理工具 {VERSION}")
        self.root.geometry("680x850")
        self.root.configure(bg="#f0f0f0")

        self.enable_money_clean = BooleanVar(value=False)
        self.enable_sum = BooleanVar(value=False)
        self.enable_internal_stats = BooleanVar(value=False)
        self.enable_cross_summary = BooleanVar(value=False)
        self.manual_mode = BooleanVar(value=False)
        self.enable_desensitize = BooleanVar(value=False)
        self.per_file_salt = BooleanVar(value=False)
        self.randomize_mode = BooleanVar(value=False)
        self.fpe_rounds = tk.IntVar(value=DESENS_ROUNDS_DEFAULT)
        self.mask_header = BooleanVar(value=False)
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
                                 "【信息脱敏】\n只处理勾选的列，没勾的原样保留。\n"
                                 "  口令加密：整表加密，凭口令还原；文件被编辑过就解不开。\n"
                                 "  密码本加密：关键信息换成代号（名-7K3MQP），不要口令；\n"
                                 "              外网随便改，回内网用密码本查表换回真身。\n\n"
                                 "【还原脱敏】\n口令加密 → 输口令；密码本加密 → 选密码本查表。\n\n"
                                 "【导出运行日志】\n当遇到异常时，可点击“导出运行日志”按钮，将详细运行信息导出供分析。")
        info_text.tag_config("red_bold", foreground="red", font=("微软雅黑", 9, "bold"))
        info_text.config(state=tk.DISABLED)
        info_text.pack(fill=tk.BOTH, expand=True)

        options_frame = tk.Frame(root, bg="#f0f0f0")
        options_frame.pack(pady=5)

        row = 0
        self.cb_money_clean = tk.Checkbutton(options_frame, text="清理不可见字符，使金额列数据可被正常统计",
                                             variable=self.enable_money_clean, bg="#f0f0f0",
                                             font=("微软雅黑", 10), command=self._sync_exclusive)
        self.cb_money_clean.grid(row=row, column=0, sticky="w", padx=10, pady=2)
        row += 1

        self.cb_sum = tk.Checkbutton(options_frame, text="在每个区块末尾插入绝对值求和行（标黄）",
                                     variable=self.enable_sum, bg="#f0f0f0",
                                     font=("微软雅黑", 10), command=self._sync_exclusive)
        self.cb_sum.grid(row=row, column=0, sticky="w", padx=10, pady=2)
        row += 1

        self.cb_stats = tk.Checkbutton(options_frame, text="在生成表格内添加对手方流水统计表",
                                       variable=self.enable_internal_stats, bg="#f0f0f0",
                                       font=("微软雅黑", 10), command=self._sync_exclusive)
        self.cb_stats.grid(row=row, column=0, sticky="w", padx=10, pady=2)
        row += 1

        self.cb_cross = tk.Checkbutton(options_frame, text="生成对手方流水统计表汇总文件（仅在勾选并选择多个文件时可用）",
                                       variable=self.enable_cross_summary, bg="#f0f0f0",
                                       font=("微软雅黑", 10), command=self._sync_exclusive)
        self.cb_cross.grid(row=row, column=0, sticky="w", padx=10, pady=2)
        row += 1

        row += 1
        self.manual_frame = Frame(options_frame, bg="#f0f0f0")
        self.manual_frame.grid(row=row, column=0, sticky="w", padx=10, pady=10)
        self.cb_manual = Checkbutton(self.manual_frame, text="手动勾选金额/对手方列（自动识别失败时用）",
                                     variable=self.manual_mode, bg="#f0f0f0",
                                     font=("微软雅黑", 10), command=self._on_manual_toggle)
        self.cb_manual.pack(anchor="w")


        row += 1
        self.desensitize_frame = Frame(options_frame, bg="#f0f0f0")
        self.desensitize_frame.grid(row=row, column=0, sticky="w", padx=10, pady=10)
        self.cb_desensitize = Checkbutton(
            self.desensitize_frame,
            text="启用信息脱敏（独立功能，不能和金额梳理类选项同时使用）",
            variable=self.enable_desensitize, bg="#f0f0f0",
            font=("微软雅黑", 10), command=self._on_desensitize_toggle)
        self.cb_desensitize.pack(anchor="w")

        self.mode_publish = BooleanVar(value=False)
        mode_row = tk.Frame(self.desensitize_frame, bg="#f0f0f0")
        mode_row.pack(anchor="w", padx=(18, 0))
        tk.Radiobutton(mode_row, text="口令加密：凭口令还原", variable=self.mode_publish,
                       value=False, bg="#f0f0f0", font=("微软雅黑", 9),
                       command=self._sync_exclusive).pack(side="left")
        tk.Radiobutton(mode_row, text="密码本加密：关键信息换代号，回内网查表还原",
                       variable=self.mode_publish, value=True, bg="#f0f0f0",
                       font=("微软雅黑", 9),
                       command=self._sync_exclusive).pack(side="left", padx=(12, 0))

        self.desensitize_pick = Label(self.desensitize_frame, text="", bg="#f0f0f0", font=("微软雅黑", 9),
                                      fg="#006400", wraplength=620, justify="left")
        self.desensitize_pick.pack_forget()
        self._sync_exclusive()


        btn_frame = tk.Frame(root, bg="#f0f0f0")
        btn_frame.pack(pady=10)
        self.btn = tk.Button(btn_frame, text="选择文件（可多选）", command=self.select_files, font=("微软雅黑", 10))
        self.btn.pack()

        restore_btn = Button(btn_frame, text="还原脱敏（口令加密 / 密码本加密，自动判断）", command=restore_file, font=("微软雅黑", 10), bg="#e0e0e0")
        restore_btn.pack(pady=5)

        restore_hint = tk.Label(btn_frame, text="口令加密的输口令；密码本加密的选密码本查表换回真身（外网改过、复制过都不影响）",
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
        # 勾选状态本身即可，界面上不再附加文字提示；选文件时会弹出勾选窗口
        pass

    def _refresh_manual_picks(self):
        pass

    def toggle_desensitize_input(self):
        pass

    def _refresh_desensitize_pick(self):
        pass

    def _open_wb(self, file_path):
        ext = str(file_path).lower()
        if ext.endswith('.csv'):
            return grid_to_workbook(read_csv_grid(file_path))
        return load_workbook(file_path, data_only=True)

    def _sheet_names(self, file_path):
        """非统计工作表名列表。"""
        try:
            wb = self._open_wb(file_path)
            try:
                return [s for s in wb.sheetnames if not is_stat_sheet_name(s)]
            finally:
                try:
                    wb.close()
                except Exception:
                    pass
        except Exception as exc:
            warn("读取工作表失败：%s" % exc)
            return []

    def _read_row_titles(self, file_path, sheet, row):
        """读某工作表第 row 行的列名；空标题给占位，方便通用表格也能勾。"""
        out = []
        try:
            wb = self._open_wb(file_path)
            try:
                ws = wb[sheet] if sheet in wb.sheetnames else wb.worksheets[0]
                for c in range(1, ws.max_column + 1):
                    v = ws.cell(row=row, column=c).value
                    t = cell_text(v) if v is not None else ""
                    out.append((c, t))
            finally:
                try:
                    wb.close()
                except Exception:
                    pass
        except Exception as exc:
            warn("读取第 %d 行表头失败：%s" % (row, exc))
        return out

    def _guess_row(self, file_path, sheet):
        try:
            wb = self._open_wb(file_path)
            try:
                ws = wb[sheet] if sheet in wb.sheetnames else wb.worksheets[0]
                return detect_header_row(ws, ws.max_row, ws.max_column)
            finally:
                try:
                    wb.close()
                except Exception:
                    pass
        except Exception:
            return 1


    def _extra_segment_titles(self, path, sheet, current_row):
        """本表里其它"表头段"的列名，标注段号，避免多表堆积时漏选。"""
        out = []
        try:
            wb = self._open_wb(path)
            try:
                ws = wb[sheet] if sheet in wb.sheetnames else wb.worksheets[0]
                segs = _sheet_segments(ws, ws.max_column)
                for i, (s, e) in enumerate(segs, start=1):
                    hdr = _segment_header_row(ws, s, e, ws.max_column)
                    if hdr == current_row:
                        continue
                    for c in range(1, ws.max_column + 1):
                        title = cell_text(ws.cell(row=hdr, column=c).value).strip()
                        if title:
                            out.append((title, i))
            finally:
                try:
                    wb.close()
                except Exception:
                    pass
        except Exception:
            pass
        return out

    def _collect_header_titles(self, file_path, sheet=None, row=None):
        """返回某工作表表头列名（去重保序）；sheet/row 省略时用自动识别。"""
        sheets = self._sheet_names(file_path)
        if not sheets:
            return []
        sheet = sheet or sheets[0]
        row = row or self._guess_row(file_path, sheet)
        titles = []
        for _c, t in self._read_row_titles(file_path, sheet, row):
            if t and t not in titles:
                titles.append(t)
        return titles

    def _sheet_signature(self, path, sheet, row):
        """工作表表头签名：把该表（含其它分段）的列名排序后拼起来，用于判断"表头是否一样"。"""
        try:
            titles = {t for _c, t in self._read_row_titles(path, sheet, row) if t}
            for title, _seg in self._extra_segment_titles(path, sheet, row):
                titles.add(title)
            return tuple(sorted(titles))
        except Exception:
            return ()

    def _split_var(self, var):
        """把「a，b、c」这种多值输入拆成列表。"""
        try:
            raw = var.get() if hasattr(var, "get") else str(var or "")
        except Exception:
            raw = ""
        return [p.strip() for p in re.split(r"[，,、]+", raw or "") if p.strip()]

    def _on_manual_toggle(self):
        try:
            self.toggle_manual_input()
        except Exception:
            pass
        self._sync_exclusive()

    def _on_desensitize_toggle(self):
        try:
            self.toggle_desensitize_input()
        except Exception:
            pass
        self._sync_exclusive()

    def _sync_exclusive(self, *_a):
        """脱敏 <-> 金额梳理 互斥。

        脱敏是独立功能：它只改标识类列、不碰金额，也产出不了求和行和统计表；
        金额梳理要读金额、算求和、出统计表。两者同时开，统计出来的就是一堆
        没有意义的数。所以一边开启，另一边整组置灰。
        """
        try:
            des = bool(self.enable_desensitize.get())
            money_vars = (self.enable_money_clean, self.enable_sum,
                          self.enable_internal_stats, self.enable_cross_summary,
                          self.manual_mode)
            money_cbs = (self.cb_money_clean, self.cb_sum, self.cb_stats,
                         self.cb_cross, self.cb_manual)
            if des:
                for v in money_vars:
                    if v.get():
                        v.set(False)
                for cb in money_cbs:
                    cb.config(state=tk.DISABLED)
                self.cb_desensitize.config(state=tk.NORMAL)
                return
            money_on = any(v.get() for v in money_vars)
            for cb in money_cbs:
                cb.config(state=tk.NORMAL)
            self.cb_desensitize.config(state=(tk.DISABLED if money_on else tk.NORMAL))
        except Exception:
            pass

    def _pick_columns(self, path, need, cur, title_prefix="", header_rows=None, sheets=None):
        """选列窗口：可切工作表、可改表头行；每行三到四列勾选框。

        返回 {"amount":[...], "cp":[...], "id":[...], "des":[...], "header_rows":{sheet:row}}
        取消返回 None。
        """
        sheets = sheets or self._sheet_names(path)
        if not sheets:
            messagebox.showwarning("读不到工作表", "这个文件里没有可处理的工作表。")
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

        hdr_state = dict(header_rows or {})
        sel_state = {}          # sheet -> {title: {key: bool}}

        win = tk.Toplevel(self.root)
        win.title(title_prefix + "选择列")
        win.geometry("780x620")
        win.transient(self.root)

        top = tk.Frame(win)
        top.pack(fill="x", padx=10, pady=(10, 4))
        tk.Label(top, text="工作表：", font=("微软雅黑", 9)).pack(side="left")
        sheet_var = tk.StringVar(value=sheets[0])
        sheet_box = ttk.Combobox(top, textvariable=sheet_var, values=sheets, state="readonly", width=20)
        sheet_box.pack(side="left")
        tk.Label(top, text="   表头在第几行：", font=("微软雅黑", 9)).pack(side="left")
        row_var = tk.IntVar(value=0)
        row_spin = tk.Spinbox(top, from_=1, to=10000, textvariable=row_var, width=6, font=("微软雅黑", 9))
        row_spin.pack(side="left")
        share_var = tk.BooleanVar(value=True)
        share_cb = tk.Checkbutton(top, text="表头相同的工作表沿用同一套勾选", variable=share_var,
                                  font=("微软雅黑", 9), bg="#f0f0f0")
        share_cb.pack(side="left", padx=10)
        ref_btn = tk.Button(top, text="重新读取", font=("微软雅黑", 9))
        ref_btn.pack(side="left", padx=8)
        hint = tk.Label(win, text="", font=("微软雅黑", 9), fg="#666666", justify="left", anchor="w", wraplength=740)
        hint.pack(fill="x", padx=10)

        holder = tk.Frame(win)
        holder.pack(fill="both", expand=True, padx=10, pady=4)
        canvas = tk.Canvas(holder, borderwidth=0, highlightthickness=0, bg="white")
        sb = tk.Scrollbar(holder, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg="white")
        canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        def _on_wheel(event, _c=canvas):
            try:
                _c.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except Exception:
                pass

        win.bind("<MouseWheel>", _on_wheel)

        current = {"sheet": sheets[0], "state": {}}

        def build():
            """按当前工作表与行号重建勾选框。"""
            for w in list(inner.children.values()):
                w.destroy()
            sheet = sheet_var.get()
            row = max(1, int(row_var.get() or 1))
            hdr_state[sheet] = row
            titles_raw = self._read_row_titles(path, sheet, row)
            # V1.8.6：同名表头合并成一行显示。
            # 原来 st 用标题当 key，表里有两列都叫「网联/银联」时，
            # 第二次 st[t] = per 会把第一次的覆盖掉，第一个勾选框就成了孤儿 ——
            # 全选怎么点它都不亮（«第三方账户名称» 同理）。
            # 脱敏本来就是按标题匹配整列的，合并成一行正好对上这个语义。
            _colmap = {}
            titles = []
            for _idx, _t in titles_raw:
                if _t:
                    if _t in _colmap:
                        _colmap[_t].append(_idx)
                    else:
                        _colmap[_t] = [_idx]
                        titles.append((_idx, _t))
                else:
                    titles.append((_idx, _t))
            prev = sel_state.setdefault(sheet, {})
            tk.Label(inner, text="表头", bg="white", font=("微软雅黑", 9, "bold")).grid(row=0, column=0, sticky="w", padx=6, pady=2)
            for i, (key, label) in enumerate(cols, start=1):
                tk.Label(inner, text=label, bg="white", font=("微软雅黑", 9, "bold")).grid(row=0, column=i, padx=8)
            st = {}
            r = 1
            for idx, t in titles:
                if t:
                    _at = _colmap.get(t) or [idx]
                    label = t if len(_at) <= 1 else "%s（第 %s 列同名）" % (
                        t, "、".join(str(x) for x in _at))
                else:
                    label = "（第 %d 列，无标题）" % idx
                tk.Label(inner, text=label, bg="white", font=("微软雅黑", 9), anchor="w",
                         justify="left").grid(row=r, column=0, sticky="w", padx=6)
                per = {}
                for i, (key, _lab) in enumerate(cols, start=1):
                    var = tk.BooleanVar(value=bool(prev.get(t, {}).get(key)) if t and key in need else False)
                    cb = tk.Checkbutton(inner, variable=var, bg="white", activebackground="white",
                                        state=("normal" if t else "disabled"))
                    cb.grid(row=r, column=i, padx=8)
                    per[key] = var
                if t:
                    st[t] = per
                r += 1
            # 其它分段的表头也列出来（多表堆在一张表里时更不容易漏）
            seen = {t for _i, t in titles if t}
            for title, seg_no in self._extra_segment_titles(path, sheet, row):
                if title in seen:
                    continue
                seen.add(title)
                tk.Label(inner, text="%s  ｜第 %d 段" % (title, seg_no), bg="white",
                         font=("微软雅黑", 9), anchor="w", justify="left",
                         fg="#555555").grid(row=r, column=0, sticky="w", padx=6)
                per = {}
                for i, (key, _lab) in enumerate(cols, start=1):
                    var = tk.BooleanVar(value=bool(prev.get(title, {}).get(key)))
                    tk.Checkbutton(inner, variable=var, bg="white",
                                   activebackground="white").grid(row=r, column=i, padx=8)
                    per[key] = var
                st[title] = per
                r += 1
            current["state"] = st
            filled = sum(1 for _i, t in titles_raw if t)
            _dup = len(titles_raw) - len(titles)
            _tail = ("，其中 %d 个同名表头已合并成一行" % _dup) if _dup else ""
            hint.config(text="本表共 %d 个表头列（有文字的 %d 个%s）。"
                             "如果上面的不是真正的表头，改「表头在第几行」再点重新读取。"
                        % (len(titles_raw), filled, _tail))

        def _save_current():
            sheet = current.get("sheet")
            if sheet and current.get("state"):
                sel_state[sheet] = {t: {k: v.get() for k, v in per.items()}
                                    for t, per in current["state"].items()}

        def refresh(*_a):
            """按输入框里写的行号重建（手改表头行后点这个 / 按回车）。"""
            _save_current()
            current["sheet"] = sheet_var.get()
            try:
                raw = str(row_spin.get()).strip()
            except Exception:
                raw = ""
            typed = _parse_row_input(raw, 0)
            if typed >= 1:
                row_var.set(typed)
                hdr_state[current["sheet"]] = typed
            build()

        def load_sheet(*_a):
            """切换工作表：载入该表已定的行号，没定过就重新猜。"""
            _save_current()
            current["sheet"] = sheet_var.get()
            row_var.set(hdr_state.get(current["sheet"]) or self._guess_row(path, current["sheet"]))
            build()

        def on_sheet_change(*_a):
            # 勾了"复用"时：切到签名相同的工作表，直接继承已配置的那一套
            if share_var.get():
                new_sheet = sheet_var.get()
                if new_sheet not in sel_state:
                    try:
                        new_row = hdr_state.get(new_sheet) or self._guess_row(path, new_sheet)
                        sig = self._sheet_signature(path, new_sheet, new_row)
                        for other, rows in hdr_state.items():
                            if other == new_sheet or other not in sel_state:
                                continue
                            if self._sheet_signature(path, other, rows) == sig and sig:
                                sel_state[new_sheet] = {k: dict(v) for k, v in sel_state[other].items()}
                                break
                    except Exception:
                        pass
            load_sheet()

        ref_btn.config(command=refresh)
        sheet_box.bind("<<ComboboxSelected>>", on_sheet_change)
        row_spin.bind("<Return>", lambda e: refresh())

        def _bulk(mode):
            """一键全选 / 全不选 / 反选「脱敏」列（只作用于当前工作表）。"""
            st = current.get("state") or {}
            for _t, per in st.items():
                var = per.get("des")
                if var is None:
                    continue
                if mode == "all":
                    var.set(True)
                elif mode == "none":
                    var.set(False)
                else:
                    var.set(not var.get())

        result = {"ok": False}

        def _ok():
            try:
                typed = _parse_row_input(str(row_spin.get()).strip(), 0)
                if typed >= 1:
                    hdr_state[sheet_var.get()] = typed
            except Exception:
                pass
            result["ok"] = True
            win.destroy()

        bulk_bar = tk.Frame(win)
        bulk_bar.pack(pady=(4, 0))
        tk.Label(bulk_bar, text="脱敏列：", font=("微软雅黑", 9)).pack(side="left")
        tk.Button(bulk_bar, text="全选", width=6, font=("微软雅黑", 9),
                  command=lambda: _bulk("all")).pack(side="left", padx=3)
        tk.Button(bulk_bar, text="全不选", width=6, font=("微软雅黑", 9),
                  command=lambda: _bulk("none")).pack(side="left", padx=3)
        tk.Button(bulk_bar, text="反选", width=6, font=("微软雅黑", 9),
                  command=lambda: _bulk("inv")).pack(side="left", padx=3)
        tk.Label(bulk_bar, text="（先全选，再把不要的几列点掉；金额列也可以勾）", font=("微软雅黑", 8),
                 fg="#888888").pack(side="left", padx=6)

        salt_bar = tk.Frame(win)
        salt_bar.pack(pady=(2, 0))
        Checkbutton(salt_bar, text="一份文件一把盐（多文件时各自独立、跨文件对不上；不勾＝同批共用一把盐，可互相对照）",
                    variable=self.per_file_salt, font=("微软雅黑", 9)).pack(side="left")

        rnd_bar = tk.Frame(win)
        rnd_bar.pack(pady=(1, 0))
        _is_pub = bool(self.mode_publish.get())
        Checkbutton(rnd_bar, text="随机化加密（同一个值每次密文都不同，抗频率分析；代价：同一对手方不再聚成一组）",
                    variable=self.randomize_mode, font=("微软雅黑", 9),
                    state=("disabled" if _is_pub else "normal")).pack(side="left")
        if _is_pub:
            tk.Label(rnd_bar, text="（发布模式用不上：假名本来就是单向的）", font=("微软雅黑", 8),
                     fg="#888888").pack(side="left", padx=6)

        hdr_bar = tk.Frame(win)
        hdr_bar.pack(pady=(1, 0))
        Checkbutton(hdr_bar, text="表头标题行也加密（勾上后整张表连列名一起变密文；还原时会一起解回来）",
                    variable=self.mask_header, font=("微软雅黑", 9)).pack(side="left")

        rnd_row = tk.Frame(win)
        rnd_row.pack(pady=(1, 0))
        tk.Label(rnd_row, text="加密轮数", font=("微软雅黑", 9)).pack(side="left")
        tk.Spinbox(rnd_row, from_=1, to=999, width=6, textvariable=self.fpe_rounds,
                   font=("微软雅黑", 9),
                   state=("disabled" if _is_pub else "normal")).pack(side="left", padx=4)
        tk.Label(rnd_row, text="轮（默认 12；爱填几轮填几轮，越高越慢越难反推，一般 12~40 够用）",
                 font=("微软雅黑", 8), fg="#888888").pack(side="left")

        btns = tk.Frame(win)
        btns.pack(pady=8)
        tk.Button(btns, text="确定", width=10, command=_ok).pack(side=tk.LEFT, padx=6)
        tk.Button(btns, text="取消", width=10, command=win.destroy).pack(side=tk.LEFT, padx=6)

        row_var.set(hdr_state.get(sheets[0]) or self._guess_row(path, sheets[0]))
        build()
        win.grab_set()
        win.lift()
        win.focus_force()
        self.root.wait_window(win)
        if not result["ok"]:
            return None

        # 合并各工作表勾选结果
        sel_state[current["sheet"]] = {t: {k: v.get() for k, v in per.items()}
                                       for t, per in current["state"].items()}
        out = {k: [] for k in ("amount", "cp", "id", "des")}
        for sheet, per_title in sel_state.items():
            for title, flags in per_title.items():
                for key, on in flags.items():
                    if on and title not in out[key]:
                        out[key].append(title)
        out["header_rows"] = dict(hdr_state)
        # 每个工作表各自一套；勾了复用的，签名相同的表并入同一套
        by_sheet = {}
        for sheet, per_title in sel_state.items():
            by_sheet[sheet] = [title for title, flags in per_title.items()
                               if any(flags.get(k) for k in ("amount", "cp", "id", "des")) and
                               flags.get("des")]
        if share_var.get():
            # 复用 = 把勾选到的那几列用到所有工作表上；
            # 某表里没有同名列表头时自然不会命中，不需要靠"表头签名完全一致"来拦。
            union = []
            for titles in by_sheet.values():
                for x in titles:
                    if x not in union:
                        union.append(x)
            for sheet in sheets:
                if sheet == "（无）":
                    continue
                by_sheet[sheet] = list(union)

        out["des_by_sheet"] = by_sheet
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
                picked = self._pick_columns(rep, need, cur, title_prefix=prefix)
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
                hr_map = picked.get("header_rows") or {}
                by_sheet = picked.get("des_by_sheet") or {}
                for fp in files:
                    overrides[_norm_key(fp)] = (list(amt), list(cp), list(cid), list(des),
                                                dict(hr_map), {k: list(v) for k, v in by_sheet.items()})
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
        try:
            set_desens_publish(self.mode_publish.get())
            set_desens_per_file_salt(self.per_file_salt.get())
            set_desens_randomize(self.randomize_mode.get())
            set_desens_rounds(self.fpe_rounds.get())
            set_desens_include_header(self.mask_header.get())
        except Exception:
            pass
        if self.enable_desensitize.get() and not self.mode_publish.get() and not DESENS_PASSWORD:
            _got = _ask_password("set")
            if not _got:
                messagebox.showinfo("已取消", "脱敏需要口令，本次不处理。")
                return
            set_desens_password(_got["password"])
            set_desens_hint(_got.get("hint", ""))
            log("本次脱敏口令已设置（不落盘、不写日志）")
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
            # V1.8.1：这条路径以前 sys.exit(0) 不写哨兵，导致每次"命令行导源码"
            # 都被下一次启动判成"上次异常退出"，白搬一遍历史
            try:
                logger.info(_EXIT_SENTINEL)
                logger.info("程序正常退出 %s（命令行导出源码）", VERSION)
                _flush_all()
            except Exception:
                pass
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