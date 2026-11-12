# -*- coding: utf-8 -*-
"""真实驱动选列窗口：改行号→重新读取→看列表是否跟着变→确定→检查返回的表头行。"""
import importlib.util
import sys
import tkinter as tk

SRC = r"C:\Users\kingo\Documents\deepseek-harness\default-workspace\bankflow-work\source.py"
XLSX = r"C:\Users\kingo\Documents\deepseek-harness\default-workspace\run13\通用表.xlsx"

spec = importlib.util.spec_from_file_location("t30", SRC)
m = importlib.util.module_from_spec(spec)
sys.modules["t30"] = m
spec.loader.exec_module(m)

root = tk.Tk()
root.withdraw()
app = m.App(root)


def walk(w):
    yield w
    for c in w.winfo_children():
        yield from walk(c)


def find(w, cls=None, text=None):
    for x in walk(w):
        if cls and x.winfo_class() != cls:
            continue
        if text is not None:
            try:
                if str(x.cget("text")) != text:
                    continue
            except Exception:
                continue
        return x
    return None


def only_toplevel():
    for w in root.winfo_children():
        if isinstance(w, tk.Toplevel):
            return w
    return None


state = {"labels_1": [], "labels_2": []}


def step1():
    win = only_toplevel()
    if not win:
        print("!! 没找到选列窗口")
        return
    spin = find(win, "Spinbox")
    print("窗口内初始行号:", spin.get())
    # 收集列表里的标签文本
    state["labels_1"] = [w.cget("text") for w in walk(win)
                         if w.winfo_class() == "Label" and str(w.cget("text")).strip()]
    spin.delete(0, "end")
    spin.insert(0, "2")
    btn = find(win, "Button", "重新读取")
    btn.invoke()
    root.after(300, step2)


def step2():
    win = only_toplevel()
    spin = find(win, "Spinbox")
    print("点过重新读取后的行号:", spin.get())
    labels = [w.cget("text") for w in walk(win)
              if w.winfo_class() == "Label" and str(w.cget("text")).strip()]
    state["labels_2"] = labels
    has_name = any("姓名" in t for t in labels)
    has_title = any("年度员工" in t for t in labels)
    print("列表里出现「姓名」:", has_name)
    print("列表里还残留标题行「年度员工…」:", has_title)
    tk.Button  # noqa
    ok = find(win, "Button", "确定")
    ok.invoke()


root.after(700, step1)
out = app._pick_columns(XLSX, ["des"], {"amount": [], "cp": [], "id": [], "des": []},
                        header_rows={"员工名单": 1}, sheets=["员工名单"])
print("返回:", None if out is None else {k: out[k] for k in ("des", "header_rows")})
print("初始行号列表片段:", state["labels_1"][:6])
print("改行号后列表片段:", state["labels_2"][:6])
root.destroy()
