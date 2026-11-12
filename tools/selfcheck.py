# -*- coding: utf-8 -*-
"""打包前自检：静态扫两类「编译能过、一跑就炸」的问题。

用法：
    python tools/selfcheck.py "..\\银行流水梳理工具\\源代码 以记事本打开修改.py"

检查项：
1. 类里 self.xxx(...) 调用了但没定义（实例属性不算）。
2. tkinter 的类名被裸用（如 IntVar(...)）但没 from tkinter import 过。

为什么需要：
- V1.7.1 重写选列窗口时把 _split_var 删了，编译不报错、启动也不报错，
  只有点「选择文件」时才抛 AttributeError。
- V1.9.1 加加密轮数输入框时写了裸 IntVar(...)，同样编译能过，
  一启动就 NameError: name 'IntVar' is not defined。
这两类都靠人肉点界面才发现，用这个脚本能在打包前拦住。
"""
import ast
import io
import sys

REQUIRED_METHODS = {
    "_split_var",
    "_pick_columns",
    "_collect_header_titles",
    "select_files",
    "start_processing",
    # V1.9.x 起脱敏相关
    "_sync_exclusive",
    "_on_desensitize_toggle",
    "_on_manual_toggle",
}

# 这些名字必须写了 tk. 前缀，或者出现在 from tkinter import 里
TK_BARE_NAMES = {
    "BooleanVar", "IntVar", "StringVar", "DoubleVar",
    "Entry", "Label", "Frame", "Checkbutton", "Button", "Toplevel",
    "Canvas", "Spinbox", "Listbox", "Text", "Menu", "PhotoImage",
}


def check(path):
    src = io.open(path, encoding="utf-8").read()
    tree = ast.parse(src)
    problems = []

    # --- 1) self.xxx 调用了但没定义
    for node in ast.walk(tree):
        if not isinstance(node, ast.ClassDef):
            continue
        if node.bases:
            # 有基类的（如 logging.Handler 子类）会用到继承来的方法，跳过
            continue
        defined = {n.name for n in node.body
                   if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))}
        assigned = set()
        for sub in ast.walk(node):
            if isinstance(sub, ast.Attribute) and isinstance(sub.ctx, ast.Store):
                assigned.add(sub.attr)
        called = set()
        for sub in ast.walk(node):
            if isinstance(sub, ast.Call) and isinstance(sub.func, ast.Attribute):
                if isinstance(sub.func.value, ast.Name) and sub.func.value.id == "self":
                    called.add(sub.func.attr)
        missing = sorted(called - defined - assigned)
        if missing:
            problems.append((node.name, missing))
        if node.name == "App":
            lack = sorted(REQUIRED_METHODS - defined)
            if lack:
                problems.append(("App(必需方法)", lack))

    # --- 2) tkinter 类名裸用但没导入
    imported_tk = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and (node.module or "") == "tkinter":
            for a in node.names:
                imported_tk.add(a.asname or a.name)
    used_bare = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Load):
            if node.id in TK_BARE_NAMES:
                used_bare.add(node.id)
    missing_tk = sorted(used_bare - imported_tk)
    if missing_tk:
        problems.append(("tkinter 未导入的名字", missing_tk))

    return problems


def main():
    if len(sys.argv) < 2:
        print("用法: python tools/selfcheck.py <源码.py>")
        return 2
    path = sys.argv[1]
    try:
        problems = check(path)
    except SyntaxError as e:
        print("语法错误：第 %s 行 %s" % (e.lineno, e.msg))
        return 1
    if not problems:
        print("自检通过：没有发现「引用了但没定义」或「tkinter 名字裸用」的问题。")
        return 0
    print("自检发现问题：")
    for cls, items in problems:
        print("  %s 缺少/未导入: %s" % (cls, ", ".join(items)))
    return 1


if __name__ == "__main__":
    sys.exit(main())
