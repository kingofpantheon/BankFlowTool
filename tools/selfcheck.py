# -*- coding: utf-8 -*-
"""打包前自检：静态检查源码里「引用了但没定义的方法」。

用法：python tools/selfcheck.py "..\源代码 V1.7.1 增加运行日志.py"

检查项：
1. 类里 self.xxx(...) 调用的方法，是否真的定义了（实例属性不算）。
2. TK 回调里常用的几个方法是否还在（防止重构时把工具方法删掉）。

为什么需要：V1.7.1 第一次发布时，重写选列窗口把 _split_var 删了，
编译不报错、启动也不报错，只有点「选择文件」时才抛 AttributeError。
这个脚本就是拦这类问题。
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
}


def check(path):
    tree = ast.parse(io.open(path, encoding="utf-8").read())
    problems = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.ClassDef):
            continue
        if node.bases:
            # 有基类的（如 logging.Handler 子类）会用到继承来的方法，跳过
            continue
        defined = {n.name for n in node.body if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))}
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
    return problems


def main():
    if len(sys.argv) < 2:
        print("用法: python tools/selfcheck.py <源码.py>")
        return 2
    problems = check(sys.argv[1])
    if not problems:
        print("自检通过：没有发现「引用了但没定义」的方法。")
        return 0
    print("自检发现问题：")
    for cls, items in problems:
        print("  类 %s 缺少: %s" % (cls, ", ".join(items)))
    return 1


if __name__ == "__main__":
    sys.exit(main())
