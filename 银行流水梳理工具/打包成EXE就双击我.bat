@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ========================================
echo   银行流水梳理工具 - 一键打包脚本
echo ========================================
echo.

:: 清理旧临时文件
if exist "_extract_ver.py" del "_extract_ver.py"
if exist "_ver.txt" del "_ver.txt"
if exist "_py_check.txt" del "_py_check.txt"

:: ---------- 查找真实 Python ----------
set "PYTHON_CMD="

:: 1. 优先便携版 python
if exist ".\python\python.exe" (
    ".\python\python.exe" -c "import sys; print(sys.executable)" > _py_check.txt 2>&1
    if !errorlevel! equ 0 (
        findstr /i /c:"python" _py_check.txt >nul
        if !errorlevel! equ 0 (
            findstr /i /c:"not found" _py_check.txt >nul
            if !errorlevel! neq 0 (
                findstr /i /c:"microsoft store" _py_check.txt >nul
                if !errorlevel! neq 0 (
                    set "PYTHON_CMD=.\python\python.exe"
                    goto :python_found
                )
            )
        )
    )
    echo [警告] 便携 python 文件夹无效，尝试系统 Python...
    del _py_check.txt
)

:: 2. 系统 python（严格排伪）
where python >nul 2>&1
if %errorlevel% equ 0 (
    python -c "import sys; print(sys.executable)" > _py_check.txt 2>&1
    if !errorlevel! equ 0 (
        findstr /i /c:"python" _py_check.txt >nul
        if !errorlevel! equ 0 (
            findstr /i /c:"not found" _py_check.txt >nul
            if !errorlevel! neq 0 (
                findstr /i /c:"microsoft store" _py_check.txt >nul
                if !errorlevel! neq 0 (
                    set "PYTHON_CMD=python"
                    goto :python_found
                )
            )
        )
    )
    echo [注意] 检测到 python 命令无效（可能为 Microsoft Store 伪装），已忽略。
    del _py_check.txt
)

:: 3. 未找到真实 Python → 询问操作
echo.
echo ========================================
echo   未找到 Python！请在根目录放置 便携版python 文件夹，或手动安装 Python 及 PyInstaller、pandas、openpyxl、xlrd 并加入 PATH，或按 Y 自动安装全部（全程免配置）。
echo ========================================
echo.
choice /c YN /n /m "是否自动下载并安装 Python 及所需依赖（Y=自动安装, N=退出）: "
if errorlevel 2 goto :no_python
if errorlevel 1 goto :install_python

:no_python
echo 您已取消安装，按任意键退出...
pause >nul
exit /b 1

:install_python
echo.
echo 正在下载 Python 安装包，请稍候...
powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.4/python-3.12.4-amd64.exe' -OutFile '%TEMP%\python_installer.exe'"
if not exist "%TEMP%\python_installer.exe" (
    echo [错误] 下载失败，请检查网络或手动安装。
    pause >nul
    exit /b 1
)
echo 正在静默安装 Python（请勿关闭窗口）...
"%TEMP%\python_installer.exe" /quiet InstallAllUsers=0 PrependPath=1
if %errorlevel% neq 0 (
    echo [错误] 安装失败，请尝试手动安装。
    del /q "%TEMP%\python_installer.exe"
    pause >nul
    exit /b 1
)
del /q "%TEMP%\python_installer.exe"

:: 定位安装后的 python
set "PYTHON_CMD=%LocalAppData%\Programs\Python\Python312\python.exe"
if not exist "!PYTHON_CMD!" set "PYTHON_CMD=%ProgramFiles%\Python312\python.exe"
if not exist "!PYTHON_CMD!" (
    set "PATH=%LocalAppData%\Programs\Python\Python312;%LocalAppData%\Programs\Python\Python312\Scripts;%PATH%"
    where python >nul 2>&1
    if %errorlevel% equ 0 (
        set "PYTHON_CMD=python"
    ) else (
        echo [错误] 安装后未找到 Python，请重启电脑后重试。
        pause >nul
        exit /b 1
    )
)
echo Python 安装成功！并已自动加入 PATH。
echo.

:python_found
if exist "_py_check.txt" del "_py_check.txt"
echo 使用 Python: %PYTHON_CMD%

:: ---------- 自动安装所需依赖 ----------
echo 正在检查打包所需依赖（PyInstaller、pandas、openpyxl、xlrd）...
%PYTHON_CMD% -m pip show pyinstaller >nul 2>&1
if !errorlevel! neq 0 (
    echo Installing PyInstaller...
    %PYTHON_CMD% -m pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple
)
%PYTHON_CMD% -m pip show pandas >nul 2>&1
if !errorlevel! neq 0 (
    echo Installing pandas...
    %PYTHON_CMD% -m pip install pandas -i https://pypi.tuna.tsinghua.edu.cn/simple
)
%PYTHON_CMD% -m pip show openpyxl >nul 2>&1
if !errorlevel! neq 0 (
    echo Installing openpyxl...
    %PYTHON_CMD% -m pip install openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple
)
%PYTHON_CMD% -m pip show xlrd >nul 2>&1
if !errorlevel! neq 0 (
    echo Installing xlrd...
    %PYTHON_CMD% -m pip install xlrd -i https://pypi.tuna.tsinghua.edu.cn/simple
)
echo 依赖检查完毕。
echo.

:: ---------- 自动查找主 .py 文件 ----------
set "MAIN_PY="
for /f "delims=" %%f in ('dir /b /a-d *.py 2^>nul') do (
    if /i not "%%f"=="_extract_ver.py" (
        if "!MAIN_PY!"=="" set "MAIN_PY=%%f"
    )
)
if "%MAIN_PY%"=="" (
    echo [错误] 当前目录未找到任何 .py 源文件。
    pause >nul
    exit /b 1
)
echo 主脚本: %MAIN_PY%

:: ---------- 自动查找图标 ----------
set "ICON_OPTION="
set "ICON_FILE="
for /f "delims=" %%i in ('dir /b /a-d *.ico 2^>nul') do (
    if "!ICON_FILE!"=="" set "ICON_FILE=%%i"
)
if not "%ICON_FILE%"=="" (
    set "ICON_OPTION=--icon="%ICON_FILE%""
    echo 图标文件: %ICON_FILE%
) else (
    echo 未找到图标文件，将不添加图标。
)

:: ---------- 版本号提取 ----------
echo import re, sys > "_extract_ver.py"
echo with open(sys.argv[1], 'r', encoding='utf-8') as f: >> "_extract_ver.py"
echo     content = f.read() >> "_extract_ver.py"
echo m = re.search(r'VERSION\s*=\s*"(.*?)"', content) >> "_extract_ver.py"
echo if m: >> "_extract_ver.py"
echo     print(m.group(1)) >> "_extract_ver.py"

%PYTHON_CMD% "_extract_ver.py" "%MAIN_PY%" > _ver.txt 2>&1
set "ver="
set /p ver=<_ver.txt
del _ver.txt

echo "%ver%" | findstr /r "V[0-9]\.[0-9]\.[0-9]" >nul
if %errorlevel% neq 0 (
    echo [错误] 未能提取有效版本号。提取内容: %ver%
    pause >nul
    exit /b 1
)
echo 版本号: %ver%
echo.

set "OUTPUT_NAME=银行流水梳理工具_%ver%"
echo 开始打包，请稍候...
echo.

:: ---------- 执行 PyInstaller ----------
%PYTHON_CMD% -m PyInstaller --onefile --windowed --name "%OUTPUT_NAME%" %ICON_OPTION% --hidden-import pandas --hidden-import openpyxl --hidden-import xlrd "%MAIN_PY%"
if %errorlevel% neq 0 (
    echo 打包失败，请检查上方日志。按任意键退出...
    pause >nul
    exit /b 1
)

:: ---------- 清理 ----------
if exist "_extract_ver.py" del "_extract_ver.py"
if exist "%OUTPUT_NAME%.spec" del "%OUTPUT_NAME%.spec"
if exist "dist\%OUTPUT_NAME%.exe" (
    move /y "dist\%OUTPUT_NAME%.exe" "%~dp0"
    echo 已将 exe 移动到当前目录。
) else (
    echo [警告] 未在 dist 目录找到 exe。
)
if exist "build" rd /s /q "build"
if exist "dist" rd /s /q "dist"

echo.
echo ========================================
echo   打包成功！
echo   输出文件: %~dp0%OUTPUT_NAME%.exe
echo ========================================
pause >nul
exit /b 0