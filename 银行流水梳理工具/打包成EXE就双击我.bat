@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ========================================
echo   银行流水梳理工具 - 一键打包脚本
echo ========================================
echo.

:: 清理旧临时文件
if exist "_extract_ver.py" del "_extract_ver.py"
if exist "_source_backup.py" del /q "_source_backup.py"
if exist "_icon_backup.ico" del /q "_icon_backup.ico"
if exist "_ver.txt" del "_ver.txt"
if exist "_py_check.txt" del "_py_check.txt"

:: ---------- 查找真实 Python ----------
set "PYTHON_CMD="
call :detect_python
if defined PYTHON_CMD goto :python_found

:: 没找到 → 先看目录里有没有现成的离线安装包
set "OFFLINE_INSTALLER="
call :find_offline_installer
if defined OFFLINE_INSTALLER (
    echo.
    echo [提示] 在本目录找到 Python 安装包：
    echo        !OFFLINE_INSTALLER!
    call :verify_installer "!OFFLINE_INSTALLER!"
    if errorlevel 1 (
        echo        这个文件看起来不是完整的安装包，已忽略。
        set "OFFLINE_INSTALLER="
    )
)
if defined OFFLINE_INSTALLER (
    echo        将直接用它安装，不需要联网下载。
    choice /c YN /n /m "是否用它安装 Python? Y=安装, N=退出: "
    if errorlevel 2 goto :no_python
    call :run_installer "!OFFLINE_INSTALLER!"
    set "RC=!errorlevel!"
    if "!RC!"=="0" goto :installed_ok
    if "!RC!"=="3010" goto :installed_ok
    set "INSTALLER_KEPT=!OFFLINE_INSTALLER!"
    goto :install_failed
)

echo.
echo ========================================
echo   未找到可用的 Python
echo ----------------------------------------
echo   三种办法，任选其一：
echo     1^) 把离线安装包 python-3.12.4-amd64.exe 放到本目录，再运行本脚本；
echo     2^) 按 Y 让本脚本自动下载安装（需要这台电脑能上外网）；
echo     3^) 在能上网的电脑上打包，把生成的 exe 拷回来用。
echo ========================================
echo.
choice /c YN /n /m "是否现在自动下载并安装 Python 及所需依赖? Y=自动安装, N=退出: "
if errorlevel 2 goto :no_python

echo.
echo 正在下载 Python 3.12.4 安装包（约 25MB），请稍候...
set "DL=%TEMP%\python_installer.exe"
call :download_installer
if not defined DL_OK goto :download_failed
call :run_installer "%DL%"
set "RC=!errorlevel!"
if "!RC!"=="0" goto :installed_ok
if "!RC!"=="3010" goto :installed_ok
set "INSTALLER_KEPT=%DL%"
goto :install_failed

:installed_ok
echo.
echo Python 安装完成，正在重新检测...
call :detect_python
if defined PYTHON_CMD goto :python_found
echo.
echo [错误] 安装完成但没找到 python.exe。
echo   1^) 重启电脑后再运行本脚本；
echo   2^) 或把 python.exe 的完整路径写进本目录的 python路径.txt（一行）。
pause >nul
exit /b 1

:download_failed
echo.
echo [错误] 下载安装包失败（python.org 和国内镜像都没下下来）。
echo   常见原因：这台电脑在内网、要走代理、或被安全软件拦了。
echo   建议：
echo     A. 在能上网的电脑下载 python-3.12.4-amd64.exe，放到本目录再运行本脚本；
echo     B. 在能上网的电脑上打包，把生成的 exe 拷回来用。
echo   手动装好 Python 后如果脚本还是找不到，把 python.exe 的完整路径写进
echo   本目录的 python路径.txt（一行，例如 C:\Python312\python.exe）。
pause >nul
exit /b 1

:install_failed
echo.
echo [错误] Python 安装失败（返回码 !RC!）。
if defined INSTALLER_KEPT (
    echo   安装包已保留：!INSTALLER_KEPT!
    echo   可以双击它手动安装，装完再运行本脚本。
    echo   如果双击提示“不是有效的 Win32 应用程序”，说明下载到的是网页而不是安装包，
    echo   请换一台能上网的电脑下载后再拷过来。
)
echo   或者在能上网的电脑上打包，把生成的 exe 拷回来用。
pause >nul
exit /b 1

:no_python
echo 您已取消安装，按任意键退出...
pause >nul
exit /b 1

:python_found
if exist "_py_check.txt" del "_py_check.txt"
echo 使用 Python: %PYTHON_CMD%

:: ---------- 自动安装所需依赖 ----------
echo 正在检查打包所需依赖（PyInstaller、pandas、openpyxl、xlrd）...
call :ensure_pkg pyinstaller
if errorlevel 1 goto :deps_failed
call :ensure_pkg pandas
if errorlevel 1 goto :deps_failed
call :ensure_pkg openpyxl
if errorlevel 1 goto :deps_failed
call :ensure_pkg xlrd
if errorlevel 1 goto :deps_failed
echo 依赖检查完毕。
echo.
goto :deps_done

:deps_failed
echo.
echo [错误] 依赖没装全，无法打包（多半是连不上 pypi 镜像站）。
echo   建议：
echo     A. 在能上网的电脑上打包，把生成的 exe 拷回来用；
echo     B. 或先手动执行：pip install pyinstaller pandas openpyxl xlrd
pause >nul
exit /b 1

:deps_done

:: ---------- 自动查找主 .py 文件 ----------
set "MAIN_PY="
set "PY_COUNT=0"
:: V1.9.9：目录里同时有旧版本源码，优先打包文件名带 V1.8.0 的那个。
:: 用纯英文通配符匹配，避免中文编码影响。
for %%f in (源代码*.py) do if not defined MAIN_PY set "MAIN_PY=%%f"
if defined MAIN_PY (
    set "PY_COUNT=1"
    goto :py_ready
)
for /f "delims=" %%f in ('dir /b /a-d *.py 2^>nul') do (
    if /i not "%%f"=="_extract_ver.py" if /i not "%%f"=="_source_backup.py" (
        if "!MAIN_PY!"=="" set "MAIN_PY=%%f"
        set /a PY_COUNT+=1
    )
)
if "%MAIN_PY%"=="" (
    echo [错误] 当前目录未找到任何 .py 源文件。
    pause >nul
    exit /b 1
)
:py_ready
echo 主脚本: %MAIN_PY%

:: 目录里有多个 .py 时，dir /b 按文件名排序取第一个，很容易把旧版本打进去；
:: 这里先列出来确认一次，避免“明明改了源码，打出来的 exe 还是旧的”。
if !PY_COUNT! leq 1 goto :skip_multi_py
echo.
echo [注意] 当前目录有 !PY_COUNT! 个 .py 文件：
for /f "delims=" %%f in ('dir /b /a-d *.py 2^>nul') do (
    if /i not "%%f"=="_extract_ver.py" if /i not "%%f"=="_source_backup.py" echo          %%f
)
echo 本脚本只打包上面“主脚本”那一行显示的 .py。
choice /c YN /n /m "确认就用 %MAIN_PY% 打包吗? Y=继续, N=退出: "
if errorlevel 2 exit /b 1
:skip_multi_py

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

"%PYTHON_CMD%" "_extract_ver.py" "%MAIN_PY%" > _ver.txt 2>&1
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

:: ---------- 把源码和图标一起打进 exe（软件里的“导出源码备份”要用） ----------
if exist "_source_backup.py" del /q "_source_backup.py"
if exist "_icon_backup.ico" del /q "_icon_backup.ico"
copy /y "%MAIN_PY%" "_source_backup.py" >nul
set "ADD_DATA_OPT="
if exist "_source_backup.py" set "ADD_DATA_OPT=--add-data "_source_backup.py;.""
if not "%ICON_FILE%"=="" copy /y "%ICON_FILE%" "_icon_backup.ico" >nul
if exist "_icon_backup.ico" set "ADD_DATA_OPT=%ADD_DATA_OPT% --add-data "_icon_backup.ico;.""
if exist "_source_backup.py" (echo 源码将一起打包: %MAIN_PY%) else (echo [警告] 源码复制失败，打出来的 exe 将没有“导出源码备份”功能。)
if exist "_icon_backup.ico" echo 图标将一起打包: %ICON_FILE%

set "OUTPUT_NAME=银行流水梳理工具_%ver%"
echo 开始打包，请稍候...
echo.

:: ---------- 执行 PyInstaller ----------
"%PYTHON_CMD%" -m PyInstaller --onefile --windowed --name "%OUTPUT_NAME%" %ICON_OPTION% %ADD_DATA_OPT% --hidden-import pandas --hidden-import openpyxl --hidden-import xlrd "%MAIN_PY%"
if %errorlevel% neq 0 (
    echo 打包失败，请检查上方日志。按任意键退出...
    pause >nul
    exit /b 1
)

:: ---------- 清理 ----------
if exist "_extract_ver.py" del "_extract_ver.py"
if exist "_source_backup.py" del /q "_source_backup.py"
if exist "_icon_backup.ico" del /q "_icon_backup.ico"
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

:: ============================================================
:: 以下是子过程（主流程执行到上面的 exit /b 0 就结束了，不会走下来）
:: ============================================================

:check_python
:: %~1 = 候选 python；可用则设置 PYTHON_CMD 并返回 0
set "PYTHON_CMD="
set "CAND=%~1"
if not defined CAND exit /b 1
if /i not "%CAND%"=="python" if not exist "%CAND%" exit /b 1
"%CAND%" -c "import sys; print(sys.executable)" > "_py_check.txt" 2>&1
if errorlevel 1 exit /b 1
findstr /i /c:"microsoft store" "_py_check.txt" >nul
if not errorlevel 1 exit /b 1
findstr /i /c:"not found" "_py_check.txt" >nul
if not errorlevel 1 exit /b 1
findstr /i /c:"python" "_py_check.txt" >nul
if errorlevel 1 exit /b 1
set "PYTHON_CMD=%CAND%"
exit /b 0

:detect_python
set "PYTHON_CMD="
if exist "_py_check.txt" del /q "_py_check.txt" >nul 2>&1
if defined PYTHON_EXE call :check_python "%PYTHON_EXE%"
if defined PYTHON_CMD exit /b 0
if exist "python路径.txt" call :read_python_path
if defined PYTHON_CMD exit /b 0
call :check_python ".\python\python.exe"
if defined PYTHON_CMD exit /b 0
call :check_python "python"
if defined PYTHON_CMD exit /b 0
for /d %%D in ("%LocalAppData%\Programs\Python\Python3*") do if not defined PYTHON_CMD call :check_python "%%~D\python.exe"
for /d %%D in ("%ProgramFiles%\Python3*") do if not defined PYTHON_CMD call :check_python "%%~D\python.exe"
for /d %%D in ("C:\Python3*") do if not defined PYTHON_CMD call :check_python "%%~D\python.exe"
if exist "_py_check.txt" del /q "_py_check.txt" >nul 2>&1
exit /b 0

:read_python_path
:: python路径.txt：一行，写 python.exe 的完整路径
set "P="
set /p P=<"python路径.txt"
if not defined P exit /b 0
set "P=!P:"=!"
if not exist "!P!" (
    echo [提示] python路径.txt 里写的路径不存在：!P!
    exit /b 0
)
call :check_python "!P!"
exit /b 0

:find_offline_installer
set "OFFLINE_INSTALLER="
for /f "delims=" %%F in ('dir /b /a-d "python-3*.exe" 2^>nul') do if not defined OFFLINE_INSTALLER set "OFFLINE_INSTALLER=%%~fF"
for /f "delims=" %%F in ('dir /b /a-d "python_installer.exe" 2^>nul') do if not defined OFFLINE_INSTALLER set "OFFLINE_INSTALLER=%%~fF"
for /f "delims=" %%F in ('dir /b /a-d "安装包\python*.exe" 2^>nul') do if not defined OFFLINE_INSTALLER set "OFFLINE_INSTALLER=%%~fF"
exit /b 0

:verify_installer
:: 检查是不是真正的安装包：大小 >= 5MB 且文件头是 MZ
:: （内网/代理常把拦截网页当安装包存下来，直接运行会报“不是有效的 Win32 应用程序”）
set "SZ=0"
for %%A in ("%~1") do set "SZ=%%~zA"
if %SZ% LSS 5000000 (
    echo      [x] 文件只有 %SZ% 字节，不是安装包
    exit /b 1
)
powershell -NoProfile -Command "try { $b = Get-Content -Encoding Byte -TotalCount 2 -LiteralPath '%~1'; if ($b[0] -eq 77 -and $b[1] -eq 90) { 'MZ' } else { 'NO' } } catch { 'NO' }" > "_mz.txt" 2>&1
set "MZ="
set /p MZ=<"_mz.txt"
del /q "_mz.txt" >nul 2>&1
if /i not "%MZ%"=="MZ" (
    echo      [x] 文件头不是 MZ，不是可执行的安装包
    exit /b 1
)
echo      [ok] 安装包校验通过（%SZ% 字节）
exit /b 0

:download_installer
set "DL_OK="
if exist "%DL%" del /q "%DL%" >nul 2>&1
call :try_download "https://www.python.org/ftp/python/3.12.4/python-3.12.4-amd64.exe"
if defined DL_OK exit /b 0
call :try_download "https://registry.npmmirror.com/-/binary/python/3.12.4/python-3.12.4-amd64.exe"
if defined DL_OK exit /b 0
call :try_download "https://mirrors.huaweicloud.com/python/3.12.4/python-3.12.4-amd64.exe"
if defined DL_OK exit /b 0
exit /b 1

:try_download
echo   尝试下载: %~1
where curl >nul 2>&1
if not errorlevel 1 (
    curl -L --fail --silent --show-error --connect-timeout 15 --max-time 900 -o "%DL%" "%~1"
    call :accept_download
    if defined DL_OK exit /b 0
)
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ProgressPreference='SilentlyContinue'; try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}; try { Invoke-WebRequest -UseBasicParsing -Uri '%~1' -OutFile '%DL%' } catch { exit 1 }"
call :accept_download
if defined DL_OK exit /b 0
exit /b 1

:accept_download
if not exist "%DL%" exit /b 1
call :verify_installer "%DL%"
if not errorlevel 1 (
    set "DL_OK=1"
    exit /b 0
)
if exist "%DL%" del /q "%DL%" >nul 2>&1
exit /b 1

:run_installer
echo.
echo 正在静默安装 Python，请稍候（约 1-3 分钟，请不要关闭窗口）...
"%~1" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_launcher=1
exit /b !errorlevel!

:ensure_pkg
"%PYTHON_CMD%" -m pip show "%~1" >nul 2>&1
if !errorlevel! equ 0 (
    echo   [ok] %~1 已安装
    exit /b 0
)
echo   正在安装 %~1 ...
"%PYTHON_CMD%" -m pip install "%~1" -i https://pypi.tuna.tsinghua.edu.cn/simple --disable-pip-version-check
if !errorlevel! equ 0 (
    echo   [ok] %~1 安装完成
    exit /b 0
)
echo   [x] %~1 安装失败
exit /b 1
