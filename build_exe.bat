@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

set "PYTHON_CMD="
set "PYTHON_ARGS="

where py >nul 2>&1
if not errorlevel 1 (
    py -3 --version >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON_CMD=py"
        set "PYTHON_ARGS=-3"
        goto :found_python
    )
)

for %%P in (python python3) do (
    where %%P >nul 2>&1
    if not errorlevel 1 (
        echo %%P | findstr /I "WindowsApps" >nul
        if errorlevel 1 (
            %%P --version >nul 2>&1
            if not errorlevel 1 (
                set "PYTHON_CMD=%%P"
                set "PYTHON_ARGS="
                goto :found_python
            )
        )
    )
)

echo.
echo [ERROR] 未检测到可用的 Python。
echo.
echo 请先安装 Python 3.10 或更高版本，然后重新运行本脚本。
echo.
echo 推荐安装方式:
echo   1. 官网: https://www.python.org/downloads/windows/
echo      安装时勾选 Add python.exe to PATH
echo   2. winget install Python.Python.3.12
echo.
echo 如果已安装但仍提示找不到 Python，请到:
echo   设置 -^> 应用 -^> 高级应用设置 -^> 应用执行别名
echo   关闭 python.exe 和 python3.exe 的 Microsoft Store 别名。
echo.
exit /b 1

:found_python
if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    %PYTHON_CMD% %PYTHON_ARGS% -m venv .venv
    if errorlevel 1 (
        echo [ERROR] 创建虚拟环境失败。
        exit /b 1
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] 虚拟环境未创建成功。
    exit /b 1
)

call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo [ERROR] 无法激活虚拟环境。
    exit /b 1
)

python -m pip install --upgrade pip
if errorlevel 1 exit /b 1

pip install -r requirements.txt
if errorlevel 1 exit /b 1

if not exist "assets\mailbatch.ico" (
    echo Generating application icon...
    python scripts\generate_icon.py
    if errorlevel 1 exit /b 1
)

set "ICON_ARGS="
set "DATA_ARGS="
if exist "assets\mailbatch.ico" (
    set ICON_ARGS=--icon assets\mailbatch.ico
    set DATA_ARGS=--add-data assets\mailbatch.ico;assets
)

set "APP_HTML_ARGS="
if exist "app.html" (
    set APP_HTML_ARGS=--add-data app.html;.
) else (
    echo [ERROR] app.html not found.
    exit /b 1
)

python -m PyInstaller --noconfirm --windowed --onefile --name MailBatch ^
    --hidden-import webview ^
    --hidden-import webview.platforms.winforms ^
    %ICON_ARGS% %DATA_ARGS% %APP_HTML_ARGS% app.py
if errorlevel 1 (
    echo [ERROR] PyInstaller 打包失败。
    exit /b 1
)

if not exist "dist\MailBatch.exe" (
    echo [ERROR] 打包结束，但未找到 dist\MailBatch.exe
    exit /b 1
)

echo.
echo Build complete: dist\MailBatch.exe
endlocal
