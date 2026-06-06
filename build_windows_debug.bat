@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

REM ============================================================
REM  报价合同生成工具 — Windows 打包 (Debug 版)
REM  每一步都有日志，失败会停住并显示原因
REM ============================================================
set LOG=%~dp0build_log.txt
echo ===== 打包日志 %DATE% %TIME% ===== > "%LOG%"

call :log "========================================"
call :log "  报价合同生成工具 - Windows 打包 (Debug)"
call :log "========================================"
call :log ""

REM ----------------------------------------------------------
REM [1/9] 检查 Python
REM ----------------------------------------------------------
call :section "1/9" "检查 Python"
py --version >> "%LOG%" 2>&1
if errorlevel 1 (
    call :error "Python 未安装或 py 启动器不可用"
    call :error "请安装 Python 3.8+ 并确保添加到 PATH"
    pause
    exit /b 1
)
call :log "Python 检查通过"

REM ----------------------------------------------------------
REM [2/9] 检查 pip
REM ----------------------------------------------------------
call :section "2/9" "检查 pip"
py -m pip --version >> "%LOG%" 2>&1
if errorlevel 1 (
    call :error "pip 不可用"
    pause
    exit /b 1
)
call :log "pip 检查通过"

REM ----------------------------------------------------------
REM [3/9] 检查 requirements.txt
REM ----------------------------------------------------------
call :section "3/9" "检查 requirements.txt"
if not exist "%~dp0requirements.txt" (
    call :error "requirements.txt 不存在"
    pause
    exit /b 1
)
call :log "requirements.txt 存在"

REM ----------------------------------------------------------
REM [4/9] 安装 Python 依赖（去掉 --quiet，显示详细输出）
REM ----------------------------------------------------------
call :section "4/9" "安装 Python 依赖"
call :log "执行: py -m pip install -r requirements.txt"
py -m pip install -r requirements.txt >> "%LOG%" 2>&1
if errorlevel 1 (
    call :error "依赖安装失败，请查看 %LOG%"
    pause
    exit /b 1
)
call :log "依赖安装成功"

REM ----------------------------------------------------------
REM [5/9] 安装 PyInstaller
REM ----------------------------------------------------------
call :section "5/9" "安装 PyInstaller"
call :log "执行: py -m pip install pyinstaller"
py -m pip install pyinstaller >> "%LOG%" 2>&1
if errorlevel 1 (
    call :error "PyInstaller 安装失败，请查看 %LOG%"
    pause
    exit /b 1
)
call :log "PyInstaller 安装成功"

REM ----------------------------------------------------------
REM [6/9] 检查 prepare_public_data.py 是否存在
REM ----------------------------------------------------------
call :section "6/9" "检查脱敏数据脚本"
if not exist "%~dp0tools\prepare_public_data.py" (
    call :error "tools\prepare_public_data.py 不存在"
    pause
    exit /b 1
)
call :log "prepare_public_data.py 存在"

REM ----------------------------------------------------------
REM [7/9] 生成脱敏公共数据包
REM ----------------------------------------------------------
call :section "7/9" "生成脱敏公共数据包（不含甲方私有数据）"
call :log "执行: py tools\prepare_public_data.py --clean"
py "%~dp0tools\prepare_public_data.py" --clean >> "%LOG%" 2>&1
if errorlevel 1 (
    call :error "脱敏数据生成失败，请查看 %LOG%"
    pause
    exit /b 1
)
call :log "脱敏数据生成成功"

REM 检查脱敏数据包是否生成
if not exist "%~dp0build_share_payload\data" (
    call :error "build_share_payload\data 未生成"
    pause
    exit /b 1
)
call :log "build_share_payload\data 已生成"

REM 检查是否包含私有数据（双重保险）
set PRIVATE_FOUND=0
for %%f in (contract_private_contacts.json contract_private_customers.json buyer_contacts.json) do (
    if exist "%~dp0build_share_payload\data\public\%%f" (
        call :error "私有数据 %%f 出现在脱敏包中！"
        set PRIVATE_FOUND=1
    )
)
if "!PRIVATE_FOUND!"=="1" (
    call :error "打包包中包含私有客户数据，终止打包"
    pause
    exit /b 1
)
call :log "私有数据检查通过（脱敏包中不包含甲方联系人）"

REM ----------------------------------------------------------
REM [8/9] 检查主程序入口
REM ----------------------------------------------------------
call :section "8/9" "检查主程序入口"
if not exist "%~dp0main.py" (
    call :error "main.py 不存在，请确认入口文件正确"
    pause
    exit /b 1
)
call :log "main.py 存在"

REM 检查 gui/app.py 是否存在（主程序依赖）
if not exist "%~dp0gui\app.py" (
    call :error "gui\app.py 不存在"
    pause
    exit /b 1
)
call :log "gui\app.py 存在"

REM 检查模板文件是否存在
if not exist "%~dp0templates\三方报价单模板-标准.docx" (
    call :error "templates\三方报价单模板-标准.docx 不存在"
    pause
    exit /b 1
)
call :log "模板文件检查通过"

REM ----------------------------------------------------------
REM [9/9] PyInstaller 打包（onedir 模式，优先保证能运行）
REM ----------------------------------------------------------
call :section "9/9" "PyInstaller 打包（onedir 模式）"
call :log "清理旧构建目录..."
if exist "%~dp0build" rmdir /s /q "%~dp0build" >> "%LOG%" 2>&1
if exist "%~dp0dist" rmdir /s /q "%~dp0dist" >> "%LOG%" 2>&1

call :log "执行 PyInstaller..."
call :log "模式: onedir（目录模式，更稳定）"
call :log "入口: main.py"

py -m PyInstaller ^
  --clean ^
  --noconsole ^
  --onedir ^
  --name "报价合同生成工具" ^
  --add-data "%~dp0build_share_payload\data;data" ^
  --add-data "%~dp0templates;templates" ^
  --exclude-module matplotlib ^
  --exclude-module scipy ^
  --log-level WARN ^
  "%~dp0main.py" >> "%LOG%" 2>&1

if errorlevel 1 (
    call :error "PyInstaller 打包失败"
    call :error "请查看 %LOG% 获取详细错误信息"
    pause
    exit /b 1
)

call :log "PyInstaller 打包成功"

REM ----------------------------------------------------------
REM 打包后检查
REM ----------------------------------------------------------
call :section "打包后检查" ""
if not exist "%~dp0dist\报价合同生成工具\报价合同生成工具.exe" (
    call :error "dist\报价合同生成工具\报价合同生成工具.exe 未生成"
    pause
    exit /b 1
)
call :log "exe 文件已生成: dist\报价合同生成工具\报价合同生成工具.exe"

REM 检查 dist 目录中是否误包含私有数据
set DIST_PRIVATE=0
for %%f in (contract_private_contacts.json contract_private_customers.json buyer_contacts.json) do (
    if exist "%~dp0dist\报价合同生成工具\data\public\%%f" (
        call :error "私有数据 %%f 出现在 dist 目录中！"
        set DIST_PRIVATE=1
    )
)
if "!DIST_PRIVATE!"=="1" (
    call :error "打包结果中包含私有客户数据，请检查 prepare_public_data.py"
    pause
    exit /b 1
)
call :log "dist 目录私有数据检查通过"

REM ----------------------------------------------------------
REM 成功
REM ----------------------------------------------------------
call :log ""
call :log "========================================"
call :log "  打包成功！"
call :log "========================================"
call :log ""
call :log "可执行文件:"
call :log "  %~dp0dist\报价合同生成工具\报价合同生成工具.exe"
call :log ""
call :log "日志文件:"
call :log "  %LOG%"
call :log ""

echo.
echo ========================================
echo   打包成功！
echo ========================================
echo.
echo 可执行文件:
echo   %~dp0dist\报价合同生成工具\报价合同生成工具.exe
echo.
echo 日志已保存到:
echo   %~dp0build_log.txt
echo.
echo 注意：此版本为 onedir 模式（目录模式）
echo       在没有 Python 的电脑上，将整个
echo       dist\报价合同生成工具\ 目录复制过去即可运行。
echo.
pause
endlocal
exit /b 0


REM ============================================================
REM  子程序
REM ============================================================

:log
echo %~1 >> "%LOG%"
echo %~1
goto :eof

:error
echo.
echo [ERROR] %~1
echo [ERROR] %~1 >> "%LOG%"
goto :eof

:section
echo.
echo [%~1] %~2
echo. >> "%LOG%"
echo [%~1] %~2 >> "%LOG%"
goto :eof
