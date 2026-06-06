@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo  报价合同生成工具 - Windows 打包
echo ========================================
echo.

echo [1/6] 清理旧构建目录...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist build_share_payload rmdir /s /q build_share_payload

echo [2/6] 安装依赖...
py -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] 依赖安装失败
    pause
    exit /b 1
)

py -m pip install pyinstaller
if errorlevel 1 (
    echo [ERROR] PyInstaller 安装失败
    pause
    exit /b 1
)

echo [3/6] 生成脱敏公共数据包（不含甲方私有数据）...
py tools\prepare_public_data.py --clean
if errorlevel 1 (
    echo [ERROR] 脱敏数据生成失败
    pause
    exit /b 1
)

echo [4/6] 检查脱敏包不包含私有数据...
set PRIVATE=0
for %%f in (contract_private_contacts.json contract_private_customers.json buyer_contacts.json) do (
    if exist "build_share_payload\data\public\%%f" (
        echo [ERROR] 私有数据 %%f 出现在脱敏包中
        set PRIVATE=1
    )
)
if "!PRIVATE!"=="1" (
    pause
    exit /b 1
)

echo [5/6] PyInstaller 打包（onedir 模式）...
py -m PyInstaller ^
  --clean ^
  --noconsole ^
  --onedir ^
  --name "报价合同生成工具" ^
  --add-data "build_share_payload\data;data" ^
  --add-data "templates;templates" ^
  --exclude-module matplotlib ^
  --exclude-module scipy ^
  main.py
if errorlevel 1 (
    echo [ERROR] PyInstaller 打包失败
    pause
    exit /b 1
)

echo [6/6] 检查输出...
if not exist "dist\报价合同生成工具\报价合同生成工具.exe" (
    echo [ERROR] exe 未生成
    pause
    exit /b 1
)

echo.
echo 打包完成: dist\报价合同生成工具\报价合同生成工具.exe
echo （将整个 dist\报价合同生成工具\ 目录分发给用户即可）
echo.
pause
endlocal
