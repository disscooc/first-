#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

echo "========================================"
echo " 报价合同生成工具 - macOS 打包"
echo "========================================"
echo

echo "[1/5] 清理旧构建目录..."
rm -rf build dist build_share_payload

echo "[2/5] 安装依赖..."
python3 -m pip install -r requirements.txt --quiet
python3 -m pip install pyinstaller --quiet

echo "[3/5] 生成脱敏公共数据包..."
python3 tools/prepare_public_data.py --clean

echo "[4/5] PyInstaller 打包..."
python3 -m PyInstaller \
  --clean \
  --windowed \
  --name "报价合同生成工具" \
  --add-data "build_share_payload/data:data" \
  --exclude-module matplotlib \
  --exclude-module scipy \
  main.py

echo "[5/5] 检查分享包不包含私有数据..."
if find dist -name "contract_private_contacts.json" -o -name "contract_private_customers.json" -o -name "buyer_contacts.json" | grep -q .; then
  echo "错误：dist 中发现私有数据文件。"
  exit 1
fi

echo
echo "打包完成: dist/报价合同生成工具.app"
