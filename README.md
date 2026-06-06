# 超算资源自动报价与合同生成系统 v2

## 部署文档

### 环境要求
- Python 3.10+
- LibreOffice 7.0+ (PDF导出时需要，非强制)
- Windows 10+ / macOS 11+ / Linux

### 快速启动（开发模式）

#### Windows
1. 双击 启动报价工具.bat
2. 或命令行: py main.py

#### macOS / Linux
1. 终端: pip3 install -r requirements.txt
2. 终端: python3 main.py

### 打包为独立exe/app

#### Windows
双击 uild_win.bat，输出在 dist\超算报价工具.exe

#### macOS
终端运行 ash build_mac.sh，输出在 dist/超算报价工具

### 目录结构
`
quote_app_v2/
├── main.py              # 入口
├── main.spec            # PyInstaller spec
├── requirements.txt     # 依赖清单
├── build_win.bat        # Windows打包
├── build_mac.sh         # Mac打包
├── 启动报价工具.bat      # 开发快速启动
├── config.py            # 全局配置
├── models.py            # 数据模型
├── utils.py             # 工具函数
├── engines/             # 业务引擎
│   ├── price_loader.py   # 价格表加载
│   ├── excel_writer.py   # Excel报价单
│   ├── word_writer.py    # Word报价单(标准+云达)
│   ├── contract_writer.py# 合同生成
│   ├── competitor.py     # 三方计算
│   └── pdf_exporter.py   # PDF导出(LibreOffice)
├── gui/                 # 界面
│   ├── app.py            # 主窗口
│   ├── panels/
│   │   ├── header_panel.py    # 基础信息
│   │   ├── resource_panel.py   # 资源库
│   │   ├── quote_panel.py      # 配置+定价
│   │   └── thirdparty_panel.py # 三方设置
│   └── dialogs/
│       └── contract_dialog.py  # 合同弹窗
├── templates/           # 模板文件
│   ├── 石油化工学院机时服务单.xlsx
│   ├── 三方报价单模板-标准.docx
│   └── 三方报价单模板-云达.docx
├── data/                # 数据文件
│   ├── unit_aliases.json
│   ├── unit_history.json
│   ├── contract_customers.json
│   ├── contract_private_contacts.json
│   └── contract_private_customers.json
└── output/              # 自动创建，输出目录
    └── YYYY-MM/         # 按月份归档
`

### 修改配置
编辑 config.py:
- COMPANIES: 报价单位列表
- MANAGERS: 业务经理列表
- PRICE_DIRS: 价格表搜索路径
- OUT_BASE: 输出目录(可自定义)
- TAX: 税率(默认1.06)

### 功能清单
1. 基础信息: 报价单位/日期/经理/客户/项目
2. 资源库: Excel价格表自动加载、搜索、分类
3. 三种定价模式: 单价/数量/总价反算
4. 折扣计算: 输入折扣率→单价 或 输入单价→折扣率
5. 赠送存储: 默认0.5T，可自定义
6. 标准报价单: Excel格式输出
7. 三方报价单: 标准模板(Word) + 云达模板(Word)
8. 两种三方模式: 同资源提价 / 总价不变增数量
9. 合同生成: Word格式，含甲乙方信息
10. PDF导出: 通过LibreOffice无头模式转换

### 故障排除
- "模板不存在": 检查templates目录，确保模板文件已放入
- "价格表为空": 将价格表Excel放入项目根目录
- "LibreOffice not found": 安装LibreOffice或跳过PDF导出
- "权限不足": 确保output目录可写

### 注意
- 原始的PowerShell版本 (quote_app.ps1) 保持不变
- 第一版Python (quote_app/) 保持不变，仅做备份
