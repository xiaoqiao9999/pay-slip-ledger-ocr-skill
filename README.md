# pay-slip-ledger-ocr（付款回单 OCR 台账）

将财务发来的付款/转账回单图片批量 OCR，提取字段生成 Excel 台账，并按收款人（供应商）
分类归档原图。增量追加、按会计流水号幂等去重、可反复运行。

## 目录结构

```
pay-slip-ledger-ocr-skill/
├── pay-slip-ledger-ocr/        ← skill 完整源码
│   ├── SKILL.md                ← WorkBuddy 技能入口（触发条件 + 执行流程）
│   ├── scripts/
│   │   └── pay_slip_ledger_scan.py    ← 主脚本（OCR + 台账 + 归档，全逻辑）
│   ├── references/
│   │   ├── ledger-spec.md      ← 台账 12 列 / 必须项 / 归档命名 / OCR 变体
│   │   └── ocr-engine.md       ← OCR 引擎定位与跨机部署
│   └── assets/
│       └── 付款回单台账_使用说明.html   ← 图文使用说明（可打印）
└── pay-slip-ledger-ocr.zip     ← 打包产物（可直接安装的 skill）
```

## 安装（WorkBuddy 本机）

```bash
# 方式一：解压 zip 到用户技能目录
unzip pay-slip-ledger-ocr.zip -d ~/.workbuddy/skills/
# 方式二：直接把 pay-slip-ledger-ocr/ 拷入 ~/.workbuddy/skills/
```

## 运行

工作区结构约定：脚本所在目录即工作区，内含 `输入/`（新回单投放）与
`已完成/`（按收款人归档），台账 `付款回单台账.xlsx` 自动维护在根目录。

```bash
python pay_slip_ledger_scan.py                 # 正式运行（识别→入账→归档）
python pay_slip_ledger_scan.py --dry-run       # 试跑预览，不写盘不移动
python pay_slip_ledger_scan.py --limit 3       # 只处理前 3 张（调试）
```

- 依赖：系统 Python ≥ 3.9 + `openpyxl`；OCR 引擎 local-ocr-npu（见 references/ocr-engine.md）
- 单张约 20 秒，37 张约 12-13 分钟
- 必须字段：收款人名称 / 金额 / 摘要（缺失才拦截）；记账日期、流水号等可选
- 归档命名：`收款人_金额两位小数_摘要[_记账日期].png` → `已完成/<收款人>/`，重名加 `_2`

## 发布到 GitHub（首次，私有仓库示例）

```bash
cd pay-slip-ledger-ocr-skill
git config user.name  "你的GitHub用户名"
git config user.email "你的GitHub邮箱"
git branch -M main
git add .
git commit -m "feat: 付款回单 OCR 台账 skill"
git remote add origin https://github.com/<你的用户名>/pay-slip-ledger-ocr-skill.git
git push -u origin main
```

> 先在 github.com 新建空仓库（建议 **Private**，勾选不添加 README/.gitignore），
> 再把上面 remote 地址换成你自己的仓库地址。

## License

MIT
