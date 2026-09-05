# OCR 引擎（本地 NPU OCR）说明

本技能依赖 `local-ocr-npu`：基于 PaddleOCR 的本地离线 OCR，调用方式为一个 PowerShell 脚本。

## 1. 脚本位置与探测顺序

`pay_slip_ledger_scan.py` 启动时按以下顺序定位 `run.ps1`：

1. 环境变量 `PAYSLIP_OCR_SCRIPT`（跨机部署推荐）
2. 本机已知安装路径：
   - `c:\Users\JQZH\.trae-cn\skills\local-ocr-npu\scripts\run.ps1`
   - `c:\Users\JQZH\.workbuddy\skills\local-ocr-npu\scripts\run.ps1`
3. 以上都找不到 → 取默认值并提示（此时必须用 `--ocr-script <path>` 显式指定）

也可直接命令行覆盖：`python pay_slip_ledger_scan.py --ocr-script "C:\path\to\run.ps1"`。

## 2. 调用契约（run.ps1）

```text
powershell -ExecutionPolicy Bypass -File <run.ps1> <图片路径> -Device cpu
```

- 参数：图片路径（**必须纯英文路径/文件名**，ppocr.exe 不支持中文路径）→ 需先复制到英文临时目录再调用
- 设备：`cpu`（NPU 可用时也兼容）
- 输出：stdout，以 **40+ 个 `-` 分隔行**包裹的块，块内每行一条 OCR 识别文本
- 解析：`---` 分隔块之间的非空行即文本行；编码 UTF-8 优先、GBK 兜底
- 单张超时 300 秒兜底

## 3. 复制本技能到新机器时的部署步骤

1. 安装 `local-ocr-npu` 技能（含 run.ps1 与其 bin/ 目录）
2. 设置环境变量 `PAYSLIP_OCR_SCRIPT` 指向该机器 run.ps1，或在运行时 `--ocr-script`
3. 确认系统 Python ≥3.9 且已装 `openpyxl`（`pip install openpyxl`）
4. 用 `python pay_slip_ledger_scan.py --dry-run --limit 1` 试跑验证

## 4. 备注

- 该引擎同时被采购文档「一键生成」工作流复用（同一 run.ps1）
- 单张识别约 20 秒（CPU），37 张约 12-13 分钟；大批量投放时设置足够超时、后台运行
