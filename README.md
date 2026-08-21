# BidBrief — 招标文件要点提取工具

投标前从几十页招标文件中快速抽取**商务 / 技术 / 资质 / 时间节点 / 废标项 / 评分标准**要点，
生成带原文页码引用的核对清单（Excel / Word / PPT / Markdown）。

## 快速上手

**方式一（推荐，免命令行）**：双击项目里的 `启动BidBrief.bat`

**方式二**：

```bash
pip install -r requirements.txt
python main.py
```

界面四步：

1. 选择招标文件所在文件夹，勾选文件类型（.pdf / .docx / .txt / .md），可选包含子文件夹
2. 模型设置里填一次 API Key（自动保存到 config.json，下次不用再填）
3. （可选）在"自定义抽取要求"里限定范围，例如：*仅做劳务分包，只关注分包资质、总包配合费、禁止转包条款及相关废标项*；支持任意筛选条件（如"只看与 BIM 相关的条款"），留空则按默认六类全量抽取
4. 点 **开始**；随时可点 **停止**（当前小块完成后中止，已完成文件照常导出），结束后一键打开输出目录

## 功能

- 每条要点附**原文摘录 + 页码**，可跳回原文核对
- 四种导出格式（可勾选）：
  - `要点报告_时间戳.xlsx`：要点总表 / 废标核对单 / 文件概览 三个 Sheet
  - `要点报告_时间戳.docx`：Word 版核对清单（废标项带勾选框，可打印逐条核对）
  - `要点汇报_时间戳.pptx`：汇报用幻灯片（每页 7 条，超长自动分页）
  - `核对清单_时间戳.md`：Markdown 勾选清单
  - `raw/*.json`：逐文件原始抽取结果
- 命令行模式：`python main.py --cli "D:\招标文件" --ext .pdf,.docx --formats xlsx,docx`
- 端到端自测：`python main.py --selftest`（生成样例 PDF 并真实调用一次 API）

## 容量与耗时

- **单个文件大小不限**：文档按约 5000 字分块依次送入模型，300 页也能处理
- 单次请求远低于 DeepSeek 上下文上限（64K~128K tokens，约可容纳 10 万汉字）
- 耗时随页数线性增长：约每块（~5000 字）30~60 秒；100 页约 10~20 分钟
- 费用很低：deepseek-chat 输入约 2 元/百万 tokens、输出 8 元/百万 tokens，100 页文件通常不到 1 元
- 想减少调用次数，可在 `config.json` 调大 `chunk_chars`（如 8000~10000），块过大可能略降抽取精度

## 配置

首次运行可在界面"模型设置"里填写 API Key，或直接编辑 `config.json`：

```json
{
  "api_key": "sk-...",
  "base_url": "https://api.deepseek.com",
  "model": "deepseek-chat",
  "custom_prompt": "可选：自定义抽取要求，与界面输入框内容相同"
}
```

`config.json` 已加入 `.gitignore`，**请勿提交或分享该文件**。

## 项目结构

```
main.py            入口（GUI / --cli / --selftest）
selftest.py        端到端自测（生成样例招标PDF并校验结果）
bidbrief/
  scanner.py       文件夹扫描 + 后缀过滤
  parser.py        PDF/DOCX/TXT 解析（识别扫描件与加密件）
  chunker.py       分块 + 原文摘录回搜定位精确页码
  extractor.py     DeepSeek 抽取（六类要点，自定义提示词，自动重试）
  pipeline.py      流水线（全程可取消，已完成结果保留）
  exporter.py      Excel / Word / PPT / Markdown 导出
  ui.py            PySide6 图形界面
```

## 打包成 exe

```bash
python -m PyInstaller --onefile --windowed --name BidBrief --collect-all pymupdf --noconfirm main.py
copy config.json dist\
```

产物为 `dist\BidBrief.exe`（约 100MB）。`config.json` 必须放在 exe 同目录——
API Key、自定义提示词在界面里修改后保存到该文件，换机器只需拷贝
`BidBrief.exe + config.json` 两个文件。杀毒软件对单文件 exe 偶有误报，加白名单即可。

## 已知限制

- **扫描件 PDF**（无文字层）会标记为"跳过(扫描件需OCR)"，本期不做 OCR
- **CA 加密 PDF** 无法读取（部分交易平台下载的文件）
- DOCX / TXT 无页码信息，页码列显示"—"
- AI 抽取仅供辅助核对，**不能替代人工确认**，投标前请逐条核对原文
