# BidBrief — 招标文件要点提取工具

投标前从几十页招标文件中快速抽取**商务 / 技术 / 资质 / 时间节点 / 废标项 / 评分标准**要点，
生成带原文页码引用的核对清单（Excel + Markdown）。

## 功能

- 图形界面：选择文件夹、勾选文件后缀（.pdf / .docx / .txt / .md）、可包含子文件夹
- **自定义抽取要求**：可填写提示词限定抽取范围（如"仅做劳务分包，只关注分包资质与相关废标项"），留空则按默认六类全量抽取
- 运行中可随时**停止**：当前小块完成后中止，已完成文件的结果照常导出
- 每条要点附**原文摘录 + 页码**，可跳回原文核对
- 四种导出格式（可勾选）：
  - `要点报告_时间戳.xlsx`：要点总表 / 废标核对单 / 文件概览 三个 Sheet
  - `要点报告_时间戳.docx`：Word 版核对清单（废标项带勾选框）
  - `要点汇报_时间戳.pptx`：汇报用幻灯片（每页7条，适合评审会议）
  - `核对清单_时间戳.md`：Markdown 勾选清单
  - `raw/*.json`：逐文件原始抽取结果

## 容量与耗时

- **单个文件大小不限**：文档按约 5000 字分块依次送入模型，300 页也能处理
- 单次请求远低于 DeepSeek 上下文上限（64K~128K tokens，约可容纳 10 万汉字）
- 耗时随页数线性增长：约每块（~5000字）30~60 秒；100 页约 10~20 分钟
- 费用很低：deepseek-chat 输入约 2 元/百万 tokens、输出 8 元/百万 tokens，100 页文件通常不到 1 元
- 想减少调用次数，可在 `config.json` 调大 `chunk_chars`（如 8000~10000），块过大可能略降抽取精度

## 安装与运行

```bash
pip install -r requirements.txt
python main.py          # 图形界面
```

命令行模式：

```bash
python main.py --cli "D:\招标文件" --ext .pdf,.docx --formats xlsx,docx
python main.py --selftest   # 端到端自测（会真实调用一次 API）
```

## 配置

首次运行可在界面"模型设置"里填写 API Key（保存到 `config.json`），或直接编辑该文件：

```json
{ "api_key": "sk-...", "base_url": "https://api.deepseek.com", "model": "deepseek-chat" }
```

`config.json` 已加入 `.gitignore`，**请勿提交或分享该文件**。

## 打包成 exe

```bash
python -m PyInstaller --onefile --windowed --name BidBrief --collect-all pymupdf --noconfirm main.py
copy config.json dist\
```

产物为 `dist\BidBrief.exe`（约100MB）。`config.json` 必须放在 exe 同目录——
API Key、自定义提示词在界面里修改后保存到该文件，换机器只需拷贝
`BidBrief.exe + config.json` 两个文件。杀毒软件对单文件 exe 偶有误报，加白名单即可。

## 已知限制

- **扫描件 PDF**（无文字层）会标记为"跳过(扫描件需OCR)"，本期不做 OCR
- **CA 加密 PDF** 无法读取（部分交易平台下载的文件）
- DOCX / TXT 无页码信息，页码列显示"—"
- AI 抽取仅供辅助核对，**不能替代人工确认**，投标前请逐条核对原文
- 处理速度取决于 DeepSeek 响应，约每块（~5000字）30~60 秒；100 页文件约 10~20 分钟
