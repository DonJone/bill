# CLAUDE.md — Bill 项目上下文

## 项目定位

个人月度账单统计 CLI 工具。用户从微信/支付宝/中国银行导出账单，导入后自动分类、去重、生成报告。

## 技术栈

- Python 3，标准库为主
- openpyxl（微信 xlsx）、pdfplumber（中行 PDF）
- SQLite（数据存储）
- Chart.js CDN（HTML 报告图表）
- argparse（CLI）

## 架构

单文件 `bill.py`（~1700 行），按区块组织：

| 区块 | 行号范围 | 职责 |
|------|---------|------|
| 配置 | 1-100 | 路径、默认分类规则、Alipay 分类映射 |
| 数据库 | 100-200 | SQLite CRUD、月度汇总/分类/趋势查询 |
| 分类引擎 | 200-350 | 关键词匹配 + Alipay 分类映射兜底 |
| 解析器 | 350-550 | 微信 xlsx / 支付宝 csv / 中行 pdf / xlsx |
| 去重 | 550-650 | 跨源去重 + 同源模糊商户名去重 |
| 导入逻辑 | 650-700 | 重叠检测、归档、插入 |
| 终端报告 | 700-800 | 表格 + 柱状图 |
| HTML 报告 | 800-1300 | Chart.js 5 图 + 2 表 |
| CLI | 1300-1700 | argparse 子命令 |

## 关键设计决策

- 分类优先级：**用户关键词 > Alipay 自带分类**（方便覆盖支付宝分类不准的情况）
- 支付宝分类存储在 remark 字段：`[ALIPAY_CAT:xxx]`，重分类时可重新读取
- 去重三层：UNIQUE 约束（插入时）→ 跨源日期+金额匹配 → 同源模糊商户名（SequenceMatcher > 0.5）
- query 管道自动 TSV：`not sys.stdout.isatty()` 时自动切换
- 错误 → stderr + exit(1)；数据 → stdout

## 数据流

```
文件 → parse_file() → detect_format() → [parse_wechat|parse_alipay|parse_boc_pdf]
     → categorize() → insert_transactions() → check_source_overlap()
     → archive_import() → find_and_mark_duplicates() → dedup_same_source()
```

## 分类规则

- `rules.json`：首次运行自动从 DEFAULT_RULES 生成，可手动编辑
- `bill rules add`：命令行添加关键词，自动重分类全部历史
- Alipay 交易分类映射表 `ALIPAY_CAT_MAP` 在代码中

## 国际化

所有用户界面为中文，包括帮助文本、报告、错误信息。

## 已知限制

- 微信日期是 Excel 序列数，需 `excel_serial_to_date()` 转换
- 中行 PDF 密码每次不同，需要 `--password` 参数
- BOC 对方账户名含换行符，需 `.replace("\n", "")` 清理
- Alipay CSV 编码为 GBK，解析时尝试多种编码
