# Bill — 个人月度账单统计工具

支持**微信**、**支付宝**、**中国银行**及**任何可导出账单银行**账单导入，自动分类、去重，生成交互式 HTML 报告。

## 安装

```bash
git clone https://github.com/DonJone/bill.git && cd bill
python3 -m venv myenv && source myenv/bin/activate
pip install -r requirements.txt
```

依赖：`openpyxl`（xlsx）、`pdfplumber`（PDF）。

## 快速开始

```bash
source myenv/bin/activate

# 1. 导出账单（详见 bill export-guide）
#    微信 xlsx / 支付宝 csv / 中国银行 pdf

# 2. 导入
bill import 微信.xlsx 支付宝.csv 中行.pdf -p 524531

# 3. 查看
bill report                  # 终端报告
bill report --html           # HTML 交互图表
bill trend -m 12             # 近 12 月趋势

# 4. 检索
bill query -m 2026-05 -c 餐饮
bill query -k 美团 --min 50
```

## 命令一览

| 命令 | 说明 |
|------|------|
| `bill import <文件...>` | 导入账单（自动识别格式，支持自定义银行） |
| `bill report [月份]` | 终端/HTML 报告 |
| `bill daily [月份]` | 每日消费分析（日均、周几、工作日/周末） |
| `bill trend [选项]` | 月度趋势（支持自定义日期范围） |
| `bill query [选项]` | 检索明细（排序、管道 TSV） |
| `bill edit <ID> -c 分类` | 修正单笔分类 |
| `bill rules [add\|add-category]` | 管理分类规则（自定义分类） |
| `bill export-guide` | 各平台账单导出教学 |

详细用法见 **[USAGE.md](USAGE.md)**。

## 特性

- 每日消费分析（日均、中位、工作日/周末对比）
- 自定义分类 + 自定义银行来源
- 自动分类（14 个默认分类 + 可编辑关键词规则 + 支付宝自带分类兜底）
- 三层去重（插入时四元组 → 跨源日期金额 → 同源模糊商户名）
- 导入归档（文件自动重命名为 `来源_起止日期.格式`）
- HTML 报告（5 个 Chart.js 图表 + 分类明细表 + 每笔支出列表）
- Unix 管道（`bill query` 管道自动 TSV，stderr 分离）

## 卸载 / 清除数据

```bash
# 清除所有导入的账单数据
rm -rf data/

# 清除分类规则（下次运行自动重建默认）
rm rules.json

# 清除虚拟环境
rm -rf myenv/

# 完全卸载
cd .. && rm -rf bill/
```

所有数据 100% 本地存储，无网络传输。

## 数据隐私

- 交易数据 → SQLite `data/bills.db`（本地）
- 导入文件 → 归档到 `data/imports/`（本地）
- HTML 报告 → 交易数据嵌入 HTML 中，Chart.js 从 CDN 加载（需联网一次）
- `.gitignore` 已排除 `data/` 和 `myenv/`

## 项目文件

| 文件 | 说明 |
|------|------|
| `bill.py` | 主程序（~1700 行单文件） |
| `rules.json` | 分类规则（首次运行自动生成，可手动编辑） |
| `requirements.txt` | Python 依赖 |
| `USAGE.md` | 详细使用手册 |
| `CLAUDE.md` | 项目上下文（给 Claude Code 用） |
