# Bill — 个人月度账单统计工具

支持**微信**、**支付宝**、**中国银行**账单导入，自动分类、去重，生成交互式 HTML 报告。

## 安装

```bash
git clone https://github.com/DonJone/bill.git && cd bill
python3 -m venv myenv && source myenv/bin/activate
pip install -r requirements.txt
```

## 快速开始

```bash
# 激活虚拟环境后
source myenv/bin/activate

# 导入账单
bill import 微信.xlsx 支付宝.csv 中行.pdf --password 123456

# 报告
bill report                  # 当月终端报告
bill report --html           # HTML 交互图表
bill trend -m 12             # 近12月趋势

# 检索 & 修正
bill query -m 2026-05 -c 餐饮
bill query -k 美团 --min 50
bill edit 42 -c 餐饮美食

# 管理分类规则
bill rules
bill rules add 餐饮美食 新商户名
```

## 支持来源

| 平台 | 格式 | 说明 |
|------|------|------|
| 微信 | xlsx | 邮箱下载，自动识别 |
| 支付宝 | csv | 邮箱下载，GBK 编码 |
| 中国银行 | pdf | 带密码保护，pdfplumber 解析 |

## 特性

- 自动分类（14 个默认分类 + 关键词规则 + 支付宝自带分类兜底）
- 三层去重（插入时跳过 + 跨源匹配 + 同源模糊商户名）
- 导入归档（自动重命名为 `来源_起止日期_截止日期.格式`）
- HTML 报告（5 个 Chart.js 图表 + 分类明细表 + 每笔支出列表）
- Unix 管道支持（`bill query` 管道自动输出 TSV，stderr 分离）
