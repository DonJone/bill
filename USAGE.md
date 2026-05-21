# Bill 使用手册

## 目录

- [安装与卸载](#安装与卸载)
- [数据与隐私](#数据与隐私)
- [命令详解](#命令详解)
  - [import — 导入账单](#import--导入账单)
  - [report — 生成报告](#report--生成报告)
  - [trend — 月度趋势](#trend--月度趋势)
  - [query — 检索明细](#query--检索明细)
  - [edit — 修正分类](#edit--修正分类)
  - [rules — 管理规则](#rules--管理规则)
  - [export-guide — 导出指南](#export-guide--导出指南)
- [账单导出教程](#账单导出教程)
- [分类体系](#分类体系)
- [去重机制](#去重机制)

---

## 安装与卸载

### 安装

```bash
git clone https://github.com/DonJone/bill.git
cd bill
python3 -m venv myenv
source myenv/bin/activate
pip install -r requirements.txt
```

依赖项：`openpyxl`（读 xlsx）、`pdfplumber`（读 PDF）。

### 卸载 / 清除数据

```bash
# 清除所有导入的账单数据（SQLite 数据库）
rm -rf data/

# 清除导入归档
rm -rf data/imports/

# 清除分类规则（下次运行自动重建默认规则）
rm rules.json

# 清除虚拟环境
rm -rf myenv/

# 完全卸载（删除整个项目目录）
cd .. && rm -rf bill/
```

数据 100% 存储在本地 `data/` 目录和 `rules.json`，无任何网络传输。

---

## 数据与隐私

- 所有交易数据存储在本地 SQLite 文件 `data/bills.db`
- 导入的账单文件自动归档到 `data/imports/`（重命名为 `来源_起止日期.格式`）
- HTML 报告使用 Chart.js CDN 加载图表库（需联网一次），交易数据嵌入 HTML 中
- `.gitignore` 已排除 `data/`、`myenv/`，不会误提交敏感数据

---

## 命令详解

### import — 导入账单

```bash
bill import <文件...> [--password PASSWORD] [--source NAME]
```

**参数**

| 参数 | 说明 |
|------|------|
| `文件` | 一个或多个账单文件路径 |
| `--password, -p` | PDF 密码 |
| `--source, -s` | 自定义来源名称（覆盖自动检测，用于非微信/支付宝/中行的银行） |

**示例**

```bash
# 单文件
bill import 微信支付账单.xlsx

# 多文件
bill import 微信.xlsx 支付宝.csv

# 带密码的中行 PDF
bill import 中行.pdf -p 524531

# 一次性导入所有来源
bill import 微信.xlsx 支付宝.csv 中行.pdf -p 524531

# 自定义银行（自动检测 CSV 列）
bill import 招商银行.csv --source 招商银行
bill import 工商银行.xlsx -s 工商银行
```

**行为**

1. 自动识别文件格式（通过文件名和扩展名）
2. 解析前检测该来源是否已有重叠数据，提示已有记录数
3. 插入时按 `(日期, 金额, 商户, 来源)` 四元组跳过完全重复行
4. 导入后自动归档到 `data/imports/`，文件重命名为 `来源_起止日期.格式`
5. 自动执行跨源去重（银行 ↔ 支付平台）
6. 自动执行同源去重（相近商户名模糊匹配）

**输出示例**

```
正在解析: 微信支付账单.xlsx ... 
  注意: wechat 在 2025-05-22 ~ 2026-05-21 区间已有 1109 条记录
  重复记录将被自动跳过（按 日期+金额+商户 去重）
[OK] 支出953 收入178 中性18 → 新增0 跳过1149
  已归档: data/imports/wechat_20250522_20260521.xlsx
跨源去重: 标记了 113 条重复（银行 ↔ 支付平台）
同源去重: 标记了 26 条重复（相近商户名）
```

---

### report — 生成报告

```bash
bill report [YYYY-MM] [--html] [--no-open]
```

**参数**

| 参数 | 说明 |
|------|------|
| `YYYY-MM` | 指定月份，省略则为最新有数据的月份 |
| `--html` | 生成 HTML 交互报告（含 Chart.js 图表） |
| `--no-open` | 不自动打开浏览器（仅 --html 时有效） |

**示例**

```bash
bill report              # 当月终端报告
bill report 2026-04      # 指定月份
bill report --html       # 当月 HTML 报告并打开
bill report 2026-05 --html --no-open  # 生成但不打开
```

**终端报告**包含：收支汇总卡片、环比变化、分类柱状图

**HTML 报告**包含：

| 图表 | 类型 | 说明 |
|------|------|------|
| 分类支出占比 | 饼图 | 当月各类占比 |
| 当月 vs 上月分类对比 | 柱状图 | 蓝色当月/灰色上月 |
| 当月每日收支趋势 | 柱状图 | 红支出/绿收入 |
| 当月每日分类支出趋势 | 折线图 | Top 5 分类各一条线 |
| 近 12 月分类趋势 | 折线图 | Top 6 分类月度走势 |

外加两张表格：分类明细（含笔数 + 环比）和当月每笔支出（高 → 低）。

---

### trend — 月度趋势

```bash
bill trend [-m N]
```

**参数**

| 参数 | 说明 |
|------|------|
| `-m, --months N` | 显示最近 N 个月（默认 6） |

**示例**

```bash
bill trend           # 近 6 月
bill trend -m 12     # 近 12 月
```

**输出**：跨月分类支出矩阵（行=分类，列=月份）+ 当月分类明细柱状图。

---

### query — 检索明细

```bash
bill query [选项...]
```

**参数**

| 参数 | 说明 |
|------|------|
| `--month, -m YYYY-MM` | 限定月份 |
| `--min N` | 最低金额 |
| `--max N` | 最高金额 |
| `--type, -t {expense,income,neutral}` | 交易类型 |
| `--source, -s {wechat,alipay,boc}` | 来源平台 |
| `--category, -c 分类` | 交易分类（支持模糊匹配，如 `-c 娱乐` 匹配"休闲娱乐"） |
| `--keyword, -k 关键词` | 搜索商户/商品/备注 |
| `--show-dup` | 包含已去重的重复交易 |
| `--plain` | 强制 TSV 输出（管道时自动启用） |

**示例**

```bash
# 查看 5 月所有交易
bill query -m 2026-05

# 5 月金额 >= 100
bill query -m 2026-05 --min 100

# 5 月 50~200 之间的支出
bill query -m 2026-05 --min 50 --max 200 -t expense

# 支付宝的所有记录
bill query -s alipay

# 模糊搜索商户名
bill query -k 美团

# 模糊匹配分类
bill query -c 娱乐          # 找到"休闲娱乐"

# 管道到 sort 按金额排序
bill query -m 2026-05 | sort -t$'\t' -k4 -rn

# 管道到 awk 计算总支出
bill query -m 2026-05 | awk -F'\t' 'NR>1{sum+=$4}END{print sum}'

# 管道到 grep 筛选
bill query -m 2026-05 | grep 餐饮
```

**输出**：终端友善表格（含 ID/日期/类型/分类/金额/商户/来源），管道时自动切换 TSV。TSV 第一行为列名，最后一行为汇总信息（输出到 stderr）。

---

### edit — 修正分类

```bash
bill edit <交易ID> [--category 新分类]
```

**参数**

| 参数 | 说明 |
|------|------|
| `交易ID` | 交易编号（从 query 输出第一列获取） |
| `--category, -c 分类` | 目标分类名称 |

**示例**

```bash
# 查看交易信息
bill edit 42

# 修改分类
bill edit 42 -c 餐饮美食
```

---

### rules — 管理规则

```bash
bill rules
bill rules add <分类> <关键词>
bill rules add-category <新分类名>
```

**示例**

```bash
# 查看当前规则
bill rules

# 添加关键词（自动重分类全部历史）
bill rules add 餐饮美食 新餐厅名
bill rules add 科技消费 ChatGPT

# 创建自定义分类
bill rules add-category 宠物
bill rules add 宠物 猫粮
```

规则保存在 `rules.json`，也可手动编辑。修改后重新运行 `bill rules add` 或删除 `rules.json` 重建默认规则。

---

### export-guide — 导出指南

```bash
bill export-guide
```

显示微信/支付宝/中国银行各平台的账单导出步骤。

---

## 账单导出教程

### 微信支付

1. 微信 → 我 → 服务 → 钱包 → 账单
2. 右上角「常见问题」→ 下载账单
3. 选择「用于个人对账」
4. 输入接收邮箱
5. 查收邮件，下载 xlsx 附件

### 支付宝

1. 支付宝 → 我的 → 账单
2. 右上角「...」→ 开具交易流水证明
3. 选择「用于个人对账」
4. 输入接收邮箱
5. 查收邮件，下载 csv 附件（解压密码在邮件中）

### 中国银行

1. 手机银行 App → 交易查询
2. 选择日期范围 → 导出
3. 选择 PDF 格式
4. PDF 需要账户密码打开（每次不同，需自行获取）

---

## 分类体系

### 默认分类（14 类）

| 分类 | 典型关键词 |
|------|-----------|
| 餐饮美食 | 餐厅、外卖、食堂、火锅、奶茶、小吃、面馆 |
| 交通出行 | 滴滴、地铁、公交、高铁、加油、停车 |
| 购物消费 | 淘宝、京东、拼多多、超市、百货 |
| 住房物业 | 房租、物业、水电、煤气、维修 |
| 休闲娱乐 | 电影、KTV、游戏、旅游、酒店、健身 |
| 医疗健康 | 医院、药店、诊所、挂号、体检 |
| 教育学习 | 学费、培训、书本、考试 |
| 科技消费 | 手机、电脑、耳机、数码、DeepSeek、话费 |
| 日用百货 | 日用品、洗衣、快递、便利店 |
| 服饰美容 | 衣服、鞋子、化妆品 |
| 人情往来 | 红包、礼物、转账、小荷包 |
| 金融服务 | 理财、提现、快捷退款、银联入账 |
| 商业服务 | 深度求索等 API/商业服务 |
| 其他 | 兜底分类 |

### 分类优先级

用户关键词规则 > 支付宝自带分类 > "其他"

---

## 去重机制

三层防护，按执行顺序：

| 层级 | 触发时机 | 匹配条件 | 处理方式 |
|------|---------|---------|---------|
| 插入层 | INSERT 时 | `(日期, 金额, 商户, 来源)` 完全相同 | `INSERT OR IGNORE` 跳过 |
| 跨源层 | import 后 | 不同来源、同日期、同金额（±0.01） | 保留 Alipay/WeChat，标记 BOC 为重复 |
| 同源层 | import 后 | 同来源、同日期、同金额、商户名相似度 > 50% | 保留先导入的，标记后来的为重复 |
