# Brief Cost Tracker Case Data Builder v1.0

## 概述

Brief Cost Tracker Case Data Builder 是一个专业的数据提取工具，用于从 Apple 零售合作伙伴的 Excel 业务案例文件中批量提取和整合财务及运营数据。该工具支持季度级别的批量处理，并提供数据合并功能。

## 主要特性

- **批量数据提取**: 从多个 Excel 文件中提取 64 个关键字段
- **VLOOKUP 功能**: 自动查询城市排名和 RTM Level 4 信息
- **智能文件名解析**: 自动从父文件夹名称提取 Case No. 和 SFOID
- **季度文件夹处理**: 自动遍历 W1-W13 周文件夹
- **数据合并**: 支持与历史数据合并，自动去重
- **统计报告**: 提供详细的处理统计和数据分析
- **蓝色主题界面**: 现代化的图形用户界面

## 使用说明

### 1. 系统要求

**必需的 Python 库**:
```bash
pip install pandas openpyxl
```

**必需的辅助文件**:
- `city ranking.xlsx` - 城市排名数据（用于 VLOOKUP）
- `Reseller list_FY26Q1W13.xlsx` - 经销商列表（用于 RTM Level 4 查询）

### 2. 文件夹结构

```
Quarter Folder/
├── W1/                    # 第1周文件夹
│   ├── 1) xxx (FY26xxx)/  # Case 文件夹（包含 Case No. 和 SFOID）
│   │   └── APR_xxx.xlsx
│   └── 2) xxx (FY26xxx)/
├── W2/
├── ...
└── W13/
```

**注意**:
- 周文件夹必须以 W 开头，后跟数字 1-13
- Case 文件夹名称格式: `数字) 名称 (SFOID)`
- 每个 Case 文件夹应包含 APR 或 APP 开头的 Excel 文件

### 3. 操作流程

#### 步骤 1: 启动程序
```bash
python "brief cost tracker case data builder"
```

#### 步骤 2: 选择必需文件
1. 点击 "Select" 按钮选择 **City Ranking File**
2. 点击 "Select" 按钮选择 **Reseller List File**
3. 点击 "Browse" 按钮选择季度文件夹

#### 步骤 3: 设置参数
- 选择年份和季度（程序会自动填充当前值）
- (可选) 点击 "Upload" 上传历史结果文件进行数据合并

#### 步骤 4: 开始提取
- 点击 "Start Extraction" 按钮
- 等待处理完成，查看汇总报告

#### 步骤 5: 保存结果
- 双击任意行查看完整的 64 个字段详情
- 点击 "Statistics" 查看数据统计
- 点击 "Save" 保存为 Excel 文件

## 提取的数据字段 (64个)

### 基本信息 (1-8)
1. RTM Level 4 - 从 Reseller list 查询
2. APR or APP - Input!M26
3. RTM# - 留空
4. Self-fund - 留空
5. Review date - 当前程序运行日期
6. New/Refit - Input!I22
7. Status - 固定为 "WW approved"
8. Comments - 留空

### 识别信息 (9-13)
9. HQ ID - Input!D14
10. Case No. - 从父文件夹名称提取
11. Apple ID - Input!D17 (带 fallback 逻辑)
12. SFOID/工单编号 - 从父文件夹名称提取
13. Appoint date - 留空

### 店铺信息 (14-19)
14. Reseller name - Input!D13
15. Store name - Input!I23
16. City - Input!D24 (处理合并单元格)
17. City rank - 从 city ranking.xlsx 查询
18. City Tier - 留空
19. Appointed week - 留空

### 商场数据 (20-23)
20. Mall sales (RMB Bn) - 计算字段
21. Mall size (sqm) - Input!Q25
22. Mall sales / sqm (RMB) - 计算字段
23. Store size - Input!I26

### 第一年运营数据 (24-30)
24. Year 1 Traffic - Input!M22
25-29. Year 1 各产品线 ST/week - Input!V51-V55
30. Year 1 third Party % of total revenue - Input!R61

### 财务数据 (31-43)
31-38. Year 1 收入、利润、费用等 - 从 Partner P&L sheets 读取
39. Credit card cost as % of revenue - Input!M80
40. Year 1 BE Revenue (USD/year) - Input!M53
41. Year 1 BE Revenue (RMB/month) - 计算字段
42. Year 1 net income (USD) - 计算字段
43. Year 1 net income % - 计算字段

### 成本数据 (44-54)
44-49. 租金、物业成本相关 - 从 Input sheet 读取或计算
50-54. 人员成本相关 - 从 Input sheet 读取或计算

### 投资回报 (55-64)
55. SFF ($) - Input!D56
56. SLF ($) - Input!D57
57. Traffic Counter ($) - Input!D58
58. ROI - 计算字段 (Apple Rev / (SFF+SLF+Traffic Counter))
59. EBITDA - Partner P&L ($USD)!H88
60. Payback Period - 计算字段
61. CY - 当前年份
62. Financial Risk - 留空
63. Risk Comment - 留空
64. HQ Burden cost as % of revenue - Input!M79

### 元数据字段
- Filename, Year, Quarter, Week, YearQuarter, ProcessingTime

## 关键功能说明

### VLOOKUP 功能
- **City Rank**: 根据城市名称自动查询城市排名
- **RTM Level 4**: 根据经销商名称自动查询 RTM Level 4
- 支持不区分大小写的模糊匹配

### 智能提取
- **Case No.**: 从父文件夹名称开头提取数字 (如 "9) xxx" → "9")
- **SFOID**: 从父文件夹名称括号中提取 (如 "xxx (FY260010740)" → "FY260010740")
- **Apple ID**: 从 D17 读取，如为空则尝试 fallback 单元格 (D16, D18, E17)

### 数据合并
- 基于 Filename 字段自动去重
- 保留第一次出现的记录
- 更新处理时间戳

## 输出文件

**文件命名**: `CostTracker_{Year}{Quarter}_Result_{Timestamp}.xlsx`
**Sheet 名称**: Cost Tracker Data
**列顺序**: 64 个数据字段 + 6 个元数据字段

## 常见问题

### 1. City Rank 显示 "Unknown"
- 检查 City Ranking 文件是否正确加载
- 确认城市名称拼写与 City Ranking 文件中一致

### 2. RTM Level 4 为空
- 检查 Reseller List 文件是否正确加载
- 确认经销商名称与 Reseller List 文件中一致
- 查看控制台调试信息

### 3. Apple ID 为空
- 查看控制台调试信息，确认 D17 单元格是否有值
- 程序会自动尝试 fallback 单元格 (D16, D18, E17)

### 4. Case No. 或 SFOID 为空
- 检查父文件夹名称格式是否正确
- 正确格式: `数字) 名称 (SFOID)`
- 查看控制台调试信息

## 版本信息

**当前版本**: v1.0
**发布日期**: 2026-02-04
**Python 版本**: Python 3.7+
**依赖库**: pandas >= 1.0.0, openpyxl >= 3.0.0

## 更新日志

### v1.0 (2026-02-04)
- 初始版本发布
- 支持 64 个字段提取
- 实现 VLOOKUP 功能
- 添加数据合并功能
- Apple ID 增强提取逻辑 (fallback 机制)
- Review date 使用当前程序运行日期
- 从父文件夹名称提取 Case No. 和 SFOID
- 蓝色主题界面
- 完整的统计报告功能

## 技术支持

如遇到问题，请检查:
1. Python 和依赖库版本是否正确
2. 辅助文件 (City Ranking, Reseller List) 是否已选择
3. Excel 文件格式是否符合要求 (包含 input 和 Partner P&L sheets)
4. 文件夹结构是否正确 (W1-W13 周文件夹)
5. 控制台调试信息

开发维护: Seiya (seiya_wu@apple.com)
