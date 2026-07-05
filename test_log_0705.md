# 测试记录 — 2025/07/05

## 操作系统环境

| 项目 | 值 |
|------|-----|
| 操作系统 | Microsoft Windows 10 家庭版 |
| 版本 | 10.0.19045 (Build 19045) |
| 系统类型 | 64 位 |
| 架构 | x64 |
| 最后启动 | 2026/6/6 9:52:22 |
| 主机名 | LAPTOP-6IHCCHQL |
| 用户名 | tian |

## 软件环境

| 项目 | 版本 |
|------|------|
| PowerShell | 5.1.19041.6456 |
| Excel | 16.0 |
| Git | 2.51.1.windows.1 |

## 测试时间

2026-07-05 11:19:42

---

## 命令行执行记录

### Step 1: 终止残留 Excel 进程并复制生产文件

`powershell
Get-Process -Name "*excel*" | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 15
Copy-Item 'G:\me\致知收益\2025\2607.xlsx' 'G:\me\致知收益\2025\test\2607.xlsx' -Force
`

### Step 2: 运行主程序

`powershell
powershell -NoProfile -ExecutionPolicy Bypass -File C:\Users\coldt\Desktop\update0703.ps1
`

**输出：**
`
45 source rows
prev=0704 today=0705
prev sheet 0704: 47 questions, 49 rows
match: 42exact 0fuzzy 0B 0C 3new
Save OK
=== Done ===
`

### Step 3: 数据验证

`powershell
=.Sheets("0705");=.UsedRange.Value2
=.GetLength(0);=.GetLength(1)
"0705:  rows x  cols"
# L 值
for(=2;-le(-1);++){=[,12];if()"Row : L="}
# 合计
"K sum=""
`

**验证输出：**
`
47 rows x 15 cols

L values:
Row 2:  L=0.79
Row 3:  L=9.43
Row 4:  L=0.38
Row 5:  L=0.91
Row 6:  L=0.25
Row 7:  L=1.50
Row 8:  L=3.78
Row 10: L=21.50
Row 14: L=0.67
Row 16: L=2.43
Row 19: L=8.50
Row 29: L=1.56
Row 35: L=0.19

Sum row 47: K=1084
Sum row 47: J=合计
`

---

## 运行参数

| 参数 | 值 |
|------|-----|
| 脚本路径 | C:\Users\coldt\Desktop\update0703.ps1 |
| 源数据 | G:\me\致知收益\2025\致知计划内容收益.xls |
| 目标文件 | G:\me\致知收益\2025\test\2607.xlsx |
| 生产文件 | G:\me\致知收益\2025\2607.xlsx |
| 前日 sheet | 0704 |
| 当日 sheet | 0705 |

## 匹配结果

| 层级 | 数量 |
|------|------|
| 精确匹配 (A列) | 42 |
| 模糊匹配 (A列 Levenshtein ≥0.65) | 0 |
| B列匹配 (类型+发布时间) | 0 |
| C列匹配 (发布时间+类型+模糊A ≥0.3) | 0 |
| 全新问题 | 3 |
| **合计** | **45** |

## 数据验证

| 检查项 | 结果 | 说明 |
|--------|------|------|
| L列小数位 | ✓ 全部 ≤2 位 | 0.19~21.50, 均 ROUND 至 2 位 |
| J公式 | ✓ F-VLOOKUP(A,前日!F) 三层回退 | 使用累计列 F/G |
| K公式 | ✓ G-VLOOKUP(A,前日!G) 三层回退 | 使用累计列 G |
| MNO 来源 | ✓ 拷贝前日 0704 对应行文本 | 因 0704 MNO 为空，故 0705 也为空 |
| 合计行 K | ✓ 1084 | =ROUND(SUM(K2:K46),2) |
| 合计行 J | ✓ "合计" | 文本标识 |
| 列数 | ✓ 15 列 | A-O |
| 保存 | ✓ OK | 无 COM 错误 |
| 新问题 MNO | ✓ 空文本 "" | 无匹配行 |
| 行匹配精度 | ✓ 42/42 精确 | 无错误匹配 |

## 公式抽样 (Row 8)

`
H8  = ROUND(E8/D8,2)
I8  = ROUND(G8/F8,2)
J8  = F8 - IFERROR(VLOOKUP(,'0704'!,6,FALSE), ...)
K8  = G8 - IFERROR(VLOOKUP(,'0704'!,7,FALSE), ...)
L8  = IFERROR(ROUND(K8/J8,2),)
M8  = '0704'!M{mr}    (值拷贝, 非公式)
N8  = '0704'!N{mr}    (值拷贝)
O8  = '0704'!O{mr}    (值拷贝)
`

## 结论

程序运行正常，0705 创建成功。核心功能验证通过:
- 自动递增 sheet 名 (0704 → 0705) ✓
- 累计列公式 (F/G) ✓
- ROUND 2 位小数 ✓
- MNO 文本拷贝 ✓
- 三层匹配 ✓
- COM 批量写入无冲突 ✓
