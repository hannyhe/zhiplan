# 项目信息

## GitHub
- 仓库: `git@github.com:hannyhe/zhiplan.git`
- 分支: `main`
- Git 路径: `C:\Users\coldt\AppData\Local\Programs\Git\bin\git.exe`

## 工作目录
- `G:\me\致知收益\2025\zhiplan`

## 项目说明
致知计划内容收益日报自动更新系统。
从 `致知计划内容收益.xls` 读取数据，写入 `26MM.xlsx` 工作簿的新 sheet，
自动匹配前日数据并计算差值/收益率。

## 关键脚本
- `update0703.ps1` — 主程序：自动文件名、模糊匹配、跨月拆分
- `AutoUpdate.ps1` — 旧版主程序
- `replace_2607.cmd` — 替换生产文件
