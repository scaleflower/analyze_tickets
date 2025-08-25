# OTRS Ticket Analysis - Windows Installation Guide

## 快速开始

### 方法一：双击运行（推荐）
1. 下载所有文件到同一文件夹
2. 双击运行 `run_analysis.bat`
3. 脚本会自动检查并安装所需依赖
4. 程序会自动分析默认的Excel文件

### 方法二：指定文件分析
1. 将Excel文件拖拽到 `run_analysis.bat` 上
2. 或者右键选择"使用此文件运行分析"

## 系统要求

- **操作系统**: Windows 7/8/10/11
- **Python**: 3.6 或更高版本（脚本会自动检测）
- **内存**: 至少 4GB RAM（处理大型Excel文件时建议8GB+）

## 手动安装步骤

如果自动脚本无法正常工作，可以手动安装：

### 1. 安装 Python
- 访问 [Python官网](https://www.python.org/downloads/)
- 下载最新版本的Python安装程序
- **重要**: 安装时勾选 "Add Python to PATH"
- 完成安装后重启命令行

### 2. 安装依赖包
打开命令提示符（CMD）并运行：
```cmd
pip install pandas openpyxl numpy
```

### 3. 运行分析
```cmd
python analyze_tickets.py [Excel文件路径]
```

## 批处理脚本说明

### install_requirements.bat
- 检查Python和pip是否安装
- 自动安装所需的Python包
- 提供详细的错误信息

### run_analysis.bat  
- 自动检查依赖是否已安装
- 支持拖拽文件运行
- 提供友好的用户界面

## 常见问题解决

### Q: 脚本提示"Python is not installed"
A: 请手动安装Python并确保勾选"Add Python to PATH"

### Q: 安装包时出现权限错误
A: 尝试以管理员身份运行命令提示符：
```cmd
pip install pandas openpyxl numpy
```

### Q: 内存不足错误
A: 关闭其他程序，或使用较小的Excel文件

### Q: Excel文件格式不支持
A: 确保文件是.xlsx或.xls格式

## 文件说明

- `analyze_tickets.py` - 主分析程序
- `requirements.txt` - Python依赖列表
- `install_requirements.bat` - 自动安装脚本
- `run_analysis.bat` - 一键运行脚本
- `README_WINDOWS.md` - 本说明文件

## 技术支持

如果遇到问题，请检查：
1. Python是否正确安装并添加到PATH
2. 网络连接是否正常（安装包需要下载）
3. Excel文件是否可访问且格式正确

脚本会自动生成带时间戳的日志文件，便于排查问题。
