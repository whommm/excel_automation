# 库存批量修改自动化工具

一个基于 Python + PyAutoGUI 的 RPA（机器人流程自动化）工具，用于批量自动化修改库存数据。

## 功能特点

- 傻瓜式图形界面，无需编程知识
- 从 Excel 读取编码和库存数据
- 支持坐标点击自动化操作
- 可自定义操作步骤流程
- 支持循环执行多条数据

## 环境要求

- Windows 10/11
- Python 3.8+
- 屏幕分辨率建议 1920x1080

## 快速部署

### 1. 克隆仓库

```bash
git clone https://github.com/whommm/excel_automation.git
cd excel_automation
```

### 2. 创建虚拟环境

```bash
# 创建虚拟环境
python -m venv venv

# 激活虚拟环境 (Windows CMD)
venv\Scripts\activate.bat

# 或者 (Windows PowerShell)
venv\Scripts\Activate.ps1

# 或者 (Git Bash / MSYS2)
source venv/Scripts/activate
```

### 3. 安装依赖

```bash
pip install -r requirements.txt
```

### 4. 启动程序

双击 `启动控制面板.bat` 或运行：

```bash
python control_panel.py
```

## 使用说明

### 第一步：配置 Excel

1. 打开控制面板，进入「1. 选择Excel」标签页
2. 点击「浏览...」选择包含库存数据的 Excel 文件
3. 设置编码列名和库存列名（默认为「编码」和「库存数」）
4. 点击「保存Excel设置」

### 第二步：配置操作步骤

1. 进入「2. 配置步骤」标签页
2. 点击「添加步骤」创建新的操作步骤
3. 配置每个步骤：
   - **步骤名称**：给步骤起个名字便于识别
   - **动作类型**：
     - 点击坐标：在指定坐标位置点击
     - 双击坐标：在指定坐标位置双击
     - 输入文本：输入文字（支持 `{code}` 和 `{quantity}` 占位符）
     - 按键：模拟键盘按键
     - 等待秒数：暂停指定时间
     - 清空输入框：清空当前输入框内容
   - **点击坐标**：点击「获取鼠标位置」按钮，3秒内将鼠标移到目标位置
   - **操作后等待**：每个操作完成后的等待时间

4. 使用「上移」「下移」调整步骤顺序
5. 点击「保存步骤」

### 第三步：运行自动化

1. 进入「3. 开始运行」标签页
2. 检查所有配置项是否显示「OK」
3. 设置循环次数（0 表示按 Excel 数据条数执行）
4. 点击「开始运行」
5. 程序会自动最小化，请切换到目标软件窗口

### 安全提示

- 运行过程中，将鼠标移到屏幕**左上角**可紧急停止程序
- 建议先用少量数据测试，确认无误后再批量执行

## 文件结构

```
excel_automation/
├── control_panel.py      # 主控制面板 GUI
├── main_bot.py           # 自动化核心逻辑
├── config.yaml           # 配置文件
├── requirements.txt      # Python 依赖
├── 启动控制面板.bat       # 一键启动脚本
├── data/                 # Excel 数据目录
│   └── inventory.xlsx    # 示例数据文件
├── assets/               # 资源文件目录
├── logs/                 # 运行日志目录
└── venv/                 # Python 虚拟环境（需自行创建）
```

## 常见问题

### Q: 点击位置不准确？
A: 确保运行时屏幕分辨率和录制坐标时一致，且目标软件窗口位置相同。

### Q: 中文输入乱码？
A: 程序使用剪贴板方式输入中文，请确保系统剪贴板正常工作。

### Q: 程序无响应？
A: 将鼠标移到屏幕左上角触发安全停止，然后检查日志文件排查问题。

## 依赖说明

- `pyautogui` - 屏幕自动化操作
- `pandas` - Excel 数据处理
- `openpyxl` - Excel 文件读取
- `pyyaml` - 配置文件解析
- `pyperclip` - 剪贴板操作
- `pillow` - 图像处理

## 许可证

MIT License
