# -*- coding: utf-8 -*-
"""
库存批量修改自动化工具
使用方法: python main_bot.py
"""

import os
import sys
import time
import logging
from datetime import datetime
from pathlib import Path

import yaml
import pandas as pd
import pyautogui
import pyperclip

# 设置 PyAutoGUI 安全模式
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1


def get_base_dir():
    """获取程序基础目录，兼容打包后的exe"""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    else:
        return Path(__file__).parent


BASE_DIR = get_base_dir()


# 日志配置 - 延迟初始化
_logger = None

def get_logger():
    """获取日志器（延迟初始化）"""
    global _logger
    if _logger is not None:
        return _logger

    try:
        log_dir = BASE_DIR / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f"bot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        handlers = [
            logging.FileHandler(str(log_file), encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    except Exception as e:
        # 如果无法创建日志文件，只输出到控制台
        print(f"警告: 无法创建日志文件: {e}")
        handlers = [logging.StreamHandler(sys.stdout)]

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=handlers
    )
    _logger = logging.getLogger(__name__)
    return _logger


class AutomationBot:
    """自动化机器人核心类"""

    def __init__(self, config_path="config.yaml"):
        # 处理配置文件路径
        config_path = Path(config_path)
        if not config_path.is_absolute():
            config_path = BASE_DIR / config_path
        self.config = self.load_config(config_path)
        self.assets_dir = BASE_DIR / "assets"
        self.stats = {"success": 0, "failed": 0, "skipped": 0}

    def load_config(self, path):
        """加载配置文件"""
        with open(path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)

    def load_excel(self):
        """读取 Excel 数据"""
        excel_cfg = self.config['excel']
        raw_path = excel_cfg.get('file_path', '')

        # 验证路径不为空
        if not raw_path or not raw_path.strip():
            raise ValueError("Excel文件路径未配置，请在控制面板中选择Excel文件")

        file_path = Path(raw_path)

        # 处理相对路径
        if not file_path.is_absolute():
            file_path = BASE_DIR / file_path

        # 验证文件存在
        if not file_path.exists():
            raise FileNotFoundError(f"Excel文件不存在: {file_path}")

        if file_path.is_dir():
            raise ValueError(f"路径是目录而非文件: {file_path}")

        get_logger().info(f"读取 Excel: {file_path}")

        df = pd.read_excel(
            file_path,
            sheet_name=excel_cfg.get('sheet_name') or 0
        )

        code_col = excel_cfg['code_column']
        qty_col = excel_cfg['quantity_column']

        # 数据校验
        if code_col not in df.columns:
            raise ValueError(f"找不到列: {code_col}")
        if qty_col not in df.columns:
            raise ValueError(f"找不到列: {qty_col}")

        # 清洗数据
        df = df[[code_col, qty_col]].dropna()
        df[code_col] = df[code_col].astype(str).str.strip()
        df[qty_col] = pd.to_numeric(df[qty_col], errors='coerce').fillna(0).astype(int)

        get_logger().info(f"共读取 {len(df)} 条有效数据")
        return df.to_dict('records')

    def execute_action(self, step, data):
        """执行单个操作步骤"""
        action = step['action']
        name = step.get('name', action)

        get_logger().info(f"  执行: {name}")

        if action == 'click':
            return self._action_click(step)

        elif action == 'double_click':
            return self._action_click(step, double=True)

        elif action == 'type_text':
            return self._action_type(step, data)

        elif action == 'press_key':
            return self._action_press_key(step)

        elif action == 'wait':
            time.sleep(step.get('seconds', 1))
            return True

        elif action == 'clear_input':
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('delete')
            return True

        else:
            get_logger().warning(f"未知动作: {action}")
            return False

    def _action_click(self, step, double=False):
        """点击操作 - 使用坐标"""
        x = step.get('x')
        y = step.get('y')

        if x is None or y is None:
            get_logger().error("未设置坐标")
            return False

        if double:
            pyautogui.doubleClick(x, y)
        else:
            pyautogui.click(x, y)

        if wait := step.get('wait_after'):
            time.sleep(wait)
        return True

    def _action_type(self, step, data):
        """输入文本"""
        text = step.get('text', '')
        code_col = self.config['excel']['code_column']
        qty_col = self.config['excel']['quantity_column']

        # 替换占位符
        text = text.replace('{code}', str(data.get(code_col, '')))
        text = text.replace('{quantity}', str(data.get(qty_col, '')))

        if step.get('clear_first'):
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.1)

        # 使用剪贴板输入中文
        pyperclip.copy(text)
        pyautogui.hotkey('ctrl', 'v')

        if wait := step.get('wait_after'):
            time.sleep(wait)
        return True

    def _action_press_key(self, step):
        """按键操作"""
        key = step.get('key', '')
        if '+' in key:
            keys = key.split('+')
            pyautogui.hotkey(*keys)
        else:
            pyautogui.press(key)

        if wait := step.get('wait_after'):
            time.sleep(wait)
        return True

    def process_single_item(self, data, index):
        """处理单条数据"""
        code_col = self.config['excel']['code_column']
        qty_col = self.config['excel']['quantity_column']
        code = data[code_col]
        qty = data[qty_col]

        get_logger().info(f"[{index}] 处理: {code} -> {qty}")

        for step in self.config['steps']:
            success = self.execute_action(step, data)
            if not success:
                get_logger().error(f"步骤失败: {step.get('name')}")
                self.stats['failed'] += 1
                return False

        self.stats['success'] += 1
        get_logger().info(f"[{index}] 完成")
        return True

    def run(self, limit=0):
        """主运行方法"""
        logger = get_logger()
        logger.info("=" * 50)
        logger.info("库存自动化程序启动")
        logger.info("安全提示: 将鼠标移到屏幕左上角可紧急停止")
        logger.info("=" * 50)

        # 倒计时
        logger.info("3秒后开始，请切换到目标软件窗口...")
        for i in range(3, 0, -1):
            logger.info(f"  {i}...")
            time.sleep(1)

        # 读取数据
        data_list = self.load_excel()

        # 限制执行条数
        if limit > 0:
            data_list = data_list[:limit]

        # 逐条处理
        total = len(data_list)
        for i, data in enumerate(data_list, 1):
            try:
                self.process_single_item(data, f"{i}/{total}")
            except pyautogui.FailSafeException:
                logger.warning("检测到鼠标移至左上角，程序终止")
                break
            except Exception as e:
                logger.error(f"处理异常: {e}")
                self.stats['failed'] += 1

        # 统计
        logger.info("=" * 50)
        logger.info("运行结束")
        logger.info(f"成功: {self.stats['success']}")
        logger.info(f"失败: {self.stats['failed']}")
        logger.info("=" * 50)


def main():
    """程序入口"""
    print("\n" + "=" * 50)
    print("  库存批量修改自动化工具")
    print("=" * 50)
    print("\n使用前请确保:")
    print("  1. assets/ 文件夹中已放入按钮截图")
    print("  2. data/ 文件夹中已放入 Excel 文件")
    print("  3. config.yaml 已正确配置")
    print("\n按 Enter 开始，按 Ctrl+C 取消...")

    try:
        input()
        bot = AutomationBot()
        bot.run()
    except KeyboardInterrupt:
        print("\n用户取消")
    except Exception as e:
        get_logger().error(f"程序错误: {e}")
        raise


if __name__ == "__main__":
    main()
