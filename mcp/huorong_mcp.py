import os
import sys
import asyncio
import argparse
import pyautogui
import time
import subprocess
from datetime import datetime  # 新增导入 for 获取当前时间
import urllib.request
import io
from PIL import Image
from typing import Optional
import shutil  # 文件复制
import sqlite3 # 数据库

# --- 输出使用utf-8编码 --
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# --- 调试开关 ---
#DEBUG_MODE = '--debug' in sys.argv
def debug_print(message: str):
    """
        仅在调试模式下输出调试信息到标准错误流
    """
    if DEBUG_MODE:
        print(message, file=sys.stderr)

# --- 导入MCP ---
try:
    from mcp.server.fastmcp import FastMCP
except ImportError:
    debug_print("[致命错误] 请先运行: uv add \"mcp[cli]\" httpx")
    sys.exit(1)

mcp = FastMCP("huorong", log_level="ERROR",port = 8888)

# --- 用户配置 ---
# HUORONG_PATH = r"C:\Program Files (x86)\Huorong\Sysdiag\bin\HipsTray.exe" # 示例
# HUORONG_PATH = "D:/Programs/HuoRong/Sysdiag/bin/HipsMain.exe"  # 用户提供的路径

# 病毒查杀功能对应的图片
QUICK_SCAN_BUTTON_IMAGE = './tag_image/huorong/huorong_quick_scan_button.png'
PAUSE_BUTTON_IMAGE = './tag_image/huorong/huorong_pause_button.png'
SCAN_COMPLETE_IMAGE = './tag_image/huorong/huorong_scan_complete.png'

# 查看隔离区
QUARANTINE_BUTTON_IMAGE = './tag_image/huorong/huorong_quarantine_button.png'
MAXIMIZE_BUTTON_IMAGE = './tag_image/huorong/huorong_maximize_button.png'

# 获取安全日志
#MENU_ICON_IMAGE = './tag_image/huorong/huorong_menu_icon.png'
SECURITY_LOG_IMAGE = './tag_image/huorong/huorong_security_log1.png'
EXPORT_LOG_BUTTON_IMAGE = './tag_image/huorong/huorong_export_log_button1.png'
FILENAME_INPUT_BOX_IMAGE = 'tag_image/huorong/huorong_filename_input_box1.png'
SAVE_BUTTON_IMAGE = './tag_image/huorong/huorong_save_button1.png'

# --- 设备性能与等待时间 ---
DEVICE_LEVEL = 1  # 1: 低性能设备，2: 中性能设备，3: 高性能设备
SLEEP_TIME_SHORT = 1 * DEVICE_LEVEL
SLEEP_TIME_MEDIUM = 3 * DEVICE_LEVEL
SLEEP_TIME_LONG = 5 * DEVICE_LEVEL

# --- 辅助函数 ---
def find_image_on_screen(image_filename, confidence_level, timeout_seconds=15, description=""):
    """
        在屏幕上查找指定的图像。
    Args:
        image_filename: 图像文件名。
        confidence_level: 图像匹配的置信度（0-1）。
        timeout_seconds: 超时时间（秒）。
        description: 图像描述（用于调试）。
    Returns:
        (x, y) 坐标元组，如果未找到则返回 None。
    """
    if not description:
        description = image_filename

    debug_print(f"正在查找图像: '{description}' (文件: {image_filename})，时间限制为{timeout_seconds}秒...")
    start_time = time.time()
    while time.time() - start_time < timeout_seconds:
        try:
            location = pyautogui.locateCenterOnScreen(image_filename, confidence=confidence_level)
            if location:
                debug_print(f"找到 '{description}'，坐标: {location}")
                return location
            else:
                time.sleep(SLEEP_TIME_SHORT)
        except pyautogui.ImageNotFoundException:
            time.sleep(SLEEP_TIME_SHORT)
        except FileNotFoundError:
            debug_print(f"严重错误：图像文件 '{image_filename}' 未找到或无法访问！")
            return None
        except Exception as e:
            if "Failed to read" in str(e) or "cannot identify image file" in str(e):
                debug_print(f"错误: 无法读取或识别图像文件 '{image_filename}'。错误详情: {e}")
                return None
            debug_print(f"查找图像 '{description}' 时发生其他错误: {e}")
            return None

    debug_print(f"超时：在 {timeout_seconds} 秒内未能找到图像 '{description}'。")
    return None

def click_image_at_location(location, description=""):
    """
        点击指定屏幕坐标处的图像。
    Args:
        location: 屏幕坐标元组 (x, y)。
        description: 图像描述。
    Returns:
        bool: 如果成功点击图像，返回True；否则返回False。
    """
    if location:
        pyautogui.click(location)
        debug_print(f"成功点击 '{description}' 在坐标: {location}")
        return True
    else:
        debug_print(f"未能点击 '{description}'，因为未找到坐标。")
        return False

def find_and_click(image_filename, confidence_level, timeout_seconds, description):
    """
        查找屏幕上的图片并点击。若成功，则返回true，否则返回false。
    """
    img_loc = find_image_on_screen(image_filename = image_filename, confidence_level=0.8, timeout_seconds=15, description =description)
    if img_loc:
        click_image_at_location(img_loc, description)
        debug_print("点击{0}成功。".format(description))
        return True
    else:
        debug_print("未找到{0}，点击失败。".format(description))
        return False

def ret2_top_page():
    """
        执行完功能后，返回首页。
    """
    global COMPLETE_BUTTON_IMAGE 
    COMPLETE_BUTTON_IMAGE = './tag_image/huorong/huorong_complete_button.png'
    img_loc = find_image_on_screen(COMPLETE_BUTTON_IMAGE, confidence_level=0.8, timeout_seconds=15, description="完成按钮")
    if img_loc:
        click_image_at_location(img_loc, description="完成按钮")
        debug_print("已经返回首页。")

def read_QuarantineEx_db(db_path, file_path):
    # 连接数据库
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # 查看所有表名
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()
    table_name = tables[0][0]  #第一张表（不准确）

    # 查询fn和vn字段
    cursor.execute(f"SELECT fn, vn FROM {table_name};")
    rows = cursor.fetchall()

    # 写入到log文件
    with open(file_path, "w", encoding="utf-8") as f:
        for fn, vn in rows:
            f.write(f"文件名: {fn}, 病毒名: {vn}\n")

    # 关闭连接
    conn.close()

def read_wlfile_db(db_path, file_path):
    # 连接数据库
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # 查询fn字段
    cursor.execute(f"SELECT fn FROM TrustRegion_60;")
    rows = cursor.fetchall()

    # 写入到log文件
    with open(file_path, "w", encoding="utf-8") as f:
        for fn in rows:
            f.write(f"{fn}\n")

    # 关闭连接
    conn.close()

# --- MCP工具 --- （异步？）
@mcp.tool()
def start_huorong(path) :
    """
        启动火绒安全软件（简称火绒）。
    Args：
        path: 火绒安全软件的完整安装路径（即HUORONG_PATH）。
    """
    try:
        if not path or not os.path.exists(path):
            debug_print(f"错误：应用程序路径 '{path}' 无效或不存在。\n\
                        请确保 HUORONG_PATH 变量已正确设置为火绒安全的启动程序路径。\n\
                        脚本将尝试在不启动新进程的情况下继续（假设火绒已打开）。")
            return None
        debug_print(f"正在尝试启动火绒: {path}")
        app = subprocess.Popen(path)
        debug_print(f"应用程序已启动 (进程ID: {app.pid})。")
        
        return app
    except Exception as e:
        debug_print(f"启动应用程序 '{path}' 时发生错误: {e}")
        debug_print("脚本将尝试在不启动新进程的情况下继续（假设火绒已打开）。")
        return None
    
@mcp.tool()
def scan_virus():
    """
        执行火绒安全软件的快速查杀功能。
    Args：
        None
    """
    # 步骤1：打开火绒安全软件
    # 不足：必须在火绒的首页
    start_huorong(HUORONG_PATH)
    debug_print(f"火绒安全软件已启动，请确保火绒处于首页。")
    time.sleep(SLEEP_TIME_LONG)  # 等待应用程序加载
    # 步骤2：点击“快速查杀”按钮
    img_loc = find_image_on_screen(QUICK_SCAN_BUTTON_IMAGE, confidence_level=0.8, timeout_seconds=15, description="快速查杀按钮")
    if img_loc:
        click_image_at_location(img_loc, description="快速查杀按钮")
        debug_print("点击快速查杀按钮成功。")
    else:
        debug_print("点击快速查杀按钮失败。")
        return "点击快速查杀按钮失败。"
    time.sleep(SLEEP_TIME_LONG)
    # 步骤3：检测是否正在查杀
    if find_image_on_screen(PAUSE_BUTTON_IMAGE, confidence_level=0.8, timeout_seconds=15, description="暂停按钮"):
        debug_print("正在执行快速查杀。")
    else:
        debug_print("未找到暂停按钮，说明未成功执行快速查杀。")
        return "未找到暂停按钮，说明未成功执行快速查杀。"
    # 步骤4：检测查杀是否完成
    start_time = time.time()
    interval = 300 # 5分钟
    while time.time() - start_time < interval:
        img_loc = find_image_on_screen(SCAN_COMPLETE_IMAGE, confidence_level=0.8, timeout_seconds=15, description="快速查杀完成")
        if img_loc:
            debug_print(f"检测到查杀完成标志，坐标为: {img_loc}")
            return "快速查杀完成。"
        time.sleep(SLEEP_TIME_MEDIUM)
    # 步骤5：返回查杀结果
    # 待补充（OCR识别查杀结果界面、联动日志查询）
    # 步骤6：点击完成，返回首页
    #ret2_top_page()

@mcp.tool()
def get_quarantine_file():
    """
        执行火绒的查看隔离区功能，具体为获取当前隔离区内的文件列表。
    Args:
        None
    """
    # 方法1：图像识别
    # 1：打开火绒安全软件
    # start_huorong(HUORONG_PATH)
    # time.sleep(SLEEP_TIME_LONG)  
    # 2：打开隔离区
    # find_and_click(QUARANTINE_BUTTON_IMAGE,"隔离区按钮")
    # time.sleep(SLEEP_TIME_SHORT)
    # find_and_click(MAXIMIZE_BUTTON_IMAGE,"最大化窗口按钮")
    # time.sleep(SLEEP_TIME_SHORT)
    # 3：获取文件列表

    # 方法2：读取数据库中信息
    # 1：复制文件到当前目录下
    source_file_path = r'C:/ProgramData/Huorong/Sysdiag/QuarantineEx.db'  
    target_dir = r'./'  # 目标目录
    shutil.copy(source_file_path, target_dir) # 复制文件到目标目录
    # 2：读取表内容
    file_path = "quarantine_files.log"
    read_QuarantineEx_db('./QuarantineEx.db',file_path)
    # 3：删除复制的数据库文件
    os.remove(os.path.join(target_dir, os.path.basename(source_file_path)))
    return f"已获取当前隔离区内的文件列表，见{file_path}"

# --- 主函数 ---
def main():
    """
        根据是否处于调试模式，执行不同的操作。
    """
    # 获取参数
    global HUORONG_PATH,DEBUG_MODE,QUARANTINE_DB_PATH
    parser = argparse.ArgumentParser(description="火绒MCP工具")
    parser.add_argument('--huorong-path', type=str, required=True, help='火绒安全软件的完整路径')
    parser.add_argument('--debug', action='store_true', help='调试模式')
    #parser.add_argument('--quarantine', type=str, help='隔离区数据库路径')
    args = parser.parse_args()
    HUORONG_PATH = args.huorong_path
    DEBUG_MODE = args.debug
    #QUARANTINE_DB_PATH = args.quarantine

    # 模式判断
    if DEBUG_MODE:
        print("--- 处于调试模式 ---")

        def main_test():
            print("[测试开始] 打开火绒安全软件，进行快速查杀...")
            result = scan_virus()
            # print("\n[测试成功] 返回结果：")
            print(result)

        main_test()
    else:
        mcp.run(transport='stdio')

# --- 主程序入口 ---
if __name__ == "__main__":
    main()
