import os
import time
import sys
import ctypes
import string
import shutil
import msvcrt 
from tkinter import filedialog, Tk

# Config
CURRENT_VERSION = "1.0"
APP_NAME = "Voice Changer"
WINDOW_TITLE = f"无畏契约配音修改器 v{CURRENT_VERSION}"


G = "\033[1;32m"; Y = "\033[1;33m"; C = "\033[1;36m"; R = "\033[1;31m"
M = "\033[1;35m"; W = "\033[1;37m"; RESET = "\033[0m"

ROOT_DIR_NAME = "ACLOS"
SUB_PATH = os.path.join("Launcher", "resources", "app.asar")
CONFIG_DIR = os.path.join(os.environ['APPDATA'], "ACLOS")
PATH_CACHE = os.path.join(CONFIG_DIR, "asar_path_v2.cfg")

LANG_MAP = {
    "1": ("zh_CN", "中文_中国"),
    "2": ("en_US", "英语_美国"),
    "3": ("de_DE", "*德语_德国"),
    "4": ("it_IT", "意语_意大利"),
    "5": ("ja_JP", "日语_日本"),
    "6": ("ko_KR", "韩语_韩国"),
    "7": ("pt_PT", "*葡萄牙语_葡萄牙"),
    "8": ("ru_RU", "*俄语_俄国")
}

def set_window_title(title):
    try: ctypes.windll.kernel32.SetConsoleTitleW(title)
    except: pass

def is_admin():
    try: return ctypes.windll.shell32.IsUserAnAdmin()
    except: return False

def log_status(symbol, color, message):
    print(f" {W}[{color}{symbol}{W}]{RESET} {W}{message}{RESET}")

def print_banner():
    os.system('cls' if os.name == 'nt' else 'clear')
    print("\n") 
    print(rf"""{C}
 __  __                                 ____     __                                               
/\ \/\ \          __                   /\  _`\  /\ \                                              
\ \ \ \ \    ___ /\_\    ___     __    \ \ \/\_\\ \ \___      __      ___      __      __   _ __  
 \ \ \ \ \  / __`\/\ \  /'___\ /'__`\   \ \ \/_/_\ \  _ `\  /'__`\  /' _ `\  /'_ `\  /'__`\/\`'__\
  \ \ \_/ \/\ \L\ \ \ \/\ \__//\  __/    \ \ \L\ \\ \ \ \ \/\ \L\.\_/\ \/\ \/\ \L\ \/\  __/\ \ \/ 
   \ `\___/\ \____/\ \_\ \____\ \____\    \ \____/ \ \_\ \_\ \__/.\_\ \_\ \_\ \____ \ \____\\ \_\ 
    `\/__/  \/___/  \/_/\/____/\/____/     \/___/   \/_/\/_/\/__/\/_/\/_/\/_/\/___L\ \/____/ \/_/ 
                                                                               /\____/            
                                                                               \_/__/             {RESET}""")
    
    print(f"\n    {M}──────────────────────────────────────────────────────────────────────────────────{RESET}")
    print(f"    {W} 模块状态:{RESET} {G}挂载中{RESET}  {M}│{RESET}  {W}内核版本:{RESET} {Y}{CURRENT_VERSION}{RESET}  {M}│{RESET}  {W}作者:{RESET} {C}ab1f5a{RESET}")
    print(f"    {M}──────────────────────────────────────────────────────────────────────────────────{RESET}")

def press_any_key_back():
    print("\n")
    log_status("ENTER", M, "点击回车键返回工具箱主菜单...")
    msvcrt.getch()

def find_asar_optimized():
    if not os.path.exists(CONFIG_DIR): os.makedirs(CONFIG_DIR, exist_ok=True)
    if os.path.exists(PATH_CACHE):
        with open(PATH_CACHE, 'r', encoding='utf-8') as f:
            path = f.read().strip()
            if os.path.exists(path): return path

    print("\n")
    log_status("*", C, f"正在自动检索 {ROOT_DIR_NAME} 根目录...")
    drives = [f"{d}:\\" for d in string.ascii_uppercase if os.path.exists(f"{d}:\\")]
    
    for drive in drives:
        search_targets = [os.path.join(drive, ROOT_DIR_NAME)]
        try:
            for item in os.listdir(drive):
                if os.path.isdir(os.path.join(drive, item)):
                    search_targets.append(os.path.join(drive, item, ROOT_DIR_NAME))
        except: continue

        for root_path in search_targets:
            if os.path.exists(root_path):
                full_asar_path = os.path.join(root_path, SUB_PATH)
                if os.path.exists(full_asar_path):
                    log_status("OK", G, f"资源定位成功: {ROOT_DIR_NAME}\{SUB_PATH}")
                    return full_asar_path

    log_status("!", Y, "检索失败，请在弹窗中手动指定 app.asar 文件")
    root = Tk(); root.withdraw(); root.attributes("-topmost", True)
    path = filedialog.askopenfilename(title="选择 app.asar", filetypes=[("ASAR File", "app.asar")])
    root.destroy()
    return path if path else None

def patch_process(file_path):
    bak_path = file_path + ".bak"
    if os.path.exists(bak_path):
        log_status("#", G, "发现备份文件，将从核心镜像 (.bak) 还原注入。")
        source_for_reading = bak_path
    else:
        log_status("!", Y, "未发现备份，正在创建原始镜像备份...")
        source_for_reading = file_path

    try:
        with open(source_for_reading, 'rb') as f: raw_content = f.read()
        if source_for_reading == file_path:
            shutil.copy2(file_path, bak_path)
            log_status("OK", G, "初始镜像备份已就绪。")

        print("\n")
        print(f"    {C}┌── 目标语言选择 ─────────────────────────────────────────┐{RESET}")
        for k, v in LANG_MAP.items():
            display_text = f" {k}. {v[0]} ({v[1]})"
            actual_w = sum(2 if ord(c) > 127 else 1 for c in display_text)
            padding = " " * (54 - actual_w)
            print(f"    {C}│{RESET}  {G}{k}{RESET}. {W}{v[0]}{RESET} ({v[1]}){padding}  {C}│{RESET}")
        print(f"    {C}└─────────────────────────────────────────────────────────┘{RESET}")
        
        print("\n")
        choice = input(f"    {M}# 请输入语言代码序号（*标语言可能无法正常使用）： >> {RESET}").strip()
        if choice not in LANG_MAP:
            log_status("FAIL", R, "指令无效，任务中止。")
            return False
        
        target_lang = LANG_MAP[choice][0].encode('utf-8')
        log_status("*", C, f"正在注入序列码: {LANG_MAP[choice][0]}")

        if b"zh_CN" not in raw_content:
            log_status("FAIL", R, "数据源校验失败，zh_CN 标识缺失。")
            return False

        final_content = raw_content.replace(b"zh_CN", target_lang)
        with open(file_path, 'wb') as f: f.write(final_content)
        
        print("\n")
        log_status("DONE", G, f"文件已成功重构: {os.path.basename(file_path)}")
        return True
    except Exception as e:
        log_status("ERR", R, f"运行时异常: {e}")
        return False

def main():
    set_window_title(WINDOW_TITLE)
    if not is_admin():
        log_status("!", Y, "权限不足，正在请求管理员提权...")
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        return

    print_banner()
    target_path = find_asar_optimized()
    
    if not target_path:
        log_status("FAIL", R, "核心路径缺失，无法继续。")
        press_any_key_back()
        return

    with open(PATH_CACHE, 'w', encoding='utf-8') as f: f.write(target_path)
    
    if patch_process(target_path):
        print(f"\n    {M}──────────────────────────────────────────────────────────────────────────────────{RESET}")
        log_status("SUCCESS", G, "配音修改已生效！请重启游戏客户端。")
    
    press_any_key_back()

if __name__ == "__main__":
    main()