import os
import sys
import ctypes
import shutil
import msvcrt
import time
import requests
import webbrowser
from ctypes import wintypes
from win32com.client import Dispatch

try:
    import vc
except ImportError:
    vc = None

#  Unicode 路径获取
def get_unicode_shell_path(csidl):
    buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
    ctypes.windll.shell32.SHGetSpecialFolderPathW(None, buf, csidl, False)
    return buf.value

CURRENT_VERSION = "1.0.2.1"
VERSION_URL = "https://ab1f5a.net/v1/version.json"
WINDOW_TITLE = f"ACLOS Tools v{CURRENT_VERSION}"

UNICODE_APPDATA = get_unicode_shell_path(0x001a)
DEPLOY_DIR = os.path.join(UNICODE_APPDATA, "ACLOS Tools")
PATCHER_EXE_NAME = "ACLOS UPDATE PATCHER.EXE"

G = "\033[1;32m"; Y = "\033[1;33m"; C = "\033[1;36m"; R = "\033[1;31m"
M = "\033[1;35m"; W = "\033[1;37m"; RESET = "\033[0m"

def set_title(title):
    try: ctypes.windll.kernel32.SetConsoleTitleW(title)
    except: pass

def check_version_sync():
    os.system('cls')
    print(f"\n {W}[{C}*{W}]{RESET} 正在初始化组件并校验版本...")
    try:
        res = requests.get(VERSION_URL, timeout=5)
        if res.status_code == 200:
            remote_data = res.json()
            remote_v = remote_data.get("version", "")
            update_log = remote_data.get("log", "无详细说明")
            if remote_v != CURRENT_VERSION:
                box_width = 50
                def get_display_width(text):
                    width = 0
                    for char in text:
                        if '\u4e00' <= char <= '\u9fff' or char in '，。！？（）《》【】':
                            width += 2
                        else:
                            width += 1
                    return width
                def split_text(text, max_width):
                    lines = []; current_line = ""; current_width = 0
                    for char in text:
                        char_width = 2 if '\u4e00' <= char <= '\u9fff' or char in '，。！？（）《》【】' else 1
                        if current_width + char_width > max_width:
                            lines.append(current_line); current_line = char; current_width = char_width
                        else:
                            current_line += char; current_width += char_width
                    lines.append(current_line)
                    return lines
                def print_bordered_line(text, color=W):
                    display_w = get_display_width(text)
                    padding = " " * (box_width - display_w)
                    print(f"    {R}│{RESET} {color}{text}{RESET}{padding} {R}│{RESET}")
                print(f"\n\n    {R}┌────────────────────────────────────────────────────┐{RESET}")
                print(f"    {R}│{RESET}           {R}        检测到有新版本        {RESET}           {R}│{RESET}")
                print(f"    {R}├────────────────────────────────────────────────────┤{RESET}")
                print_bordered_line(f"本地版本: {CURRENT_VERSION}", Y)
                print_bordered_line(f"最新版本: {remote_v}", G)
                log_lines = split_text(update_log, box_width - 12) 
                for i, line in enumerate(log_lines):
                    if i == 0: print_bordered_line(f"更新日志: {line}", C)
                    else: print_bordered_line(f"          {line}", C)
                print(f"    {R}└────────────────────────────────────────────────────┘{RESET}")
                webbrowser.open("https://github.com/ab1f5a/ACLOS-UPDATE-PACHER/releases/")
                print(f"\n    {W}[{G}#{W}]{RESET} 浏览器已弹出下载页面。")
                print(f"    {W}[{C}@{W}]{RESET} 请按 {G}任意键{RESET} 退出程序并手动安装新版本...")
                msvcrt.getch()
                os._exit(0)
            else:
                print(f" {W}[{G}OK{W}]{RESET} 核心版本校验通过。")
                time.sleep(0.8)
        else:
            raise Exception()
    except:
        print(f"\n\n    {R}┌──────────────────────────────────────────────────┐{RESET}")
        print(f"    {R}│{RESET}            {R}致命错误: 无法验证软件版本{RESET}            {R}│{RESET}")
        print(f"    {R}├──────────────────────────────────────────────────┤{RESET}")
        print(f"    {R}│{RESET} {W}错误说明: 无法连接服务器, 请检查网络。    {RESET}       {R}│{RESET}")
        print(f"    {R}└──────────────────────────────────────────────────┘{RESET}")
        print(f"\n    {W}[{C}@{W}]{RESET} 请按 {G}任意键{RESET} 退出程序...")
        msvcrt.getch()
        os._exit(0)

def deploy_patcher():
    os.system('cls')
    print(f"\n\n {W}[{G}#{W}]{RESET} {W}启动部署程序...{RESET}")
    print(f" {M} ──────────────────────────────────{RESET}")
    
    try:
        if not os.path.exists(DEPLOY_DIR): 
            os.makedirs(DEPLOY_DIR, exist_ok=True)
    except:
        print(f"\n {W}[{R}!{W}]{RESET} {R}权限错误:{RESET} 无法创建部署目录")
        msvcrt.getch(); return

    src = os.path.join(sys._MEIPASS if hasattr(sys, '_MEIPASS') else os.path.abspath("."), PATCHER_EXE_NAME)
    dst = os.path.join(DEPLOY_DIR, PATCHER_EXE_NAME)
    
    try:
        shutil.copy2(src, dst)
        
        desktop = get_unicode_shell_path(0x0000)
        shortcut_target = os.path.join(desktop, "ACLOS2.13.1.lnk")
        
        if os.path.exists(shortcut_target):
            try:
                os.chmod(shortcut_target, 0o777)
                os.remove(shortcut_target)
            except: pass
            
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(str(shortcut_target))
        shortcut.TargetPath = str(dst)
        shortcut.WorkingDirectory = str(DEPLOY_DIR)
        shortcut.IconLocation = str(dst)
        shortcut.save()
        
        print(f"\n {W}[{G}+{W}]{RESET} {G}部署成功:{RESET} 桌面入口已更新。")
    except Exception as e:
        print(f"\n {W}[{R}!{W}]{RESET} {R}部署失败:{RESET} {e}")
        if "2147352567" in str(e) or "2147024893" in str(e):
            print(f"    {Y}提示: 桌面路径可能被锁定或杀毒软件拦截。{RESET}")
            print(f"    {Y}请尝试暂时关闭火绒/360或检查OneDrive同步状态。{RESET}")
    
    print(f"\n\n {W}[{M}ENTER{W}]{RESET} 点击回车键返回菜单...")
    while True:
        char = msvcrt.getch()
        if char in [b'\r', b'\n']:
            break

def print_home():
    os.system('cls')
    print("\n")
    print(rf"""{C}
 ______  ____     __       _____   ____        ______                ___             
/\  _  \/\  _`\  /\ \     /\  __`\/\  _`\     /\__  _\              /\_ \            
\ \ \L\ \ \ \/\_\\ \ \    \ \ \/\ \ \,\L\_\   \/_/\ \/   ___     ___\//\ \     ____  
 \ \  __ \ \ \/_/_\ \ \  __\ \ \ \ \/_\__ \      \ \ \  / __`\  / __`\\ \ \   /',__\ 
  \ \ \/\ \ \ \L\ \\ \ \L\ \\ \ \_\ \/\ \L\ \     \ \ \/\ \L\ \/\ \L\ \\_\ \_/\__, `\
   \ \_\ \_\ \____/ \ \____/ \ \_____\ `\____\     \ \_\ \____/\ \____//\____\/\____/
    \/_/\/_/\/___/   \/___/   \/_____/\/_____/      \/_/\/___/  \/___/ \/____/\/___/ 
                                                                                                                                            
         {RESET}""")
    print(f"\n    {M}──────────────────────────────────────────────────────────────────────────────────{RESET}")
    print(f"    {W} 系统状态:{RESET} {G}就绪{RESET}   {M}│{RESET}   {W}当前版本:{RESET} {Y}{CURRENT_VERSION}{RESET}   {M}│{RESET}   {W}开发维护:{RESET} {C}ab1f5a{RESET}")
    print(f"    {M}──────────────────────────────────────────────────────────────────────────────────{RESET}")
    print("\n")
    print(f"      {W}[{G} 1 {W}]{RESET}   {W}一键部署绕过补丁{RESET}")
    print("\n")
    print(f"      {W}[{G} 2 {W}]{RESET}   {W}无畏契约配音修改{RESET}")
    print("\n")
    print(f"      {W}[{R} Q {W}]{RESET}   {W}退出工具箱系统{RESET}")
    print("\n")
    print(f"    {M}──────────────────────────────────────────────────────────────────────────────────{RESET}\n")

def main():
    if not ctypes.windll.shell32.IsUserAnAdmin():
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        return
    set_title(WINDOW_TITLE)
    check_version_sync()
    while True:
        print_home()
        char = msvcrt.getch()
        try: cmd = char.decode('utf-8').upper()
        except: continue
        if cmd == '1': deploy_patcher()
        elif cmd == '2':
            if vc: vc.main(); set_title(WINDOW_TITLE)
            else: print(f"\n   {W}[{R}!{W}]{RESET} 错误: vc.py 模块未就绪。"); time.sleep(2)
        elif cmd == 'Q': break

if __name__ == "__main__":
    main()