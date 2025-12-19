import os
import subprocess
import sys
import psutil, logging, time
if getattr(sys, 'frozen', False):
    script_directory = os.path.dirname(sys.executable)
else:
    script_directory = os.path.dirname(os.path.abspath(__file__))
log_file = os.path.join(script_directory, "app.log")
prompt_file = os.path.join(script_directory, "prompts.txt")
image_folder = os.path.join(script_directory, "image")
logger = logging.getLogger("customLogger")
logger.setLevel(logging.INFO)
file_handler = logging.FileHandler(log_file, encoding="utf-8")
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S'))
logger.addHandler(file_handler)
def pwrite(message, p=False):
    message = f'{message}'
    logger.info(message)
    if p:
        print(message)
def is_chrome_running_with_port(port):
    """Check if Chrome is already running with the specified debugging port."""
    for proc in psutil.process_iter(['name', 'cmdline']):
        try:
            if proc.info['name'] and 'chrome' in proc.info['name'].lower():
                cmdline = proc.info['cmdline']
                if cmdline and any(f'--remote-debugging-port={port}' in arg for arg in cmdline):
                    return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False

def start_chrome_with_debugging():
    # Get script folder
    script_folder = os.path.dirname(os.path.abspath(__file__))
    
    # Find Chrome executable
    chrome_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ]
    
    chrome_path = None
    for path in chrome_paths:
        if os.path.exists(path):
            chrome_path = path
            break
    
    if not chrome_path:
        pwrite("Chrome executable not found.")
        input("Press Enter to exit...")
        sys.exit(1)
    
    # Set user data directory and port
    user_data_dir = os.path.join(script_folder, "chrome")
    custom_port = 9233
    
    # Check if Chrome is already running with this port
    if is_chrome_running_with_port(custom_port):
        pwrite(f"Chrome is already running with debugging port {custom_port}")
        pwrite("Not starting a new instance.")
        return
    
    # Create user data directory if it doesn't exist
    os.makedirs(user_data_dir, exist_ok=True)
    
    # Start Chrome
    cmd = [
        chrome_path,
        f"--remote-debugging-port={custom_port}",
        f"--user-data-dir={user_data_dir}",
        "--disable-popup-blocking"
    ]
    
    pwrite(f"Starting Chrome with debugging port {custom_port}...")
    subprocess.Popen(cmd)
    pwrite("Chrome started successfully.")

start_chrome_with_debugging()