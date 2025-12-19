from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import psutil
import traceback
import time,os,sys, logging, glob, traceback, subprocess
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

if getattr(sys, 'frozen', False):
    script_directory = os.path.dirname(sys.executable)
else:
    script_directory = os.path.dirname(os.path.abspath(__file__))
log_file = os.path.join(script_directory, "app.log")
prompt_file = os.path.join(script_directory, "prompts.txt")
reports_file = os.path.join(script_directory, "reports.txt")
leads_file = os.path.join(script_directory, "TestVideoask1.csv")
if not os.path.exists(reports_file):
    with open(reports_file, "w"):
        pass
    
logger = logging.getLogger("customLogger")
logger.setLevel(logging.INFO)
file_handler = logging.FileHandler(log_file, encoding="utf-8")
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S'))
logger.addHandler(file_handler)

def pwrite(message, p=True):
    et = time.time() 
    message = f'{message}'
    logger.info(message)
    if p:
        print(message)

user_data_dir = os.path.join(script_directory, "chrome")
os.environ['WDM_LOG_LEVEL'] = '0'

def kill_processes_using_dir(user_data_dir):
    user_data_dir = os.path.abspath(user_data_dir)

    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            cmdline = proc.info['cmdline']
            if cmdline and any(user_data_dir in arg for arg in cmdline):
                proc.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            cmdline = proc.info['cmdline']
            if not cmdline:
                continue

            cmdline_str = " ".join(cmdline).lower()

            if "--user-data-dir=" in cmdline_str and user_data_dir in cmdline_str:
                proc.kill()

        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

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
    script_folder = os.path.dirname(os.path.abspath(__file__))
    
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
    
    user_data_dir = os.path.join(script_folder, "chrome")
    custom_port = 9233
    
    if is_chrome_running_with_port(custom_port):
        return
    
    os.makedirs(user_data_dir, exist_ok=True)
    
    cmd = [
        chrome_path,
        f"--remote-debugging-port={custom_port}",
        f"--user-data-dir={user_data_dir}",
        "--disable-popup-blocking"
    ]
    
    pwrite(f"Starting Chrome with debugging port {custom_port}...")
    subprocess.Popen(cmd)
    time.sleep(3)  # Give Chrome time to start
    pwrite("Chrome started successfully.")

start_chrome_with_debugging()

def create_driver():
    options = webdriver.ChromeOptions()
    options.add_experimental_option('debuggerAddress', 'localhost:9233')
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    pwrite(f"Connected to Chrome: {driver.title}")
    return driver

driver = create_driver()
driver.get('https://app.videoask.com/app/organizations/0790e150-965d-4181-b745-7ab50d376715/respondents')
time.sleep(3)

def make_click(xpathe, t=10, sleep_time=None):
    WebDriverWait(driver, t).until(EC.presence_of_element_located((By.XPATH, xpathe)))
    if sleep_time is not None:
            time.sleep(sleep_time)
    element =WebDriverWait(driver, t).until(EC.element_to_be_clickable((By.XPATH, xpathe)))
    try:
        try:
            element.click()            
        except:                
            element.send_keys(Keys.ENTER)
    except:
        try:
            driver.execute_script("arguments[0].scrollIntoView({ behavior: 'auto', block: 'center', inline: 'center' });", element)
            time.sleep(1)
            element.click()
        except:
            driver.execute_script("arguments[0].click();", element)


def write_t(xpathe, text, t=10, sleep_time=None):
    
        WebDriverWait(driver, t).until(EC.presence_of_element_located((By.XPATH, xpathe)))
        
        if sleep_time is not None:
            time.sleep(sleep_time)
        
        element = WebDriverWait(driver, t).until(EC.presence_of_element_located((By.XPATH, xpathe)))
        time.sleep(0.5)
        
        if len(text) < 8:
            for char in text:
                element.send_keys(char)
                time.sleep(0.3)
        else:
            element.send_keys(text)


def extract_text(xpathe, t=10):
        element_text = WebDriverWait(driver, t).until(EC.presence_of_element_located((By.XPATH, xpathe))).text
        return element_text.strip()
    

def close_extra_tabs():
    try:
        tabs = driver.window_handles
        for tab in tabs[1:]:
            driver.switch_to.window(tab)
            driver.close()
        driver.switch_to.window(tabs[0])
    except Exception as e:
        pwrite(f"Error closing extra tabs: {e}")

try:
    while True:
        with open(reports_file, "r") as record:
            records = set(line.strip() for line in record.readlines())
        
        try:
            # Wait for contacts to load
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//*[contains(@class, "list-item-subtitle__CardSubtitle")]'))
            )
            
            contacts = driver.find_elements(By.XPATH, '//*[contains(@class, "list-item-subtitle__CardSubtitle")]')
            pwrite(f"Found {len(contacts)} contacts on page")
            
            if not contacts:
                pwrite("No more contacts found. Exiting...")
                break
            
            for contact in contacts:
                try:
                    try:
                        email = contact.text.strip()
                    except:
                        driver.execute_script("arguments[0].scrollIntoView({ behavior: 'auto', block: 'center', inline: 'center' });", contact)
                        email = contact.text.strip()
                    
                    if not email:
                        continue
                    
                    if email in records:
                        pwrite(f"Skipping existing contact: {email}")
                        continue
                    
                    pwrite(f"Processing contact: {email}")
                    
                    try:
                        # Open new tab
                        driver.execute_script("window.open('https://app.videoask.com/app/organizations/0790e150-965d-4181-b745-7ab50d376715/respondents', '_blank');")
                        time.sleep(2)
                        driver.switch_to.window(driver.window_handles[-1])
                        time.sleep(2)
                        
                        # Search for contact
                        write_t('//*[@placeholder="Find a contact"]', email, sleep_time=4)
                        make_click(f'//*[contains(@class,"ListItem")]//*[contains(text(), "{email}")]', sleep_time=2)
                        
                        # Check if "fall prevention" already exists
                        try:
                            WebDriverWait(driver, 3).until(
                                EC.presence_of_element_located((By.XPATH, '//*[contains(@class, "list-item-title__CardTitle") and text()="fall prevention"]'))
                            )
                            pwrite(f"'fall prevention' already exists for {email}")
                            driver.close()
                            driver.switch_to.window(driver.window_handles[0])
                            
                            with open(reports_file, "a") as record:
                                record.write(str(email) + "\n")

                            continue
                        except TimeoutException:
                            pwrite(f"Creating new VideoAsk for {email}")
                        
                        # Create VideoAsk
                        make_click('//*[contains(@aria-label, "Start a new video interaction with")]', sleep_time=2)
                        make_click('//*[@aria-label="Pick a video from the library"]', sleep_time=2)
                        make_click('(//*[@data-testid="media-thumbnail"])[1]', sleep_time=2)
                        make_click('//*[@aria-label="YES"]', sleep_time=2)
                        write_t('//*[@aria-label="title"]', "fall prevention", sleep_time=2)
                        make_click('//*[@aria-label="Fit video"]', sleep_time=2)
                        make_click('//button//*[text()="Button"]', sleep_time=2)
                        write_t('//*[@aria-label="button label"]', 
                                "Learn more about Mann Materials Anti-Fall Flooring and how it can protect seniors during falls", 
                                sleep_time=2)
                        write_t('//*[@aria-label="button link"]', "https://mannmaterials.com", sleep_time=2)
                        make_click('//button[@type="submit"]', sleep_time=3)
                        
                        pwrite(f"Successfully created VideoAsk for {email}")
                        time.sleep(3)
                        
                        # Close tab and return to main
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                        
                        # Save to records
                        with open(reports_file, "a") as record:
                            record.write(str(email) + "\n")
                    except:
                        pwrite(traceback.format_exc(), p=False)
                        continue
                        pass

                    
                except Exception as e:
                    pwrite(f" Error processing contact {email}: {traceback.format_exc()}")
                    pwrite(traceback.format_exc())
                    
                    # Clean up tabs
                    close_extra_tabs()
                    time.sleep(2)
            
            # Scroll to load more contacts
            try:
                last_contact = driver.find_element(By.XPATH, "(//*[contains(@class, \"list-item-subtitle__CardSubtitle\")])[last()]")
                driver.execute_script("arguments[0].scrollIntoView(true);", last_contact)
                time.sleep(3)

            except NoSuchElementException:
                pwrite("No more contacts to scroll to")
                break
                
        except TimeoutException:
            pwrite(" Timeout waiting for contacts. Breaking loop...")
            break
        except Exception as e:
            pwrite(f"Error in main loop: {e}")
            pwrite(traceback.format_exc())
            break

except KeyboardInterrupt:
    pwrite("Script interrupted by user")
except Exception as e:
    pwrite(traceback.format_exc())
finally:
    pwrite("Script completed. Browser will remain open.")
    driver.quit()