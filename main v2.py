from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import psutil
import traceback
import time,os,sys, logging, glob, traceback, subprocess, csv, random
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import configparser
if getattr(sys, 'frozen', False):
    script_directory = os.path.dirname(sys.executable)
else:
    script_directory = os.path.dirname(os.path.abspath(__file__))
log_file = os.path.join(script_directory, "app.log")
prompt_file = os.path.join(script_directory, "prompts.txt")
reports_file = os.path.join(script_directory, "reports.txt")
#leads_file = os.path.join(script_directory, "TestVideoask1.csv")
config_file = os.path.join(script_directory, "config.ini")
config = configparser.ConfigParser()
config.read(config_file)
login = config.get('auth', 'login')
password = config.get('auth', 'password')
button_text = config.get('auth', 'button_text')
button_url = config.get('auth', 'button_url')
leads_file = config.get('auth', 'contacts_file')
if '.csv' not in leads_file:
    leads_file += '.csv'
print(f"Using contacts file: {leads_file}")
skip_if_message_exist = config.get('auth', 'skip_if_message_exists').lower()
title = config.get('auth', 'title')
speed = config.get('auth', 'speed').lower()
leads_file_exist = False
if not os.path.exists(leads_file):
    leads_file_exist = False
elif not os.path.isfile(leads_file):
    leads_file_exist = False
elif not os.access(leads_file, os.R_OK):
    leads_file_exist = False
else:
    leads_file_exist = True
def speed_to_time(speed=speed):
    if speed == 'slow':
        time.sleep(random.uniform(3, 6))
        return
    elif speed == 'normal':
        time.sleep(random.uniform(1, 4))
        return 3
    elif speed == 'fast':
        time.sleep(random.uniform(0.5, 2))
        return
    else:
        return 3  # default to normal if unrecognized
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
def create_driver():
    options = webdriver.ChromeOptions()    
    options.add_argument('--no-sandbox')  # Disable the sandbox
    prefs = {"credentials_enable_service": False,
        "profile.password_manager_enabled": False,
        "profile.default_content_setting_values.media_stream_camera": 1,
    "profile.default_content_setting_values.media_stream_mic": 1,
    "profile.default_content_setting_values.geolocation": 1,
    "profile.default_content_setting_values.notifications": 1,}
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--use-fake-ui-for-media-stream")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-infobars")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("disable-infobars")
    options.add_argument('--log-level=3')
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    pwrite(f"Connected to Chrome: {driver.title}")
    return driver

driver = create_driver()


def make_click(xpathe, t=10, sleep_time=None):
    WebDriverWait(driver, t).until(EC.presence_of_element_located((By.XPATH, xpathe)))
    speed_to_time()
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
        
        speed_to_time()
        
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
driver.get('https://auth.videoask.com/login')
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="email"]'))).clear()
    time.sleep(1)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="password"]'))).clear()
    time.sleep(1)
    write_t('//*[@name="email"]', login)
    write_t('//*[@name="password"]', password)
    make_click('//*[@type="submit"]')
except:
    pass
try:
    make_click("//a[contains(@href, 'respondents')]")
except:
    href = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href, 'respondents')]"))).get_attribute("href")
    driver.get(href)
if leads_file_exist:
    try:
        with open(leads_file, newline='', encoding='utf-8', errors='ignore') as f:
            reader = csv.DictReader(f)  # uses headers automatically
            
            for i, row in enumerate(reader, start=1):
                email = row.get("Email")
                with open(reports_file, "r") as record:
                    records = set(line.strip() for line in record.readlines())
                    try:
                        
                        if email in records:
                            pwrite(f"Skipping existing contact: {email}")
                            continue
                        
                        pwrite(f"Processing contact: {email}")
                        
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@placeholder="Find a contact"]'))).send_keys(Keys.CONTROL + "a")
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@placeholder="Find a contact"]'))).send_keys(Keys.BACKSPACE)
                            write_t('//*[@placeholder="Find a contact"]', email, sleep_time=4)
                            make_click(f'//*[contains(@class,"ListItem")]//*[contains(text(), "{email}")]')
                            if skip_if_message_exist == 'yes':
                                try:
                                    WebDriverWait(driver, 3).until(
                                        EC.presence_of_element_located((By.XPATH, f'//*[contains(@class, "list-item-title__CardTitle") and text()="{title}"]'))
                                    )
                                    pwrite(f"'{title}' already exists for {email}")
                                    driver.close()
                                    driver.switch_to.window(driver.window_handles[0])
                                    
                                    with open(reports_file, "a") as record:
                                        record.write(str(email) + "\n")

                                    continue
                                except TimeoutException:
                                    pwrite(f"Creating new VideoAsk for {email}")
                            
                            # Create VideoAsk
                            make_click('//*[contains(@aria-label, "Start a new video interaction with")]')
                            make_click('//*[@aria-label="Pick a video from the library"]')
                            make_click('(//*[@data-testid="media-thumbnail"])[1]')
                            make_click('//*[@aria-label="YES"]')
                            write_t('//*[@aria-label="title"]', title)
                            make_click('//*[@aria-label="Fit video"]')
                            make_click('//button//*[text()="Button"]')
                            write_t('//*[@aria-label="button label"]', button_text)
                            write_t('//*[@aria-label="button link"]', button_url)
                            make_click('//button[@type="submit"]')
                            
                            pwrite(f"Successfully created VideoAsk for {email}")
                            speed_to_time()
                            
                            
                            with open(reports_file, "a") as record:
                                record.write(str(email) + "\n")
                            try:
                                make_click("//a[contains(@href, 'respondents')]")
                            except:
                                href = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href, 'respondents')]"))).get_attribute("href")
                                driver.get(href)
                        except:
                            pwrite(traceback.format_exc(), p=False)
                            continue
                            pass

                        
                    except Exception as e:
                        pwrite(f" Error processing contact {email}: {traceback.format_exc()}")
                        pwrite(traceback.format_exc())
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="email"]'))).send_keys(Keys.CONTROL + "a")
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="email"]'))).send_keys(Keys.BACKSPACE)
                            time.sleep(1)
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="password"]'))).send_keys(Keys.CONTROL + "a")
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="password"]'))).send_keys(Keys.BACKSPACE)
                            time.sleep(1)
                            write_t('//*[@name="email"]', login)
                            write_t('//*[@name="password"]', password)
                            make_click('//*[@type="submit"]')
                        except:
                            pass
                        # Clean up tabs
                        close_extra_tabs()
                        time.sleep(2)
            
                    
            

    except KeyboardInterrupt:
        pwrite("Script interrupted by user")
    except Exception as e:
        pwrite(traceback.format_exc())
    finally:
        pwrite("Script completed. Browser will remain open.")
        driver.quit()
else:
    
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
                            speed_to_time()
                            driver.switch_to.window(driver.window_handles[-1])
                            speed_to_time()
                            
                            # Search for contact
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@placeholder="Find a contact"]'))).send_keys(Keys.CONTROL + "a")
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@placeholder="Find a contact"]'))).send_keys(Keys.BACKSPACE)
                            write_t('//*[@placeholder="Find a contact"]', email, sleep_time=4)
                            make_click(f'//*[contains(@class,"ListItem")]//*[contains(text(), "{email}")]')
                            if skip_if_message_exist == 'yes':
                                try:
                                    WebDriverWait(driver, 3).until(
                                        EC.presence_of_element_located((By.XPATH, f'//*[contains(@class, "list-item-title__CardTitle") and text()="{title}"]'))
                                    )
                                    pwrite(f"'{title}' already exists for {email}")
                                    driver.close()
                                    driver.switch_to.window(driver.window_handles[0])
                                    
                                    with open(reports_file, "a") as record:
                                        record.write(str(email) + "\n")

                                    continue
                                except TimeoutException:
                                    pwrite(f"Creating new VideoAsk for {email}")
                            
                            # Create VideoAsk
                            make_click('//*[contains(@aria-label, "Start a new video interaction with")]')
                            make_click('//*[@aria-label="Pick a video from the library"]')
                            make_click('(//*[@data-testid="media-thumbnail"])[1]')
                            make_click('//*[@aria-label="YES"]')
                            write_t('//*[@aria-label="title"]', title)
                            make_click('//*[@aria-label="Fit video"]')
                            make_click('//button//*[text()="Button"]')
                            write_t('//*[@aria-label="button label"]',button_text)
                            write_t('//*[@aria-label="button link"]', button_url)
                            make_click('//button[@type="submit"]')
                            
                            pwrite(f"Successfully created VideoAsk for {email}")
                            speed_to_time()
                            
                            # Close tab and return to main
                            driver.close()
                            driver.switch_to.window(driver.window_handles[0])
                            try:
                                make_click("//a[contains(@href, 'respondents')]")
                            except:
                                href = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href, 'respondents')]"))).get_attribute("href")
                                driver.get(href)
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
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="email"]'))).send_keys(Keys.CONTROL + "a")
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="email"]'))).send_keys(Keys.BACKSPACE)
                            time.sleep(1)
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="password"]'))).send_keys(Keys.CONTROL + "a")
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="password"]'))).send_keys(Keys.BACKSPACE)
                            time.sleep(1)
                            write_t('//*[@name="email"]', login)
                            write_t('//*[@name="password"]', password)
                            make_click('//*[@type="submit"]')
                        except:
                            pass
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
