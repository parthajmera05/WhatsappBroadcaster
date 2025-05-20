import time
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Load contacts from Excel
excel_file = "Phone.xlsx"
df = pd.read_excel(excel_file)

# Setup Chrome options
options = Options()
options.add_argument(r"user-data-dir=C:\SeleniumProfile")  # Change this to your path
options.add_argument("--profile-directory=Default")
options.add_experimental_option("detach", True)

# Initialize Chrome driver
service = Service(executable_path="chromedriver.exe")
driver = webdriver.Chrome(service=service, options=options)

# Function to send image with caption
def send_whatsapp_media_with_caption(phone, caption, media_path):
    print(f"\nSending to {phone}...")

    url = f"https://web.whatsapp.com/send?phone={phone}"
    driver.get(url)

    try:
        # Wait for chat to load
        # Wait for chat to load
        attach_button = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//button[@title="Attach"]'))
        )
        time.sleep(1)
        attach_button.click()
        time.sleep(1)

        # Just send the file to the hidden input (skip menu click)
        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'))
        )
        file_input.send_keys(os.path.abspath(media_path))
        time.sleep(3)


        # Add caption
        caption_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[@aria-label="Add a caption"][@contenteditable="true"]'))
        )
        caption_box.click()
        caption_box.send_keys(caption)
        time.sleep(1)

        # Click send
        send_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
        )
        send_button.click()

        print(f"✅ Media with caption sent to {phone}")

    except Exception as e:
        print(f"❌ Failed to send to {phone}: {e}")

# Send to all contacts
for index, row in df.iterrows():
    phone = str(row['Phone']).strip()
    caption = "Hello, this is a test image with caption."
    media_path ="image.jpg"
    send_whatsapp_media_with_caption(phone, caption, media_path)
    time.sleep(10)

driver.quit()
