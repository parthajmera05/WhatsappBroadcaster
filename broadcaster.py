import time
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# Load contacts from Excel
excel_file = "Phone.xlsx"  # Make sure it's in the same directory
df = pd.read_excel(excel_file)

# Set up Chrome options with your existing profile (macOS path)
options = Options()
# options.add_argument("user-data-dir=/Users/parthajmera/Library/Application Support/Google/Chrome")  # Root user data dir
# options.add_argument("--profile-directory=Default")
options.add_argument(r"user-data-dir=C:\SeleniumProfile")  # Use the custom profile
options.add_argument("--profile-directory=Default")  # Default inside the custom dir
options.add_experimental_option("detach", True)  # Optional: Keep window open after script ends

# Path to your ChromeDriver (replace with actual path if different)
service = Service(executable_path="chromedriver.exe") # or wherever your chromedriver is installed

# Start the driver
driver = webdriver.Chrome(service=service, options=options)

# Function to send WhatsApp message
def send_whatsapp_message(phone, message):
    print(f"Sending to {phone}...")
    url = f"https://web.whatsapp.com/send?phone={phone}&text={message}"
    driver.get(url)
    time.sleep(10)  # Wait for WhatsApp Web to load

    try:
        # Wait for chat to open and click send
        send_button = driver.find_element(By.XPATH, '//button[@aria-label="Send"]')
        send_button.click()
        print(f"✅ Message sent to {phone}")
    except Exception as e:
        print(f"❌ Failed to send to {phone}: {e}")

# Loop through all contacts and send messages
for index, row in df.iterrows():
    phone = str(row['Phone']).strip()
    message = "Hi, This is a Test Message"
    send_whatsapp_message(phone, message)
    time.sleep(10)  # Delay to avoid spam detection

driver.quit()
