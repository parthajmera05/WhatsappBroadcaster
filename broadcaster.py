import time
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# Function to send WhatsApp message
def send_whatsapp_message(phone, message, driver):
    print(f"Sending to {phone}...")
    url = f"https://web.whatsapp.com/send?phone={phone}&text={message}"
    driver.get(url)
    time.sleep(10)

    try:
        send_button = driver.find_element(By.XPATH, '//button[@aria-label="Send"]')
        send_button.click()
        print(f"✅ Message sent to {phone}")
    except Exception as e:
        print(f"❌ Failed to send to {phone}: {e}")

# Function to select Excel file
def select_excel():
    path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if path:
        excel_path.set(path)

# Function to start sending messages
def send_messages():
    excel_file = excel_path.get()
    message = message_entry.get("1.0", tk.END).strip()

    if not excel_file or not message:
        messagebox.showerror("Error", "Please select an Excel file and enter a message.")
        return

    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel file: {e}")
        return

    # Setup Chrome with user profile
    options = Options()
    options.add_argument(r"user-data-dir=C:\SeleniumProfile")  # Change this path if needed
    options.add_argument("--profile-directory=Default")
    options.add_experimental_option("detach", True)

    service = Service(executable_path="chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=options)

    for index, row in df.iterrows():
        phone = str(row['Phone']).strip()
        send_whatsapp_message(phone, message, driver)
        time.sleep(10)

    driver.quit()
    messagebox.showinfo("Done", "All messages have been sent!")

# Build GUI
root = tk.Tk()
root.title("WhatsApp Message Sender")
root.geometry("550x350")

excel_path = tk.StringVar()

# Excel file selection
frame_excel = tk.Frame(root)
frame_excel.pack(pady=10, fill='x')
tk.Label(frame_excel, text="Select Excel File:").pack(side='left', padx=5)
tk.Entry(frame_excel, textvariable=excel_path, width=40).pack(side='left', padx=5)
tk.Button(frame_excel, text="Browse", command=select_excel).pack(side='left', padx=5)

# Message input
tk.Label(root, text="Enter Message:").pack(pady=5)
message_entry = tk.Text(root, height=6, width=60)
message_entry.pack(pady=5)

# Send button
tk.Button(root, text="Send Messages", command=send_messages, bg="green", fg="white", height=2, width=20).pack(pady=20)

root.mainloop()
