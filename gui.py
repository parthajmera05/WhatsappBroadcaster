import time
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Function to send image with caption
def send_whatsapp_media_with_caption(phone, caption, media_path, driver):
    print(f"\nSending to {phone}...")

    url = f"https://web.whatsapp.com/send?phone={phone}"
    driver.get(url)

    try:
        attach_button = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//button[@title="Attach"]'))
        )
        time.sleep(1)
        attach_button.click()
        time.sleep(1)

        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'))
        )
        file_input.send_keys(os.path.abspath(media_path))
        time.sleep(3)

        caption_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[@aria-label="Add a caption"][@contenteditable="true"]'))
        )
        caption_box.click()
        caption_box.send_keys(caption)
        time.sleep(1)

        send_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
        )
        send_button.click()

        print(f"✅ Media with caption sent to {phone}")
    except Exception as e:
        print(f"❌ Failed to send to {phone}: {e}")

# GUI functions
def select_excel():
    path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if path:
        excel_path.set(path)

def select_media():
    path = filedialog.askopenfilename(filetypes=[("Image/Video Files", "*.jpg;*.png;*.mp4;*.mov;*.3gp")])
    if path:
        media_path.set(path)

def send_messages():
    excel_file = excel_path.get()
    media = media_path.get()
    caption = caption_entry.get("1.0", tk.END).strip()

    if not all([excel_file, media, caption]):
        messagebox.showerror("Error", "Please select all inputs and enter a message.")
        return

    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel file: {e}")
        return

    options = Options()
    options.add_argument(r"user-data-dir=C:\SeleniumProfile")  # Change path if needed
    options.add_argument("--profile-directory=Default")
    options.add_experimental_option("detach", True)

    service = Service(executable_path="chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=options)

    for index, row in df.iterrows():
        phone = str(row['Phone']).strip()
        send_whatsapp_media_with_caption(phone, caption, media, driver)
        time.sleep(10)

    driver.quit()
    messagebox.showinfo("Done", "All messages have been sent!")

# Create GUI
root = tk.Tk()
root.title("WhatsApp Media Sender")
root.geometry("550x400")  # Increased height to show all widgets

excel_path = tk.StringVar()
media_path = tk.StringVar()

# Excel File Input
frame_excel = tk.Frame(root)
frame_excel.pack(pady=10, fill='x')
tk.Label(frame_excel, text="Select Excel File:").pack(side='left', padx=5)
tk.Entry(frame_excel, textvariable=excel_path, width=40).pack(side='left', padx=5)
tk.Button(frame_excel, text="Browse", command=select_excel).pack(side='left', padx=5)

# Media File Input
frame_media = tk.Frame(root)
frame_media.pack(pady=10, fill='x')
tk.Label(frame_media, text="Select Media File:").pack(side='left', padx=5)
tk.Entry(frame_media, textvariable=media_path, width=40).pack(side='left', padx=5)
tk.Button(frame_media, text="Browse", command=select_media).pack(side='left', padx=5)

# Caption Text Box
tk.Label(root, text="Enter Caption:").pack(pady=5)
caption_entry = tk.Text(root, height=5, width=60)
caption_entry.pack(pady=5)

# Send Button
tk.Button(root, text="Send Messages", command=send_messages, bg="green", fg="white", height=2, width=20).pack(pady=20)

root.mainloop()

