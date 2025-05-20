import pandas as pd
from tkinter import *
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import threading
import time
import os
import traceback

class WhatsAppBroadcaster:
    def __init__(self, root):
        self.root = root
        self.root.title("WhatsApp Broadcaster")
        self.root.geometry("500x400")

        self.excel_path = ""
        self.media_path = ""

        Button(root, text="Select Excel File", command=self.load_excel).pack(pady=10)
        self.excel_label = Label(root, text="No Excel selected")
        self.excel_label.pack()

        Label(root, text="Enter Message (Unicode supported):").pack(pady=5)
        self.message_text = Text(root, height=5, width=50)
        self.message_text.pack()

        Button(root, text="Select Media File (optional)", command=self.load_media).pack(pady=10)
        self.media_label = Label(root, text="No media selected")
        self.media_label.pack()

        Button(root, text="Start Broadcast", command=self.start_broadcast_thread).pack(pady=20)

        self.log_output = Text(root, height=8, width=60, state=DISABLED)
        self.log_output.pack(pady=10)

    def log(self, text):
        self.log_output.config(state=NORMAL)
        self.log_output.insert(END, text + "\n")
        self.log_output.config(state=DISABLED)
        self.log_output.see(END)

    def load_excel(self):
        self.excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        self.excel_label.config(text=os.path.basename(self.excel_path) if self.excel_path else "No Excel selected")

    def load_media(self):
        self.media_path = filedialog.askopenfilename(filetypes=[("Media files", "*.*")])
        self.media_label.config(text=os.path.basename(self.media_path) if self.media_path else "No media selected")

    def start_broadcast_thread(self):
        threading.Thread(target=self.broadcast, daemon=True).start()

    def broadcast(self):
        try:
            if not self.excel_path:
                messagebox.showerror("Missing File", "Please select an Excel file.")
                return

            message = self.message_text.get("1.0", END).strip()
            if not message:
                messagebox.showerror("Empty Message", "Please enter a message.")
                return

            self.log("Reading Excel...")
            df = pd.read_excel(self.excel_path)
            if "Phone" not in df.columns:
                self.log("Excel must have a 'Phone' column.")
                return

            self.log("Launching Chrome...")
            options = webdriver.ChromeOptions()
            options.add_argument("user-data-dir=C:/temp/whatsapp_profile")
            driver = webdriver.Chrome(options=options)
            driver.get("https://web.whatsapp.com")
            self.log("Waiting for WhatsApp Web to load...")
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "canvas[aria-label='Scan me!']"))
            )
            self.log("Please scan the QR code if prompted...")

            # Wait until WhatsApp loads main UI (chat list)
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='button'][title*='New chat']"))
            )

            for idx, row in df.iterrows():
                try:
                    phone = str(row["Phone"]).strip()
                    if not phone.isdigit():
                        self.log(f"Invalid phone at row {idx+2}: {phone}")
                        continue

                    self.log(f"Sending to {phone}...")
                    driver.get(f"https://web.whatsapp.com/send?phone={phone}&text&app_absent=0")
                    time.sleep(8)

                    try:
                        # Wait until message input box is available
                        input_box = WebDriverWait(driver, 15).until(
                            EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='10']"))
                        )
                        input_box.click()
                        input_box.send_keys(message)
                        time.sleep(1)
                    except Exception as e:
                        self.log(f"Failed to find input box for {phone}: {e}")
                        continue

                    # Send text or media
                    if self.media_path and os.path.exists(self.media_path):
                        try:
                            attach_btn = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, "span[data-icon='clip']"))
                            )
                            attach_btn.click()
                            time.sleep(1)

                            file_input = driver.find_element(By.CSS_SELECTOR, "input[type='file']")
                            file_input.send_keys(self.media_path)
                            time.sleep(3)

                            send_btn = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, "span[data-icon='send']"))
                            )
                            send_btn.click()
                            self.log(f"Media + text sent to {phone}")
                        except Exception as e:
                            self.log(f"Media failed for {phone}: {e}")
                    else:
                        send_btn = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, "span[data-icon='send']"))
                        )
                        send_btn.click()
                        self.log(f"Text sent to {phone}")

                    time.sleep(5)
                except Exception as inner_e:
                    self.log(f"Error with {phone}: {inner_e}")
                    traceback.print_exc()

            self.log("Broadcast complete.")
            driver.quit()

        except Exception as e:
            self.log(f"Fatal Error: {e}")
            traceback.print_exc()

if __name__ == "__main__":
    root = Tk()
    app = WhatsAppBroadcaster(root)
    root.mainloop()
