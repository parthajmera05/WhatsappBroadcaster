import pandas as pd
from tkinter import *
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import threading
import time
import urllib.parse
import os
import traceback

class WhatsAppBroadcaster:
    def __init__(self, root):
        self.root = root
        self.root.title("WhatsApp Broadcaster")
        self.root.geometry("550x500")

        self.excel_path = ""
        self.media_path = ""

        # GUI Components
        Button(root, text="Select Excel File", command=self.load_excel).pack(pady=10)
        self.excel_label = Label(root, text="No Excel selected")
        self.excel_label.pack()

        Label(root, text="Enter Message (Unicode supported):").pack(pady=5)
        self.message_text = Text(root, height=5, width=60)
        self.message_text.pack()

        Button(root, text="Select Media File (optional)", command=self.load_media).pack(pady=10)
        self.media_label = Label(root, text="No media selected")
        self.media_label.pack()

        Button(root, text="Start Broadcast", command=self.start_broadcast_thread, bg="green", fg="white").pack(pady=20)

        self.log_output = Text(root, height=10, width=65, state=DISABLED, bg="black", fg="lime")
        self.log_output.pack(pady=10)

    def log(self, text):
        self.log_output.config(state=NORMAL)
        self.log_output.insert(END, text + "\n")
        self.log_output.config(state=DISABLED)
        self.log_output.see(END)

    def load_excel(self):
        try:
            self.excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
            self.excel_label.config(text=os.path.basename(self.excel_path) if self.excel_path else "No Excel selected")
        except Exception as e:
            self.log("Error selecting Excel file: " + str(e))

    def load_media(self):
        try:
            self.media_path = filedialog.askopenfilename(filetypes=[("Media files", "*.*")])
            self.media_label.config(text=os.path.basename(self.media_path) if self.media_path else "No media selected")
        except Exception as e:
            self.log("Error selecting media file: " + str(e))

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

            self.log("Reading Excel file...")
            df = pd.read_excel(self.excel_path)
            if "Phone" not in df.columns:
                self.log("Excel must have a 'Phone' column.")
                return

            self.log("Connecting to existing Chrome session...")
            options = webdriver.ChromeOptions()
            options.add_argument("--user-data-dir=/Users/parthajmera/Library/Application Support/Google/Chrome")
            options.add_argument("--profile-directory=Default")
            options.add_argument("--headless=new")  # New headless mode (Chrome 109+)
            options.add_argument("--window-size=1200,1000")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")

            driver = webdriver.Chrome(options=options)
            self.log("Driver connected successfully.")

            for idx, row in df.iterrows():
                try:
                    phone = str(row["Phone"]).strip()
                    if not phone.isdigit():
                        self.log(f"Invalid phone number at row {idx + 2}: {phone}")
                        continue

                    text = urllib.parse.quote(message)
                    url = f"https://web.whatsapp.com/send?phone={phone}&text={text}"
                    driver.get(url)
                    self.log(f"Opened chat for {phone}")
                    time.sleep(8)

                    try:
                        send_btn = WebDriverWait(driver, 15).until(
                            EC.element_to_be_clickable((By.XPATH, '//*[@data-testid="compose-btn-send"]'))
                        )
                        send_btn.click()
                        self.log(f"Text sent to {phone}")
                    except Exception as e:
                        self.log(f"Failed to send message to {phone}: {e}")

                    time.sleep(5)
                except Exception as inner_e:
                    self.log(f"Error with row {idx + 2}: {inner_e}")
                    traceback.print_exc()

            self.log("Broadcast completed.")
            messagebox.showinfo("Done", "Broadcast completed successfully.")
        except Exception as e:
            self.log("Fatal error: " + str(e))
            traceback.print_exc()


if __name__ == "__main__":
    root = Tk()
    app = WhatsAppBroadcaster(root)
    root.mainloop()
