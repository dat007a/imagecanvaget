import os
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class CanvaAIAutoTool:
    def __init__(self, master):
        self.master = master
        self.master.title("Canva AI Image Generator")

        self.excel_file = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.status = tk.StringVar(value="Ready")
        self.driver = None
        self.running = False

        self.build_gui()

    def build_gui(self):
        padding = {'padx': 10, 'pady': 5}

        tk.Label(self.master, text="Excel File:").grid(row=0, column=0, sticky='e', **padding)
        tk.Entry(self.master, textvariable=self.excel_file, width=50).grid(row=0, column=1, **padding)
        tk.Button(self.master, text="Browse", command=self.select_excel).grid(row=0, column=2, **padding)

        tk.Label(self.master, text="Output Folder:").grid(row=1, column=0, sticky='e', **padding)
        tk.Entry(self.master, textvariable=self.output_folder, width=50).grid(row=1, column=1, **padding)
        tk.Button(self.master, text="Browse", command=self.select_output).grid(row=1, column=2, **padding)

        self.start_btn = tk.Button(self.master, text="Start Action", command=self.start)
        self.start_btn.grid(row=2, column=1, **padding)
        self.stop_btn = tk.Button(self.master, text="Stop", command=self.stop, state='disabled')
        self.stop_btn.grid(row=2, column=2, **padding)

        tk.Label(self.master, textvariable=self.status, fg='blue').grid(row=3, column=0, columnspan=3, **padding)
        self.log_box = tk.Text(self.master, height=15, width=80, state='disabled')
        self.log_box.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

        tk.Label(self.master, text="🟡 Vui lòng tự mở Chrome bằng lệnh sau trước khi nhấn Start:").grid(row=5, column=0, columnspan=3, **padding)
        tk.Label(self.master, text="chrome.exe --remote-debugging-port=9222 --user-data-dir=C:/CanvaProfile https://www.canva.com/ai", fg='green', wraplength=700, justify='left').grid(row=6, column=0, columnspan=3, **padding)

    def log(self, message):
        self.log_box.config(state='normal')
        self.log_box.insert('end', message + '\n')
        self.log_box.see('end')
        self.log_box.config(state='disabled')
        self.status.set(message)

    def select_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.excel_file.set(file_path)

    def select_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder.set(folder)

    def start(self):
        if not self.excel_file.get() or not self.output_folder.get():
            messagebox.showwarning("Missing Info", "Please select Excel file and Output folder.")
            return

        self.running = True
        self.start_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        threading.Thread(target=self.connect_and_run, daemon=True).start()

    def stop(self):
        self.running = False
        self.status.set("Stopping... Please wait")

    def connect_and_run(self):
        try:
            self.log("Đang kết nối tới Chrome qua cổng 9222...")
            chrome_options = Options()
            chrome_options.debugger_address = "127.0.0.1:9222"
            self.driver = webdriver.Chrome(options=chrome_options)
            self.log("Đã kết nối Chrome thành công. Bắt đầu xử lý ảnh...")
            self.run_automation()
        except Exception as e:
            import traceback
            self.log(f"Không kết nối được Chrome: {e}")
            self.log(traceback.format_exc())
        finally:
            self.start_btn.config(state='normal')
            self.stop_btn.config(state='disabled')
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass

    def run_automation(self):
        try:
            df = pd.read_excel(self.excel_file.get())
            self.log(f"Loaded {len(df)} prompts from Excel file.")

            for index, row in df.iterrows():
                if not self.running:
                    break
                self.log(f"Creating image {index + 1} of {len(df)}: {row[0]}")

                try:
                    self.driver.execute_script("window.scrollTo(0, 0);")
                    prompt_box = WebDriverWait(self.driver, 15).until(
    EC.element_to_be_clickable((
        By.XPATH,
        "//textarea[@placeholder=\"Describe the image in your mind, and I’ll bring it to life\"] | //textarea[@placeholder=\"Describe your idea, and I’ll bring it to life\"]"
    ))
                    )
                    self.driver.execute_script("arguments[0].value = ''", prompt_box)
                    prompt_box.clear()
                    prompt_box.click()
                    prompt_box.send_keys(str(row[1]))
                    time.sleep(1.5)

                    submit_btn = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and @aria-label='Submit']"))
                    )
                    submit_btn.click()
                    self.log("Prompt submitted. Waiting for images to generate...")
                    time.sleep(12)
                except Exception as e:
                    self.log(f"❌ Lỗi khi nhập prompt hoặc ấn submit: {e}")
                    continue

                try:
                    image_section = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_all_elements_located((By.XPATH, "//button[@aria-label='Download Image']"))
                    )
                    for i, btn in enumerate(image_section[:4]):
                        if not self.running:
                            break
                        try:
                            btn.click()
                            self.log(f"Downloading image {i + 1}...")
                            time.sleep(2)
                            filename = f"{row[0]}-{i + 1}.png"
                            default_download = os.path.expanduser("~/Downloads")
                            dst = os.path.join(self.output_folder.get(), filename)

                            newest_file = max(
                                [os.path.join(default_download, f) for f in os.listdir(default_download)],
                                key=os.path.getctime
                            )
                            if os.path.isfile(newest_file):
                                os.rename(newest_file, dst)
                                self.log(f"Saved as {filename}")
                        except Exception as e:
                            self.log(f"Download failed: {e}")
                except Exception as e:
                    self.log(f"❌ Lỗi khi tìm hoặc tải ảnh: {e}")

            self.log("✅ Hoàn thành toàn bộ quá trình.")
        except Exception as e:
            import traceback
            self.log(f"Error: {e}")
            self.log(traceback.format_exc())

if __name__ == "__main__":
    root = tk.Tk()
    app = CanvaAIAutoTool(root)
    root.mainloop()
