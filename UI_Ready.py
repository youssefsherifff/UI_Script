# Updated

# Updated

import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from PIL import Image, ImageTk
import sys
import os

stop_flag = False
driver = None
results = []

root = None

# --- Globals from UI ---
EMAIL_ADDRESS_UI = ""
EMAIL_PASSWORD_UI = ""
INPUT_FILE = ""
OUTPUT_FILE = "output.xlsx"

# --- Redirect print to log window ---
class TextRedirector(object):
    def __init__(self, widget):
        self.widget = widget

    def write(self, msg):
        self.widget.insert(tk.END, msg)
        self.widget.see(tk.END)

    def flush(self):
        pass

# --- Core runner function (your automation script goes here) ---
def run_main():
    import pandas as pd
    import traceback
    from datetime import datetime
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.keys import Keys
    import time
    import imaplib
    import email
    from email.header import decode_header
    import re
    import requests

    global EMAIL_ADDRESS, EMAIL_PASSWORD, INPUT_FILE, OUTPUT_FILE
    global stop_flag, driver, results

    EMAIL_ADDRESS = EMAIL_ADDRESS_UI
    EMAIL_PASSWORD = EMAIL_PASSWORD_UI

    def extract_otp(text):
        match = re.search(r"\b\d{4}\b", text)
        if match:
            return match.group(0)
        return None

    def wait_for_otp(timeout=120):
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        mail.select('"[Gmail]/All Mail"')
        status, messages = mail.search(None, 'UNSEEN')
        for num in messages[0].split():
            mail.store(num, '+FLAGS', '\\Seen')
        print("📩 Waiting for new OTP email...")
        start_time = time.time()
        while time.time() - start_time < timeout:
            mail.select('"[Gmail]/All Mail"')
            status, messages = mail.search(None, 'UNSEEN')
            if status == "OK":
                unread_count = len(messages[0].split())
                print(f"🔎 Found {unread_count} unread emails")
                for num in messages[0].split():
                    res, msg_data = mail.fetch(num, "(RFC822)")
                    for response_part in msg_data:
                        if isinstance(response_part, tuple):
                            msg = email.message_from_bytes(response_part[1])
                            subject, encoding = decode_header(msg["Subject"])[0]
                            if isinstance(subject, bytes):
                                subject = subject.decode(encoding or "utf-8", errors="ignore")
                            print(f"📨 New email subject: {subject}")
                            body = None
                            if msg.is_multipart():
                                for part in msg.walk():
                                    ctype = part.get_content_type()
                                    disp = str(part.get("Content-Disposition"))
                                    if ctype == "text/plain" and "attachment" not in disp:
                                        body = part.get_payload(decode=True).decode(errors="ignore")
                                        break
                            else:
                                body = msg.get_payload(decode=True).decode(errors="ignore")
                            if body:
                                preview = body[:200].replace("\n", " ")
                                print(f"📜 Email body preview: {preview}")
                                otp_code = extract_otp(body)
                                if otp_code:
                                    print(f"✅ OTP found: {otp_code}")
                                    mail.store(num, '+FLAGS', '\\Seen')
                                    mail.logout()
                                    return otp_code
                                else:
                                    print("❌ No OTP found in this email body")
            time.sleep(5)
        mail.logout()
        raise TimeoutError("❌ OTP email not received in time")

    input_file = INPUT_FILE
    output_file = "output.xlsx"
    data = pd.read_excel(input_file)
    total_instances = len(data)
    print(f"📂 Loaded {total_instances} instances from {input_file}")
    results = []
    global driver
    driver = webdriver.Chrome()
    for idx, row in data.iterrows():
        if stop_flag:
            print("⚠️ Stop requested, exiting loop...")
            break
        print(f"\n--- Processing instance {idx+1}/{total_instances} ---")
        try:
            visa_number = int(str(row["رقم التاشيرة"]).split('.')[0])
            nationality = row["الجنسية"]
            passport_number = str(row["رقم الجواز"])
            dob = str(row.get("تاريخ الميلاد", "01/01/1990"))
            dob_formatted = dob.replace("/", "-")
            gender_raw = row["النوع"]
            gender = gender_raw.replace("أنثي", "ذكر")
            code = "20"
            mobile_number = row["رقم الجوال"]
            assistance_map = {"نعم": "1", "لا": "0"}
            assistance_value = assistance_map["لا"]
            email_address = str(row["الايميل"])
            password = str(row.get("كلمة السر", "Aa@1234567"))
            driver.get("https://services.nusuk.sa")
            first_button = driver.find_element(By.XPATH, "//h1[text()='الدخول بحساب نسك']/..")
            first_button.click()
            print("First Button clicked")
            wait = WebDriverWait(driver, 10)
            Second_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[text()='إنشاء حساب']/.."))
            )
            Second_button.click()
            print("Second Button clicked")
            time.sleep(1)
            Third_button = wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='radio' and @value='V']"))
            )
            driver.execute_script("arguments[0].click();", Third_button)
            print("Third button clicked")
           
            time.sleep(1)
            dropdown = driver.find_element(By.ID, "selectedNationalityV")
            dropdown.click()
            print("Dropdown opened")
            wait = WebDriverWait(driver, 10)
            options_list = wait.until(EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, "li[role='option']")
            ))
            print("Options list appeared")
            for option in options_list:
                if option.text.strip() == nationality:
                    option.click()
                    print(f"Nationality '{nationality}' selected")
                    break

            passport_input = wait.until(
                EC.presence_of_element_located((By.ID, "vPassport"))
            )
            passport_input.clear()
            passport_input.send_keys(str(passport_number))
            print(f"Passport number '{passport_number}' entered")

            next_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='التالي']]"))
            )
            next_button.click()
            print("Next button clicked")

            visa_input = wait.until(
                EC.presence_of_element_located((By.ID, "visa"))
            )
            visa_input.clear()
            visa_input.send_keys(str(visa_number))
            print("Visa number entered")

            time.sleep(0.5)

            next_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='التالي']]"))
            )
            next_button.click()
            print("Next button clicked")

            dob_input = wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@placeholder='حدد تاريخ الميلاد']"))
            )
            dob_input.clear()
            dob_input.send_keys(dob_formatted)
            print(f"DOB filled: {dob_formatted}")
            dob_input.send_keys(Keys.ESCAPE)
            time.sleep(0.5)
            gender_map = {
                "ذكر": "1",
                "أنثى": "2"
            }
            gender_value = gender_map[gender]
            gender_radio = wait.until(
                EC.presence_of_element_located((By.XPATH, f"//input[@type='radio' and @value='{gender_value}']"))
            )
            driver.execute_script("arguments[0].click();", gender_radio)
            print(f"Gender '{gender}' selected")
            dropdown_trigger = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//span[@role='combobox']"))
            )
            dropdown_trigger.click()
            print("Dropdown opened")
            option = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, f"//li[@role='option']//span[normalize-space(text())='{code}']/..")
                )
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", option)
            driver.execute_script("arguments[0].click();", option)
            print(f"Code '{code}' selected")
            mobile_input = wait.until(
                EC.presence_of_element_located((By.ID, "mobile"))
            )
            mobile_input.clear()
            mobile_input.send_keys(mobile_number)
            print(f"Mobile number '{mobile_number}' entered successfully")
            assistance_radio = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, f"//input[@type='radio' and @name='needAssistance' and @value='{assistance_value}']")
                )
            )
            driver.execute_script("arguments[0].click();", assistance_radio)
            print(f"Assistance option 'لا' selected")
            next_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='التالي']]"))
            )
            next_button.click()
            print("Next button clicked after Assistance selection")
            email_input = wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@id='email']"))
            )
            email_input.clear()
            email_input.send_keys(email_address)
            print(f"Email '{email_address}' entered successfully")
            password_input = wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='password']"))
            )
            password_input.clear()
            password_input.send_keys(password)
            print("Password entered successfully")
            confirm_password_input = wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@placeholder=' تأكيد كلمة السر']"))
            )
            confirm_password_input.clear()
            confirm_password_input.send_keys(password)
            print("Confirm password entered successfully")
            create_account_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='إنشاء حساب']]"))
            )
            driver.execute_script("arguments[0].click();", create_account_button)
            print("Create Account button clicked")
            try:
                error_box = WebDriverWait(driver, 2).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "div.p-message-wrapper span.p-message-detail")
                    )
                )
                error_text = error_box.text.strip()
                print(f"❌ Error box appeared: {error_text}")
                results.append({
                    "name": row.get("اسم المعتمر", ""),
                    "passport_number": passport_number,
                    "visa_number": visa_number,
                    "email": email_address,
                    "password": password,
                    "notes": error_text
                })
                continue
            except:
                print("✅ No error box, proceeding to OTP step")
            otp_code = wait_for_otp(timeout=20)
            otp_str = str(otp_code).zfill(4)
            otp_container = wait.until(
                EC.presence_of_element_located((By.ID, "inputs"))
            )
            otp_inputs = otp_container.find_elements(By.CSS_SELECTOR, "input[formcontrolname]")
            if len(otp_inputs) != 4:
                raise Exception(f"Expected 4 OTP input fields, found {len(otp_inputs)}")
            for i, digit in enumerate(otp_str):
                otp_inputs[i].clear()
                otp_inputs[i].send_keys(digit)
            time.sleep(0.5)
            print(f"✅ OTP '{otp_str}' entered successfully")
            time.sleep(2)
            verify_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space(text())='تحقق']"))
            )
            driver.execute_script("arguments[0].click();", verify_button)
            print("🔘 Verify button clicked")
            time.sleep(3)
            results.append({
                "name": row.get("اسم المعتمر", ""),
                "passport_number": passport_number,
                "visa_number": visa_number,
                "email": email_address,
                "password": password
            })
            print(f"✅ Completed instance {idx+1}/{total_instances}")
        except Exception as e:
            error_message = f"Issue with this account: {str(e)}"
            print(f"❌ Error at instance {idx+1}: {error_message}")
            results.append({
                "name": row.get("اسم المعتمر", ""),
                "passport_number": row.get("رقم الجواز", ""),
                "visa_number": row.get("رقم التاشيرة", ""),
                "email": row.get("الايميل", ""),
                "password": row.get("كلمة السر", "Aa@1234567"),
                "notes": error_message
            })
            continue

    def save_progress():
        try:
            pd.DataFrame(results).to_excel(output_file, index=False)
            print(f"💾 Progress saved to {output_file}")
        except Exception as e:
            print(f"❌ Failed to save progress: {e}")

    save_progress()
    if driver:
        try:
            driver.quit()
        except:
            pass
    driver = None
    print("✅ Cleanup done.")
    print(f"📩 Using {EMAIL_ADDRESS} / {EMAIL_PASSWORD}")
    print(f"📂 Input file: {INPUT_FILE}")

# --- Background worker ---
def start_task():
    thread = threading.Thread(target=run_main, daemon=True)
    thread.start()

def stop_task():
    global stop_flag, driver, root
    stop_flag = True
    print("🛑 Stop button pressed. Finishing current task and cleaning up...")
    print("📤 Updating server status to 'closed'...")
    try:
        if driver:
            print("🔧 Attempting to quit browser now...")
            driver.quit()
    except Exception as e:
        print(f"⚠️ Could not quit browser immediately: {e}")
    if root:
        print("🪟 Closing UI window in 1 second...")
        root.after(1000, root.destroy)  # Delay UI close

# --- UI Screens ---
def start_ui():
    global root, frame1, frame2, email_entry, pass_entry, file_var, log_box
    root = tk.Tk()
    root.title("Automation Tool")
    frame1 = tk.Frame(root)
    frame1.pack(fill="both", expand=True)
    spoiler_frame = tk.Frame(frame1)
    try:
        img = Image.open("columns.png")
        photo = ImageTk.PhotoImage(img)
        img_label = tk.Label(spoiler_frame, image=photo)
        img_label.image = photo
        img_label.pack()
    except Exception:
        img_label = tk.Label(spoiler_frame, text="(Column format image missing)")
        img_label.pack()
    def toggle_spoiler():
        if spoiler_frame.winfo_ismapped():
            spoiler_frame.pack_forget()
            toggle_btn.config(text="Show Column Format")
        else:
            spoiler_frame.pack(pady=5)
            toggle_btn.config(text="Hide Column Format")
    toggle_btn = tk.Button(frame1, text="Show Column Format", command=toggle_spoiler)
    toggle_btn.pack(pady=5)
    tk.Label(frame1, text="Email Address").pack()
    email_entry = tk.Entry(frame1, width=40)
    email_entry.pack()
    tk.Label(frame1, text="Email Password").pack()
    pass_entry = tk.Entry(frame1, show="*", width=40)
    pass_entry.pack()
    def browse_file():
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            file_var.set(path)
    file_var = tk.StringVar()
    tk.Label(frame1, text="Excel Input File").pack()
    file_entry = tk.Entry(frame1, textvariable=file_var, width=40)
    file_entry.pack()
    tk.Button(frame1, text="Browse", command=browse_file).pack()
    frame2 = tk.Frame(root)
    log_box = scrolledtext.ScrolledText(frame2, wrap=tk.WORD, width=80, height=25)
    log_box.pack(fill="both", expand=True)
    stop_button = tk.Button(frame2, text="Stop", command=stop_task)
    stop_button.pack(pady=5)
    sys.stdout = TextRedirector(log_box)
    sys.stderr = TextRedirector(log_box)
    def go_next():
        global EMAIL_ADDRESS_UI, EMAIL_PASSWORD_UI, INPUT_FILE
        EMAIL_ADDRESS_UI = email_entry.get().strip()
        EMAIL_PASSWORD_UI = pass_entry.get().strip()
        INPUT_FILE = file_var.get().strip()
        if not EMAIL_ADDRESS_UI or not EMAIL_PASSWORD_UI or not INPUT_FILE:
            messagebox.showerror("Missing Input", "⚠️ Please fill in all fields and select a file before proceeding.")
            return
        if not os.path.isfile(INPUT_FILE):
            messagebox.showerror("Invalid File", f"⚠️ The file '{INPUT_FILE}' does not exist.")
            return
        frame1.pack_forget()
        frame2.pack(fill="both", expand=True)
        start_task()
    tk.Button(frame1, text="Next", command=go_next).pack(pady=10)
    root.protocol("WM_DELETE_WINDOW", stop_task)
    root.mainloop()

if __name__ == "__main__":
    start_ui()


