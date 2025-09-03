import pandas as pd
import subprocess
import time
import urllib.parse
import random
from datetime import datetime
import pyautogui

# Configurable safety limits
MAX_MESSAGES_PER_HOUR = 50
BATCH_SIZE = 30
COOLDOWN_MINUTES = 12

chrome_path = 'C:/Program Files/Google/Chrome/Application/chrome.exe'  # Adjust path if different

# Load Excel file
file_path = 'C:/Users/YourUsername/Desktop/whatsapp_sender/tshirt.xlsx'  # Update your path
sheet_name = 'Sheet1'
phone_col = 'Phone'
msg_col = 'Message'

df = pd.read_excel(file_path, sheet_name=sheet_name)
print("Columns in Excel:", df.columns)

sent_count = 0

# Create or overwrite log file
log_file_path = 'failed_log.txt'
log_file = open(log_file_path, 'w')

for index, row in df.iterrows():
    if sent_count >= MAX_MESSAGES_PER_HOUR:
        print(f"Reached hourly limit of {MAX_MESSAGES_PER_HOUR} messages. Stopping...")
        break

    phone = str(row[phone_col]).strip().replace(' ', '').replace('-', '')
    if not phone.startswith('+20'):
        phone = '+20' + phone

    message = str(row[msg_col]).strip()
    encoded_msg = urllib.parse.quote(message)
    url = f"https://wa.me/{phone}?text={encoded_msg}"

    try:
        subprocess.Popen([chrome_path, url])
        time.sleep(random.uniform(12, 20))  # Let the page load

        # Simulate pressing 'Enter' to send message
        pyautogui.press('enter')

        print(f"[✓] Sent to {phone} at {datetime.now().strftime('%H:%M:%S')}")
        sent_count += 1
    except Exception as e:
        error_message = f"[✗] Failed to send to {phone}: {e}"
        print(error_message)
        log_file.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {error_message}\n")

    # Safety delay before next message
    delay = random.uniform(20, 35)
    print(f"Waiting {delay:.1f} seconds before next...")
    time.sleep(delay)

    # Pause after batch
    if sent_count % BATCH_SIZE == 0 and sent_count != 0:
        print(f"Completed {BATCH_SIZE} messages. Cooling down for {COOLDOWN_MINUTES} minutes...")
        time.sleep(COOLDOWN_MINUTES * 60)

# Close log file
log_file.close()
print("All done. Log saved to failed_log.txt.")