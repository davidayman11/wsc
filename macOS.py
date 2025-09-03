import pandas as pd
import subprocess
import time
import urllib.parse
import random
from datetime import datetime
import os

# Configurable safety limits
MAX_MESSAGES_PER_HOUR = 50
BATCH_SIZE = 30               # Messages per batch
COOLDOWN_MINUTES = 12         # Pause duration between batches

chrome_path = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'

# AppleScript to press the Return key
applescript = '''
osascript -e '
tell application "WhatsApp"
    activate
    delay 1
    tell application "System Events"
        keystroke return
    end tell
end tell

delay 3 -- Wait a few seconds to ensure the message is sent

tell application "Google Chrome"
    activate
    delay 1
    tell application "System Events"
        keystroke "w" using {command down}
    end tell
end tell
'
'''

# Excel file setup
file_path = '/Users/davidayman/Desktop/wsc/pythonProject/lectures.xlsx'
sheet_name = 'Sheet16'
phone_col = 'Phone'
msg_col = 'Message'

# Validate file
if not os.path.exists(file_path):
    print(f"[✗] File not found: {file_path}")
    exit(1)

# Load Excel
df = pd.read_excel(file_path, sheet_name=sheet_name)
print("Loaded Excel. Columns found:", df.columns)

sent_count = 0
log_file_path = 'failed_log.txt'

with open(log_file_path, 'w') as log_file:
    for index, row in df.iterrows():
        if sent_count >= MAX_MESSAGES_PER_HOUR:
            print(f"Reached hourly limit of {MAX_MESSAGES_PER_HOUR} messages. Stopping...")
            break

        # --- Clean phone number ---
        raw_phone = row[phone_col]
        if isinstance(raw_phone, float):
            raw_phone = str(int(raw_phone))  # Strip .0 if float
        else:
            raw_phone = str(raw_phone).strip()

        # Remove all non-digit characters
        phone = ''.join(filter(str.isdigit, raw_phone))

        # Add Egypt country code if missing
        if not phone.startswith('20'):
            phone = '20' + phone

        # --- Prepare message ---
        message = str(row[msg_col]).strip()
        encoded_msg = urllib.parse.quote(message)
        url = f"https://wa.me/{phone}?text={encoded_msg}"

        print(f"Opening: {url}")
        try:
            subprocess.Popen([chrome_path, url])
            time.sleep(random.uniform(12, 20))  # Time to load the page

            subprocess.run(applescript, shell=True)
            print(f"[✓] Sent to {phone} at {datetime.now().strftime('%H:%M:%S')}")
            sent_count += 1

        except Exception as e:
            error_message = f"[✗] Failed to send to {phone}: {e}"
            print(error_message)
            log_file.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {error_message}\n")

        delay = random.uniform(20, 35)
        print(f"Waiting {delay:.1f} seconds before next...")
        time.sleep(delay)

        if sent_count % BATCH_SIZE == 0 and sent_count != 0:
            print(f"Completed {BATCH_SIZE} messages. Cooling down for {COOLDOWN_MINUTES} minutes...")
            time.sleep(COOLDOWN_MINUTES * 60)

print("✅ All done. Log saved to failed_log.txt.")