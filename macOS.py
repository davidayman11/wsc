import pandas as pd
import subprocess
import time
import urllib.parse
import os
import re

# Define the path to Google Chrome
chrome_path = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

# AppleScript to press Return key and close tab
applescript = '''
osascript -e '
tell application "System Events"
    delay 3
    keystroke return
end tell
delay 5
tell application "Google Chrome"
    keystroke "w" using {command down}
end tell
'
'''

# Read data from Excel
file_path = '/Users/davidayman/Desktop/whatsapp sender code /pythonProject/script.xlsx'  # Replace with your Excel file path
sheet_name = 'Data'  # Ensure this matches your sheet name
phone_column = 'Phone'  # Ensure this matches your Excel column
message_column = 'Message'  # Ensure this matches your Excel column

# Check if file exists
if not os.path.exists(file_path):
    print(f"❌ Error: File not found at {file_path}")
    exit()

# Load the Excel file
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Debug: Print column names
print("📄 Columns in the Excel file:", df.columns)

# Delay between messages
delay = 10  # Increased delay for safety

# Iterate over phone numbers and send messages
for index, row in df.iterrows():
    phone_number = str(row[phone_column]).strip()
    message = str(row[message_column]).strip()

    # Ensure valid phone number format
    phone_number = phone_number.replace(' ', '').replace('-', '')
    if not phone_number.startswith('+20'):
        phone_number = '+20' + phone_number  # Ensure it starts with '+20'

    # Validate number format
    if not re.match(r'^\+20\d{9,10}$', phone_number):
        print(f"⚠️ Skipping invalid phone number: {phone_number}")
        continue

    # URL encode the message
    encoded_message = urllib.parse.quote(message)

    # Generate WhatsApp Web URL
    url = f"https://wa.me/{phone_number}?text={encoded_message}"

    try:
        print(f"📨 Sending message to {phone_number}...")

        # Open WhatsApp Web
        subprocess.Popen([chrome_path, url])

        # Allow time for page to load
        time.sleep(15)

        # Execute AppleScript to press Return and close the tab
        subprocess.run(applescript, shell=True)

        print(f"✅ Message sent to {phone_number}")
    except Exception as e:
        print(f"❌ Failed to send message to {phone_number}. Error: {str(e)}")

    # Wait before sending the next message
    time.sleep(delay)