# 🎯 US Visa Slot Availability Monitor (F-1 Category)

This Python script monitors the latest US F-1 (Regular) visa slot availability from [checkvisaslots.com](https://checkvisaslots.com/latest-us-visa-availability.html#F1Regular), detects any updates, logs changes in an Excel file, and sends email alerts upon detecting new availability.

## 📌 Features

- Headless monitoring using Selenium and Firefox
- Extracts F-1 Regular table data from the website
- Logs new availability data in an Excel file
- Sends email alerts using Gmail when slots change
- Keeps a historical log of slot changes

## 📁 Output Files

- `visa_slot_log.xlsx`: Logs with timestamps
- `text.txt`: Last fetched table content

## 🧰 Prerequisites

- Python 3.8+
- Firefox Browser installed
- [GeckoDriver](https://github.com/mozilla/geckodriver/releases) added to PATH

## 🛠️ Installation

```bash
git clone https://github.com/your-username/visa-slot-monitor.git
cd visa-slot-monitor
pip install -r requirements.txt
```

### `requirements.txt`

```txt
selenium
beautifulsoup4
openpyxl
```

## 📧 Gmail Setup for Alerts

1. Enable 2-Step Verification for your Gmail account.
2. Generate an **App Password**:

   - Go to [https://myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
   - Select **Mail** as the app and **Other (Custom)** for device name like `VisaMonitor`
   - Copy the 16-character app password

3. Update `visa_slot_monitor.py`:
   ```python
   SENDER_EMAIL = "your-email@gmail.com"
   SENDER_PASSWORD = "your-app-password"  # paste the generated password here
   ```

## 🚀 Run the Script

```bash
python visa_slot_monitor.py
```

*Keep the script running in background (or use a cloud VM) to continuously monitor the page.*

## 🧪 Sample Output

See `/sample_output/` for:
- Example Excel logs (`visa_slot_log.xlsx`)
- Sample last fetched table content (`text.txt`)

## 🛑 To Stop Monitoring

Press `Ctrl+C` in the terminal.

## 📄 License

MIT License – free to use, modify, and share with attribution.

---

### 🙋‍♂️ Created by Abhishek Tyagi
