# 🎯 US Visa Slot Availability Monitor (F-1 & J-1)

This Python script monitors the latest US visa slot availability from [checkvisaslots.com](https://checkvisaslots.com/latest-us-visa-availability.html) for **F-1 (Regular)** and **J-1 (Regular)** visas. It detects updates, logs changes in an Excel file, and sends email alerts when new availability appears.

> **Note:** checkvisaslots.com now has a separate page for each visa type (for example `/latest-us-visa-availability/f-1-regular/`). The old `#F1Regular` link no longer works, and the script has been updated for the new layout.

## 📌 Features

- Monitors **F-1** or **J-1** with one script (pick the visa when you run it)
- Uses Selenium and Firefox (can run headless)
- Reads the latest availability table: location, visa type, earliest date, slots on the earliest date, and last seen time
- Logs new availability in an Excel file, one file per visa type
- Sends email alerts through Gmail when slots change
- Keeps a history of slot changes
- Easy to extend to other visa types (H-1B, B1/B2, etc.)

## 📁 Output Files

Each visa type gets its own files, so F-1 and J-1 can run side by side:

| File | Description |
|------|-------------|
| `f1_visa_slot_log.xlsx` / `j1_visa_slot_log.xlsx` | Log of changes with timestamps |
| `f1_text.txt` / `j1_text.txt` | Last fetched table content (used to detect changes) |

**Upgrading from the old version?** You can delete the old `visa_slot_log.xlsx` and `text.txt`, because the script no longer uses them.

## 🧰 Prerequisites

- Python 3.8+
- Firefox browser installed
- [GeckoDriver](https://github.com/mozilla/geckodriver/releases): Selenium 4.6+ downloads it automatically. Only install it yourself if that fails, and add it to your PATH.

## 🛠️ Installation

```bash
git clone https://github.com/your-username/visa-slot-monitor.git
cd visa-slot-monitor
pip install -r requirements.txt
```

### `requirements.txt`

```txt
selenium>=4.6
beautifulsoup4
openpyxl
```

## 📧 Gmail Setup for Alerts

1. Turn on 2-Step Verification for your Gmail account.
2. Generate an **App Password**:

   - Go to [https://myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
   - Enter an app name such as `VisaMonitor` and click **Create**
   - Copy the 16-character app password

3. Update `visa_slot_monitor.py`:
   ```python
   SENDER_EMAIL = "your-email@gmail.com"
   SENDER_PASSWORD = "your-app-password"  # paste the generated password here
   RECEIVER_EMAILS = [
       "receiver-email1@gmail.com",
       "receiver-email2@gmail.com"
   ]
   ```

> ⚠️ **Never commit your real email or app password to GitHub.** Keep the placeholder values in the repository and only fill in your real details in your local copy.

## 🚀 Run the Script

Choose the visa type to monitor:

```bash
python visa_slot_monitor.py F1    # monitor F-1 (Regular)
python visa_slot_monitor.py J1    # monitor J-1 (Regular)
```

If you don't give a visa type, the script uses `DEFAULT_VISA` (F-1). To monitor both at once, run each command in its own terminal.

*Keep the script running in the background (or on a cloud VM) to monitor the page continuously.*

### ⚙️ Optional Settings

These are near the top of `visa_slot_monitor.py`:

| Setting | Default | Description |
|---------|---------|-------------|
| `DEFAULT_VISA` | `"F1"` | Visa type used when you don't give one |
| `CHECK_INTERVAL` | `60` | Seconds between checks. The site updates about every 2 minutes, so checking much more often doesn't help. |
| `COLUMNS_TO_KEEP` | location, type, earliest date, slots, last seen | Table columns to track |

To run without opening a browser window, uncomment this line in `main()`:

```python
options.add_argument('--headless')
```

### ➕ Adding Another Visa Type

Add an entry to `VISA_TYPES` using the slug from the site's URL:

```python
VISA_TYPES = {
    "F1":  {"slug": "f-1-regular", "label": "F-1 (Regular)", "filter": "F-1"},
    "J1":  {"slug": "j-1-regular", "label": "J-1 (Regular)", "filter": "J-1"},
    "H1B": {"slug": "h-1b-regular", "label": "H-1B (Regular)", "filter": "H-1B"},
}
```

Then run `python visa_slot_monitor.py H1B`.

## 🧪 Sample Output

See `/sample_output/` for:
- Example Excel log (`f1_visa_slot_log.xlsx`)
- Sample last fetched table content (`f1_text.txt`)

Example email alert:

```
Subject: F-1 (Regular) Slot Changed: CHENNAI VAC - earliest 08 Oct, 26 (seen 25 Sep 2026, 06:05 PM)

Changes detected in F-1 (Regular) visa slots:

Visa Location | Visa Type | Earliest Date | Slots On Earliest Date | Last Seen At
CHENNAI VAC | F-1 (Regular) | 08 Oct, 26 | 3 | 25 Sep 2026, 06:05 PM
```

## 🛑 To Stop Monitoring

Press `Ctrl+C` in the terminal.

## 📄 License

MIT License – free to use, modify, and share with attribution.

---

### 🙋‍♂️ Created by Abhishek Tyagi
