# AccEquity Excel Automation

Automated workflow that processes companies from `ace.csv`, refreshes data in an Excel template using the **AccEquity XL NXT** add-in via COM automation, saves company-specific Excel files, and uploads them to **tickercharcha.com** via API.

---

## 📁 Project Structure

```
ace_data/
├── .env                    ← Configuration (fill this in)
├── main.py                 ← Entry point — run this
├── config.py               ← Reads settings from .env
├── excel_automation.py     ← Excel COM automation (pywin32)
├── api_client.py           ← API auth + file upload
├── logger_setup.py         ← Logging configuration
├── requirements.txt        ← Python dependencies
├── ace.csv                 ← Your company list (you provide this)
├── template.xlsx           ← Your Excel template (you provide this)
└── output/                 ← Generated Excel files (auto-created)
```

---

## 🚀 Setup (Windows Machine)

### 1. Prerequisites

- **Windows 10/11** with **Microsoft Excel** installed
- **AccEquity XL NXT** Excel add-in installed and working
- **Python 3.10+** installed

### 2. Install Dependencies

```cmd
cd C:\ace_data
pip install -r requirements.txt
```

### 3. Configure `.env`

Open `.env` and fill in **all values**:

| Variable | Description |
|---|---|
| `API_LOGIN_URL` | Login endpoint (e.g. `https://tickercharcha.com/api/login`) |
| `API_UPLOAD_URL` | Upload endpoint (e.g. `https://tickercharcha.com/api/upload`) |
| `API_USER_ID` | Your login username/ID |
| `API_PASSWORD` | Your login password |
| `CSV_FILE_PATH` | Full path to `ace.csv` |
| `TEMPLATE_FILE_PATH` | Full path to `template.xlsx` |
| `OUTPUT_FOLDER_PATH` | Folder for generated files |
| `ACCORD_CODE_CELL` | Cell to write code into (default: `A1`) |
| `REFRESH_WAIT_SECONDS` | Wait time after refresh (default: `30`) |

### 4. Place Your Files

- Put `ace.csv` at the path specified in `CSV_FILE_PATH`
- Put `template.xlsx` at the path specified in `TEMPLATE_FILE_PATH`

### 5. Run

```cmd
python main.py
```

---

## 📋 ace.csv Format

```csv
Accord Code,Company Name
12345,ABC Ltd
67890,XYZ Industries
```

- **Accord Code** — numeric company code written to Excel
- **Company Name** — used for the output filename

---

## 🔄 Workflow Per Company

```
1. Copy template.xlsx → working copy
2. Open in Excel via COM (headless)
3. Write Accord Code to cell A1
4. Trigger RefreshAll → AccEquity XL NXT fetches data
5. Wait 30 seconds (configurable)
6. Save as "<Company Name>.xlsx" in output folder
7. Upload file to tickercharcha.com API
8. Log result → move to next company
```

---

## 🔐 Authentication

- Logs in once at startup using `API_USER_ID` + `API_PASSWORD`
- Token is auto-detected from response (`token`, `access_token`, `auth_token`)
- On HTTP 401/403 during upload → **automatically re-authenticates**

---

## ⚠️ Error Handling

| Scenario | Behavior |
|---|---|
| Excel fails to open | Logged, skips to next company |
| Refresh fails | Retries up to `EXCEL_MAX_RETRIES` times |
| Token expires | Auto re-authenticates |
| Upload fails | Retries up to `UPLOAD_MAX_RETRIES` times |
| Any company fails | Continues with remaining companies |

---

## 📝 Logs

| File | Contents |
|---|---|
| `automation.log` | Full audit trail (every action, timestamp, result) |
| `errors.log` | Only failures (for quick review) |
| Console | Real-time progress |

---

## 🛠 Customization

- **Change refresh wait time**: Edit `REFRESH_WAIT_SECONDS` in `.env`
- **Change target cell**: Edit `ACCORD_CODE_CELL` in `.env`
- **Adjust retry counts**: Edit `EXCEL_MAX_RETRIES` / `UPLOAD_MAX_RETRIES` in `.env`
- **API payload fields**: Edit `api_client.py` → `upload_file()` method if the API expects different form fields
