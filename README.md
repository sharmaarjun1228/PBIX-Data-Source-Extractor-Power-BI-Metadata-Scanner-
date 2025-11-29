
---

# ✅ **README.md**

# PBIX Data Source Extractor (Power BI Metadata Scanner)

A simple web-based tool that scans any folder containing **Power BI (.pbix)** files and automatically extracts:

- SQL Stored Procedures used in Power BI reports  
- SQL Views / Tables referenced  
- Native SQL queries (`SELECT`, `FROM`, `JOIN`, `EXEC`)  
- Data source information (Server, Database)  
- Power Query (M) navigation paths (`Schema`, `Item`)  

The tool generates a clean **Excel report** containing:

- **Detailed** sheet → All data source occurrences  
- **UniqueObjects** sheet → Unique views/procs with list of PBIX files where they are used  

Designed to save hours of manual work for BI teams maintaining many PBIX files.

---

## 🚀 Features

- ✔ Extract all SQL objects used inside `.pbix`
- ✔ Detect **Stored Procedures**, **Views**, and **Tables**
- ✔ Parse:
  - `Sql.Database(...)`
  - `Query = "SELECT..."`
  - `Value.NativeQuery(...)`
  - `Schema="dbo", Item="View_Name"`
- ✔ Auto-generate timestamped Excel reports
- ✔ Web UI built with **Flask**
- ✔ No need to upload PBIX files — scans your **local folder**
- ✔ Clean HTML/CSS interface  
- ✔ Portable: works on Windows, macOS, Linux

---




## 📁 Project Structure

```

pbix_extractor_ui/
│
├── app.py                 # Flask application (web server)
├── extractor.py           # Core PBIX extraction logic
│
├── templates/
│   └── index.html         # UI page
│
└── static/
└── style.css          # Stylesheet

````

---

## 🛠 Requirements

- Python 3.9+  
- Power BI Desktop installed (required by pbixray)  
- Packages:
  - Flask  
  - pandas  
  - openpyxl  
  - pbixray  

Install dependencies:

```bash
pip install flask pandas openpyxl pbixray
````

---

## ▶️ How to Run

1. Clone the repo:

```bash
git clone https://github.com/your-user-name/pbix-extractor-ui.git
cd pbix-extractor-ui
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Start the web app:

```bash
python app.py
```

4. Open your browser and go to:

```
http://127.0.0.1:5000
```

5. Enter any folder path (e.g. `D:\Reports\PowerBI`)
6. Click **Extract Data Sources**
7. Your Excel file will download automatically.

---

## 📊 Output Example (Excel)

**Sheet 1: Detailed**

| pbix_file    | table_name | source_type              | server   | database | object_name  | object_type      |
| ------------ | ---------- | ------------------------ | -------- | -------- | ------------ | ---------------- |
| report1.pbix | FactSales  | Navigation (Schema/Item) | SQLSRV01 | ProdDB   | dbo.vw_Sales | View             |
| report1.pbix | ProcData   | SQL (Query option)       | SQLSRV01 | ProdDB   | usp_GetSales | Stored Procedure |

**Sheet 2: UniqueObjects**

| object_name  | object_type      | pbix_files                 |
| ------------ | ---------------- | -------------------------- |
| usp_GetSales | Stored Procedure | report1.pbix; report2.pbix |
| vw_Sales     | View             | report1.pbix               |

---

## 📸 Screenshots (optional)

> Add screenshots of your UI here
> Example:
> `![Screenshot](screenshots/home.png)`

---

## 🙌 Credits

**Made with 💞 by Arjun Sharma**
If this tool saved your time, please ⭐ the repo on GitHub!

---

## 📄 License

This project is open-source under the **MIT License**.

```

---

If you want, I can also create:

✅ `requirements.txt`  
✅ `LICENSE` file  
✅ GitHub-friendly folder structure  
✅ A GitHub banner (header image)  
✅ Badges (Python version, License, Stars)  

Just tell me — I'll prepare everything.
```
