# 🤖 RobotSpareBin Sales Automation

Automated workflow built using **Robocorp RPA Framework** to process weekly sales data from an Excel file, submit records via a web interface, and generate a PDF report.

---

## 📌 Overview

This robot performs an end-to-end automation process:

1. Opens the RobotSpareBin intranet portal
2. Logs in with user credentials
3. Downloads weekly sales data (Excel)
4. Iterates through each record and submits a sales form
5. Captures results
6. Exports results into a PDF report
7. Logs out of the system

---

## ⚙️ Tech Stack

- **Robocorp**
- **Python**
- **RPA Framework Libraries**
  - `RPA.HTTP`
  - `RPA.Excel.Files`
  - `RPA.PDF`
- **Browser Automation (Playwright via Robocorp)**

---

## 📂 Project Structure

.
├── tasks.py # Main automation script  
├── output/  
│ ├── sales_summary.png # Screenshot of results  
│ └── sales_results.pdf # Generated PDF report  
├── SalesData.xlsx # Downloaded input file  
└── README.md # Project documentation

---

## 🚀 How It Works

### 1. Browser Initialization

Configures the browser with slow motion for stability and debugging.

### 2. Authentication

Logs into the system using predefined credentials.

### 3. Data Retrieval

Downloads the Excel file containing sales records:  
https://robotsparebinindustries.com/SalesData.xlsx

### 4. Data Processing

- Reads Excel data into a structured format
- Iterates through each row
- Submits the form dynamically

### 5. Result Capture

Takes a screenshot of the final summary page.

### 6. PDF Export

Extracts HTML content and converts it into a PDF report.

### 7. Logout

Terminates the session securely.

---

## 🧠 Key Functions

| Function                       | Description                  |
| ------------------------------ | ---------------------------- |
| `robot_spare_bin_python()`     | Main task orchestrator       |
| `open_the_intranet_website()`  | Opens target web application |
| `log_in()`                     | Authenticates user           |
| `download_excel_file()`        | Downloads sales data         |
| `fill_form_with_excel_data()`  | Processes Excel records      |
| `fill_and_submit_sales_form()` | Submits individual entries   |
| `collect_results()`            | Captures screenshot          |
| `export_as_pdf()`              | Generates PDF report         |
| `log_out()`                    | Ends session                 |

---

## 🔐 Security Note

Credentials are currently hardcoded:

username: maria  
password: thoushallnotpass

⚠️ For production:

- Use **Robocorp Vault** or environment variables
- Avoid storing credentials in code

---

## ▶️ Running the Robot

1. Install dependencies:  
   pip install -r requirements.txt

2. Run the robot:  
   rcc run

---

## 📸 Output

After execution, the following files are generated:

- `output/sales_summary.png` → Screenshot of results
- `output/sales_results.pdf` → Final PDF report

---

## 🧩 Future Improvements

- Integrate secure credential management
- Add retry/error handling for form submissions
- Validate Excel data before processing
- Improve logging and reporting
- Parameterize inputs (URL, file path, etc.)

---

## 👨‍💻 Author

**Siddhartha Thapa**
