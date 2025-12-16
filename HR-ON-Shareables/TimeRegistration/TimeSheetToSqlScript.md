***

# 📅 HR-ON Schedule Automation Dashboard

This tool automates the process of scheduling employees, assigning shifts by department or bulk patterns, and converting that data into SQL queries for the HR-ON database.

## 📋 Prerequisites

Before you begin, ensure you have:
1.  **Google Sheets Access** (Permission to create scripts).
2.  **HR-ON API Credentials** (Client ID and Client Secret).
3.  **Reference IDs** (UUIDs for Presence Types, Break Types, and Projects from HR-ON).

---

## 🚀 Phase 1: Installation & Initialization

### 1. Install the Script
1.  Open a new Google Sheet.
2.  Go to **Extensions > Apps Script**.
3.  Paste the provided code into the editor (replace any existing code).
4.  Click **Save** (💾).
5.  Refresh the Google Sheet tab.

### 2. Run the Initializer
You will see a new menu called **HR-ON Automation** at the top.
1.  Click **HR-ON Automation > 🚀 INITIALIZE SHEET (Run First)**.
2.  Grant the necessary permissions if asked.
3.  **Result:** The script will automatically create four color-coded sheets and apply professional formatting.

---

## ⚙️ Phase 2: Configuration (The Settings Sheet)

Go to the **`Settings`** sheet. This is your control center. It is divided into three color-coded blocks:

### 🔵 1. System Configuration (Blue Table)
Enter your API credentials here.

| Key | Value |
| :--- | :--- |
| **CLIENT_ID** | *(Paste your HR-ON Client ID)* |
| **CLIENT_SECRET** | *(Paste your HR-ON Client Secret)* |
| **TOKEN_URL** | `https://auth.hr-on.com/oauth2/token` |
| **USERS_API_URL** | `https://api.hr-on.com/v1/staff/employees?size=1000` |

### 🟡 2. Reference Data (Yellow Table)
You must manually paste the names and IDs from HR-ON here. This allows the script to map a human name (e.g., "Lunch") to a database ID.

* **Columns A & B:** Presence Types (e.g., "Normal Work")
* **Columns C & D:** Break Types (e.g., "Lunch")
* **Columns E & F:** Projects (e.g., "Math Class")

### 🟢 3. Shift Definitions (Green Table)
Define the shifts you want to use in your calendar.

1.  **Shift Name:** Give it a short code (e.g., "Morning", "Night").
2.  **Start/End Times:** Format as `HH:mm`.
3.  **Dropdowns:** Once you fill in the Yellow table, run **Step 1: Data Setup > 2. Refresh Dropdowns**. You can then select Presence, Break, and Projects from the dropdowns in columns F, G, and H.

---

## 👥 Phase 3: Employee Setup

1.  Click **HR-ON Automation > Step 1: Data Setup > 1. Refresh Employees**.
2.  The script will pull data from the API.
3.  Go to the **`Employee_Database`** sheet. You will see your staff list, sorted automatically by **Department**, then by **Name**.

---

## 📅 Phase 4: Scheduling Workflow

### 1. Create the Calendar
1.  Click **HR-ON Automation > Step 2: Scheduling > Create/Reset Calendar**.
2.  Enter the Start and End date.
3.  **Result:** A `Planning_Calendar` sheet is created with Dates on the rows (Y-axis) and Employees on the columns (X-axis).

> **Pro Tip:** For a better visual experience, highlight the calendar grid, go to **Data > Data validation**, and enable **Chip** display style and **Allow multiple selections**.

### 2. Assigning Shifts
You have three ways to schedule:

**A. Manual Selection**
* Click any cell in the grid and select a shift from the dropdown.

**B. 📅 Bulk Assign (Pattern-Based)**
* Click **Step 2: Scheduling > Bulk Assign Shifts**.
* A popup will appear. Select a Shift, choose Days of the Week (e.g., Mon-Fri), and check the specific employees you want to assign.
* Click **Apply**.

**C. 🏢 Assign by Department**
* Click **Step 2: Scheduling > Assign Shifts by Department**.
* Select a Department (e.g., "Science Dept") and a Shift.
* The script will find all employees in that department and assign the shift to them for the selected days.

### 3. Copying Patterns
If you have a perfect week set up:
1.  Click **Step 2: Scheduling > 📋 Copy Week Pattern**.
2.  Enter the date of the **Monday** you want to copy.
3.  The script will replicate that week's schedule across the rest of the calendar.

---

## ✅ Phase 5: Validation & Export

### 1. Validate Data
Before exporting, ensure you haven't typed a shift name that doesn't exist in your Settings.
1.  Click **Step 2: Scheduling > ✅ Validate Calendar Shifts**.
2.  The script scans every cell.
3.  If it finds invalid data, it will highlight the bad cells in **Red** and ask if you want to add those shift names to your Settings automatically.

### 2. Process to SQL
1.  Click **HR-ON Automation > Step 3: Export > Process Calendar to SQL**.
2.  The script reads the visual calendar, matches the Shift Names to the IDs in Settings, and generates the data rows.
3.  Check the **`SQL_Output`** sheet to review the data.

### 3. Email IT
1.  Click **HR-ON Automation > Step 3: Export > Email SQL to IT**.
2.  Enter the recipient address.
3.  The script formats the data into `INSERT` SQL statements and emails them.
4.  Rows are marked as "Sent" in the sheet.

---

## ❓ Troubleshooting

| Issue | Solution |
| :--- | :--- |
| **"User API URL missing"** | Check the Blue table in the `Settings` sheet. Ensure column names match exactly. |
| **Dropdowns are empty in Settings** | Run **Step 1 > 2. Refresh Dropdowns** after pasting data into the Yellow table. |
| **"No employees found"** | Run **Step 1 > 1. Refresh Employees** before creating the calendar. |
| **Dates are wrong in SQL** | The script auto-converts dates to UTC. Ensure your spreadsheet timezone (File > Settings) is correct. |

---

### 🎨 Visual Guide to Sheets

* 🟦 **Settings:** Configuration & Reference Data.
* teal **Employee_Database:** Read-only list of staff.
* 🟨 **Planning_Calendar:** Where you do the work.
* 🟥 **SQL_Output:** The final data for export.
