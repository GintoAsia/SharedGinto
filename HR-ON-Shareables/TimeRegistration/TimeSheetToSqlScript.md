---

# HR-ON Schedule Automation Tool: Setup Guide

This tool automates the process of creating employee schedules in Google Sheets and converting them into SQL queries for the HR-ON database.

## 📋 Prerequisites

Before you begin, ensure you have:
1.  **Google Sheets Access** (Permission to create scripts).
2.  **HR-ON API Credentials** (Client ID and Client Secret).
3.  **HR-ON Admin Access** (To find Project/Break IDs manually).

---

## 🚀 Phase 1: Prepare the Spreadsheet

1.  Create a **New Google Sheet**.
2.  Click **HR-ON Automation > 🚀 INITIALIZE SHEET (Run First)**.
3.  This will automatically create and style all required sheets:
    * `Settings` - Configuration and shift definitions
    * `Employee_Database` - Employee list with departments
    * `Planning_Calendar` - Schedule calendar
    * `SQL_Output` - Generated SQL queries

---

## ⚙️ Phase 2: Install the Script

1.  In your Google Sheet, go to **Extensions > Apps Script**.
2.  **Delete** any code currently in the `Code.gs` file.
3.  **Paste** the "Master Script" provided by your developer.
4.  Click the **Save** icon (💾).
5.  **Refresh** your Google Sheet browser tab.
    * *You should now see a custom menu called **"HR-ON Automation"** at the top.*

---

## 🛠 Phase 3: Configuration

### 1. Setup API Keys
Go to the **`Settings`** sheet. The System Configuration section contains the following fields:

| Row | Column A (Key) | Column B (Value) |
| :--- | :--- | :--- |
| **2** | `CLIENT_ID` | *(Paste your HR-ON Client ID)* |
| **3** | `CLIENT_SECRET` | *(Paste your HR-ON Client Secret)* |
| **4** | `TOKEN_URL` | `https://auth.hr-on.com/oauth2/token` |
| **5** | `USERS_API_URL` | `https://api.hr-on.com/v1/staff/employees?size=1000` |
| **6** | `COMPANY_API_URL` | `https://api.hr-on.com/v1/staff/company` |
| **7** | `DEPARTMENTS_API_URL` | `https://api.hr-on.com/v1/staff/departments` |
| **8** | `Default Company ID` | *(Your Company UUID, e.g., 3f66...)* |

### 2. Fill Reference IDs
In the **Reference Data** section (starting at row 12), enter your Presence Types, Break Types, and Project IDs:
* **Column A:** Presence Name → **Column B:** Presence ID
* **Column C:** Break Name → **Column D:** Break ID  
* **Column E:** Project Name → **Column F:** Project ID

> **Tip:** You can find these IDs by looking at the URL when editing a project in the HR-ON web portal.

### 3. Fetch Employees
1.  Click **HR-ON Automation > Step 1: Data Setup > 1. Refresh Employees (Sorted by Dept)**.
2.  This will pull all current staff into the `Employee_Database` sheet, **sorted by department**.

### 4. Update Dropdowns
1.  Click **HR-ON Automation > Step 1: Data Setup > 2. Refresh Dropdowns (Internal)**.
2.  This links the Reference Data to your Shift Definitions table.

---

## 📅 Phase 4: Setup Templates & Calendars

### 1. Define Your Shifts
Go to the **Shift Definitions** section in the Settings sheet (starting at row 35):

| Column | Description |
| :--- | :--- |
| **A** | Shift Name (e.g., "Morning A") |
| **B** | Start Time (e.g., "08:00") |
| **C** | End Time (e.g., "16:00") |
| **D** | Break Start |
| **E** | Break End |
| **F** | Presence Type (dropdown) |
| **G** | Break Type (dropdown) |
| **H** | Project (dropdown) |

### 2. Generate the Calendar
1.  Click **HR-ON Automation > Step 2: Scheduling > Create/Reset Calendar**.
2.  Enter the Start and End date for your schedule period.
3.  The script will create the `Planning_Calendar` sheet with:
    * **Row 1:** Department headers
    * **Row 2:** Employee names (sorted by department)
    * **Row 3+:** Date rows with weekend highlighting

### 3. Assign Shifts by Department
1.  Click **HR-ON Automation > Step 2: Scheduling > 🏢 Assign Shifts by Department**.
2.  Select a department, shift, and days of the week.
3.  All employees in that department will be assigned the selected shift on the specified days.

### 4. Clear Calendar
1.  Click **HR-ON Automation > Step 2: Scheduling > 🗑️ Clear Calendar**.
2.  Confirm to remove all shift assignments (headers and structure remain).

### 5. Validate Shifts
1.  Click **HR-ON Automation > Step 2: Scheduling > ✅ Validate Calendar Shifts**.
2.  The tool will check all assigned shifts and highlight any that aren't defined in Settings.
3.  Option to automatically add missing shift names to Settings.

### 6. Cleanup Sheet
1.  Click **HR-ON Automation > Step 2: Scheduling > 🧹 Cleanup Sheet**.
2.  This will:
    * Remove unused rows and columns
    * Auto-resize columns
    * Apply consistent styling
    * Re-apply weekend highlighting

### ⚠️ Optional: Enable Multi-Select for Manual Edits
*This step is only needed if you want to manually edit cells with multiple shifts.*

1.  Go to the `Planning_Calendar` sheet.
2.  **Select the grid area** (from cell B3 down to the end).
3.  Go to **Data > Data validation**.
4.  Click the rule on the right sidebar.
5.  Under **Advanced options**:
    * Display style: **Chip**.
    * Check box: **Allow multiple selections**.
6.  Click **Done**.

---

## ✅ Phase 5: Daily Usage Workflow

### 1. Plan Schedules
* Use **🏢 Assign by Department** to quickly assign shifts to entire departments.
* Make individual adjustments as needed by clicking cells directly.
* Use **🧹 Cleanup Sheet** periodically to keep everything tidy.

### 2. Process Data
* When you are ready to finalize, click **HR-ON Automation > Step 3: Export > Process Calendar to SQL**.
* This reads your visual calendar and converts it into data rows in the `SQL_Output` sheet.

### 3. Send to IT/HR
* Click **HR-ON Automation > Step 3: Export > Email SQL to IT**.
* Confirm the email address.
* The system will convert the data into SQL queries and email them to the developer.
* The rows in the sheet will be marked as "Sent".

---

## ❓ Troubleshooting

**Error: "API response was not a list"**
* Check the `USERS_API_URL` in the Settings sheet. It might have a typo.

**Error: "No calendar found" or "Please create a calendar first"**
* You haven't run the "Create/Reset Calendar" step yet.

**Dropdowns disappeared in the Calendar**
* Run **Step 2: Create/Reset Calendar** again. This repairs the data validation links.

**Employees not appearing in correct department order**
* Refresh the employee data using Step 1, then recreate the calendar.

**Shift validation showing errors**
* Add the missing shift names to the Shift Definitions section in Settings.

**Email not sending**
* Ensure you typed a valid email address in the popup box.
