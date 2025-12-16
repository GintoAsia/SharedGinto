***

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
2.  Rename the first tab to **`Settings`**.
3.  Create three more blank tabs and name them **exactly** as follows (Case Sensitive):
    * `EmployeeData`
    * `ShiftTemplates`
    * `Schedule SQL Generator`

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
Go to the **`Config`** sheet. Enter the following information in **Column A** and **Column B**.

| Row | Column A (Key) | Column B (Value) |
| :--- | :--- | :--- |
| **1** | `CLIENT_ID` | *(Paste your HR-ON Client ID)* |
| **2** | `CLIENT_SECRET` | *(Paste your HR-ON Client Secret)* |
| **3** | `TOKEN_URL` | `https://auth.hr-on.com/oauth2/token` |
| **4** | `USERS_API_URL` | `https://api.hr-on.com/v1/staff/employees?size=1000` |
| **5** | `DeveloperEmail` | *(Email address where SQL queries should be sent)* |
| **6** | `DEFAULT_COMPANY_ID` | *(Your Company UUID, e.g., 3f66...)* |

### 2. Initialize Reference Data
1.  Click **HR-ON Automation > Step 1: Setup > 1. Create Blank Reference Sheets**.
2.  This will create 3 new tabs: `Ref_PresenceTypes`, `Ref_BreakTypes`, and `Ref_Projects`.

### 3. Fill Reference IDs (Manual Step)
Because these settings are static, you must fill them once manually.
* **Column A:** The Human Name (e.g., "Lunch", "Math Class").
* **Column B:** The System UUID (e.g., `cafae750-0827...`).

> **Tip:** You can find these IDs by looking at the URL when editing a project in the HR-ON web portal.

### 4. Fetch Employees
1.  Click **HR-ON Automation > Step 1: Setup > 2. Refresh Employee Data**.
2.  This will pull all current staff into the `EmployeeData` sheet.

---

## 📅 Phase 4: Setup Templates & Calendars

### 1. Define Your Shifts
Go to the **`ShiftTemplates`** sheet. Define your shift types (e.g., Morning, Late).

* **Column A (Shift_Name):** Give it a name (e.g., "Morning").
* **Columns B, E, H:** These will be blank initially.
* **Columns C, D, F, G, I, J:** Enter the times (Format: `HH:mm:ss`).

### 2. Activate Dropdowns
1.  Click **HR-ON Automation > Step 1: Setup > 3. Update Template Dropdowns**.
2.  Go back to **`ShiftTemplates`**. You can now select the Presence, Break, and Project types from the dropdown menus in Columns B, E, and H.

### 3. Generate the Calendars
1.  Click **HR-ON Automation > Step 2: Scheduling > Create/Reset Calendar**.
2.  Enter the Start and End date for your school year.
3.  The script will create the `Planning_Calendar` sheet with all employees and dates.

### 4. Bulk Assign Shifts (NEW - Recommended!)
Instead of manually clicking each cell, use these powerful bulk assignment tools:

#### 📅 Bulk Assign Shifts (Pattern-Based)
1.  Click **HR-ON Automation > Step 2: Scheduling > 📅 Bulk Assign Shifts (Pattern-Based)**.
2.  A dialog will appear with the following options:
    * **Select Shift**: Choose from your defined shift types.
    * **Select Days of Week**: Check/uncheck days (e.g., weekdays only, weekends only).
    * **Date Range (Optional)**: Limit the assignment to a specific period.
    * **Select Employees**: Choose which employees to assign.
3.  Click **Apply Shifts** to bulk-fill the calendar.

#### 🏢 Assign Shifts by Department
1.  Click **HR-ON Automation > Step 2: Scheduling > 🏢 Assign Shifts by Department**.
2.  Select a department, shift, and days of the week.
3.  All employees in that department will be assigned the selected shift on the specified days.

#### 📋 Copy Week Pattern
1.  First, manually set up one week of schedules in the calendar.
2.  Click **HR-ON Automation > Step 2: Scheduling > 📋 Copy Week Pattern**.
3.  Enter the Monday date of the week you configured.
4.  The script will copy that week's pattern to all other weeks in the calendar.

#### 🗑️ Clear Calendar
1.  Click **HR-ON Automation > Step 2: Scheduling > 🗑️ Clear Calendar**.
2.  Confirm to remove all shift assignments (employee names remain).

### ⚠️ Optional: Enable Multi-Select for Manual Edits
*This step is only needed if you want to manually edit cells with multiple shifts.*

1.  Go to the `Planning_Calendar` sheet.
2.  **Select the grid area** (from cell B2 down to the end).
3.  Go to **Data > Data validation**.
4.  Click the rule on the right sidebar.
5.  Under **Advanced options**:
    * Display style: **Chip**.
    * Check box: **Allow multiple selections**.
6.  Click **Done**.

---

## ✅ Phase 5: Daily Usage Workflow

### 1. Plan Schedules (Recommended: Use Bulk Tools)
* For a year's worth of schedules, use **📅 Bulk Assign Shifts** or **🏢 Assign by Department**.
* Set up one typical week, then use **📋 Copy Week Pattern** to replicate it.
* Make individual adjustments as needed by clicking cells directly.

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
* Check the `USERS_API_URL` in the Config sheet. It might have a typo.

**Error: "No calendar found" or "Please create a calendar first"**
* You haven't run the "Create/Reset Calendar" step yet.

**Dropdowns disappeared in the Calendar**
* Run **Step 2: Create/Reset Calendar** again. This repairs the data validation links.

**Bulk assign not working for some employees**
* Ensure employee names in the calendar match exactly with the Employee_Database sheet.

**Email not sending**
* Check if you have a `DeveloperEmail` set in the Config sheet, or ensure you typed a valid email in the popup box.
