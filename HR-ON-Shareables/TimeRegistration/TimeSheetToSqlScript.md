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
2.  Rename the first tab to **`Config`**.
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
1.  Click **HR-ON Automation > Step 2: Calendars > Create/Reset Shift Calendars**.
2.  Enter the Start and End date for your school year.
3.  The script will create tabs like `Plan_Morning`, `Plan_Late`.

### ⚠️ Critical Final Step: Enable Multi-Select
*This step must be done manually for each Plan sheet to enable the "Chip" view.*

1.  Go to a planner sheet (e.g., `Plan_Morning`).
2.  **Select the entire grid** (from cell B2 down to the end).
3.  Go to **Data > Data validation**.
4.  Click the rule on the right sidebar.
5.  Under **Advanced options**:
    * Display style: **Chip**.
    * Check box: **Allow multiple selections**.
6.  Click **Done**.

---

## ✅ Phase 5: Daily Usage Workflow

### 1. Plan Schedules
* Open a Plan sheet (e.g., `Plan_Morning`).
* Find the Date (Column) and Day (Row).
* Select employees from the dropdown. You can select multiple people per cell.

### 2. Process Data
* When you are ready to finalize, click **HR-ON Automation > Step 3: Process > Process All Calendars to SQL**.
* This reads your visual calendar and converts it into data rows in the `Schedule SQL Generator` sheet.

### 3. Send to IT/HR
* Click **HR-ON Automation > Generate & Email SQL**.
* Confirm the email address.
* The system will convert the data into SQL queries and email them to the developer.
* The rows in the sheet will be marked as "Sent".

---

## ❓ Troubleshooting

**Error: "API response was not a list"**
* Check the `USERS_API_URL` in the Config sheet. It might have a typo.

**Error: "No sheets found starting with Plan_"**
* You haven't run the "Create/Reset Shift Calendars" step yet.

**Dropdowns disappeared in the Calendar**
* Run **Step 2: Create/Reset Shift Calendars** again. This repairs the data validation links.

**Email not sending**
* Check if you have a `DeveloperEmail` set in the Config sheet, or ensure you typed a valid email in the popup box.
