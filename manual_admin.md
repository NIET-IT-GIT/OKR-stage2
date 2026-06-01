# NIET OKR Platform — Admin User Manual

## Overview

As an Admin, you have full access to the NIET OKR Performance Tracker. You can manage users, configure the organisational structure, monitor performance across all departments, and publish company-wide reports.

---

## Logging In

1. Open the application in your browser.
2. Click **Sign in with Microsoft** to authenticate with your NIET account.
3. Alternatively, use the admin credentials (username: `admin`) for system-level access.
4. Once authenticated, you will land on the **Company Overview** dashboard.

---

## Navigation

The sidebar on the left contains all pages available to you:

| Page | Purpose |
|---|---|
| Company Overview | Organisation-wide performance at a glance |
| Departments | Manage department structure and KRs |
| OKR / KPI Setup | Configure key results for departments and teams |
| Weekly Submissions | Review all staff weekly submissions |
| Monthly Reports | Generate and publish company-wide reports |
| Projects | View all manager-submitted projects |
| Leaderboard | Company-wide staff rankings |
| User Management | Create, edit, and delete user accounts |

---

## Company Overview

This is your main dashboard. It shows:

- **Company completion rate** — the aggregated KPI completion percentage across all departments.
- **Department cards** — each card shows the department name, completion rate, and status indicator (green = on track, yellow = at risk, red = behind).
- **Summary metrics** — total departments, teams, and tracked staff.

Click any department card to navigate directly to that department's detail view.

---

## Departments

### Viewing Departments

Each department row shows the department name, college, completion rate, team count, and member count. Click a department row to expand it and see:

- Department objective
- Completion rate, team count, and member count
- Key results (KRs) with targets, actuals, units, and data sources
- Team structure with leads and team-level KRs
- Individual member performance ranked by completion rate

### Creating a Department

1. Click **+ Add Department** (or the equivalent add button at the top of the page).
2. Enter: department name, head, college, and objective.
3. Save to create the department.

### Editing a Department

Click the edit icon next to a department to update its name, objective, head, or college.

### Deleting a Department

Click the delete (✕) button next to the department. You will be asked to confirm before deletion. This cannot be undone.

### Adding Key Results to a Department

Within an expanded department, use the **Add KR** control to create a new key result. The **description/label** and **target value** are required. **Unit** and **Data Source** are optional and can be left blank — you can fill them in later via the OKR / KPI Setup page.

---

## OKR / KPI Setup

Use this page to configure key results for departments and their teams.

1. Select a **department** from the list on the left.
2. Select a **team** within that department (or work at the department level).
3. The KR table displays all current key results across these columns: **ID**, **Key Result**, **Target**, **Actual**, **Unit**, and **Data Source**.
4. To **add a KR**: fill in the row at the bottom of the table and click Add. Only the **Key Result** (description) and **Target** fields are required — **Unit** and **Data Source** can be left blank and filled in later.
5. To **edit a KR**: click the value in the table to modify target, unit, or data source inline.
6. To **delete a KR**: click the delete button on the KR row.
7. To **update actuals**: edit the actual value field directly.

### Adding Custom Columns

Click **+ Add Column** (top-right of the department section) to add an extra column to all KR tables on this page. Enter a column name and press **Add** or hit Enter. The new column appears in both the Department KR table and the Team KR table.

- To **rename** a custom column, click its header label and type the new name.
- To **remove** a custom column, click the small ✕ next to its header label.
- Custom column values are saved per KR — type directly into the cell for each row.

### Adjusting Column Widths

Every column header has a drag handle on its right edge (shown as a faint vertical bar). Drag it left or right to resize that column. This applies to all built-in columns (ID, Key Result, Target, Actual, Unit, Data Source) and any custom columns you have added. If the total table width exceeds the page width, the table scrolls horizontally.

---

## Weekly Submissions

This page shows all staff weekly submissions across the entire company.

### Filtering

Use the filter tabs to show: **All**, **Pending**, **Approved**, or **Rejected** submissions.

### What You Can See

Each submission shows:
- Staff member name, title, and department
- Their manager's name
- Submission week and date
- Work outcome items (what they reported doing)
- Current approval status

**Note:** Admins can view all submissions, but approvals are performed by managers on their own Approve Submissions page.

---

## Monthly Reports

### Generating a Report

1. Click **Generate Report** to create a report for the current month. The system will automatically calculate the company completion rate, identify the top 3 performers, and flag staff performing below 60%.
2. To generate a report for a custom period, click **Custom Report**, enter a period label (e.g. Q1 FY2026) and date range, then generate.

### Published Reports

Each published report shows:
- Department rankings by completion rate
- Top performers
- Staff who require action (below threshold)

### Deleting a Report

Click the delete button on a report. Reports are visible to all staff once published.

---

## Projects

This page shows all projects submitted by department managers, grouped by department.

### What You Can See

- Total projects, active projects, and completed projects (summary metrics at top)
- Projects listed under their department heading
- Each project shows: name, manager name, due date, last updated date, completion %, status, and any notes

### Editing a Project

1. Click **Edit** on a project card.
2. Update the project name, completion %, status (Active/Completed), due date, or log/notes.
3. Click **Save** to confirm.

### Deleting a Project

Click the **✕** button on a project card and confirm the deletion prompt.

---

## Leaderboard

Displays all staff ranked by KPI completion rate from highest to lowest.

Each entry shows: rank, avatar, name, title, department, completion rate, progress bar, and status tag (On Track / At Risk / Behind).

---

## User Management

### Viewing Users

The user list shows all accounts with their name, email, title, role badge, and department. Summary metrics at the top show total users broken down by role.

Use the **search bar** to filter users by name or email.

### Creating a User

1. Click **+ Add User**.
2. Fill in: full name, email, role (Admin / Manager / Member), job title.
3. If the role is Manager or Member, assign a department.
4. If the role is Manager, assign one or more teams.
5. Save to create the account.

### Editing a User

Click the **Edit** button next to a user to update any of their details inline:

- **Name**, **Email**, **Title**, **Role**
- **Department** — reassign the user to a different department (not applicable to Admin accounts)
- **Team** — once a department is selected, a team selector appears:
  - For **Members**: choose a single team from a dropdown
  - For **Managers**: tick one or more teams using checkboxes

Changing a user's role or department automatically clears their previous team assignment so you can make a fresh selection.

### Deleting a User

Click the **Delete** button next to a user and confirm. You cannot delete your own account or the system admin account.

---

## Signing Out

Click **Sign Out** at the bottom of the sidebar. This clears your session and returns you to the login screen.
