# 📘 Intune App Group Assignment Manager - PowerShell GUI Tool

## Overview

This PowerShell-based GUI tool allows administrators to **assign or remove Azure AD groups** as *inclusions* or *exclusions* to **Intune-deployed apps**. It leverages the Microsoft Graph API to automate and streamline app assignment management, and supports both selective and bulk operations.

<img width="755" height="760" alt="image" src="https://github.com/user-attachments/assets/5fb3b0cc-634f-44cc-be04-c4f1f3e824b7" />

---

## ✨ Features

- ✅ Assign **AAD Groups** as **inclusion** targets to selected apps.
- ✅ Assign **AAD Groups** as **exclusion** targets to selected apps.
- ✅ Remove group assignments from selected apps.
- ✅ Bulk exclude a group from **all** Intune apps.
- ✅ Bulk remove a group from **all** apps.
- ✅ Export all current app assignments to CSV.
- ✅ Export assignment details of selected apps only.
- 🪟 GUI built using **Windows Forms** for user-friendly operation.

---

## 🛠 Prerequisites

- PowerShell 5.1 or later
- Microsoft Graph PowerShell SDK:
  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser
  ```
- Required Graph API Permissions (delegated or app):
  - `DeviceManagementApps.ReadWrite.All`
  - `Group.Read.All`

---

## 🔐 Authentication

The script uses:
```powershell
Connect-MgGraph -Scopes "DeviceManagementApps.ReadWrite.All", "Group.Read.All"
```
If authentication fails, a GUI message is shown and execution stops.

---

## 🖥️ GUI Layout

| Section | Description |
|--------|-------------|
| **Group Name Input** | Textbox to enter the target AAD group name |
| **App Name Filter** | Optional comma-separated app names to filter which apps to process |
| **Log Viewer** | Real-time console-style output showing action results |
| **Buttons** | Each triggers a specific action (see below) |

---

## 🚀 Button Functions

| Button | Function |
|--------|----------|
| 🛑 **Stop** | Aborts long-running operations |
| ➕ **Add Inclusion (Selected Apps)** | Assigns group as required target to filtered apps |
| 🚫 **Add Exclusion (Selected Apps)** | Assigns group as exclusion target to filtered apps |
| 🚫 **Bulk Exclude Group (All Apps)** | Adds exclusion assignment to all apps for the given group |
| 🧹 **Bulk Remove Group (All Apps)** | Removes all group assignments (inclusion or exclusion) from all apps |
| 🧹 **Remove Group (Selected Apps)** | Removes group assignment from specified apps only |
| 📤 **Export All App Assignments** | Outputs a CSV of all app-to-group assignments |
| 📤 **Export Selected App Details** | Outputs assignment details of only selected apps to CSV |

---

## 🧩 Function Definitions

### `Resolve-AADGroup`
Finds the Azure AD group object based on display name.

### `Get-AppAssignments`
Retrieves current group assignments for a given app.

### `Add-GroupAsInclusionToApps`
Adds the group as a **required** target to filtered apps.

### `Add-GroupAsExclusionToApps`
Adds the group as an **exclusion** target to filtered apps.

### `Remove-GroupExclusionFromApps`
Removes group assignment from filtered apps.

### `Add-GroupAsBulkExclusionToAllApps`
Adds group as exclusion target to **all** apps.

### `Remove-GroupAssignmentFromAllApps`
Removes any assignment (inclusion/exclusion) of the group from all apps.

### `Export-AllAppAssignments`
Exports all app assignments across tenant to CSV.

### `Export-SelectedAppDetails`
Exports assignment data of specific apps to CSV.

---

## 📄 Output Files

- `Intune-All-App-Assignments.csv`
- `Intune-Selected-App-Details.csv`

> These are saved to the current user’s **Desktop** by default.

---

## 📌 Tips

- App names do **partial match**, so `Excel` matches `Excel 365`.
- Assignments respect **Graph API limits**, so large tenants may take a few minutes.
- `Stop` button is responsive and halts processing immediately.

---

## 🔧 Customization Ideas

- Add filtering by **App Type** (e.g., `Win32`, `iOS`).
- Add logging to a file (e.g., `C:\Logs\IntuneAssignmentTool.log`).
- Add profile-based configuration saving for repeatable operations.

---

## 📎 Sample Use Case

### Scenario: Add a group `Finance Users` to all apps containing `Office`
- Enter `Finance Users` in the group name field.
- Enter `Office` in the app name filter.
- Click `➕ Add Inclusion (Selected Apps)`.

---

## 👨‍💻 Author & Maintenance

This tool was designed for use by **Intune administrators** managing large-scale app assignments across environments.

For updates or contributions, package the GUI as `.ps1`, or compile with `ps2exe` to a standalone `.exe`.
