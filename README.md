# AD Manager — PowerShell Active Directory Management Tool

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-0078D4?style=for-the-badge&logo=powershell&logoColor=white)](https://github.com/PowerShell/PowerShell)
[![Windows](https://img.shields.io/badge/Windows-Server%202016%2B-0078D4?style=for-the-badge&logo=windows&logoColor=white)]()
[![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)]()
[![Status](https://img.shields.io/badge/Status-Simulation%20Ready-yellow?style=for-the-badge)]()

Interactive PowerShell tool for managing Active Directory users, groups, and OUs. 17-option menu, `-WhatIf` simulation mode, audit logging, and bulk CSV operations.

---

## Why I built this

During my ASIR degree I spent a lot of time working with Windows Server and Active Directory. Managing things through `dsa.msc` gets slow fast — creating 20 users or reorganizing OUs from the GUI is tedious work that shouldn't require clicking through menus.

I built AD Manager to practice automating those tasks properly with PowerShell: a single script that covers the most common AD operations, logs everything it does, and has a `-WhatIf` mode so you can safely rehearse before touching a real environment.

> All Active Directory cmdlets are included as comments. The script runs in simulation mode out of the box — no AD module required to test it.

---

## Features

| Feature | Description |
|---|---|
| Interactive Menu | 17 options organized in 4 categories |
| User Management | Bulk import, create, search, enable/disable, reset password, move, delete |
| Group Management | Bulk import, add/remove members, list members, delete |
| OU Management | Bulk import, list, delete with recursive safety check |
| Export | Export user list or last search result to CSV |
| Audit Log | Every action logged with timestamp, mode and result |
| WhatIf Mode | Simulate all actions — zero real changes applied |
| Admin Check | Detects privilege level at startup |
| Smart CSV Search | Finds CSV files on the system, warns on duplicates |
| Error Handling | try/catch on every operation with detailed logging |

---

## Requirements

- Windows PowerShell 5.1+ or PowerShell 7+
- Windows Server 2016+ *(for live AD operations only)*
- RSAT Active Directory module *(only needed for real execution, not for simulation)*

---

## Getting Started

### Run in simulation mode (no AD needed)

```powershell
.\AD_Manager.ps1
```

### Run with WhatIf — no changes will be applied

```powershell
.\AD_Manager.ps1 -WhatIf
```

### Custom log file name

```powershell
.\AD_Manager.ps1 -LogFileName "MyAudit.log"
```

---

## WhatIf Mode

WhatIf mode lets you rehearse every action safely before touching a real environment. When active, a warning banner appears on every menu load:

```
[!] WHATIF MODE ACTIVE - No changes applied
```

Instead of executing, every modifying action outputs:

```
What if: Performing the operation "Delete User" on target "AD User: asmith"
```

**Covered by WhatIf** (no changes made): creating, moving and deleting users · enabling and disabling accounts · resetting passwords · bulk importing users, groups and OUs · deleting groups and OUs · exporting data

**Not affected** (always read-only): searching users · listing group members · listing OUs · viewing the audit log

> Always run `-WhatIf` first when deploying in a new environment.

---

## Menu

```
=========================================
        ACTIVE DIRECTORY MANAGER
=========================================
 -- USER MANAGEMENT --
  1.  Bulk import users from CSV
  2.  Create single user manually
  3.  Search user (name, login or email)
  4.  Enable / Disable user account
  5.  Reset user password
  6.  Move user to another OU
  7.  Delete user permanently

 -- GROUP MANAGEMENT --
  8.  Bulk import groups from CSV
  9.  Add / Remove user from group
  10. List all members of a group
  11. Delete group

 -- ORGANIZATIONAL UNITS --
  12. Bulk import OUs from CSV
  13. List all OUs and their users
  14. Delete OU with safety check

 -- REPORTS & EXPORT --
  15. Export full user list to CSV
  16. Export last search result to CSV
  17. Show audit log (last 50 entries)

  X.  Exit
=========================================
```

---

## CSV Format

### users.csv

```csv
Name,SamAccountName,GivenName,Surname,Path,EmailAddress,Department,OfficePhone
"Alice Smith",asmith,Alice,Smith,"OU=Users,DC=contoso,DC=com",asmith@contoso.com,IT,555-0101
```

| Column | Description |
|---|---|
| Name | Full display name |
| SamAccountName | Login / account name |
| GivenName | First name |
| Surname | Last name |
| Path | OU distinguished name |
| EmailAddress | Corporate email address |
| Department | Department name |
| OfficePhone | Office phone number |

### groups.csv

```csv
Name,GroupCategory,GroupScope,Path,Description
"IT Admins",Security,Global,"OU=Groups,DC=contoso,DC=com","Information Technology Administrators"
```

### ous.csv

```csv
Name,Path,Description
"Users","DC=contoso,DC=com","Standard Corporate Users"
```

The `sample-data/` folder includes ready-to-use examples: 10 users, 6 groups, and 5 OUs.

---

## Audit Log

Every action is automatically written to `AD_Manager.log` in the script directory:

```
[2026-03-18T10:23:01Z] | Mode: REAL   | Action: Bulk Import User | Target: asmith    | Result: Success
[2026-03-18T10:23:45Z] | Mode: WHATIF | Action: Delete User      | Target: bjohnson  | Result: Simulated (WhatIf)
[2026-03-18T10:24:10Z] | Mode: REAL   | Action: Reset Password   | Target: cwilliams | Result: Success (Masked)
```

---

## Project Structure

```
AD-Manager/
│
├── AD_Manager.ps1        # Main script
├── README.md             # Documentation
├── .gitignore            # Excludes *.log files
└── sample-data/
    ├── users.csv         # 10 sample users
    ├── groups.csv        # 6 sample groups
    └── ous.csv           # 5 sample OUs
```

---

## Using in a Live AD Environment

1. Install the AD PowerShell module:

```powershell
Install-WindowsFeature -Name RSAT-AD-PowerShell
```

2. Run the script on a domain-joined machine.
3. Test with `-WhatIf` first.
4. Uncomment the AD cmdlet lines inside each function.

---

## Security Notes

- Passwords are generated via `System.Web.Security.Membership.GeneratePassword` and handled as `SecureString` — never logged or displayed in plain text.
- The script detects administrator privileges at startup and warns if not elevated.
- Deleting an OU triggers an explicit safety warning due to the recursive nature of the operation.

---

## Disclaimer

This script is intended for learning and simulation purposes. Always validate in a test environment before running against a production Active Directory infrastructure.

---

## License

MIT — free to use, modify and distribute.
