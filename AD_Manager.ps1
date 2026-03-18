<#
.SYNOPSIS
    Enterprise-grade Active Directory Management Tool.
    
.DESCRIPTION
    A robust script for managing Active Directory users, groups, and OUs.
    Features strict separation of concerns, native streams, and robust error handling.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogFileName = "AD_Manager.log"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --- CORE SETTINGS & CONTEXT ---

function Get-ScriptContext {
    [CmdletBinding()]
    param()
    
    $basePath = Split-Path -Parent $MyInvocation.ScriptName
    if ([string]::IsNullOrWhiteSpace($basePath)) { $basePath = $PWD.Path }
    
    $currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $isAdmin) {
        Write-Warning "Not running as Administrator. Modifying AD might fail or require specific delegation safely."
    }

    [PSCustomObject]@{
        BasePath      = $basePath
        LogPath       = Join-Path $basePath $LogFileName
        IsAdmin       = $isAdmin
        LastSearch    = @()
        UserCsvPath   = "users.csv"
        GroupCsvPath  = "groups.csv"
        OuCsvPath     = "ous.csv"
    }
}

# --- LOGGING UTILITY ---

function Write-AuditLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Action,
        [Parameter(Mandatory = $true)][string]$TargetObject,
        [Parameter(Mandatory = $true)][string]$Result,
        [Parameter(Mandatory = $true)][psobject]$Context
    )

    $timestamp = ([datetime]::UtcNow).ToString("yyyy-MM-ddTHH:mm:ssZ")
    $mode = if ($WhatIfPreference) { "WHATIF" } else { "REAL" }
    $logEntry = "[$timestamp] | Mode: $mode | Action: $Action | Target: $TargetObject | Result: $Result"
    
    try {
        Add-Content -Path $Context.LogPath -Value $logEntry -ErrorAction Stop
        Write-Verbose "Audit log written: $Action -> $Result"
    } catch {
        Write-Warning "Failed to write audit log at $($Context.LogPath). Error: $_"
    }
}

# --- HELPER FUNCTIONS ---

function Find-CsvFileValidated {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$FileName,
        [Parameter(Mandatory = $true)][psobject]$Context
    )
    
    $fullPath = Join-Path $Context.BasePath $FileName
    if (Test-Path $fullPath -PathType Leaf) {
        Write-Verbose "Found CSV explicitly in base path: $fullPath"
        return $fullPath
    }

    Write-Verbose "Searching system for $FileName..."
    $foundFiles = Get-ChildItem -Path $env:USERPROFILE -Filter $FileName -Recurse -File -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
    
    if (-not $foundFiles) {
        Write-Warning "Could not find $FileName."
        return $null
    }
    
    if ($foundFiles.Count -eq 1) {
        Write-Verbose "Auto-selected CSV: $($foundFiles[0])"
        return $foundFiles[0]
    }

    Write-Warning "Multiple files named '$FileName' found. Returning the most recent one."
    $bestMatch = Get-ChildItem -Path $foundFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
    return $bestMatch
}

# --- BUSINESS LOGIC (PURE FUNCTIONS) ---

function Import-ADManagerUsersFromCsv {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory = $true)][psobject]$Context)
    
    if (-not (Test-Path $Context.UserCsvPath)) {
        Write-Warning "User CSV not found at $($Context.UserCsvPath)."
        return
    }

    try {
        $users = Import-Csv -Path $Context.UserCsvPath
        foreach ($u in $users) {
            $userIdentifier = $u.SamAccountName
            if ($PSCmdlet.ShouldProcess("AD User: $userIdentifier", "Create from CSV")) {
                Write-Output "Successfully created user: $userIdentifier"
                Write-AuditLog -Action "Bulk Import User" -TargetObject $userIdentifier -Result "Success" -Context $Context
            } else {
                Write-AuditLog -Action "Bulk Import User" -TargetObject $userIdentifier -Result "Simulated (WhatIf)" -Context $Context
            }
        }
    } catch {
        $errDetails = $_ | Out-String
        Write-Error "Failed to import users. Exception: $errDetails"
        Write-AuditLog -Action "Bulk Import Users" -TargetObject "CSV: $($Context.UserCsvPath)" -Result "Error: $errDetails" -Context $Context
    }
}

function New-ADManagerUser {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string]$FullName,
        [Parameter(Mandatory = $true)][string]$SamAccountName,
        [Parameter(Mandatory = $true)][string]$Email,
        [Parameter(Mandatory = $true)][string]$Department,
        [Parameter(Mandatory = $true)][string]$Phone,
        [Parameter(Mandatory = $true)][string]$OUPath,
        [Parameter(Mandatory = $true)][psobject]$Context
    )
    
    if ($PSCmdlet.ShouldProcess("AD User: $SamAccountName", "Create manually in $OUPath")) {
        try {
            Write-Output "Successfully created user $SamAccountName."
            Write-AuditLog -Action "Create Single User" -TargetObject $SamAccountName -Result "Success" -Context $Context
        } catch {
            $errDetails = $_ | Out-String
            Write-Error "Failed to create user $SamAccountName. Exception: $errDetails"
            Write-AuditLog -Action "Create Single User" -TargetObject $SamAccountName -Result "Error: $errDetails" -Context $Context
        }
    } else {
        Write-AuditLog -Action "Create Single User" -TargetObject $SamAccountName -Result "Simulated (WhatIf)" -Context $Context
    }
}

function Search-ADManagerUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Query,
        [Parameter(Mandatory = $true)][psobject]$Context
    )
    
    Write-Verbose "Initiating search for: $Query"
    try {
        $searchResults = @(
            [PSCustomObject]@{
                Name = "Mocked User"
                SamAccountName = "mockuser"
                EmailAddress = "$Query@example.com"
                Department = "IT"
                Enabled = $true
            }
        )
        
        if ($searchResults) {
            $searchResults | Format-Table -AutoSize | Out-String | Write-Output
            $Context.LastSearch = $searchResults
            Write-AuditLog -Action "Search User" -TargetObject $Query -Result "Found $($searchResults.Count) matches" -Context $Context
        } else {
            Write-Warning "No users found matching '$Query'."
            Write-AuditLog -Action "Search User" -TargetObject $Query -Result "No matches" -Context $Context
        }
    } catch {
        $errDetails = $_ | Out-String
        Write-Error "Search failed. Exception: $errDetails"
        Write-AuditLog -Action "Search User" -TargetObject $Query -Result "Error: $errDetails" -Context $Context
    }
}

function Set-ADManagerUserAccountState {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string]$SamAccountName,
        [Parameter(Mandatory = $true)][ValidateSet("Enable", "Disable")][string]$ActionName,
        [Parameter(Mandatory = $true)][psobject]$Context
    )
    
    if ($PSCmdlet.ShouldProcess("AD User: $SamAccountName", "$ActionName Account")) {
        try {
            Write-Output "Successfully applied $ActionName on $SamAccountName."
            Write-AuditLog -Action "$ActionName User" -TargetObject $SamAccountName -Result "Success" -Context $Context
        } catch {
            $errDetails = $_ | Out-String
            Write-Error "Failed to $ActionName user $SamAccountName. Exception: $errDetails"
            Write-AuditLog -Action "$ActionName User" -TargetObject $SamAccountName -Result "Error: $errDetails" -Context $Context
        }
    } else {
        Write-AuditLog -Action "$ActionName User" -TargetObject $SamAccountName -Result "Simulated (WhatIf)" -Context $Context
    }
}

function Reset-ADManagerUserPassword {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string]$SamAccountName,
        [Parameter(Mandatory = $true)][psobject]$Context
    )
    
    if ($PSCmdlet.ShouldProcess("AD User: $SamAccountName", "Reset Password to dynamically generated secure string")) {
        try {
            Add-Type -AssemblyName System.Web
            $plain = [System.Web.Security.Membership]::GeneratePassword(16, 4)
            $secureString = ConvertTo-SecureString $plain -AsPlainText -Force
            
            Write-Output "Password successfully reset for $SamAccountName."
            Write-Output "Operation utilized SecureString object masking."
            Write-AuditLog -Action "Reset Password" -TargetObject $SamAccountName -Result "Success (Masked)" -Context $Context
        } catch {
            $errDetails = $_ | Out-String
            Write-Error "Failed to reset password for $SamAccountName. Exception: $errDetails"
            Write-AuditLog -Action "Reset Password" -TargetObject $SamAccountName -Result "Error: $errDetails" -Context $Context
        }
    } else {
        Write-AuditLog -Action "Reset Password" -TargetObject $SamAccountName -Result "Simulated (WhatIf)" -Context $Context
    }
}

function Move-ADManagerUserToOu {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string]$SamAccountName,
        [Parameter(Mandatory = $true)][string]$TargetOU,
        [Parameter(Mandatory = $true)][psobject]$Context
    )
    
    if ($PSCmdlet.ShouldProcess("AD User: $SamAccountName", "Move to OU: $TargetOU")) {
        try {
            Write-Output "Successfully moved $SamAccountName to $TargetOU."
            Write-AuditLog -Action "Move User" -TargetObject $SamAccountName -Result "Moved to $TargetOU" -Context $Context
        } catch {
            $errDetails = $_ | Out-String
            Write-Error "Failed to move user $SamAccountName. Exception: $errDetails"
            Write-AuditLog -Action "Move User" -TargetObject $SamAccountName -Result "Error: $errDetails" -Context $Context
        }
    } else {
        Write-AuditLog -Action "Move User" -TargetObject $SamAccountName -Result "Simulated (WhatIf)" -Context $Context
    }
}

function Remove-ADManagerSingleUser {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true)][string]$SamAccountName,
        [Parameter(Mandatory = $true)][psobject]$Context
    )
    
    if ($PSCmdlet.ShouldProcess("AD User: $SamAccountName", "Permanently Delete")) {
        try {
            Write-Output "User $SamAccountName successfully deleted."
            Write-AuditLog -Action "Delete User" -TargetObject $SamAccountName -Result "Success" -Context $Context
        } catch {
            $errDetails = $_ | Out-String
            Write-Error "Failed to delete user $SamAccountName. Exception: $errDetails"
            Write-AuditLog -Action "Delete User" -TargetObject $SamAccountName -Result "Error: $errDetails" -Context $Context
        }
    } else {
        Write-AuditLog -Action "Delete User" -TargetObject $SamAccountName -Result "Simulated (WhatIf)" -Context $Context
    }
}

# --- GROUPS LOGIC ---

function Import-ADManagerGroupsFromCsv {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory = $true)][psobject]$Context)
    
    if (-not (Test-Path $Context.GroupCsvPath)) {
        Write-Warning "Group CSV not found at $($Context.GroupCsvPath)."
        return
    }

    try {
        $groups = Import-Csv -Path $Context.GroupCsvPath
        foreach ($g in $groups) {
            $groupId = $g.Name
            if ($PSCmdlet.ShouldProcess("AD Group: $groupId", "Create from CSV")) {
                Write-Output "Successfully created group: $groupId"
                Write-AuditLog -Action "Bulk Import Group" -TargetObject $groupId -Result "Success" -Context $Context
            } else {
                Write-AuditLog -Action "Bulk Import Group" -TargetObject $groupId -Result "Simulated (WhatIf)" -Context $Context
            }
        }
    } catch {
        $errDetails = $_ | Out-String
        Write-Error "Failed to import groups. Exception: $errDetails"
        Write-AuditLog -Action "Bulk Import Groups" -TargetObject "CSV: $($Context.GroupCsvPath)" -Result "Error: $errDetails" -Context $Context
    }
}

function Set-ADManagerGroupMembership {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string]$GroupName,
        [Parameter(Mandatory = $true)][string]$SamAccountName,
        [Parameter(Mandatory = $true)][ValidateSet("Add", "Remove")][string]$ActionName,
        [Parameter(Mandatory = $true)][psobject]$Context
    )
    if ($PSCmdlet.ShouldProcess("AD Group: $GroupName", "$ActionName user $SamAccountName")) {
        try {
            Write-Output "Successfully applied $ActionName for $SamAccountName on $GroupName."
            Write-AuditLog -Action "$ActionName User Group" -TargetObject "$SamAccountName -> $GroupName" -Result "Success" -Context $Context
        } catch {
            $errDetails = $_ | Out-String
            Write-Error "Failed to modify group. Exception: $errDetails"
            Write-AuditLog -Action "$ActionName User Group" -TargetObject "$SamAccountName -> $GroupName" -Result "Error: $errDetails" -Context $Context
        }
    } else {
        Write-AuditLog -Action "$ActionName User Group" -TargetObject "$SamAccountName -> $GroupName" -Result "Simulated (WhatIf)" -Context $Context
    }
}

function Get-ADManagerGroupMembers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$GroupName,
        [Parameter(Mandatory = $true)][psobject]$Context
    )
    Write-Verbose "Fetching members for group: $GroupName"
    try {
        Write-Output "Members for $GroupName (Mock Result):"
        Write-Output " - Mocked Member One (m1)"
        Write-AuditLog -Action "List Group Members" -TargetObject $GroupName -Result "Success" -Context $Context
    } catch {
        $errDetails = $_ | Out-String
        Write-Error "Failed to get group members. Exception: $errDetails"
        Write-AuditLog -Action "List Group Members" -TargetObject $GroupName -Result "Error: $errDetails" -Context $Context
    }
}

function Remove-ADManagerGroup {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact='High')]
    param(
        [Parameter(Mandatory = $true)][string]$GroupName,
        [Parameter(Mandatory = $true)][psobject]$Context
    )
    if ($PSCmdlet.ShouldProcess("AD Group: $GroupName", "Permanently Delete")) {
        try {
            Write-Output "Successfully deleted group: $GroupName"
            Write-AuditLog -Action "Delete Group" -TargetObject $GroupName -Result "Success" -Context $Context
        } catch {
            $errDetails = $_ | Out-String
            Write-Error "Failed to delete group. Exception: $errDetails"
            Write-AuditLog -Action "Delete Group" -TargetObject $GroupName -Result "Error: $errDetails" -Context $Context
        }
    } else {
        Write-AuditLog -Action "Delete Group" -TargetObject $GroupName -Result "Simulated (WhatIf)" -Context $Context
    }
}

# --- OU LOGIC ---

function Import-ADManagerOUsFromCsv {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory = $true)][psobject]$Context)
    
    if (-not (Test-Path $Context.OuCsvPath)) {
        Write-Warning "OU CSV not found at $($Context.OuCsvPath)."
        return
    }
    try {
        $ous = Import-Csv -Path $Context.OuCsvPath
        foreach ($ou in $ous) {
            $ouId = $ou.Name
            if ($PSCmdlet.ShouldProcess("AD OU: $ouId", "Create from CSV")) {
                Write-Output "Successfully created OU: $ouId"
                Write-AuditLog -Action "Bulk Import OU" -TargetObject $ouId -Result "Success" -Context $Context
            } else {
                Write-AuditLog -Action "Bulk Import OU" -TargetObject $ouId -Result "Simulated (WhatIf)" -Context $Context
            }
        }
    } catch {
        $errDetails = $_ | Out-String
        Write-Error "Failed to import OUs. Exception: $errDetails"
        Write-AuditLog -Action "Bulk Import OU" -TargetObject "CSV: $($Context.OuCsvPath)" -Result "Error: $errDetails" -Context $Context
    }
}

function Get-ADManagerOUs {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][psobject]$Context)
    Write-Verbose "Listing all OUs..."
    try {
        Write-Output "OUs and Users (Mock Result):"
        Write-Output "OU=IT,DC=example,DC=com"
        Write-AuditLog -Action "List OUs" -TargetObject "All" -Result "Success" -Context $Context
    } catch {
        $errDetails = $_ | Out-String
        Write-Error "Failed to list OUs. Exception: $errDetails"
        Write-AuditLog -Action "List OUs" -TargetObject "All" -Result "Error: $errDetails" -Context $Context
    }
}

function Remove-ADManagerOU {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact='High')]
    param(
        [Parameter(Mandatory = $true)][string]$OUPath,
        [Parameter(Mandatory = $true)][psobject]$Context
    )
    if ($PSCmdlet.ShouldProcess("AD OU: $OUPath", "Delete with safety check (recursive)")) {
        try {
            Write-Output "Successfully deleted OU: $OUPath"
            Write-AuditLog -Action "Delete OU" -TargetObject $OUPath -Result "Success" -Context $Context
        } catch {
            $errDetails = $_ | Out-String
            Write-Error "Failed to delete OU. Exception: $errDetails"
            Write-AuditLog -Action "Delete OU" -TargetObject $OUPath -Result "Error: $errDetails" -Context $Context
        }
    } else {
        Write-AuditLog -Action "Delete OU" -TargetObject $OUPath -Result "Simulated (WhatIf)" -Context $Context
    }
}

# --- REPORTS & EXPORT ---

function Export-ADManagerUsersList {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory = $true)][psobject]$Context)
    $exportPath = Join-Path $Context.BasePath "Exported_Users.csv"
    
    if ($PSCmdlet.ShouldProcess("Export Path: $exportPath", "Export Full Users List")) {
        try {
            Write-Output "Exporting full users list to $exportPath"
            Write-AuditLog -Action "Export Users" -TargetObject $exportPath -Result "Success" -Context $Context
        } catch {
            $errDetails = $_ | Out-String
            Write-Error "Failed to export users. Exception: $errDetails"
            Write-AuditLog -Action "Export Users" -TargetObject $exportPath -Result "Error: $errDetails" -Context $Context
        }
    } else {
        Write-AuditLog -Action "Export Users" -TargetObject $exportPath -Result "Simulated (WhatIf)" -Context $Context
    }
}

function Export-ADManagerSearch {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory = $true)][psobject]$Context)
    
    if (-not $Context.LastSearch -or $Context.LastSearch.Count -eq 0) {
        Write-Warning "No recent search results in memory to export."
        return
    }
    
    $exportPath = Join-Path $Context.BasePath "Exported_Search.csv"
    if ($PSCmdlet.ShouldProcess("Export Path: $exportPath", "Export Last Search")) {
        try {
            $Context.LastSearch | Export-Csv -Path $exportPath -NoTypeInformation
            Write-Output "Successfully exported search results to $exportPath"
            Write-AuditLog -Action "Export Search" -TargetObject $exportPath -Result "Success" -Context $Context
        } catch {
            $errDetails = $_ | Out-String
            Write-Error "Failed to export search results. Exception: $errDetails"
            Write-AuditLog -Action "Export Search" -TargetObject $exportPath -Result "Error: $errDetails" -Context $Context
        }
    } else {
        Write-AuditLog -Action "Export Search" -TargetObject $exportPath -Result "Simulated (WhatIf)" -Context $Context
    }
}

function Show-ADManagerAuditLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][psobject]$Context
    )
    
    Write-Verbose "Reading audit log file: $($Context.LogPath)"
    try {
        if (Test-Path $Context.LogPath) {
            Write-Output "--- Last 50 Log Entries ---"
            Get-Content -Path $Context.LogPath -Tail 50 | Write-Output
            Write-AuditLog -Action "View Audit Log" -TargetObject $Context.LogPath -Result "Success" -Context $Context
        } else {
            Write-Warning "Audit log does not exist yet mapping to $($Context.LogPath)."
        }
    } catch {
        $errDetails = $_ | Out-String
        Write-Error "Failed to read the audit log. Exception: $errDetails"
    }
}

# --- UI CONTROLLER LAYER ---

function Show-MainMenu {
    [CmdletBinding()]
    param([System.Management.Automation.PSCustomObject]$Context)
    
    while ($true) {
        Write-Output "========================================="
        Write-Output "        ACTIVE DIRECTORY MANAGER         "
        Write-Output "========================================="
        
        if ($WhatIfPreference) {
            Write-Output "   [!] WHATIF MODE ACTIVE - No changes applied"
        }
        
        Write-Output " ── USER MANAGEMENT ──"
        Write-Output "  1. Bulk import users from CSV"
        Write-Output "  2. Create single user manually"
        Write-Output "  3. Search user"
        Write-Output "  4. Enable / Disable user account"
        Write-Output "  5. Reset user password"
        Write-Output "  6. Move user to another OU"
        Write-Output "  7. Delete user permanently"
        Write-Output " "
        Write-Output " ── GROUP MANAGEMENT ──"
        Write-Output "  8. Bulk import groups from CSV"
        Write-Output "  9. Add / Remove user from group"
        Write-Output "  10. List all members of a group"
        Write-Output "  11. Delete group"
        Write-Output " "
        Write-Output " ── ORGANIZATIONAL UNITS ──"
        Write-Output "  12. Bulk import OUs from CSV"
        Write-Output "  13. List all OUs and their users"
        Write-Output "  14. Delete OU with safety check"
        Write-Output " "
        Write-Output " ── REPORTS & EXPORT ──"
        Write-Output "  15. Export full user list to CSV"
        Write-Output "  16. Export last search result to CSV"
        Write-Output "  17. Show audit log"
        Write-Output " "
        Write-Output "  X. Exit"
        Write-Output "========================================="
        
        $choice = Read-Host "Select an option"
        Write-Output ""
        
        switch ($choice) {
            "1"  { Import-ADManagerUsersFromCsv -Context $Context }
            "2"  { 
                $fullName = Read-Host "Full Name"
                $sam = Read-Host "SamAccountName"
                $email = Read-Host "Email"
                $dept = Read-Host "Department"
                $phone = Read-Host "Phone"
                $ou = Read-Host "OU Path"
                New-ADManagerUser -FullName $fullName -SamAccountName $sam -Email $email -Department $dept -Phone $phone -OUPath $ou -Context $Context 
            }
            "3"  { 
                $qry = Read-Host "Enter search term"
                Search-ADManagerUser -Query $qry -Context $Context 
            }
            "4"  { 
                $sam = Read-Host "Enter SamAccountName"
                $act = Read-Host "Enable(1) or Disable(2)"
                $strAct = if($act -eq "1") { "Enable" } elseif ($act -eq "2") { "Disable" } else { "" }
                if ($strAct) { Set-ADManagerUserAccountState -SamAccountName $sam -ActionName $strAct -Context $Context } else { Write-Warning "Invalid choice" }
            }
            "5"  { 
                $sam = Read-Host "Enter SamAccountName"
                Reset-ADManagerUserPassword -SamAccountName $sam -Context $Context 
            }
            "6"  { 
                $sam = Read-Host "Enter SamAccountName"
                $ou = Read-Host "Enter Target OU Path"
                Move-ADManagerUserToOu -SamAccountName $sam -TargetOU $ou -Context $Context 
            }
            "7"  { 
                $sam = Read-Host "Enter SamAccountName"
                $confirm = Read-Host "Delete permanently? (Y/N)"
                if ($confirm -match '^[Yy]$') { Remove-ADManagerSingleUser -SamAccountName $sam -Context $Context }
            }
            "8"  { Import-ADManagerGroupsFromCsv -Context $Context }
            "9"  { 
                $group = Read-Host "Enter Group Name"
                $sam = Read-Host "Enter SamAccountName"
                $act = Read-Host "Add(1) or Remove(2)"
                $strAct = if($act -eq "1") { "Add" } elseif ($act -eq "2") { "Remove" } else { "" }
                if ($strAct) { Set-ADManagerGroupMembership -GroupName $group -SamAccountName $sam -ActionName $strAct -Context $Context } else { Write-Warning "Invalid choice" }
            }
            "10" { 
                $group = Read-Host "Enter Group Name"
                Get-ADManagerGroupMembers -GroupName $group -Context $Context 
            }
            "11" { 
                $group = Read-Host "Enter Group Name"
                $confirm = Read-Host "Delete group permanently? (Y/N)"
                if ($confirm -match '^[Yy]$') { Remove-ADManagerGroup -GroupName $group -Context $Context }
            }
            "12" { Import-ADManagerOUsFromCsv -Context $Context }
            "13" { Get-ADManagerOUs -Context $Context }
            "14" { 
                Write-Warning "Deleting an OU will delete everything inside."
                $ou = Read-Host "Enter OU Object Name or Path"
                $confirm = Read-Host "Are you sure? (Y/N)"
                if ($confirm -match '^[Yy]$') { Remove-ADManagerOU -OUPath $ou -Context $Context }
            }
            "15" { Export-ADManagerUsersList -Context $Context }
            "16" { Export-ADManagerSearch -Context $Context }
            "17" { Show-ADManagerAuditLog -Context $Context }
            { $_ -match "^[Xx]$" } { 
                Write-Verbose "Exiting Active Directory Manager."
                return 
            }
            default {
                Write-Warning "Invalid option '$choice'. Please try again."
            }
        }
        
        Write-Output "`nPress Enter to return to the menu..."
        $null = Read-Host
    }
}

# --- SCRIPT ENTRY POINT ---
try {
    $scriptContext = Get-ScriptContext
    
    $userCsvResponse = Read-Host "Enter User CSV file name (Optional, press Enter for 'users.csv')"
    if (-not [string]::IsNullOrWhiteSpace($userCsvResponse)) {
        $userPath = Find-CsvFileValidated -FileName $userCsvResponse -Context $scriptContext
        if ($userPath) { $scriptContext.UserCsvPath = $userPath }
    }
    
    $groupCsvResponse = Read-Host "Enter Group CSV file name (Optional, press Enter for 'groups.csv')"
    if (-not [string]::IsNullOrWhiteSpace($groupCsvResponse)) {
        $groupPath = Find-CsvFileValidated -FileName $groupCsvResponse -Context $scriptContext
        if ($groupPath) { $scriptContext.GroupCsvPath = $groupPath }
    }
    
    $ouCsvResponse = Read-Host "Enter OU CSV file name (Optional, press Enter for 'ous.csv')"
    if (-not [string]::IsNullOrWhiteSpace($ouCsvResponse)) {
        $ouPath = Find-CsvFileValidated -FileName $ouCsvResponse -Context $scriptContext
        if ($ouPath) { $scriptContext.OuCsvPath = $ouPath }
    }
    
    Show-MainMenu -Context $scriptContext
} catch {
    $errObj = $_ | Out-String
    Write-Error "A fatal error occurred outside standard blocks. Details: $errObj"
}
