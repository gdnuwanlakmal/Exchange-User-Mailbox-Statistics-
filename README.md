# 📬 Mailbox Size and Folder Statistics Script
This PowerShell script retrieves detailed folder-level statistics for an Exchange 2019 on-premises mailbox, including:

- Folder sizes (in both MB and GB)

- Total mailbox size used

- Mailbox quota information

- Warning quota

- Send limit

- Remaining/free space before limit

## ✅ Features
- Lists all mailbox folders with item counts and storage size

- Automatically detects and converts sizes in MB and GB

- Calculates total mailbox usage from folder sizes

- Retrieves quota values from Exchange mailbox settings

- Highlights how much free space is left before hitting the quota

## PowerShell script

```shell
# ===============================
# Mailbox Size & Quota Report
# ===============================

# >>>>>> EDIT THIS USER <<<<<<
$user = "user@domain.lk"
# ===============================

Write-Host "`n=== Mailbox Size & Quota Report ===" -ForegroundColor Cyan
Write-Host "Target mailbox: $user" -ForegroundColor Yellow

# --- Get mailbox ---
try {
    $mailbox = Get-Mailbox -Identity $user -ErrorAction Stop
} catch {
    Write-Host "Error: Could not retrieve mailbox $user" -ForegroundColor Red
    exit
}

# --- Get mailbox folder statistics ---
$allFolderStats = Get-MailboxFolderStatistics -Identity $mailbox.DistinguishedName

# --- Build folder info list with size conversion ---
$folderSizeTable = foreach ($folderStats in $allFolderStats) {
    $sizeString = $folderStats.FolderAndSubfolderSize.ToString()
    $sizeMB = 0
    $sizeDisplay = "N/A"

    if ($sizeString -match "([\d\.]+)\s+GB") {
        $sizeMB = [math]::Round([double]$matches[1] * 1024, 2)
        $sizeDisplay = "$($matches[1]) GB"
    }
    elseif ($sizeString -match "([\d\.]+)\s+MB") {
        $sizeMB = [math]::Round([double]$matches[1], 2)
        $sizeDisplay = "$sizeMB MB"
    }
    elseif ($sizeString -match "([\d\.]+)\s+KB") {
        $sizeMB = [math]::Round([double]$matches[1] / 1024, 2)
        $sizeDisplay = "$($matches[1]) KB"
    }

    [PSCustomObject]@{
        FolderPath        = $folderStats.FolderPath
        FolderType        = $folderStats.FolderType
        ItemsInFolder     = $folderStats.ItemsInFolder
        ItemsInSubfolders = $folderStats.ItemsInFolderAndSubfolders
        SizeDisplay       = $sizeDisplay
        SizeMB            = $sizeMB
    }
}

# --- Display folder-level table ---
Write-Host "`nFolder Size Breakdown (largest first):" -ForegroundColor Green
$folderSizeTable |
    Sort-Object SizeMB -Descending |
    Format-Table FolderPath, FolderType, ItemsInFolder, ItemsInSubfolders, SizeDisplay -AutoSize

# --- Total mailbox size calculation ---
$totalSizeMB = ($folderSizeTable | Measure-Object -Property SizeMB -Sum).Sum
$totalSizeGB = [math]::Round($totalSizeMB / 1024, 2)
Write-Host "`nTotal mailbox size used: $totalSizeMB MB ($totalSizeGB GB)" -ForegroundColor Cyan

# --- Get quota information ---
$mailboxStats = Get-MailboxStatistics -Identity $mailbox.DistinguishedName
$usedSize = $mailboxStats.TotalItemSize.Value.ToString()

try { $db = Get-MailboxDatabase -Identity $mailbox.Database -ErrorAction Stop } catch { $db = $null }

# --- Issue Warning Quota ---
if ($null -ne $mailbox.IssueWarningQuota -and $mailbox.IssueWarningQuota.Value -ne $null) {
    $issueWarningQuota = $mailbox.IssueWarningQuota.Value.ToMB()
}
elseif ($db -and $null -ne $db.IssueWarningQuota -and $db.IssueWarningQuota.Value -ne $null) {
    $issueWarningQuota = $db.IssueWarningQuota.Value.ToMB()
}
else {
    $issueWarningQuota = "Default (N/A)"
}

# --- Prohibit Send Quota ---
if ($null -ne $mailbox.ProhibitSendQuota -and $mailbox.ProhibitSendQuota.Value -ne $null) {
    $prohibitSendQuota = $mailbox.ProhibitSendQuota.Value.ToMB()
}
elseif ($db -and $null -ne $db.ProhibitSendQuota -and $db.ProhibitSendQuota.Value -ne $null) {
    $prohibitSendQuota = $db.ProhibitSendQuota.Value.ToMB()
}
else {
    $prohibitSendQuota = "Default (N/A)"
}

# --- Calculate free space if numeric ---
if ([double]::TryParse($prohibitSendQuota, [ref]$null)) {
    $freeSpaceMB = $prohibitSendQuota - $totalSizeMB
    $freeSpaceGB = [math]::Round($freeSpaceMB / 1024, 2)
} else {
    $freeSpaceMB = "N/A"
    $freeSpaceGB = "N/A"
}

# --- Display quota info ---
Write-Host "`nMailbox quota info:" -ForegroundColor Yellow
Write-Host "Used (Exchange reported): $usedSize"
Write-Host "Issue warning at: $issueWarningQuota MB"
Write-Host "Prohibit send at: $prohibitSendQuota MB"
Write-Host "Free space remaining: $freeSpaceMB MB ($freeSpaceGB GB)" -ForegroundColor Green

Write-Host "`n=== Completed successfully ===" -ForegroundColor Cyan

```
