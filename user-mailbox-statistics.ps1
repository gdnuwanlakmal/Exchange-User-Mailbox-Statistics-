# Get the mailbox
$mailbox = Get-Mailbox -Identity "user@domain.lk"

# Get mailbox folder statistics
$allFolderStats = Get-MailboxFolderStatistics -Identity $mailbox.DistinguishedName

# Build folder info list with size conversion
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

    [PSCustomObject]@{
        FolderPath        = $folderStats.FolderPath
        FolderType        = $folderStats.FolderType
        ItemsInFolder     = $folderStats.ItemsInFolder
        ItemsInSubfolders = $folderStats.ItemsInFolderAndSubfolders
        SizeDisplay       = $sizeDisplay
        SizeMB            = $sizeMB
    }
}

# Display folder-level table
$folderSizeTable | Sort-Object SizeMB -Descending | Format-Table FolderPath, FolderType, ItemsInFolder, ItemsInSubfolders, SizeDisplay -AutoSize

# --- Total mailbox size calculation ---
$totalSizeMB = ($folderSizeTable | Measure-Object -Property SizeMB -Sum).Sum
$totalSizeGB = [math]::Round($totalSizeMB / 1024, 2)

Write-Host "`nTotal mailbox size used: $totalSizeMB MB ($totalSizeGB GB)" -ForegroundColor Cyan

# --- Get quota information ---
$mailboxStats = Get-MailboxStatistics -Identity $mailbox.DistinguishedName

$usedSize = $mailboxStats.TotalItemSize.Value.ToString()
$issueWarningQuota = $mailbox.IssueWarningQuota.Value.ToMB()
$prohibitSendQuota = $mailbox.ProhibitSendQuota.Value.ToMB()

$freeSpaceMB = $prohibitSendQuota - $totalSizeMB
$freeSpaceGB = [math]::Round($freeSpaceMB / 1024, 2)

Write-Host "`nMailbox quota info:" -ForegroundColor Yellow
Write-Host "Used (Exchange reported): $usedSize"
Write-Host "Issue warning at: $issueWarningQuota MB"
Write-Host "Prohibit send at: $prohibitSendQuota MB"
Write-Host "Free space remaining: $freeSpaceMB MB ($freeSpaceGB GB)" -ForegroundColor Green
