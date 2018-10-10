#requires -version 2
<#
.SYNOPSIS
  Create-MailboxStatistics - Script collect and export mailbox statistics.

.DESCRIPTION
  None

.PARAMETER <Parameter_Name>
  None

.INPUTS
  None

.OUTPUTS
  <Outputs if any, otherwise state None - example: Log file stored in C:\Windows\Temp\<name>.log>

.NOTES
  Version:        1.1
  Author:         Michele Nappa <mnappa@microsoft.com>
  Creation Date:  12/05/2016
  Purpose/Change: Corrected bugs and minor improvements
  
.EXAMPLE
  .\Create-MailboxStatistcs

.CREDITS
  Powershell template http://9to5it.com/powershell-script-template/
  Generate Mailbox Size and Information Reports using PowerShell https://gallery.technet.microsoft.com/exchange/Generate-Mailbox-Size-and-3f408172
  PowerShell: Logging Functions https://gist.github.com/9to5IT/9620565
  AnalyzeMoveRequestStats.ps1 script https://gallery.technet.microsoft.com/scriptcenter/AnalyzeMoveRequestStatsps1-6c71167a
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

# Report variable
$report = @()

#----------------------------------------------------------[Dependecies]-----------------------------------------------------------

Import-Module ActiveDirectory -ErrorAction STOP

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
$sScriptVersion = "1.1"

#-----------------------------------------------------------[Functions]------------------------------------------------------------


#-----------------------------------------------------------[Execution]------------------------------------------------------------

#Log File Info
$sLogPath = Split-Path $MyInvocation.MyCommand.Path 
$random = -join(48..57+65..90+97..122 | ForEach-Object {[char]$_} | Get-Random -Count 6)
$timestamp = Get-Date -UFormat %Y%m%d-%H%M
$sLogName = "Mailbox Statistics-$timestamp-$random.log"
$sReportfileName = "Mailbox Statistics-$timestamp.csv"
$sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName
$sReportFile = Join-Path -Path $sLogPath -ChildPath $sReportfileName


#Dot Source required Function Libraries
$incFunctions = $sLogPath + "\Logging_Functions.ps1"
. $incFunctions

Write-Host $sReportfileName

$Recipients = @("Name Surname <Name.Surname1@contoso.com>", "Name Surname <Name.Surname2@contoso.com>")
 
Log-Start -LogPath $sLogPath -LogName $sLogName -ScriptVersion $sScriptVersion

Log-Write -LogPath $sLogFile -LineValue "[$([DateTime]::Now)] Start collecting mailboxes..."
$Mailboxes = Get-Mailbox -resultsize Unlimited| Where-Object {$_.LitigationHoldEnabled -eq $true}
#$Mailboxes = get-mailbox "Eagan-greenwell April (FCA)"

$MailboxCount = $Mailboxes.count
Log-Write -LogPath $sLogFile -LineValue "[$([DateTime]::Now)] Found $MailboxCount mailbox(es)."

Log-Write -LogPath $sLogFile -LineValue "[$([DateTime]::Now)] Start collecting mailbox database(s)..."
$Databases = @(Get-MailboxDatabase)
$DatabasesCount = $Databases.Count 
Log-Write -LogPath $sLogFile -LineValue "[$([DateTime]::Now)] Found $DatabasesCount database(s)."

#Loop through mailbox list and collect the mailbox statistics
Log-Write -LogPath $sLogFile -LineValue "[$([DateTime]::Now)] Starting mailbox statistics collection..."
$i = 0
foreach ($mb in $Mailboxes)
{
    Log-Write -LogPath $sLogFile -LineValue "[$([DateTime]::Now)] Collecting mailbox $mb details..."
    
    $i = $i + 1
    $pct = $i/$MailboxCount * 100
    Write-Progress -Activity "Collecting mailbox details" -Status "Processing mailbox $i of $MailboxCount - $mb" -PercentComplete $pct

    $stats = $mb | Get-MailboxStatistics | Select-Object TotalItemSize,TotalDeletedItemSize,ItemCount,LastLogonTime,LastLoggedOnUserAccount
    
    if ($mb.ArchiveDatabase){
        $archivestats = $mb | Get-MailboxStatistics -Archive | Select-Object TotalItemSize,TotalDeletedItemSize,ItemCount
    }
    else{
        $archivestats = "n/a"
    }

    $inboxstats = Get-MailboxFolderStatistics $mb -FolderScope Inbox | Select -First 1
    $sentitemsstats = Get-MailboxFolderStatistics $mb -FolderScope SentItems | Select -First 1
    $deleteditemsstats = Get-MailboxFolderStatistics $mb -FolderScope DeletedItems | Select -First 1
    $RecoverableItems = Get-MailboxFolderStatistics $mb -FolderScope RecoverableItems

    $lastlogon = $stats.LastLogonTime

    $user = Get-User $mb
    $aduser = Get-ADUser $mb.samaccountname -Properties Enabled,AccountExpirationDate
    
    $primarydb = $Databases | where {$_.Name -eq $mb.Database.Name}
    $archivedb = $Databases | where {$_.Name -eq $mb.ArchiveDatabase.Name}

    #Create a custom PS object to aggregate the data we're interested in

    $userObj = New-Object PSObject
    $userObj | Add-Member NoteProperty -Name "DisplayName" -Value $mb.DisplayName
    $userObj | Add-Member NoteProperty -Name "Mailbox Type" -Value $mb.RecipientTypeDetails
    $userObj | Add-Member NoteProperty -Name "Primary Email Address" -Value $mb.PrimarySMTPAddress
    $userObj | Add-Member NoteProperty -Name "Organizational Unit" -Value $user.OrganizationalUnit

    $userObj | Add-Member NoteProperty -Name "Mailbox Size (Mb)" -Value $stats.TotalItemSize.Value.ToMB()
    $userObj | Add-Member NoteProperty -Name "Mailbox Recoverable Item Size (MB)" -Value $RecoverableItems[0].FolderAndSubfolderSize.ToMB()
    $userObj | Add-Member NoteProperty -Name "Total Mailbox Size (MB)" -Value ($stats.TotalItemSize.Value.ToMB() + $RecoverableItems[0].FolderAndSubfolderSize.ToMB())

    $userObj | Add-Member NoteProperty -Name "Inbox Folder Size (MB)" -Value $inboxstats.FolderandSubFolderSize.ToMB()
    $userObj | Add-Member NoteProperty -Name "Sent Items Folder Size (MB)" -Value $sentitemsstats.FolderandSubFolderSize.ToMB()
    $userObj | Add-Member NoteProperty -Name "Deleted Items Folder Size (MB)" -Value $deleteditemsstats.FolderandSubFolderSize.ToMB()
    #$userObj | Add-Member NoteProperty -Name "Recoverable Item Folder Size (MB)" -Value $RecoverableItems[0].FolderandSubFolderSize.ToMB()

    $userObj | Add-Member NoteProperty -Name "Audits Folder Size (MB)" -Value ($RecoverableItems|?{$_.Name -eq "Audits"}).FolderSize.ToMB()
    $userObj | Add-Member NoteProperty -Name "Calendar Logging Folder Size (MB)" -Value ($RecoverableItems|?{$_.Name -eq "Calendar Logging"}).FolderSize.ToMB()
    $userObj | Add-Member NoteProperty -Name "Deletions Folder Size (MB)" -Value ($RecoverableItems|?{$_.Name -eq "Deletions"}).FolderSize.ToMB()
    $userObj | Add-Member NoteProperty -Name "DiscoveryHolds Folder Size (MB)" -Value ($RecoverableItems|?{$_.Name -eq "DiscoveryHolds"}).FolderSize.ToMB()
    $userObj | Add-Member NoteProperty -Name "Search Discovery Holds Folder Size (MB)" -Value ($RecoverableItems|?{$_.Name -eq "SearchDiscoveryHoldsFolder"}).FolderSize.ToMB()
    $userObj | Add-Member NoteProperty -Name "Purges Folder Size (MB)" -Value ($RecoverableItems|?{$_.Name -eq "Purges"}).FolderSize.ToMB()
    $userObj | Add-Member NoteProperty -Name "Versions Folder Size (MB)" -Value ($RecoverableItems|?{$_.Name -eq "Versions"}).FolderSize.ToMB()
      
    if ($archivestats -eq "n/a"){
        $userObj | Add-Member NoteProperty -Name "Archive Size (MB)" -Value "n/a"
        $userObj | Add-Member NoteProperty -Name "Archive Deleted Item Size (MB)" -Value "n/a"
        $userObj | Add-Member NoteProperty -Name "Total Archive Size (MB)" -Value "n/a"
    }
    else{
        $userObj | Add-Member NoteProperty -Name "Archive Size (MB)" -Value $archivestats.TotalItemSize.Value.ToMB()
        $userObj | Add-Member NoteProperty -Name "Archive Deleted Item Size (MB)" -Value $archivestats.TotalDeletedItemSize.Value.ToMB()
        $userObj | Add-Member NoteProperty -Name "Archive Items" -Value $archivestats.ItemCount
        $userObj | Add-Member NoteProperty -Name "Total Archive Size (MB)" -Value ($archivestats.TotalItemSize.Value.ToMB() + $archivestats.TotalDeletedItemSize.Value.ToMB())
    }

    $userObj | Add-Member NoteProperty -Name "Audit Enabled" -Value $mb.AuditEnabled
    $userObj | Add-Member NoteProperty -Name "LitigationHold Enabled" -Value $mb.LitigationHoldEnabled
    $userObj | Add-Member NoteProperty -Name "LitigationHoldDate " -Value $mb.LitigationHoldDate
    $userObj | Add-Member NoteProperty -Name "LitigationHoldOwner " -Value $mb.LitigationHoldOwner
    $userObj | Add-Member NoteProperty -Name "LitigationHoldDuration " -Value $mb.LitigationHoldDuration
    $userObj | Add-Member NoteProperty -Name "InPlace Holds" -Value ($mb.InPlaceHolds -join "`t ")
    $userObj | Add-Member NoteProperty -Name "Email Address Policy Enabled" -Value $mb.EmailAddressPolicyEnabled
    $userObj | Add-Member NoteProperty -Name "Hidden From Address Lists" -Value $mb.HiddenFromAddressListsEnabled
    $userObj | Add-Member NoteProperty -Name "Use Database Quota Defaults" -Value $mb.UseDatabaseQuotaDefaults
    
    if ($mb.UseDatabaseQuotaDefaults -eq $true){
        $userObj | Add-Member NoteProperty -Name "Issue Warning Quota" -Value $primarydb.IssueWarningQuota.Value.ToMB()
        $userObj | Add-Member NoteProperty -Name "Prohibit Send Quota" -Value $primarydb.ProhibitSendQuota.Value.ToMB()
        $userObj | Add-Member NoteProperty -Name "Prohibit Send Receive Quota" -Value $primarydb.ProhibitSendReceiveQuota.Value.ToMB()
    }
    elseif ($mb.UseDatabaseQuotaDefaults -eq $false){
        $userObj | Add-Member NoteProperty -Name "Issue Warning Quota" -Value $mb.IssueWarningQuota.IssueWarningQuota.Value.ToMB()
        $userObj | Add-Member NoteProperty -Name "Prohibit Send Quota" -Value $mb.ProhibitSendQuota.IssueWarningQuota.Value.ToMB()
        $userObj | Add-Member NoteProperty -Name "Prohibit Send Receive Quota" -Value $mb.ProhibitSendReceiveQuota.Value.ToMB()
    }

    $userObj | Add-Member NoteProperty -Name "Account Enabled" -Value $aduser.Enabled
    $userObj | Add-Member NoteProperty -Name "Account Expires" -Value $aduser.AccountExpirationDate
    $userObj | Add-Member NoteProperty -Name "Last Mailbox Logon" -Value $lastlogon
    $userObj | Add-Member NoteProperty -Name "Last Logon By" -Value $stats.LastLoggedOnUserAccount

    $userObj | Add-Member NoteProperty -Name "Primary Mailbox Database" -Value $mb.Database
    $userObj | Add-Member NoteProperty -Name "Primary Server/DAG" -Value $primarydb.MasterServerOrAvailabilityGroup

    $userObj | Add-Member NoteProperty -Name "Archive Mailbox Database" -Value $mb.ArchiveDatabase
    $userObj | Add-Member NoteProperty -Name "Archive Server/DAG" -Value $archivedb.MasterServerOrAvailabilityGroup

	#Add the object to the report
    $report = $report += $userObj
}
Log-Write -LogPath $sLogFile -LineValue "[$([DateTime]::Now)] Finished mailbox statistics collection..."


$reportcount = $report.count

if ($reportcount -eq 0){
    Write-Host -ForegroundColor Yellow "No mailboxes were found matching that criteria."
}
else{
    Log-Write -LogPath $sLogPath -LineValue "[$([DateTime]::Now)] Exporting mailbox statistics collection..."
    $report | Export-Csv -Path $sReportfileName -NoTypeInformation -Encoding UTF8
    Write-Host -ForegroundColor White "Report written to $sReportfileName in current path."
    Log-Write -LogPath $sLogPath -LineValue "[$([DateTime]::Now)] Finished mailbox statistics export..."
    Log-Write -LogPath $sLogPath -LineValue "[$([DateTime]::Now)] Sending email to $Recipients"
    Log-Write -LogPath $sLogPath -LineValue "[$([DateTime]::Now)] Email Subject: CNH Report del $timestamp" 
    Log-Write -LogPath $sLogPath -LineValue "[$([DateTime]::Now)] Email Body: In allegato il report per $($report.count) mailbox."
    Send-MailMessage -From "Application Sender <Application.Sender@contoso.com>" -To $Recipients -Subject "Report del $timestamp" -Body "In allegato il report per $($report.count) mailbox." -Attachments $sReportfileName -dno onSuccess, onFailure -SmtpServer "smtprelay.fgremc.it"
}

Log-Finish -LogPath $sLogFile
