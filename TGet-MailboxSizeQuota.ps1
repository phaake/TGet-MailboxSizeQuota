<#
.SYNOPSIS
    TGet-MailboxSizeQuota.ps1 returns all mailboxes above a certain quota (ProhibitSendQuota) usage.
.DESCRIPTION
    TGet-MailboxSizeQuota.ps1 returns all mailboxes with a quota usage of 80% and above.
    If you want to check for another quota limit, pass the percentage as a parameter.

    TGet-MailboxSizeQuota.ps1 <percentage>
.EXAMPLE
    TGet-MailboxSizeQuota.ps1
    The above command will return all mailboxes with a quota usage of 80% and above
.EXAMPLE
    TGet-MailboxSizeQuota.ps1 85
    The above command will return all mailboxes with a quota usage of 85% and above
.NOTES
    Author:  Peter Haake
    Version: 0.3
    Date:    2015-03-26    
#>
#
# Load the Exchange Management Module
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
# Make sure en-us locale is used no matter what the server/user has configured
# Keep this section if you output decimal numbers
# Slightly modified from From http://occasionalutility.blogspot.com.au/2014/03/everyday-powershell-part-17-using-new.html
[System.Reflection.Assembly]::LoadWithPartialName("System.Threading")>$null
[System.Reflection.Assembly]::LoadWithPartialName("System.Globalization")>$null
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::CreateSpecificCulture("en-us")

#Set quotalimit to the first parameter passed. If no parameter is passed, set it at 80%
if ([INT]$args[0] -gt "") {
    $QuotaLimit = [INT]$args[0]
    }
else {
    $QuotaLimit = 80
}


# Get all mailboxes
$Mailboxes = @(Get-Mailbox -ResultSize Unlimited | select-object DisplayName, Identity, ProhibitSendQuota, ProhibitSendReceiveQuota)
# Clear the report object variable
$Report =@()

# Loop through all mailboxes
foreach ($usr_mailbox in $Mailboxes)
{
    # Clear variables (add all...)
    $usr_quota_default = "N/A"
    # Get statistics for all mailboxes
    $usr_mailboxstats = Get-MailboxStatistics -identity $usr_mailbox.Identity | select-object Displayname,Identity,Database,TotalItemSize,TotalDeletedItemSize,DatabaseIssueWarningQuota,DatabaseProhibitSendQuota

    #Convert TotalItemSize to INT64 and remove crap (looks like this initially "1.123 GB (1,205,513,370 bytes)" and comes out as a numeric 1205513370)
    [int64]$usr_mailboxstats_totalitemsize = [convert]::ToInt64(((($usr_mailboxstats.TotalItemSize.ToString().split("(")[-1]).split(")")[0]).split(" ")[0]-replace '[,]',''))
    #Convert TotalDeletedItemSize to INT and remove crap (looks like this initially "1.123 GB (1,205,513,370 bytes)" and comes out as a numeric 1205513370)
    [int64]$usr_mailboxstats_totaldeleteditemsize = [convert]::ToInt64(((($usr_mailboxstats.TotalDeletedItemSize.ToString().split("(")[-1]).split(")")[0]).split(" ")[0]-replace '[,]',''))

    # If the mailbox quota is Unlimited, then the database defaults are used.
    if ($usr_mailbox.ProhibitSendQuota -eq "Unlimited") {
        # Get quota from Database
        [INT64]$usr_quota = [convert]::ToInt64(((($usr_mailboxstats.DatabaseProhibitSendQuota.ToString().split("(")[-1]).split(")")[0]).split(" ")[0]-replace '[,]',''))
        $usr_quota_default = "Yes"
        }
    else {
        # Get quota from user mailbox
        [INT64]$usr_quota = [convert]::ToInt64(((($usr_mailbox.ProhibitSendQuota.ToString().split("(")[-1]).split(")")[0]).split(" ")[0]-replace '[,]',''))
        $usr_quota_default = "No"
    }
    # Calculate the quota percentage
    $usr_quota_percentage = [INT]((($usr_mailboxstats_totalitemsize + $usr_mailboxstats_totaldeleteditemsize) / $usr_quota)*100)

    # Add to report object
    if ($usr_quota_percentage -ge $QuotaLimit) {
        $usr_reportObject = New-Object PSObject
        $usr_reportObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $usr_mailboxstats.DisplayName
        $usr_reportObject | Add-Member -MemberType NoteProperty -Name "TotalItemSize" -Value $usr_mailboxstats_totalitemsize
        $usr_reportObject | Add-Member -MemberType NoteProperty -Name "TotalDeletedItemSize" -Value $usr_mailboxstats_totaldeleteditemsize
        $usr_reportObject | Add-Member -MemberType NoteProperty -Name "ProhibitSendQuota" -Value $usr_quota
        $usr_reportObject | Add-Member -MemberType NoteProperty -Name "QuotaPercent" -Value $usr_quota_percentage
        $usr_reportObject | Add-Member -MemberType NoteProperty -Name "DBDefaultQuota" -Value $usr_quota_default
        $report += $usr_reportObject
    }
}
# Output the report, sorted with the highest quota percentage at the top
$Report | Sort-Object QuotaPercent -Descending
