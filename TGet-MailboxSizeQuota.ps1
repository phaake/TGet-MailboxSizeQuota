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
    Version: 0.4
    Date:    2015-03-30
#>

<#
Revision History

0.3 2015-03-26
First version in revision control

0.4 2015-03-30
Removed the work-around for number format
Changed number format in the output, from bytes to MB

#>

#
# Load the Exchange Management Module
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

#Set quotalimit to the first parameter passed. If no parameter is passed, set it at 80%
if ([INT]$args[0] -gt "") {
    $QuotaLimit = [INT]$args[0]
    }
else {
    $QuotaLimit = 80
}


# Get all mailboxes
$Mailboxes = @(Get-Mailbox -ResultSize Unlimited | select-object DisplayName, Identity, ProhibitSendQuota, ProhibitSendReceiveQuota, UseDatabaseQuotaDefaults)
# Clear the report object variable
$Report =@()

# Loop through all mailboxes
foreach ($usr_mailbox in $Mailboxes)
{
    # Get statistics for all mailboxes
    $usr_mailboxstats = Get-MailboxStatistics -identity $usr_mailbox.Identity | select-object Displayname,Identity,Database,TotalItemSize,TotalDeletedItemSize,DatabaseIssueWarningQuota,DatabaseProhibitSendQuota

    #Convert TotalItemSize to INT64 and remove crap (looks like this initially "1.123 GB (1,205,513,370 bytes)" and comes out as a numeric 1205513370)
    $usr_mailboxstats_totalitemsize = $usr_mailboxstats.TotalItemSize.Value.ToBytes()
    #Convert TotalDeletedItemSize to INT and remove crap (looks like this initially "1.123 GB (1,205,513,370 bytes)" and comes out as a numeric 1205513370)
    $usr_mailboxstats_totaldeleteditemsize = $usr_mailboxstats.TotalDeletedItemSize.Value.ToBytes()

    # If the mailbox quota is Unlimited, then the database defaults are used.
    if ($usr_mailbox.UseDatabaseQuotaDefaults -eq "True") {
        # Get quota from Database
        $usr_quota = $usr_mailboxstats.DatabaseProhibitSendQuota.Value.ToBytes()
        $usr_quota_MB = $usr_mailboxstats.DatabaseProhibitSendQuota.Value.ToMB()
        $usr_quota_default = "Yes"
        }
    else {
        # Get quota from user mailbox
        $usr_quota = $usr_mailbox.ProhibitSendQuota.Value.ToBytes()
        $usr_quota_MB = $usr_mailbox.ProhibitSendQuota.Value.ToMB()
        $usr_quota_default = "No"
    }
    # Calculate the quota percentage
    $usr_quota_percentage = [INT]((($usr_mailboxstats_totalitemsize + $usr_mailboxstats_totaldeleteditemsize) / $usr_quota)*100)

    # Add to report object
    if ($usr_quota_percentage -ge $QuotaLimit) {
        $usr_reportObject = New-Object PSObject
        $usr_reportObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $usr_mailboxstats.DisplayName
        $usr_reportObject | Add-Member -MemberType NoteProperty -Name "TotalItemSize (MB)" -Value $usr_mailboxstats.TotalItemSize.Value.ToMB()
        $usr_reportObject | Add-Member -MemberType NoteProperty -Name "TotalDeletedItemSize (MB)" -Value $usr_mailboxstats.TotalDeletedItemSize.Value.ToMB()
        $usr_reportObject | Add-Member -MemberType NoteProperty -Name "ProhibitSendQuota (MB)" -Value $usr_quota_MB
        $usr_reportObject | Add-Member -MemberType NoteProperty -Name "QuotaPercent" -Value $usr_quota_percentage
        $usr_reportObject | Add-Member -MemberType NoteProperty -Name "DBDefaultQuota" -Value $usr_quota_default
        $report += $usr_reportObject
    }
}
# Output the report, sorted with the highest quota percentage at the top
$Report | Sort-Object QuotaPercent -Descending
