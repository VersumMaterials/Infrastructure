

$list = Import-Csv C:\scripts\UserProvisioning-03.17-GoldList.csv
$MbxSize = @()

Foreach ($l in $lists){
$MbxSize += (
	Get-Mailbox -resultSize Unlimited $l.UserPrincipalNameAP | foreach{
$mbx = $_ | select Name,UserPrincipalName,PrimarySMTPAddress
Get-MailboxStatistics $_ | foreach{

$mbx | add-member -type noteProperty -name TotalItemSizeMB -value $_.TotalItemSize -PassThru |`
add-member -type noteProperty -name ItemCount -value $_.ItemCount -PassThru
  }
 }
)
} 

$MbxSize | select Name,UserPrincipalName,PrimarySMTPAddress,ItemCount,TotalItemSizeMB | export-csv C:\Scripts\VersumMailboxSizes.csv -NoTypeInformation