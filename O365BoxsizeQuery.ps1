get-mailbox -ResultSize unlimited | foreach {
  $MailUser = $_.Distinguishedname
  $Mailadd = $_.UserPrincipalName
  $stats= Get-MailboxStatistics $MailUser
  $usagelo = $_.Office
  $boxtype =$_.RecipientTypeDetails
    New-Object -TypeName PSObject -Property @{
      DisplayName = $stats.DisplayName
      MailboxSize = [math]::Round(($stats.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)
      Email = $Mailadd
      MailboxType = $boxtype
      Office = $usagelo
}} | Export-Csv -NoTypeInformation -Path "C:\temp\mailbox_size_gb.csv"
