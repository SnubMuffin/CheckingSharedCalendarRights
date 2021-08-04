#Run the Set-ExecutionPolicy part first

Set-ExecutionPolicy RemoteSigned -Force
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session


$Result=@()
$allMailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object -Property Displayname,PrimarySMTPAddress
$totalMailboxes = $allMailboxes.Count
$i = 1 
$allMailboxes | ForEach-Object {
$mailbox = $_
Write-Progress -activity "Processing $($_.Displayname)" -status "$i out of $totalMailboxes completed"
$folderPerms = Get-MailboxFolderPermission -Identity "$($_.PrimarySMTPAddress):\Calendar"
$folderPerms | ForEach-Object {
$Result += New-Object PSObject -property @{ 
MailboxName = $mailbox.DisplayName
User = $_.User
Permissions = $_.AccessRights
}}
$i++
}
$Result | Select MailboxName, User, Permissions |
Export-CSV "C:\CalendarPermissions.CSV" -NoTypeInformation -Encoding UTF8