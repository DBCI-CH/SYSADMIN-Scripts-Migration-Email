
set-executionpolicy remotesigned

$usercredential=get-credential
$CSVfileandpath = "c:\temp\delegation_comma.csv"

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
import-pssession $Session

import-csv $CSVfileandpath | foreach object {New-Mailbox -Name $_.primarymailboxemail -Alias $_.alias –Shared -PrimarySmtpAddress $_.primarymailboxemail }