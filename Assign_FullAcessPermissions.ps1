
set-executionpolicy remotesigned



#functions

Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

$usercredential=get-credential
#Sets path (default c:\temp to csv file input)
$inputfile = Get-FileName "C:\temp"
$users = get-content $inputfile
#$CSVfileandpath = "c:\temp\delegation_comma.csv"

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
import-pssession $Session

#add full access permissions to csv members
import-csv $inputfile | foreach {Add-MailboxPermission -identity $_.PrimaryMailboxEmail -User $_.delegate -AccessRights FullAccess -InheritanceType all –AutoMapping $False }

#add sendas permission to csv members 
import-csv $inputfile | foreach {Add-RecipientPermission -identity $_.PrimaryMailboxEmail -Trustee $_.delegate -AccessRights SendAs -confirm:$False}

#add copy for sendonbehalf and sentas for sent items
import-csv $inputfile | foreach {Set-Mailbox $_.primarymailboxemail -MessageCopyForSendOnBehalfEnabled $true}
import-csv $inputfile | foreach {Set-Mailbox $_.primarymailboxemail -MessageCopyForSentAsenabled $true}


