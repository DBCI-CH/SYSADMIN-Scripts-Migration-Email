#Reapply Full Permission on Mailbox
set-executionpolicy remotesigned
$usercredential=get-credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session
$DelegateMailboxEmail = Read-Host -Prompt 'Input your delegate or share mailbox email'
$User = Read-Host -Prompt 'Input User Name Email that must be given rights'


#remove full access permissions to csv members
remove-MailboxPermission -identity $_.delegatemailboxemail -User $_.user -AccessRights FullAccess -InheritanceType all
Write-Host "full access rights removed on mailbox $_.delegatemailboxemail for user $_.user" 

#remove sendas permission to csv members 
Remove-RecipientPermission -identity $_.delegatemailboxemail -Trustee $_.user -AccessRights SendAs -confirm:$False
Write-Host "sendas rights removed on mailbox $_.delegatemailboxemail for user $_.user" 

#Reapply Permission to user 
#add full access permissions to csv members
Add-MailboxPermission -identity $_.delegatemailboxemail -User $_.user -AccessRights FullAccess -InheritanceType all –AutoMapping $False

#add sendas permission to csv members 
Add-RecipientPermission -identity $_.delegatemailboxemail -Trustee $_.user -AccessRights SendAs -confirm:$False