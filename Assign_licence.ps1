set-executionpolicy remotesigned
$usercredential=Get-Credential -message "please confirm your Office365 Admin account (if different)" ("$env:username" + "@fia.com")

#Sets path (default c:\temp to csv file input)
$inputfile = Get-FileName "C:\temp"
$users = get-content $inputfile


$LicenseName="FIAO365:STANDARDPACK"
$LicenseNameE3="FIAO365:ENTERPRISEPACK"
$LicenseNameE1="FIAO365:STANDARDPACK"

$l = $users.count
$i = 1

install-module msonline
import-module activedirectory

Connect-MsolService -Credential $UserCredential

<#
Import-Csv $CSVfileandpath | ForEach {
  $csv = $_
  Get-ADUser -Filter "UserPrincipalName -eq '$($csv.UPN)'" |
   
}
#>

for each ($user in $users)
{
$upn=$users.Primarymailboxemail
$License=$users.licencetype
$location=$users.location }

import-csv -path C:\temp\UserLicenceList.csv -Delimiter "," | forEach-Object {
  $UserEmailAddress = $_.PrimaryMailboxEmail
  $License=$_.licenseType
  $Location = $_.Location
                
  if($license -eq "E3" ) {$LicenseName=$LicenseNameE3}
  if($license -eq "E1" ) {$LicenseName=$LicenseNameE1}
  
  Write-host "Assignment of $LicenseName for $UserEmailAddress...$i/$l"
  Set-MsolUser -UserPrincipalName $UserEmailAddress -UsageLocation $Location
  Set-MsolUserLicense -UserPrincipalName $UserEmailAddress -AddLicenses $LicenseName
  Write-host "                 Done."
  $i += 1
}

Write-Host "License assignment finished. Wait for at least 5 minute before language configuration. Press any key to continue..."
$tt=read-host


$i=1
$Users | forEach-Object {
  $UserEmailAddress = $_.PrimarySmtpAddress
  $SpecialMailbox=$_.RecipientTypeDetails
    
  if($SpecialMailbox -eq "SharedMailbox"){$IsShared = $True} else {$IsShared = $False}
  if($SpecialMailbox -eq "RoomMailbox"){$IsRoom   = $true} else {$IsRoom = $False}
  if($SpecialMailbox -eq "EquipmentMailbox"){$IsResource   = $true} else {$IsResource = $False}

  $Language = $_.Language
  $TimeZone = $_.TimeZone
  
  Write-host "Configuration of regitonal settings for $UserEmailAddress - TimeZone-$TimeZone | Language-$Language...$i/$l"
      
  Set-MailboxRegionalConfiguration $UserEmailAddress -TimeZone $TimeZone –Language $Language -LocalizeDefaultFolderName:$true
 
  if($IsShared)
    {
    write-host "	Set Shared Mailbox for $UserEmailAddress"
    Set-MailBox $UserEmailAddress -type shared
    Start-Sleep -s 10
    Set-MsolUserLicense -UserPrincipalName $UserEmailAddress -RemoveLicenses $LicenseName
    }

  if($IsRoom)
    {
     write-host "	Set Room Mailbox for $UserEmailAddress"
     Set-MailBox $UserEmailAddress -type Room
     Start-Sleep -s 10
     Set-MsolUserLicense -UserPrincipalName $UserEmailAddress -RemoveLicenses $LicenseName
     }
  if($IsResource)
    {
    write-host "	Set Resource Mailbox for $UserEmailAddress"
    Set-MailBox $UserEmailAddress -type Equipment
    Start-Sleep -s 10
    Set-MsolUserLicense -UserPrincipalName $UserEmailAddress -RemoveLicenses $LicenseName
    }
  $i += 1
}



Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}