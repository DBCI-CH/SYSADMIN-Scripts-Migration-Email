
#requires -runasadministrator
#requires -module msonline
install-module msonline
#import-module activedirectory


#set window #defines system box windows assembly type without input
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

set-executionpolicy remotesigned
$usercredential=Get-Credential -message "please confirm your Office365 Admin account (if different)" ("$env:username" + "@fia.com")
Connect-MsolService -Credential $UserCredential

$LicenseNameE3="FIAO365:ENTERPRISEPACK"
$LicenseNameE1="FIAO365:STANDARDPACK"
#giveslicenseoptions to disable (all but exchange online) ref : https://docs.microsoft.com/en-us/office365/enterprise/powershell/disable-access-to-services-while-assigning-user-licenses
$LO=new-msollicenseoptions -Accountskuid FIAO365:ENTERPRISEPACK -disabledplans "BPOS_S_TODO_2","FORMS_PLAN_E3","STREAM_O365_E3","Deskless","FLOW_O365_P2","POWERAPPS_O365_P2","Teams1","PROJECTWORKMANAGEMENT","YAMMER_ENTERPRISE","RMS_S_ENTERPRISE","OFFICESUBSCRIPTION","SWAY","INTUNE_O365","MCOSTANDARD","SHAREPOINTWAC","SHAREPOINTENTERPRISE"                      
           

$l = $users.count
$i = 1

#Sets path (default c:\temp to csv file input)
$inputfile = Get-FileName "C:\temp"
$users = get-content $inputfile

#example ew-MsolUser -UserPrincipalName allieb@litwareinc.com -DisplayName "Allie Bellew" -FirstName Allie -LastName Bellew -LicenseAssignment litwareinc:ENTERPRISEPACK -LicenseOptions $LO -UsageLocation US
import-csv $inputfile | foreach {New-MsolUser -UserPrincipalName <Account> -DisplayName <DisplayName> -FirstName <FirstName> -LastName <LastName> -LicenseAssignment <AccountSkuId> -LicenseOptions $LO -UsageLocation <CountryCode>}

Write-Host "License assignment finished. We wait 3 minutes (sleep in script) before following to shared section (mailbox creation). "
Start-countdown -seconds 180 - Message "Please wait"


write-host "connect to exchange online over powershell"
write-host "set mailbox to shared and remove license"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session

import-csv $inputfile | foreach {
 write-host "	Set Shared Mailbox Type for $_.Primarymailboxemail "
    Set-MailBox $_.Primarymailboxemail -type shared
    Start-Sleep -s 10
    write-host "Remove license for $_.Primarymailboxemail "
    Set-MsolUserLicense -UserPrincipalName $_.Primarymailboxemail  -RemoveLicenses $LicenseNameE3

  $i += 1}

Write-host "check script result for shared mailboxes"
import-csv $inputfile | foreach {get-mailbox -identity $_.primarymailboxemail | select name,userprincipalname,recipienttypedetails}


#function to browse for CSV file from Windows Explorer (cleaner input)

Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}



#source : https://community.spiceworks.com/scripts/show/1712-start-countdown 
Function Start-Countdown 
{   <#
    .SYNOPSIS
        Provide a graphical countdown if you need to pause a script for a period of time
    .PARAMETER Seconds
        Time, in seconds, that the function will pause
    .PARAMETER Messge
        Message you want displayed while waiting
    .EXAMPLE
        Start-Countdown -Seconds 30 -Message Please wait while Active Directory replicates data...
    .NOTES
        Author:            Martin Pugh
        Twitter:           @thesurlyadm1n
        Spiceworks:        Martin9700
        Blog:              www.thesurlyadmin.com
       
        Changelog:
           2.0             New release uses Write-Progress for graphical display while couting
                           down.
           1.0             Initial Release
    .LINK
        http://community.spiceworks.com/scripts/show/1712-start-countdown
    #>
    Param(
        [Int32]$Seconds = 10,
        [string]$Message = "Pausing for 10 seconds..."
    )
    ForEach ($Count in (1..$Seconds))
    {   Write-Progress -Id 1 -Activity $Message -Status "Waiting for $Seconds seconds, $($Seconds - $Count) left" -PercentComplete (($Count / $Seconds) * 100)
        Start-Sleep -Seconds 1
    }
    Write-Progress -Id 1 -Activity $Message -Status "Completed" -PercentComplete 100 -Completed
}