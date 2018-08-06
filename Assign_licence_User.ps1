<# Assign_licence_user.ps1 V2.0
David Blum
This script does the following :
- Get Office365 Admin credentials
- Connects to MSOnline with credentials provided 
- Gathers CSV Input with required user list (userprincipalname (email), License, Localisation)
- Assign E3 License to users and disables all services but Exchange Online
Required : 
#>


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

#sets required conditions for script 
install-module msonline
#requires -runasadministrator
#requires -module msonline
#import-module activedirectory

#set window #defines system box windows assembly type without input
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

#enable execution policy for powershell script (Local script allowed)
set-executionpolicy remotesigned

#Gets Office 365 credentials
$usercredential=Get-Credential -message "please confirm your Office365 Admin account (if different)" ("$env:username" + "@fia.com")

#connects to office 365
Connect-MsolService -Credential $UserCredential

#Sets path (default c:\temp to csv file input) and open explorer to browse for file 
$inputfile = Get-FileName "C:\temp"
$users = get-content $inputfile

#defines license for tenant
$LicenseNameE3="FIAO365:ENTERPRISEPACK"
$LicenseNameE1="FIAO365:STANDARDPACK"

#giveslicenseoptions to disable all but exchange online ref : https://docs.microsoft.com/en-us/office365/enterprise/powershell/disable-access-to-services-while-assigning-user-licenses
$options=new-msollicenseoptions -Accountskuid FIAO365:ENTERPRISEPACK -disabledplans "BPOS_S_TODO_2","FORMS_PLAN_E3","STREAM_O365_E3","Deskless","FLOW_O365_P2","POWERAPPS_O365_P2","Teams1","PROJECTWORKMANAGEMENT","YAMMER_ENTERPRISE","RMS_S_ENTERPRISE","OFFICESUBSCRIPTION","SWAY","INTUNE_O365","MCOSTANDARD","SHAREPOINTWAC","SHAREPOINTENTERPRISE"                      
           
#count users and set counter for script operations
$l = $users.count
$i = 1

#import-csv file for users, add usage location (Country ISO Code),

import-csv $inputfile | foreach {Set-MsolUser -UserPrincipalName $_.Primarymailboxemail  -UsageLocation $_.Location

    Write-host "Assignment of $LicenseNameE3 for entry ...$i/$l"
  Set-MsolUserLicense -UserPrincipalName $_.Primarymailboxemail  -AddLicenses $LicenseNameE3 
  #disable unwanted plans 
  write-host "disable unused plans for entry... $i/$l"
  set-msoluserlicense -UserPrincipalName $_.Primarymailboxemail -licenseoptions $options
  Write-host "Done."
  $i += 1

  }

#restart counter at 1
$i =1

#Runs housekeeping and validation 
Write-host "check script result for user mailboxes"
import-csv $inputfile | foreach {get-mailbox -identity $_.primarymailboxemail | select name,userprincipalname,recipienttypedetails}
  

Write-Host "License assignment finished. Wait for at least 5 minute before language configuration. Press any key to continue..."
$tt=read-host

# Functions used in script 

