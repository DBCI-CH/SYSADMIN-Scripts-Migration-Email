<#Installer wrapper for DMA Agent (Deployment Pro
Dblum 09.02.2018
FIA Project 
#>
#set source agent location and exe name
$CurrentLocation = '\\aitfia.com\netlogon\DMA_Agent'
$exe = 'BitTitanDMASetup_E2F274A493CB8301__.exe'

#checks if c:\migrationO365 if not creates it and copies agent sources from netlogon

$path = "C:\MigrationO365\"

If(!(test-path $path))
{
      New-Item -ItemType Directory -Force -Path $path
      Copy-Item -Path $currentlocation\$exe  -destination $env:systemdrive\MigrationO365\BitTitanDMASetup_E2F274A493CB8301__.exe
     
}

#Check for Bittitan Presence if not install agent

$path = "C:\ProgramData\BitTitan"

If(!(test-path $path))
{
     
c:\MigrationO365\BitTitanDMASetup_E2F274A493CB8301__.exe -silent

     

}
