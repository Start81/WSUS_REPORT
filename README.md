# WSUS_REPORT
This script provide Wsus reporting by email. the boy is an html documment and there is an csv attachement

Original scipt from https://learn-powershell.net/2014/05/04/wsus-server-html-report/

#### Set Execution policy 

Tu run powershell script you must set executon policy:

```powershell
Set-ExecutionPolicy RemoteSigned
```

### Use case 

The following areas (Configuration.ps1) require some user interaction in order to properly run the script. 
Some may not be required, but should still be reviewed prior to running the script:

```powershell
#region User Specified WSUS Information
$WSUSServer = 'WSUS FQDN'
#Accepted values are "80","443","8530" and "8531"
$Port = 8531
$UseSSL = $true

#Specify when a computer is considered stale
$DaysComputerStale = 7 
#Maximun des Maj manquantes cible :
$NbUpdatesCible = 10
#Optional report part 
$ShowFailedIntallationbyUpdate = $false
$ShowUnassignedComputers = $false

#Send email of report
[bool]$SendEmail = $false
#Display HTML file
[bool]$ShowFile = $true
#endregion User Specified WSUS Information
#region User Specified Email Information
$EmailParams = @{
    To = 'Mailbox@toto.fr','Mailbox1@toto.fr','Mailbox2@toto.fr'
    From = 'WSUSReport@Toto.fr'    
    Subject = "$WSUSServer WSUS Report"
    SMTPServer = 'smtp.toto.fr'
    BodyAsHtml = $True
}
#endregion User Specified Email Information
```
We can run the scrit manualy or by a schedule task 

```Powershell
cd ScriptPath
.\WSUS_Report.ps1
```
