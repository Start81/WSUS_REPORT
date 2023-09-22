#region User Specified WSUS Information
$WSUSServer = 'WSUS FQDN'
$Port = 'PORT NUMBER'
#Accepted values are "80","443","8530" and "8531"
$UseSSL = $True

#Specify when a computer is considered stale
$DaysComputerStale = 7 
#Max Target uptates to install :
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