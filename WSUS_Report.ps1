
<#
.SYNOPSIS
    Audit and reporting for a WSUS server 
.NOTES
    Version: 1.0.3
    Name: WSUS_Report
    Updatedby : Start81 (DESMAREST Julien) 
    LastUpdate : 29/07/2022
    Changelog
        1.0.0 22/09/2020 : Initial release from https://learn-powershell.net/2014/05/04/wsus-server-html-report/
        1.0.1 25/09/2020 : Fix automatic approval list in case of $Null properties 
        1.0.2 04/11/2020 : Remove alias use in the script and remove the use of Shlwapi.dll
        1.0.3 29/07/2022 : Force usage of Measure-Object with count in all report
    ** Requires WSUS Administrator Console Installed or UpdateServices Module available **        
.DESCRIPTION
    This script create an html report about the Wsus server and produce a csv file with clients statistics
    The parameters of this script are in the configuration.ps1 file
.EXAMPLE
.\WSUS_Report.ps1
#>


#Region read Parameter
If ($MyInvocation.InvocationName) {
    $ScriptPath = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)
} else {
    $ScriptPath = $(Get-Location).Path
}
$MyConf = $ScriptPath + "\Configuration.ps1"
Try {
    .$MyConf
} Catch {
    $_
}
$WSUSServer = $WSUSServer.ToUpper()
#End region read Parameter
#region Helper Functions
Function Set-AlternatingCSSClass {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [string]$HTMLFragment,
        [Parameter(Mandatory=$True)]
        [string]$CSSEvenClass,
        [Parameter(Mandatory=$True)]
        [string]$CssOddClass
    )
    [xml]$xml = $HtmlFragment
    $Table = $xml.SelectSingleNode('table')
    $Classname = $CSSOddClass
    If ($Table.tr)
    {
        $Table.tr | ForEach-Object -Process{
            If ($Classname -Eq $CSSEvenClass) {
                $Classname = $CssOddClass
            } else {
                $Classname = $CSSEvenClass
            }
            $Class = $xml.CreateAttribute('class')
            $Class.value = $Classname
            $_.attributes.append($Class) | Out-null
        }
    }    
    $xml.innerxml | out-string
}
#endregion Helper Functions

#region Load WSUS Required Assembly
If (-Not (Get-Module -ListAvailable -Name UpdateServices)) {
    #Add-Type "$Env:ProgramFiles\Update Services\Api\Microsoft.UpdateServices.Administration.dll"
    $Null = [reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
} Else {
    Import-Module -Name UpdateServices
}
#endregion Load WSUS Required Assembly

#region CSS Layout
$Head = @"
    <style> 
        h1 {
            text-align:center;
            border-bottom:1px solid #666666;
            color:#009933;
        }
        TABLE {
            TABLE-LAYOUT: fixed; 
            FONT-SIZE: 100%; 
            WIDTH: 100%
        }
        * {
            margin:0
        }

        .pageholder {
            margin: 0px auto;
        }
                    
        td {
            VERTICAL-ALIGN: TOP; 
            FONT-FAMILY: Tahoma
        }
                    
        th {
            VERTICAL-ALIGN: TOP; 
            COLOR: #018AC0; 
            TEXT-ALIGN: left;
            background-color:DarkGrey;
            color:Black;
        }
        body {
            text-align:left;
            font-smoothing:always;
            width:100%;
        }
        .odd { background-color:#ffffff; }
        .even { background-color:#dddddd; }               
    </style>
"@
#endregion CSS Layout

#region Initial WSUS Connection
$ErrorActionPreference = 'Stop'
Try {
    $Wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($WSUSServer, $UseSSL, $Port)
    [Microsoft.UpdateServices.Administration.IUpdateServer] $Server = $Wsus
    $ClassIfications = $Server.GetUpdateClassIfications() | Where-Object {(($_.Title -like "*Security*") -or ($_.Title -like "*Critical*") -or ($_.Title -like"*Service Packs*"))}
} Catch {
    $_
    exit 3
}
$ErrorActionPreference = 'Continue'
#endregion Initial WSUS Connection

#region Pre-Stage -- Used in more than one location
$HtmlFragment = ''
$WSUSConfig = $Wsus.GetConfiguration()
$WSUSStats = $Wsus.GetStatus()
$TargetGroups = $Wsus.GetComputerTargetGroups()
$EmptyTargetGroups = $TargetGroups | Where-Object {
    ( $_.GetComputerTargets() | Measure-Object ).Count -Eq 0 -AND $_.Name -ne 'Unassigned Computers'
}

#Stale Computers
$ComputerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
$ComputerScope.ToLastReportedStatusTime = (Get-Date).AddDays(-$DaysComputerStale)
$StaleComputers = $Wsus.GetComputerTargets($ComputerScope) | ForEach-Object -Process {
    New-Object PSCustomObject -Property @{
        Computername = $_.FullDomainName
        ID =  $_.Id
        IpAddress = $_.IpAddress
        LastReported = $_.LastReportedStatusTime
        LastSync = $_.LastSyncTime
        #TargetGroups = ($_.GetComputerTargetGroups() | Select -Expand Name) -join ', '
        TargetGroups =[String]::Join(", ", ($_.GetComputerTargetGroups() |  Select-Object -Expand Name) )
    }
}
#Pending Reboots
$UpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
$UpdateScope.IncludedInstallationStates = 'InstalledPendingReboot'
$ClassIfications | ForEach-Object -Process {
    $Null = $UpdateScope.ClassIfications.add($_)
}
$ComputerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
$ComputerScope.IncludedInstallationStates = 'InstalledPendingReboot'
$GroupRebootHash = @{}
$ComputerPendingReboot = $Wsus.GetComputerTargets($ComputerScope) | ForEach-Object -Process {
    If($_.GetUpdateInstallationInfoPerUpdate($UpdateScope))
    {
       $Update = [String]::Join(", ",($_.GetUpdateInstallationInfoPerUpdate($UpdateScope) | ForEach-Object -Process {
       $Update = $_.GetUpdate()
       $Update.Title}))
       $UpdatesCount = ( $_.GetUpdateInstallationInfoPerUpdate($UpdateScope) | Measure-Object).Count
    }
    
    If ($Update) {
        $TempTargetGroups = ($_.GetComputerTargetGroups() |  Select-Object -Expand Name)
        $TempTargetGroups | ForEach-Object -Process {
            $GroupRebootHash[$_]++
        }
        New-Object PSCustomObject -Property  @{
            Computername = $_.FullDomainName
            #ID = $_.Id
            IpAddress = $_.IpAddress
            TargetGroups =[String]::Join(", ", $TempTargetGroups)
            Updates = $Update
            UpdatesCount = $UpdatesCount
        }
    }
} | Sort-Object Computername | Select-Object Computername, IpAddress, TargetGroups, UpdatesCount, Updates  
#Failed Installations
$UpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
$UpdateScope.IncludedInstallationStates = 'Failed'
#$UpdateScope.ApprovedStates='HasStaleUpdateApprovals'
$ClassIfications | ForEach-Object -Process {
    $Null = $UpdateScope.ClassIfications.add($_)
}
$ComputerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
$ComputerScope.IncludedInstallationStates = 'Failed'
$GroupFailHash = @{}
$ComputerHash = @{}
$UpdateHash = @{}
$ComputerFailInstall = $Wsus.GetComputerTargets($ComputerScope) | ForEach-Object -Process {
    $Computername = $_.FullDomainName
    $GetUpdateInstallationInfoPerUpdate=$_.GetUpdateInstallationInfoPerUpdate($UpdateScope)
    If ($GetUpdateInstallationInfoPerUpdate) 
    {
        $UpdatesCount = ( $GetUpdateInstallationInfoPerUpdate | Measure-Object ).Count
        $Update = [String]::Join(", ",($GetUpdateInstallationInfoPerUpdate | ForEach-Object -Process {
            $Update = $_.GetUpdate()
            $Update.title
            $ComputerHash[$Computername] += ,$Update.title
            $UpdateHash[$Update.title] += ,$Computername
        }))
        If ($Update) {
            $TempTargetGroups = ($_.GetComputerTargetGroups() |  Select-Object -Expand Name)
            $TempTargetGroups | ForEach-Object -Process {
                $GroupFailHash[$_]++
            }
            New-Object PSCustomObject -Property  @{
                Computername = $_.FullDomainName
                #ID = $_.Id
                IpAddress = $_.IpAddress
                TargetGroups = [String]::Join(", ",$TempTargetGroups -join ', ')
                Updates = $Update
                UpdatesCount = $UpdatesCount
            }
        }
    }
} | Sort-Object Computername  | Select-Object Computername,IpAddress, TargetGroups, UpdatesCount, Updates


#NotInstalled
$UpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
$UpdateScope.IncludedInstallationStates = 'NotInstalled'
$UpdateScope.ExcludeOptionalUpdates = $True
$ClassIfications | ForEach-Object -Process {
    $Null = $UpdateScope.ClassIfications.add($_)
}
#$UpdateScope.ApprovedStates='HasStaleUpdateApprovals'
$ComputerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
$ComputerScope.IncludedInstallationStates = 'NotInstalled'
$GroupNotInstalledHash = @{}
$ComputerNotInstalledHash = @{}
$UpdateNotInstalledHash = @{}#Pour pouvoir generer la listes not intalled by update (Pas encore implementé)
$ComputerNotInstalled = $Wsus.GetComputerTargets($ComputerScope) | ForEach-Object -Process {
    $Computername = $_.FullDomainName
    
    $GetUpdateInstallationInfoPerUpdate = $_.GetUpdateInstallationInfoPerUpdate($UpdateScope)
    If ($GetUpdateInstallationInfoPerUpdate) 
    {
        $UpdatesCount = ( $GetUpdateInstallationInfoPerUpdate | Measure-Object ).Count
        $Update = [String]::Join(", ",($GetUpdateInstallationInfoPerUpdate | ForEach-Object -Process {
            $Update = $_.GetUpdate()
            $Update.title
            $ComputerNotInstalledHash[$Computername] += ,$Update.title
            $UpdateNotInstalledHash[$Update.title] += ,$Computername
        }))
        If ($Update) {
            $TempTargetGroups = ($_.GetComputerTargetGroups() | Select-Object -Expand Name)
            $TempTargetGroups | ForEach-Object -Process {
                $GroupNotInstalledHash[$_]++
            }
            New-Object PSCustomObject -Property  @{
                Computername = $_.FullDomainName
                #ID = $_.Id
                IpAddress = $_.IpAddress
                TargetGroups = [String]::Join(", ",$TempTargetGroups -join ', ')
                Updates = $Update
                UpdatesCount = $UpdatesCount
            }
        }
    }
} |Sort-Object UpdatesCount -Descending | Select-Object Computername, IpAddress, TargetGroups, UpdatesCount, Updates
$NeedingUpdates = $($computerNotInstalled |Measure-Object | Select-Object -expand count)
$NeedingUpdatesLeXX =$($computerNotInstalled | Where-Object {$_.UpdatesCount -le $NbUpdatesCible} |Measure-Object | Select-Object -expand count)


#Export-CSV

$UpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope

$UpdateScope.ExcludeOptionalUpdates=$true
$ClassIfications | ForEach-Object -Process {
    $Null = $UpdateScope.ClassIfications.add($_)
}

$ComputerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope


$FileName = "RapportWsus_" + ((get-date).tostring('dd-MM-yyyy_HH_mm')) + ".csv"
$Computerlist = $Wsus.GetComputerTargets($ComputerScope) | ForEach-Object -Process {    
    $GetUpdateInstallationInfoPerUpdate = $_.GetUpdateInstallationInfoPerUpdate($UpdateScope)
    If ($GetUpdateInstallationInfoPerUpdate) 
    {
        $Downloaded = 0
        $Failed = 0
        $Installed = 0
        $InstalledPendingReboot = 0
        $NotApplicable = 0
        $NotInstalled = 0
        $Unknown = 0
        
        $Temp = $GetUpdateInstallationInfoPerUpdate | Group-Object UpdateInstallationState 
        $temp | ForEach-Object -Process {
            $cpt = $_
            switch ($cpt.Name)
            {
                Downloaded { $Downloaded = $_.Count + $Downloaded; break }
                Failed { $Failed = $cpt.Count + $Failed; break }
                Installed { $Installed = $cpt.Count + $Installed; break }
                InstalledPendingReboot { $InstalledPendingReboot = $cpt.Count + $InstalledPendingReboot; break }
                NotApplicable { $NotApplicable = $cpt.Count + $NotApplicable; break }
                NotInstalled { $NotInstalled = $cpt.Count + $NotInstalled; break }
                Unknown { $Unknown = $cpt.Count + $Unknown; break }
            }        
        }
        $Updatelvl = 0 
        $NbUpdates = $Unknown + $NotInstalled + $Failed + $Downloaded + $InstalledPendingReboot + $Installed + $NotApplicable
        If ($NbUpdates -Ne 0) { $UpdateLvl = (($Installed + $NotApplicable) * 100) / $NbUpdates }        
        $TempTargetGroups = ($_.GetComputerTargetGroups() | Select-Object -Expand Name)
        New-Object PSCustomObject -Property  @{
            Computername = $_.FullDomainName
            LastSync = $_.LastSyncTime
            LastReported =$_.LastReportedStatusTime
            ID = $_.Id
            IpAddress = $_.IpAddress
            TargetGroups = [String]::Join(", ",$TempTargetGroups -join ', ')
            Downloaded=$Downloaded
            Failed = $Failed
            Installed = $Installed
            InstalledPendingReboot = $InstalledPendingReboot
            NotApplicable = $NotApplicable
            NotInstalled = $NotInstalled
            Unknown = $Unknown
            UpdateLevel = $UpdateLvl
            OS = $_.OSDescription
                
        }
        
    }
} | Sort-Object NotInstalled -Descending
$Computerlist | Select-Object  Computername, OS, IpAddress, TargetGroups, LastSync, LastReported, Downloaded, Installed, InstalledPendingReboot, 
   NotInstalled, NotApplicable, Failed, Unknown, UpdateLevel | Export-Csv -NoTypeInformation -path $FileName -delimiter ";" -Encoding Utf8 
#endregion Export-CSV

#region WSUS SERVER INFORMATION
$Pre = @"
<div style='margin: 0px auto; BACKGROUND-COLOR:Blue;Color:White;font-weight:bold;FONT-SIZE: 16pt;'>
    WSUS Server Information
</div>
"@
    #region WSUS Version
    $WSUSVersion =New-Object PSCustomObject -Property @{
        Computername = $Wsus.ServerName
        Version = $Wsus.Version
        Port = $Wsus.PortNumber
        ServerProtocolVersion = $Wsus.ServerProtocolVersion
    }
    $Pre += @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
            WSUS Information
        </div>

"@
    $Body = $WSUSVersion | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $HtmlFragment += $Pre,$Body,$Post
    #endregion WSUS Version

    #region WSUS Server Content
    $Drive = $WSUSConfig.LocalContentCachePath.Substring(0,3)
    #$Data = Get-CIMInstance -ComputerName $WSUSServer -ClassName Win32_LogicalDisk -Filter "DeviceID='$drive'"
    #$UsedSpace = $Data.Size - $Data.Freespace
    #$PercentFree = "{0:P}" -f ($Data.Freespace / $Data.Size)

    $Data = Get-WMIObject -Class Win32_Volume | Where-Object {($_.name -Eq $Drive.toupper())} 
    $UsedSpace = $Data.Capacity - $Data.Freespace
    If ($Data) { #Si $Data est $null c'est que la requette wmi a echoué. sans doute pour un problème de droits
        $PercentFree = "{0:P}" -f ($Data.Freespace / $Data.Capacity) 
    } Else { 
        $PercentFree = 0
    }
    
    $Pre = @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
            WSUS Server Content Drive
        </div>

"@
    $WSUSDrive = New-Object PSCustomObject -Property @{
        LocalContentPath = $WSUSConfig.LocalContentCachePath
        TotalSpace = $([math]::Round(($Data.Capacity) / (1024*1024*1024),1)).ToString() + " Go"
        UsedSpace = $([math]::Round($UsedSpace / (1024*1024*1024),1)).ToString() + " Go"
        FreeSpace = $([math]::Round($Data.freespace / (1024*1024*1024),1)).ToString() + " Go"
        PercentFree = $PercentFree
    }
    $Body = $WSUSDrive | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $HtmlFragment += $Pre,$Body,$Post
    #endregion WSUS Server Content

    #region Last Synchronization
    $Synch = $Wsus.GetSubscription()
    $SynchHistory = $Synch.GetSynchronizationHistory()[0]
    $WSUSSynch = New-Object PSCustomObject -Property @{
        IsAuto = $Synch.SynchronizeAutomatically
        SynchTime = $Synch.SynchronizeAutomaticallyTimeOfDay
        LastSynch = $Synch.LastSynchronizationTime
        Result = $SynchHistory.Result
    }
    If ($SynchHistory.Result -Eq 'Failed') {
        $WSUSSynch = $WSUSSynch | Add-Member -MemberType NoteProperty -Name ErrorType -Value $SynchHistory.Error -PassThru |
        Add-Member -MemberType NoteProperty -Name ErrorText -Value $SynchHistory.ErrorText -PassThru
    }
    $Pre = @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
            Last Server Synchronization
        </div>

"@
    $Body = $WSUSSynch | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $HtmlFragment += $Pre,$Body,$Post
    #endregion Last Synchronization

    #region Upstream Server Config
    $WSUSUpdateConfig = New-Object PSCustomObject -Property @{
        SyncFromMU = $WSUSConfig.SyncFromMicrosoftUpdate
        UpstreamServer = $WSUSConfig.UpstreamWsusServerName
        UpstreamServerPort = $WSUSConfig.UpstreamWsusServerPortNumber
        SSLConnection = $WSUSConfig.UpstreamWsusServerUseSsl
    }
    $Pre = @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
            Upstream Server Information
        </div>

"@
    $Body = $WSUSUpdateConfig | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $HtmlFragment += $Pre,$Body,$Post
    #endregion Upstream Server Config

    #region Automatic Approvals
    $Rules = $Wsus.GetInstallApprovalRules()
    $ApprovalRules = $Rules | ForEach-Object -Process {
        If ($_.GetCategories()) {
            $MyCategories = [String]::Join(", ",($_.GetCategories() | Select-Object -ExpandProperty Title))
        }
        If ($_.GetUpdateClassIfications()) {
            $MyClassIfications = [String]::Join(", ",($_.GetUpdateClassIfications() | Select-Object -ExpandProperty Title))
        }
        If ($_.GetComputerTargetGroups()) {
            $MyTargetGroups = [String]::Join(", ",($_.GetComputerTargetGroups() | Select-Object -ExpandProperty Name) )
        }
        New-Object PSCustomObject -Property @{
            Name = $_.Name
            #ID = $_.ID
            Enabled = $_.Enabled
            Action = $_.Action
            Categories = $MyCategories
            ClassIfications = $MyClassIfications
            TargetGroups = $MyTargetGroups
        }

    }
    #$ApprovalRules
    $Pre = @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
            Automatic Approvals
        </div>

"@
    $Body = $ApprovalRules | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $HtmlFragment += $Pre,$Body,$Post
    #endregion Automatic Approvals

    #region WSUS Child Servers
    $ChildUpdateServers = $Wsus.GetChildServers()
    If ($ChildUpdateServers) {
        $ChildServers =  $ChildUpdateServers | ForEach-Object -Process {
            New-Object PSCustomObject -Property @{
                ChildServer = $_.FullDomainName
                Version = $_.Version
                UpstreamServer = $_.UpdateServer.Name
                LastSyncTime = $_.LastSyncTime
                SyncsFromDownStreamServer = $_.SyncsFromDownStreamServer
                LastRollUpTime = $_.LastRollupTime
                IsReplica = $_.IsReplica
            }
        }
    }
    $Pre = @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
            Child Servers
        </div>

"@
    $Body = $ChildServers | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $HtmlFragment += $Pre,$Body,$Post
    #endregion WSUS Child Servers

    #region Database Information
    $WSUSDB = $Wsus.GetDatabaseConfiguration()
    $DBInfo = New-Object PSCustomObject -Property @{
        DatabaseName = $WSUSDB.databasename
        Server = $WSUSDB.ServerName
        IsDatabaseInternal = $WSUSDB.IsUsingWindowsInternalDatabase
        Authentication = $WSUSDB.authenticationmode
    }
    $Pre = @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
            WSUS Database
        </div>

"@
    $Body = $DBInfo | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $HtmlFragment += $Pre,$Body,$Post
    #endregion Database Information

#endregion WSUS SERVER INFORMATION

#region CLIENT INFORMATION
$Pre = @"
<div style='margin: 0px auto; BACKGROUND-COLOR:Blue;Color:White;font-weight:bold;FONT-SIZE: 16pt;'>
    WSUS Client Information
</div>
"@
    #region Computer Statistics
    $WSUSComputerStats = New-Object PSCustomObject -Property @{
        TotalComputers = [int]$WSUSStats.ComputerTargetCount    
        "Stale($DaysComputerStale Days)" = ($StaleComputers | Measure-Object).count
        NeedingUpdates = $NeedingUpdates #[int]$WSUSStats.ComputerTargetsNeedingUpdatesCount
        "NeedingUpdates -le $NbUpdatesCible" = $NeedingUpdatesLeXX
        FailedInstall = [int]$WSUSStats.ComputerTargetsWithUpdateErrorsCount
        PendingReboot = ($ComputerPendingReboot | Measure-Object).Count
    }

    $Pre += @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
            Computer Statistics
        </div>

"@
    $Body = $WSUSComputerStats | Select-Object NeedingUpdates, "NeedingUpdates -le $NbUpdatesCible", PendingReboot, 
        FailedInstall, "Stale($DaysComputerStale Days)", TotalComputers | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $HtmlFragment += $Pre,$Body,$Post
    #endregion Computer Statistics

    #region Operating System
    $Pre = @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
            By Operating System
        </div>

"@
    $Body = $Wsus.GetComputerTargets() | Group-Object OSDescription |
         Select-Object @{ L = 'OperatingSystem';E = {$_.Name} }, Count | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'Odd'
    $Post = "<br>"
    $HtmlFragment += $Pre,$Body,$Post    
    #endregion Operating System

    #region Stale Computers
    $Pre = @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
            Stale Computers ($DaysComputerStale Days)
        </div>

"@
    $Body = $StaleComputers | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $HtmlFragment += $Pre, $Body, $Post
    #endregion Stale Computers

    #region Unassigned Computers
If ($ShowUnassignedComputers)
    {    
        $Unassigned = ($TargetGroups | Where-Object {
            $_.Name -Eq 'Unassigned Computers'
        }).GetComputerTargets() | ForEach-Object -Process {
            New-Object PSCustomObject -Property @{
                Computername = $_.FullDomainName
                OperatingSystem = $_.OSDescription
                #ID = $_.Id
                IpAddress = $_.IpAddress
                LastReported = $_.LastReportedStatusTime
                LastSync = $_.LastSyncTime
            }    
        }
        $Pre = @"
            <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
                Unassigned Computers (in Unassigned Target Group)
            </div>

"@
        $Body = $Unassigned | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
        $Post = "<br>"
        $HtmlFragment += $Pre,$Body,$Post
    }
    #endregion Unassigned Computers

    #region Failed Update Install
    If ($ComputerFailInstall) {
        $Pre = @"
            <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
                Failed Update Installations By Computer
            </div>

"@
        $Body = $ComputerFailInstall | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
        $Post = "<br>"
        $HtmlFragment += $Pre, $Body, $Post
    }
    #endregion Failed Update Install
    
    #region NotInstalled
    If ($ComputerNotInstalled) {
        $Pre = @"
            <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
                Not Installed Update By Computer
            </div>

"@
        $Body = $ComputerNotInstalled  | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
        $Post = "<br>"
        $HtmlFragment += $Pre, $Body, $Post
    }
    #endregion NotInstalled
    
    #region Pending Reboot 
    If ($ComputerPendingReboot) {
        $Pre = @"
            <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
                Computers with Pending Reboot
            </div>

"@
        $Body = $ComputerPendingReboot | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
        $Post = "<br>"
        $HtmlFragment += $Pre, $Body, $Post
    }
    #endregion Pending Reboot

#endregion CLIENT INFORMATION

#region UPDATE INFORMATION
$Pre = @"
<div style='margin: 0px auto; BACKGROUND-COLOR:Blue;Color:White;font-weight:bold;FONT-SIZE: 16pt;'>
    WSUS Update Information
</div>
"@
    #region Update Statistics
    $WSUSUpdateStats = New-Object PSCustomObject -Property @{
        TotalUpdates = [int]$WSUSStats.UpdateCount    
        Needed = [int]$WSUSStats.UpdatesNeededByComputersCount
        Approved = [int]$WSUSStats.ApprovedUpdateCount
        NotApprovedUpdate=[int]$WSUSStats.NotApprovedUpdateCount
        Declined = [int]$WSUSStats.DeclinedUpdateCount
        ClientInstallError = [int]$WSUSStats.UpdatesWithClientErrorsCount
        UpdatesNeedingFiles = [int]$WSUSStats.ExpiredUpdateCount    
    }
    $Pre += @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
            Update Statistics
        </div>

"@
    $Body = $WSUSUpdateStats | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $HtmlFragment += $Pre, $Body, $Post
    #endregion Update Statistics
    If ($ShowFailedIntallationbyUpdate)
    {
        #region Failed Update Installations
        $FailedUpdateInstall = $UpdateHash.GetEnumerator() | ForEach-Object -Process {
            New-Object PSCustomObject -Property @{
                Update = $_.Name
                NbComputer = ( $_.Value | Measure-Object ).Count
                #Computername = [String]::Join(",`r`n",($_.Value))
                Computername = [String]::Join(", ",($_.Value))
            }
        }
        If($FailedUpdateInstall){
        $Pre = @"
            <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
                Failed Update Installations By Update
            </div>

"@
        $Body = $FailedUpdateInstall | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
        $Post = "<br>"
        $HtmlFragment += $Pre,$Body,$Post}
        #endregion Failed Update Installations
    }
#endregion UPDATE INFORMATION

#region TARGET GROUP INFORMATION
$Pre = @"
<div style='margin: 0px auto; BACKGROUND-COLOR:Blue;Color:White;font-weight:bold;FONT-SIZE: 16pt;'>
    WSUS Target Group Information
</div>
"@
    #region Target Group Statistics
    $GroupStats = New-Object PSCustomObject -Property @{
        TotalGroups = [int]( $TargetGroups | Measure-Object ).Count
        TotalEmptyGroups = [int]( $EmptyTargetGroups | Measure-Object ).Count
    }
    $Pre += @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
            Target Group Statistics
        </div>

"@
    $Body = $GroupStats | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $HtmlFragment += $Pre,$Body,$Post
    #endregion Target Group Statistics

    #region Empty Groups
    If ($EmptyTargetGroups ){
    $Pre = @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:LightBlue;Color:Black;font-weight:bold;FONT-SIZE: 14pt;'>
            Empty Target Groups
        </div>

"@
    $Body = $EmptyTargetGroups | Select-Object Name, ID | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $HtmlFragment += $Pre,$Body,$Post}
    #endregion Empty Groups

#endregion TARGET GROUP INFORMATION

#region Compile HTML Report
$HTMLParams = @{
    Head = $Head
    Title = "WSUS Report for $WSUSServer"
    PreContent = "<H1><font color='white'>Please view in html!</font><br>$WSUSServer WSUS Report</H1>"
    PostContent = "$($htmlFragment)<i>Report generated on $((Get-Date).ToString())</i>" 
}
$Report = ConvertTo-Html @HTMLParams | Out-String
#endregion Compile HTML Report

If ($ShowFile) {
    $Report | Out-File WSUSReport.html
    Invoke-Item WSUSReport.html
}

#region Send Email
If ($SendEmail) {
    $EmailParams.Body = $Report
    $EmailParams.Attachment=$FileName
    Send-MailMessage @EmailParams -Encoding([System.Text.Encoding]::UTF8) -Verbose -Debug
    Remove-Item $FileName 
}
#endregion Send Email
