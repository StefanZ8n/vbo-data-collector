<# 
.NAME
    Veeam Backup for Microsoft Office 365 Data Collector
.SYNOPSIS
    Script to use for getting a detailed report per backup job modification
.DESCRIPTION
    Get relevant statistics data of Veeam Backup for Microsoft Office 365 v4 installations    
    
    Created for Veeam Backup for Microsoft Office 365 v4
    Released under the MIT license.
.LINK
    https://github.com/StefanZi
.VERSION
    v0.4.1, 07.02.2020
#>
[CmdletBinding()]
Param(
    [System.IO.FileInfo]$tmpPath = [System.IO.Path]::GetTempPath() + "vbo-data-collector",
    [bool]$Quick = $true
)
DynamicParam {
    Import-Module Veeam.Archiver.PowerShell
    $ParameterName = 'Organization'
    $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary 
    $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
    $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
    $ParameterAttribute.Mandatory = $false
    $AttributeCollection.Add($ParameterAttribute)
    $arrSet = Get-VBOOrganization | select -ExpandProperty Name
    $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
    $AttributeCollection.Add($ValidateSetAttribute)
    $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string[]], $AttributeCollection)
    $RuntimeParameterDictionary.Add($ParameterName, $RuntimeParameter)
    return $RuntimeParameterDictionary
}

begin {
    $Organization = $PsBoundParameters[$ParameterName]
    Import-Module Veeam.Archiver.PowerShell
    $result = @{ }

    class OrgInfo {
        [int]$Users
        [int]$LocalUsers
        [int]$ProtectedUsers
        [int]$ProtectedLocalUsers
        [int]$Mailboxes
        [int]$ProtectedMailboxes
        [int]$ProtectedArchives
        [int]$SPSites
        [int]$LocalSPSites
        [int]$ProtectedSPSites
        [int]$ProtectedLocalSPSites
        [int]$OneDrives
        [int]$ProtectedOneDrives
        [int]$ProtectedPersonalSites
        [System.Object[]] $Jobs
        $DataSize
    }

    class HWInfo {
        [string]$hostname
        $computerSystem
        $processor
        HWInfo(
            [string]$hostname
        ) {
            $this.hostname = $hostname
            $this.computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $hostname | select -Property Manufacturer, Model, NumberOfLogicalProcessors, NumberOfProcessors, TotalPhysicalMemory
            $this.processor = Get-WmiObject -Class Win32_Processor -ComputerName $hostname | select -Property DeviceId, MaxClockSpeed, Name
        }
    }

    class ProxyInfo : HWInfo {
        $repositories

        ProxyInfo(
            [string]$hostname
        ) : base($hostname) {            
            $proxy = Get-VBOProxy -Hostname $hostname
            $this.repositories = Get-VBORepository -Proxy $proxy | Select-Object Name, RetentionType ,Capacity, FreeSpace, RetentionPeriod, RetentionFrequencyType, EnableObjectStorageEncryption, ObjectStorageRepository.Name

        }
    }

    class JobInfo {
        [string]$jobname
        [string]$jobType
        [int]$objects
        #[int]$mailboxItems
        #[int]$archiveItems
        #[int]$sharePointItems
        #[int]$oneDriveItems
        $statistics

        JobInfo(
            $job
        ) {
            
            $this.jobname = $job.Name
            $this.jobType = $job.JobBackupType
            $this.objects = $this.getLastRunObjects($job)
            $this.statistics = $this.getStatistics($job)
        }

        [int] getLastRunObjects($job) {
            return (Get-VBOJobSession -Job $job -Last).Progress
        }

        [System.Object[]] getStatistics($job) {
            return Get-VBOJobSession -Job $job | Select-Object -last 10 Status, CreationTime -ExpandProperty Statistics
        }
    }


    New-Item -ItemType Directory -Force -Path $tmpPath > $Null
    $tmp = Get-Item -Path $tmpPath
}

process {
    if ($Organization) {
        $organizations = @()
        foreach ($Org in $Organization) {
            $organizations += (Get-VBOOrganization -Name $Org)
        }
    } else {
        $organizations = Get-VBOOrganization
    } 

    $result.add("orgs", @{ })

    foreach ($org in $organizations) {
        $thisOrg = [OrgInfo]::new()
        $users = Get-VBOOrganizationUser -Organization $org -Type User
        $thisOrg.Users = $users.length
        $thisOrg.LocalUsers = ($users | ? { $_.OnPremiseId }).length

        $protectedUsers = Get-VBOOrganizationUser -Organization $org -Type User | ? { $_.IsBackedUp -eq $true }
        $thisOrg.ProtectedUsers = $protectedUsers.length
        $thisOrg.ProtectedLocalUsers = $protectedUsers | ? { $_.OnPremiseId }
        
        
        $children = $tmp.getFiles()
        Get-VBOMailboxProtectionReport -Path $tmpPath -Organization $org -Format csv
        $childrenChanged = $tmp.getFiles()
        $report = (Compare-Object -ReferenceObject $children -DifferenceObject $childrenChanged).InputObject
        $mailboxes = Import-Csv -Path $report.Fullname
        # TODO: Cleanup Report File or cleanup temp folder at the end!!
        # TODO: Determine local mailboxes

        $thisOrg.Mailboxes = $mailboxes.Length
        $thisOrg.ProtectedMailboxes = ($mailboxes | ? { $_.'Protection Status' -eq "Protected" }).length
        
        $sites = Get-VBOOrganizationSite -Organization $org
        $thisOrg.SPSites = $sites.length
        $thisOrg.LocalSPSites = ($sites | ? { $_.IsCloud -eq $false }).length
        $thisOrg.ProtectedSPSites = ($sites | ? { $_.IsBackedUp -eq $true }).length
        $thisOrg.ProtectedLocalSPSites = ($sites | ? { $_.IsCloud -eq $false -and $_.IsBackedUp -eq $true }).length

        $thisOrg.Jobs = (Get-VBOJob -Organization $org | ? { $_.IsEnabled -eq $true }) | % { [JobInfo]::new($_); }

        if ($Quick -eq $false) {
            $m = Measure-VBOOrganizationFullBackupSize -Organization $org
            $thisOrg.DataSize = @{
                Mailboxes = "{0:N2}" -f ($m.MailboxSize/(1024*1024*1024));
                Archives = "{0:N2}" -f ($m.ArchiveMailboxSize/(1024*1024*1024));
                Sites = "{0:N2}" -f ($m.SiteSize/(1024*1024*1024));
                PersonalSites = "{0:N2}" -f ($m.PersonalSiteSize/(1024*1024*1024));
                OneDrives = "{0:N2}" -f ($m.OneDriveSize/(1024*1024*1024));
            }
        }

        $entitydata = Get-VBORepository | % { 
            $repo = $_;
            Get-VBOEntityData -Type User -Repository $repo | ? { $_.Email -in $users.UserName }
        }
        $thisOrg.ProtectedArchives = ($entitydata | ? { $_.IsArchiveBackedUp -eq $true }).length
        $thisOrg.ProtectedOneDrives = ($entitydata | ? { $_.IsOneDriveBackedUp -eq $true }).length
        $thisOrg.ProtectedPersonalSites = ($entitydata | ? { $_.IsPersonalSiteBackedUp -eq $true}).length

        $result.orgs.add($org.Name, $thisOrg)
    }
    
    $protectedSum = (($result.orgs.Values | % { $_.ProtectedUsers }) | Measure-Object -Sum).Sum
    $localSum = (($result.orgs.Values | % { $_.ProtectedLocalUsers }) | Measure-Object -Sum).Sum
    $onlineSum = $protectedSum - $localSum

    $SPSum = (($result.orgs.Values | % { $_.ProtectedSPSites }) | Measure-Object -Sum).Sum
    $localSPSum = (($result.orgs.Values | % { $_.ProtectedLocalSPSites }) | Measure-Object -Sum).Sum
    $onlineSPSum = $SPSum - $localSPSum

    $setup = @{ online = @{ 
            exchangeusers = $onlineSum;
            spsites       = $onlineSPSum;
        };
        local          = @{ 
            exchangeusers = $localSum;
            spsites       = $localSPSum;
        }
    }

    $result.Add("setup", $setup)

    # Setup & User numbers

    $protected = @{ 
        users     = (($result.orgs.Values | % { $_.ProtectedUsers }) | Measure-Object -Sum).Sum;
        mailboxes = (($result.orgs.Values | % { $_.ProtectedMailboxes }) | Measure-Object -Sum).Sum;
        archives = (($result.orgs.Values | % { $_.ProtectedArchives }) | Measure-Object -Sum).Sum;
        spsites = (($result.orgs.values | % { $_.ProtectedSPSites }) | Measure-Object -Sum).Sum;
        onedrives = (($result.orgs.values | % { $_.ProtectedOnedrives }) | Measure-Object -Sum).Sum;

    }
    $result.Add("protected", $protected)

    # Object Numbers


    # General VBO architecture data

    $result.Add("architecture", @{
            vboVersion  = (Get-WmiObject -Class Win32_Product | ? { $_.Caption -match "Veeam Backup for Microsoft Office 365.*" }).Version;
            controllers = @(Get-VBOServerComponents | ? { $_.Name -match ".*Server.*" } | % { [HWInfo]::new($_.ServerName) });
            proxies     = @(Get-VBOProxy | % { [ProxyInfo]::new($_.Hostname) });        
        })

    
    $result | ConvertTo-Json -Depth 10
}