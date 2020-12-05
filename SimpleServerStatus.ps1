
<#PSScriptInfo

.VERSION 0.1.0

.GUID a5cd742b-853b-4f2e-a462-685d13b1048a

.AUTHOR Alex Ferguson

.COMPANYNAME 

.COPYRIGHT 2020

.TAGS 

.LICENSEURI 

.PROJECTURI https://github.com/Arramol/SimpleServerStatus

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES


#>

<# 

.DESCRIPTION 
 PowerShell script that sends an email report on the status of the server. 

#> 

Param(
    [ValidateScript({Test-Path -Path $_})]
    [System.IO.FileInfo]$ConfigFilePath = "$PSScriptRoot\config.xml"
)

Function Get-ScriptSettings{
    $Settings = Import-Clixml -Path $ConfigFilePath -ErrorAction Stop
    $PropertiesToValidate = @('EmailCredentialsFilePath', 'EmailTo', 'SmtpServer', 'SmtpPort', 'UseSsl', 'IncludeBackups')
    ForEach($Property in $PropertiesToValidate){
        If($Settings.PSObject.Properties.Name -notcontains $Property){
            Throw "Failed to validate imported settings. Property $Property was not found in $ConfigFilePath."
        }
    }

    Return $Settings
}

Function Send-EmailReport{
    Param(
        [Parameter(Mandatory=$true)]
        [PSObject]$Settings,

        [Parameter(Mandatory=$true)]
        [pscredential]$Credentials,

        [Parameter(Mandatory=$true)]
        $ReportData
    )

    Send-MailMessage -From $Credentials.UserName -To $Settings.EmailTo -Subject 'Server Health Report' -Body $ReportData -SmtpServer $Settings.SmtpServer -Port $Settings.SmtpPort -UseSsl:$Settings.UseSsl -Credential $Credentials
}

Function Measure-StorageUsage{
    $StorageProblems = New-Object -TypeName System.Collections.Generic.List[String]
    $DetailedReport = "Storage Status:`n"
    [Array]$Volumes = Get-Volume | Where-Object {$_.DriveType -eq 'Fixed'} | ForEach-Object {
        $UsedSpaceBytes = $_.Size - $_.SizeRemaining
        $PercentUsage = [Math]::Round(($UsedSpaceBytes / $_.Size * 100), 2)
        $Problems = 'None'
        If($_.Size -gt (1TB)){
            $Capacity = "$([Math]::Round(($_.Size / 1TB), 2)) TB"
            $FreeSpace = "$([Math]::Round(($_.SizeRemaining / 1TB), 2)) TB"
        }
        Else{
            $Capacity = "$([Math]::Round(($_.Size / 1GB), 2)) GB"
            $FreeSpace = "$([Math]::Round(($_.SizeRemaining / 1GB), 2)) GB"
        }

        If($_.HealthStatus -ne 'Healthy'){
            $Problems = 'Drive health'
        }
        If($PercentUsage -gt 80){
            If($Problems -ne 'None'){
                $Problems += "`nLow free space"
            }
            Else{
                $Problems = 'Low free space'
            }
        }

        [PSCustomObject] @{
            'Drive Letter' = $_.DriveLetter
            Name = $_.FriendlyName
            'File System' = $_.FileSystemType
            'Health Status' = $_.HealthStatus
            'Percent Used' = $PercentUsage
            'Free Space' = $FreeSpace
            Capacity = $Capacity
            Problems = $Problems
        }
    }

    If(($Volumes | Where-Object {$_.HealthStatus -ne 'Healthy'}).Count -gt 0){
        $StorageProblems.Add('One or more volumes has a problem health status.')
    }
    If(($Volumes | Where-Object {$_.'Percent Used' -gt 80}).Count -gt 0){
        $StorageProblems.Add('One or more volumes has low remaining free space.')
    }

    $Healthy = $StorageProblems.Count -gt 0
    $DetailedReport += $Volumes | Format-Table 'Drive Letter', Name, 'File System', 'Health Status', 'Percent Used', 'Free Space', Capacity, Problems | Out-String
    $SummaryObject = [PSCustomObject] @{
        Category = 'Storage'
        Healthy = $Healthy
    }

    $StorageReport = [PSCustomObject] @{
        Volumes = $Volumes
        Problems = $StorageProblems
        Healthy = $Healthy
        DetailedReport = $DetailedReport
        SummaryObject = $SummaryObject
    }

    Return $StorageReport
}

Function Get-MissingServices{
    Param(
        [Parameter(Mandatory=$true)]
        [PSObject]$Settings
    )
    
    $DetailedReport = "Services Status:`n"
    [Array]$ServicesToIgnore = @()
    If($Settings.IgnoredServicesListPath){
        Try{
            [Array]$ServicesToIgnore = Get-Content -Path $Settings.IgnoredServicesListPath -ErrorAction Stop
        }
        Catch{
            Write-Error "Failed to read list of services to ignore from $($Settings.IgnoredServicesListPath)."
        }
    }

    [Array]$MissingServices = Get-Service | Where-Object {$_.StartType -eq 'Automatic' -and $_.Status -ne 'Running' -and $ServicesToIgnore -notcontains $_.Name}
    If($MissingServices.Count -eq 0){
        $Healthy = $true
        $DetailedReport += 'Services are healthy. All auto-start services are running.'
    }
    Else{
        $Healthy = $false
        $DetailedReport += "The following auto-start services are not running:`n"
        $DetailedReport += $MissingServices | Sort-Object DisplayName | Format-Table DisplayName, Name, Status | Out-String
    }

    $SummaryObject = [PSCustomObject] @{
        Category = 'Services'
        Healthy = $Healthy
    }

    $ServicesStatus = [PSCustomObject] @{
        'Missing Services' = $MissingServices
        Healthy = $Healthy
        DetailedReport = $DetailedReport
        SummaryObject = $SummaryObject
    }

    Return $ServicesStatus
}

Function Get-BackupResult{
    If((Get-Module).Name -notcontains 'WindowsServerBackup'){
        Import-Module -Name WindowsServerBackup -ErrorAction Stop
    }

    $DetailedReport = "Backup Status:`n"
    $BackupSummary = Get-WBSummary -ErrorAction Stop
    $BackupHealthy = $true
    $BackupProblems = New-Object -TypeName System.Collections.Generic.List[String]
    If($BackupSummary.LastSuccessfulBackupTime){
        If($BackupSummary.LastSuccessfulBackupTime -ne $BackupSummary.LastBackupTime){
            $BackupHealthy = $false
            $BackupProblems.Add("Most recent backup job failed at $($BackupSummary.LastBackupTime).")
        }
        If($BackupSummary.LastSuccessfulBackupTime -lt (Get-Date).AddDays(-1)){
            $BackupHealthy = $false
            $BackupProblems.Add("No successful backups in the last 24 hours. Last successful backup was at $($BackupSummary.LastSuccessfulBackupTime).")
        }
    }
    Else{
        $BackupHealthy = $false
        $BackupProblems.Add("No record of any successful backups!")
    }

    If($BackupHealthy){
        $DetailedReport += "Backups are healthy. Last backup completed successfully at $($BackupSummary.LastSuccessfulBackupTime)."
    }
    Else{
        $DetailedReport += "Error with backups. Last successful backup: $($BackupSummary.LastSuccessfulBackupTime)"
        ForEach($Problem in $BackupProblems){
            $DetailedReport += "`n$Problem"
        }
    }

    $SummaryObject = [PSCustomObject] @{
        Category = 'Backups'
        Healthy = $Healthy
    }

    $BackupResult = [PSCustomObject] @{
        Healthy = $BackupHealthy
        'Last Successful Backup' = $BackupSummary.LastSuccessfulBackupTime
        Problems = $BackupProblems
        DetailedReport = $DetailedReport
        SummaryObject = $SummaryObject
    }

    Return $BackupResult
}

Function Get-UpdateStatus{
    $DetailedReport = "Update Status:`n"
    [Array]$RecentUpdates = Get-Hotfix | Where-Object {$_.InstalledOn -ge (Get-Date).AddDays(-45)} | Sort-Object InstalledOn -Descending
    If($RecentUpdates.Count -gt 0){
        $Healthy = $true
        $DetailedReport += "Updates are healthy. The following updates were installed in the past 45 days:`n"
        $DetailedReport += $RecentUpdates | Select-Object HotfixID, Description, InstalledOn | Format-Table | Out-String
    }
    Else{
        $Healthy = $false
        $DetailedReport += 'PROBLEM: No updates have been installed in the past 45 days.'
    }
    
    $SummaryObject = [PSCustomObject] @{
        Category = 'Updates'
        Healthy = $Healthy
    }

    $UpdateStatus = [PSCustomObject] @{
        Healthy = $Healthy
        'Updates installed in the last 45 days' = $RecentUpdates
        DetailedReport = $DetailedReport
        SummaryObject = $SummaryObject
    }

    Return $UpdateStatus
}

$Settings = Get-ScriptSettings
$Credentials = Import-CliXml -Path $Settings.EmailCredentialsFilePath -ErrorAction Continue
$Summary = New-Object -TypeName System.Collections.Generic.List[Object]
If($Credentials.UserName -notlike "*@*"){
    Throw "$($Settings.EmailCredentialsFilePath) does not appear to contain valid credentials, or the credentials were encrypted by a different user account."
}

Clear-Host
Write-Host "Getting storage status." -ForegroundColor Cyan
$StorageUsage = Measure-StorageUsage
$Summary.Add($StorageUsage.SummaryObject)
$DetailedReport = $StorageUsage.DetailedReport + "============================================================`n`n"

Write-Host "Checking for missing services." -ForegroundColor Cyan
$ServicesStatus = Get-MissingServices -Settings $Settings
$Summary.Add($ServicesStatus.SummaryObject)
$DetailedReport += $ServicesStatus.DetailedReport + "============================================================`n`n"
If($Settings.IncludeBackups){
    Write-Host "Checking backup status." -ForegroundColor Cyan
    Try{
        $BackupResult = Get-BackupResult
        $Summary.Add($BackupResult.SummaryObject)
        $DetailedReport += $BackupResult.DetailedReport + "============================================================`n`n"
    }
    Catch{
        $BackupResult = 'Failed to retrieve backup summary. Check that the WindowsServerBackup PowerShell module is installed and that a backup schedule is configured.'
        $DetailedReport += $BackupResult + "============================================================`n`n"
    }
}
Else{
    Write-Host "Skipping backup status check." -ForegroundColor Cyan
    $BackupResult = 'Not included'
}
Write-Host "Checking update status." -ForegroundColor Cyan
$UpdateStatus = Get-UpdateStatus
$Summary.Add($UpdateStatus.SummaryObject)
$DetailedReport += $UpdateStatus.DetailedReport

$EmailBody = "Summary:`n"
$EmailBody += $Summary | Format-Table Category, Healthy | Out-String
$EmailBody += "============================================================`n"
$EmailBody += '============================================================'
$EmailBody += "`n`nDetailed Reports:`n"
$EmailBody += $DetailedReport

If($Credentials){
    Send-EmailReport -Settings $Settings -Credentials $Credentials -ReportData $EmailBody
}
Else{
    $EmailBody
}