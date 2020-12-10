
<#PSScriptInfo

.VERSION 0.3.0

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

    Send-MailMessage -From $Credentials.UserName -To $Settings.EmailTo -Subject 'Server Health Report' -Body $ReportData -BodyAsHtml -SmtpServer $Settings.SmtpServer -Port $Settings.SmtpPort -UseSsl:$Settings.UseSsl -Credential $Credentials
}

Function Measure-StorageUsage{
    $StorageProblems = New-Object -TypeName System.Collections.Generic.List[String]
    $DetailedReport = "Storage Status:`n"
    [Array]$Volumes = Get-Volume | Where-Object {($_.DriveType -eq 'Fixed') -and (($_.DriveLetter -or $_.FriendlyName))} | ForEach-Object {
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
        #Problems = $StorageProblems
        #Healthy = $Healthy
        DetailedReport = $DetailedReport
        HtmlReport = (Format-HtmlStorageDetails -Volumes $Volumes)
        SummaryObject = $SummaryObject
    }

    Return $StorageReport
}

Function Format-HtmlStorageDetails{
    Param(
        [Array]$Volumes = @()
    )

    $HtmlStorage = New-TableHeader -SectionName 'Storage Details' -ColumnNames @('Drive Letter', 'Drive Name', 'File System', 'Health Status', 'Percent Used', 'Free Space', 'Capacity')

    ForEach($Volume in $Volumes){
        If($Volume.'Health Status' -eq 'Healthy'){
            $HealthColor = $BgColorHealthy
        }
        Else{
            $HealthColor = $BgColorError
        }
        
        If($Volume.'Percent Used' -lt 80){
            $UsageColor = $BgColorHealthy
        }
        ElseIf($Volume.'Percent Used' -lt 90){
            $UsageColor = $BgColorWarning
        }
        Else{
            $UsageColor = $BgColorError
        }

        $HtmlStorage += @"
  <tr>
    <td>$($Volume.'Drive Letter')</td>
    <td>$($Volume.Name)</td>
    <td>$($Volume.'File System')</td>
    <td bgcolor="$HealthColor;">$($Volume.'Health Status')</td>
    <td bgcolor="$UsageColor;">$($Volume.'Percent Used')</td>
    <td>$($Volume.'Free Space')</td>
    <td>$($Volume.Capacity)</td>
  </tr>
"@
    }

    $HtmlStorage += "</table>"
    Return $HtmlStorage
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

    $HtmlReport = Format-HtmlServiceDetails -Services $MissingServices

    $SummaryObject = [PSCustomObject] @{
        Category = 'Services'
        Healthy = $Healthy
    }

    $ServicesStatus = [PSCustomObject] @{
        DetailedReport = $DetailedReport
        SummaryObject = $SummaryObject
        HtmlReport = $HtmlReport
    }

    Return $ServicesStatus
}

Function Format-HtmlServiceDetails{
    Param(
        [Array]$Services = @()
    )

    $HtmlServices = New-TableHeader -SectionName 'Services Details' -ColumnNames @('Display Name', 'Name', 'Status')
    If($Services.Count -eq 0){
        $HtmlServices += @"
    <tr>
        <td colspan="3" bgcolor="$BgColorHealthy;">No missing services</td>
    </tr>
"@
    }
    Else{
        ForEach($Service in $Services){
            $HtmlServices += @"
    <tr>
      <td>$($Service.DisplayName)</td>
      <td>$($Service.Name)</td>
      <td bgcolor="$BgColorError;">$($Service.Status)</td>
"@
        }
    }

    $HtmlServices += "</table>"
    Return $HtmlServices
}

Function Get-BackupResult{
    If((Get-Module).Name -notcontains 'WindowsServerBackup'){
        Import-Module -Name WindowsServerBackup -ErrorAction Stop
    }

    $DetailedReport = "Backup Status:`n"
    $BackupSummary = Get-WBSummary -ErrorAction Stop
    $BackupHealthy = $true
    $LastSuccessColor = $BgColorHealthy
    $BackupProblems = New-Object -TypeName System.Collections.Generic.List[String]
    If($BackupSummary.LastSuccessfulBackupTime){
        If($BackupSummary.LastSuccessfulBackupTime -ne $BackupSummary.LastBackupTime){
            $BackupHealthy = $false
            $LastSuccessColor = $BgColorWarning
            $BackupProblems.Add("Most recent backup job failed at $($BackupSummary.LastBackupTime).")
        }
        If($BackupSummary.LastSuccessfulBackupTime -lt (Get-Date).AddDays(-1)){
            $BackupHealthy = $false
            $LastSuccessColor = $BgColorError
            $BackupProblems.Add("No successful backups in the last 24 hours. Last successful backup was at $($BackupSummary.LastSuccessfulBackupTime).")
        }
    }
    Else{
        $BackupHealthy = $false
        $LastSuccessColor = $BgColorError
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

    $HtmlReport = Format-HtmlBackupDetails -Problems $BackupProblems -BackupSummary $BackupSummary -LastSuccessColor $LastSuccessColor

    $BackupResult = [PSCustomObject] @{
        DetailedReport = $DetailedReport
        SummaryObject = $SummaryObject
        HtmlReport = $HtmlReport
    }

    Return $BackupResult
}

Function Format-HtmlBackupDetails{
    Param(
        [Array]$Problems = @(),

        $BackupSummary,

        $LastSuccessColor
    )

    $HtmlBackup = @"
    <table style="width: 33%" style="border-collapse: collapse; border: 1 px solid #000000;">
      <tr>
        <td colspan="2" bgcolor="$BgColorSectionHeader" style="color: #FFFFFF; font-size: large; height: 35px;">
          'Backup Details'
        </td>
      </tr>
      <tr>
"@

    If($Problems.Count -gt 0){
        $HtmlBackup += @"
        <td colspan="2" style=`"text-align: center;`"><b>Problems</b></td>
      </tr>
"@
        ForEach($Problem in $Problems){
            $HtmlBackup += @"
      <tr>
        <td colspan="2;">$Problem</td>
      </tr>
"@
        }
    $HtmlBackup += "  <tr>"
    }

    $HtmlBackup += @"
        <td colspan=`"2`" style=`"text-align: center;`"><b>Backup Job(s)</b></td>"
      </tr>
      <tr>
        <td style = "text-align: center;"><b>Job Timestamp</b></td>
        <td style = "text-align: center;"><b>Result</b></td>
      </tr>
"@
    If($BackupSummary.LastBackupTime){
        If($BackupSummary.LastSuccessfulBackupTime -ne $BackupSummary.LastBackupTime){
            $HtmlBackup += @"
      <tr>
        <td>$($BackupSummary.LastBackupTime)</td>
        <td bgcolor="$BgColorError;">Failed</td>
      </tr>
"@
            
        }
    }
    Else{
        $HtmlBackup += @"
      <tr>
        <td colspan="2" bgcolor="$BgColorError;">No backup jobs found</td>
      </tr>
"@
    }
    If($BackupSummary.LastSuccessfulBackupTime){
        $HtmlBackup += @"
      <tr>
        <td>$($BackupSummary.LastSuccessfulBackupTime)</td>
        <td bgcolor="$LastSuccessColor;">Succeeded</td>
      </tr>
"@
    }

    $HtmlBackup += "</table>"
    Return $HtmlBackup
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

    $HtmlReport = Format-HtmlUpdateDetails -RecentUpdates $RecentUpdates

    $UpdateStatus = [PSCustomObject] @{
        DetailedReport = $DetailedReport
        SummaryObject = $SummaryObject
        HtmlReport = $HtmlReport
    }

    Return $UpdateStatus
}

Function Format-HtmlUpdateDetails{
    Param(
        [Array]$RecentUpdates = @()
    )

    $HtmlUpdates = New-TableHeader -SectionName 'Update Details (Last 45 Days)' -ColumnNames @('Hotfix ID', 'Description', 'Installed On')
    If($RecentUpdates.Count -eq 0){
        $HtmlUpdates += @"
      <tr>
        <td colspan="3" bgcolor="$BgColorError;">No updates installed in the last 45 days</td>
      </tr>
"@
    }
    Else{
        ForEach($Update in $RecentUpdates){
            $HtmlUpdates += @"
      <tr>
        <td>$($Update.HotfixId)</td>
        <td>$($Update.Description)</td>
        <td>$($Update.InstalledOn)</td>
      </tr>
"@
        }
    }

    $HtmlUpdates += "</table>"
    Return $HtmlUpdates
}

Function New-TableHeader{
    Param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$SectionName,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String[]]$ColumnNames
    )

    $TableHeader = @"
    <table style="width: 33%" style="border-collapse: collapse; border: 1 px solid #000000;">
      <tr>
        <td colspan="$($ColumnNames.Count)" bgcolor="$BgColorSectionHeader" style="color: #FFFFFF; font-size: large; height: 35px;">
          $SectionName
        </td>
      </tr>
      <tr>
"@

    ForEach($ColumnName in $ColumnNames){
        $TableHeader += "    <td style=`"text-align: center;`"><b>$ColumnName</b></td>"
    }
    
    $TableHeader += "  </tr>"
    Return $TableHeader
}
Function Format-HtmlSummary{
    Param(
        [System.Collections.Generic.List[Object]]$SummaryObjects
    )

    $HtmlSummary = New-TableHeader -SectionName 'Summary' -ColumnNames @('Category', 'Healthy')

    ForEach($SummaryObject in $SummaryObjects){
        If($SummaryObject.Healthy){
            $HealthText = 'Yes'
            $HealthColor = "$BgColorHealthy"
        }
        Else{
            $HealthText = 'No'
            $HealthColor = "$BgColorError"
        }
        $HtmlSummary += @"
  <tr style="border-bottom-style: solid; border-bottom-width: 1px; padding-bottom: 1px">
    <td>$($SummaryObject.Category)</td>
    <td bgcolor="$HealthColor;">$HealthText</td>
  </tr>
"@
    }

    $HtmlSummary += "</table>"
    Return $HtmlSummary
}

Clear-Host
Write-Host "$(Get-Date) - Beginning script run."
$Summary = New-Object -TypeName System.Collections.Generic.List[Object]
$HtmlDetails = New-Object -TypeName System.Collections.Generic.List[Object]

#Set global color options for HTML report
New-Variable -Name BgColorHealthy -Scope Script -Value "#99FF99" -Option ReadOnly
New-Variable -Name BgColorWarning -Scope Script -Value "#FFFF99" -Option ReadOnly
New-Variable -Name BgColorError -Scope Script -Value "#FF6666" -Option ReadOnly
New-Variable -Name BgColorSectionHeader -Scope Script -Value "#0000AA" -Option ReadOnly

#Import settings and email credentials
$Settings = Get-ScriptSettings
$Credentials = Import-CliXml -Path $Settings.EmailCredentialsFilePath -ErrorAction Continue
If($Credentials.UserName -notlike "*@*"){
    Throw "$($Settings.EmailCredentialsFilePath) does not appear to contain valid credentials, or the credentials were encrypted by a different user account."
}

#Collect report data
Write-Host "Getting storage status." -ForegroundColor Cyan
$StorageUsage = Measure-StorageUsage
$Summary.Add($StorageUsage.SummaryObject)
$HtmlDetails.Add($StorageUsage.HtmlReport)

Write-Host "Checking for missing services." -ForegroundColor Cyan
$ServicesStatus = Get-MissingServices -Settings $Settings
$Summary.Add($ServicesStatus.SummaryObject)
$HtmlDetails.Add($ServicesStatus.HtmlReport)

If($Settings.IncludeBackups){
    Write-Host "Checking backup status." -ForegroundColor Cyan
    Try{
        $BackupResult = Get-BackupResult
        $Summary.Add($BackupResult.SummaryObject)
        $HtmlDetails.Add($BackupResult.HtmlReport)
    }
    Catch{
        $BackupResult = 'Failed to retrieve backup summary. Check that the WindowsServerBackup PowerShell module is installed and that a backup schedule is configured.'
    }
}
Else{
    Write-Host "Skipping backup status check." -ForegroundColor Cyan
    $BackupResult = 'Not included'
}

Write-Host "Checking update status." -ForegroundColor Cyan
$UpdateStatus = Get-UpdateStatus
$Summary.Add($UpdateStatus.SummaryObject)
$HtmlDetails.Add($UpdateStatus.HtmlReport)

#Generate the email report
Write-Host "Generating email report." -ForegroundColor Cyan
$HtmlSummary = Format-HtmlSummary -SummaryObjects $Summary
$EmailBody = $HtmlSummary
ForEach($Item in $HtmlDetails){
    $EmailBody += "<br>"
    $EmailBody += "<br>"
    $EmailBody += $Item
}

If($Credentials){
    Write-Host "Sending email report." -ForegroundColor Cyan
    Send-EmailReport -Settings $Settings -Credentials $Credentials -ReportData $EmailBody
}
Else{
    $EmailBody
}

Write-Host "$(Get-Date) - Script execution complete." -ForegroundColor Green