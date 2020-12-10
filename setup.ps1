Function Get-PreviousSettings{
    $ImportSuccess = $true
    $PropertiesToValidate = @('EmailCredentialsFilePath', 'EmailTo', 'SmtpServer', 'SmtpPort', 'UseSsl', 'IncludeBackups')
    $Settings = [PSCustomObject] @{
        EmailCredentialsFilePath = ''
        EmailTo = New-Object -TypeName System.Collections.Generic.List[String]
        SmtpServer = ''
        SmtpPort = 587
        UseSsl = $true
        IncludeBackups = $false
        IgnoredServicesListPath = ''
    }
    Write-Verbose "Checking for existing settings file at $PSScriptRoot\config.xml"
    If(Test-Path -Path "$PSScriptRoot\config.xml"){
        Try{
            Write-Verbose "Settings file found. Loading settings."
            $PreviousSettings = Import-Clixml -Path "$PSScriptRoot\config.xml" -ErrorAction Stop
        }
        Catch{
            Write-Error "File exists at $PSScriptRoot\config.xml but it could not be read."
            $ImportSuccess = $false
        }
        If($PreviousSettings){
            Write-Verbose "Validating imported settings."
            ForEach($Property in $PropertiesToValidate){
                If($PreviousSettings.PSObject.Properties.Name -notcontains $Property){
                    Write-Error "Failed to validate imported settings. Property $Property was not found in $PSScriptRoot\config.xml."
                    $ImportSuccess = $false
                    Break
                }
                Else{
                    Write-Verbose "Successfully validated imported settings."
                }
            }
        }
        Else{
            Write-Error "Failed to read file $PSScriptRoot\config.xml."
            $ImportSuccess = $false
        }
    }
    Else{
        Write-Verbose "No existing settings file was found."
        $ImportSuccess = $false
    }

    If($ImportSuccess){
        ForEach($Property in $PreviousSettings.PSObject.Properties.Name){
            Write-Verbose "Assigning value of property $Property"
            $Settings.$Property = $PreviousSettings.$Property
        }
        If(Test-Path -Path $Settings.EmailCredentialsFilePath){
            Write-Verbose "Attempting to load email credentials from credentials file $($Settings.EmailCredentialsFilePath)."
            Try{
                [pscredential]$EmailCredentials = Import-Clixml -Path $Settings.EmailCredentialsFilePath -ErrorAction Stop
                Write-Verbose "Successfully loaded email credentials with username $($EmailCredentials.UserName)."
            }
            Catch{
                Write-Error "Failed to read email credentials from $($Settings.EmailCredentialsFilePath). They may have been saved with a different user account."
            }
        }
        Else{
            Write-Verbose "No email credentials file was found."
        }
    }

    Return $Settings
}

Function Format-CommaList{
    Param(
        [Array]$ListItems = @()
    )

    $FormattedList = ''
    ForEach($Item in $ListItems){
        If($FormattedList.Length -gt 0){
            $FormattedList += ', '
        }
        $FormattedList += $Item
    }

    Return $FormattedList
}

Function Write-Menu{
    Param(
        $Settings
    )

    $EmailRecipientList = Format-CommaList -ListItems $Settings.EmailTo

    $ValidChoices = New-Object -TypeName System.Collections.Generic.List[Object]
    @('E', 'S', 'X') | ForEach-Object {$ValidChoices.Add($_)}
    1..7  | ForEach-Object {$ValidChoices.Add($_)}
    $Choice = ''

    Clear-Host
    Write-Host "Current Settings:" -ForegroundColor Magenta
    Write-Host "(1) Email Credentials File Path: $($Settings.EmailCredentialsFilePath)" -ForegroundColor Cyan
    Write-Host "(2) Email Recipient(s): $EmailRecipientList" -ForegroundColor Cyan
    Write-Host "(3) SMTP Server: $($Settings.SmtpServer)" -ForegroundColor Cyan
    Write-Host "(4) SMTP Port: $($Settings.SmtpPort)" -ForegroundColor Cyan
    Write-Host "(5) Use SSL for SMTP Connection (recommended): $($Settings.UseSsl)" -ForegroundColor Cyan
    Write-Host "(6) Include Windows Backup status: $($Settings.IncludeBackups)" -ForegroundColor Cyan
    Write-Host "(7) Excluded Services List File Path: $($Settings.IgnoredServicesListPath)" -ForegroundColor Cyan
    Write-Host "(E) Create email credentials file" -ForegroundColor Cyan
    Write-Host "(S) Save configuration and exit" -ForegroundColor Green
    Write-Host "(X) Exit without making changes" -ForegroundColor Yellow
    Do{
        $Choice = Read-Host -Prompt "Please enter the number of a setting to configure, or the letter of another action"
    }While($ValidChoices -notcontains $Choice)

    Return $Choice
}

Function Get-NewCredsFilePath{
    Param(
        $Settings
    )

    Clear-Host
    Write-Host "Current credentials file path is: $($Settings.EmailCredentialsFilePath)" -ForegroundColor Magenta
    $Choice = Read-Host -Prompt "Please enter the new credentials file path, or C to cancel."
    If($Choice -ne 'C'){
        $Settings.EmailCredentialsFilePath = $Choice
    }

    Return $Settings
}

Function Get-NewRecipients{
    Param(
        $Settings
    )

    Clear-Host
    $Recipients = Format-CommaList -ListItems $Settings.EmailTo
    Write-Host "Current recipient email addresses are: $Recipients" -ForegroundColor Magenta
    $Choice = Read-Host -Prompt "Please enter a new comma-separated list of email addresses, or C to cancel."
    If($Choice -ne 'C'){
        $Choice.Replace(' ', '')
        $Settings.EmailTo = $Choice.Split(',')
    }

    Return $Settings
}

Function Get-NewSmtpServer{
    Param(
        $Settings
    )

    Clear-Host
    Write-Host "Current SMTP server is: $($Settings.SmtpServer)" -ForegroundColor Magenta
    $Choice = Read-Host -Prompt "Please enter a new SMTP server address (ex: smtp.gmail.com), or C to cancel."
    If($Choice -ne 'C'){
        $Settings.SmtpServer = $Choice
    }

    Return $Settings
}

Function Get-NewSmtpPort{
    Param(
        $Settings
    )

    Clear-Host
    Write-Host "Current SMTP port is: $($Settings.SmtpPort)" -ForegroundColor Magenta
    [Int]$NewPort = 0
    Do{
        $Choice = Read-Host -Prompt "Please enter a new SMTP port, or C to cancel. Port must be a whole number from 1-65535. Most secure email providers use port 587, although some may use 25 or another"
        If($Choice -ne 'C'){
            Try{
                [Int]$NewPort = [Int]$Choice
                If($NewPort -eq 0 -or $NewPort -gt 65535){
                    Write-Host "Port must be a whole number from 1-65535. Please enter a valid port number, or C to cancel." -ForegroundColor Red
                }
            }
            Catch{
                Write-Host "Port must be a whole number from 1-65535. Please enter a valid port number, or C to cancel." -ForegroundColor Red
            }
        }
    }While($Choice -ne 'C' -and !($NewPort -gt 0 -and $NewPort -le 65535))
    If($Choice -ne 'C'){
        $Settings.SmtpServer = $NewPort
    }

    Return $Settings
}

Function Get-NewSslSetting{
    Param(
        $Settings
    )

    Clear-Host
    If($Settings.UseSsl){
        Write-Host "SSL encryption for SMTP connection is currently enabled. It is recommended that you leave it enabled." -ForegroundColor Magenta
        $Prompt = "Leave SSL enabled? (y/n)"
    }
    Else{
        Write-Host "SSL encryption for SMTP connection is currently disabled. It is recommended that enable it." -ForegroundColor Magenta
        $Prompt = "Enable SSL? (y/n)"
    }
    Do{
        $Choice = Read-Host -Prompt $Prompt
        If($Choice -ne 'y' -and $Choice -ne 'n'){
            Write-Host "Please enter y or n." -ForegroundColor Red
        }
    }While($Choice -ne 'y' -and $Choice -ne 'n')
    
    If($Choice -eq 'y'){
        $Settings.UseSsl = $true
    }
    Else{
        $Settings.UseSsl = $false
    }

    Return $Settings
}

Function Get-NewBackupSetting{
    Param(
        $Settings
    )

    Clear-Host
    If($Settings.IncludeBackups){
        Write-Host "Windows Backup status is currently included in the report." -ForegroundColor Magenta
        $Prompt = "Continue including? (y/n)"
    }
    Else{
        Write-Host "Windows Backup status is not currently included in the report." -ForegroundColor Magenta
        $Prompt = "Add backup status to the report? (y/n)"
    }
    Do{
        $Choice = Read-Host -Prompt $Prompt
        If($Choice -ne 'y' -and $Choice -ne 'n'){
            Write-Host "Please enter y or n." -ForegroundColor Red
        }
    }While($Choice -ne 'y' -and $Choice -ne 'n')
    
    If($Choice -eq 'y'){
        $Settings.IncludeBackups = $true
    }
    Else{
        $Settings.IncludeBackups = $false
    }

    Return $Settings
}

Function Get-NewServiceListPath{
    Param(
        $Settings
    )

    Clear-Host
    Write-Host "Current path to list of services to ignore is: $($Settings.IgnoredServicesListPath)" -ForegroundColor Magenta
    $Choice = Read-Host -Prompt "Please enter a new file path, or C to cancel."
    If($Choice -ne 'C'){
        $Settings.IgnoredServicesListPath = $Choice
    }

    Return $Settings
}

Function New-EmailCredsFile{
    Param(
        $Settings
    )

    Clear-Host
    $SpecifyNewPath = $false
    $NewPathMandatory = $false
    Write-Host "Please enter the desired email address and password." -ForegroundColor Magenta
    $Credentials = Get-Credential
    If($Settings.EmailCredentialsFilePath){
        Write-Host "Current credentials file path is $($Settings.EmailCredentialsFilePath)"
        Do{
            $Choice = Read-Host -Prompt 'Use this path? (y/n)'
            If($Choice -ne 'y' -and $Choice -ne 'n'){
                Write-Host "Please enter y or n." -ForegroundColor Red
            }
        }While($Choice -ne 'y' -and $Choice -ne 'n')
        If($Choice -eq 'n'){
            $SpecifyNewPath = $true
        }
    }
    Else{
        $SpecifyNewPath = $true
        $NewPathMandatory = $true
    }

    If($SpecifyNewPath){
        $Instructions = "Please enter the new credentials file path"
        If($NewPathMandatory){
            $Instructions += '.'
        }
        Else{
            $Instructions += ", or C to cancel and use the previous path of $($Settings.EmailCredentialsFilePath)."
        }

        Do{
            $Finished = $false
            $PathChoice = Read-Host -Prompt $Instructions
            If($PathChoice -ne 'C' -or $NewPathMandatory){
                $Finished = Test-Path -Path $PathChoice -IsValid
                If(-not $Finished){
                    Write-Host "$PathChoice does not appear to be a valid path." -ForegroundColor Red
                }
            }
            Else{
                $Finished = $true
            }
        }While(-not $Finished)

        If($PathChoice -ne 'C'){
            $Settings.EmailCredentialsFilePath = $PathChoice
        }
    }

    $Credentials | Export-Clixml -Path $Settings.EmailCredentialsFilePath
    Return $Settings
}

Function Write-SettingsFile{
    Param(
        $Settings
    )

    Clear-Host
    $ConfigPath = "$PSScriptRoot\config.xml"
    $Done = $false
    Do{
        Write-Host "Target config file path is $ConfigPath." -ForegroundColor Magenta
        Write-Host "(1) Specify a new path" -ForegroundColor Cyan
        Write-Host "(2) Save to this path and exit this script" -ForegroundColor Cyan
        Write-Host "(3) Cancel and return to menu" -ForegroundColor Cyan
        Do{
            $Choice = Read-Host -Prompt "Please input the number of your selection"
        }While($Choice -ne '1' -and $Choice -ne '2' -and $Choice -ne '3')

        Switch($Choice){
            '1' {
                $ConfigPath = Read-Host -Prompt "Please enter the new path to save the config file"
                Clear-Host
            }
            '2' {
                $WriteError = $null
                $Settings | Export-Clixml -Path $ConfigPath -ErrorVariable WriteError
                If($WriteError){
                    Clear-Host
                    Write-Host "Failed to write to path $ConfigPath with error $WriteError. Please choose another option or path." -ForegroundColor Red
                }
                Else{
                    Write-Host "Config file was successfully written to path $ConfigPath" -ForegroundColor Green
                    $Done = $true
                }
            }
            '3' {$Done = $true}
        }
    }While(-not $Done)

    If($Choice -eq '2'){
        $ReturnValue = 'X'
    }
    Else{
        $ReturnValue = 'S'
    }

    Return $ReturnValue
}

Clear-Host
$Settings = Get-PreviousSettings
Do{
    $Choice = Write-Menu -Settings $Settings
    Switch($Choice){
        '1' {$Settings = Get-NewCredsFilePath -Settings $Settings}
        '2' {$Settings = Get-NewRecipients -Settings $Settings}
        '3' {$Settings = Get-NewSmtpServer -Settings $Settings}
        '4' {$Settings = Get-NewSmtpPort -Settings $Settings}
        '5' {$Settings = Get-NewSslSetting -Settings $Settings}
        '6' {$Settings = Get-NewBackupSetting -Settings $Settings}
        '7' {$Settings = Get-NewServiceListPath -Settings $Settings}
        'E' {$Settings = New-EmailCredsFile -Settings $Settings}
        'S' {$Choice = Write-SettingsFile -Settings $Settings}
    }
}While($Choice -ne 'X')