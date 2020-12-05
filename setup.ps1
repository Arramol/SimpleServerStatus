Function Get-PreviousSettings{
    $ImportSuccess = $true
    $CredentialsFound = $false
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
            $Settings.PSObject.Properties[$Property] = $PreviousSettings.PSObject.Properties[$Property]
        }
        If(Test-Path -Path $Settings.EmailCredentialsFilePath){
            Write-Verbose "Attempting to load email credentials from credentials file $($Settings.EmailCredentialsFilePath)."
            Try{
                [pscredential]$EmailCredentials = Import-Clixml -Path $Settings.EmailCredentialsFilePath -ErrorAction Stop
                $CredentialsFound = $true
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
}

[String]$EmailFrom = ''
[String[]]$EmailTo = @()
[String]$SmtpServer = ''
$UseSSL = $true
$PortSpecified = $false
