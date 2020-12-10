# SimpleServerStatus
PowerShell script that sends an email report on the status of the server.

## Getting Started
The script requires a configuration file and an encrypted email credentials file. To create these,
run the setup.ps1 script. Note that you must run this with the same account that will be running
the status report script itself, otherwise it will not be able to read the encrypted credentials.

Once you have created these files, use the Windows Task Scheduler to schedule the main script to
run whenever you would like. For the action, use the following:
powershell.exe -C "{path to the folder here}\SimpleServerStatus.ps1"

Example:
powershell.exe -C "C:\Scripts\SimpleServerStatus.ps1"