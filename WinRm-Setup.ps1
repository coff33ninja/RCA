# Function to check if the script is running with admin rights
function Test-AdminRights {
    $currentUser = New-Object Security.Principal.WindowsPrincipal ([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to restart the script with admin privileges if not already running as admin
function Confirm-AdminPrivileges {
    if (-not (Test-AdminRights)) {
        Write-Output 'This script needs to be run with administrative privileges.'
        Write-Output 'Relaunching the script with admin rights...'
        $scriptPath = $MyInvocation.MyCommand.Definition
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
        exit
    }
}

# Call function to ensure the script runs with admin privileges
Ensure-AdminPrivileges

# Function to clear the screen
function ClearScreen {
    Clear-Host
}

# Function to temporarily disable VM adapters
function DisableVMAdapters {
    try {
        $vmAdapters = Get-NetAdapter | Where-Object { $_.InterfaceDescription -match 'Virtual|Hyper-V|VMware' }
        foreach ($adapter in $vmAdapters) {
            Disable-NetAdapter -Name $adapter.Name -Confirm:$false
            Write-Output "Disabled adapter $($adapter.Name)"
        }
    }
    catch {
        Write-Error "Error disabling VM adapters: $_"
    }
}

# Function to enable VM adapters
function EnableVMAdapters {
    try {
        $vmAdapters = Get-NetAdapter | Where-Object { $_.InterfaceDescription -match 'Virtual|Hyper-V|VMware' }
        foreach ($adapter in $vmAdapters) {
            Enable-NetAdapter -Name $adapter.Name -Confirm:$false
            Write-Output "Enabled adapter $($adapter.Name)"
        }
    }
    catch {
        Write-Error "Error enabling VM adapters: $_"
    }
}

# Function to check if the current user has a password
function HasPassword {
    param (
        [string]$username
    )

    try {
        $user = Get-LocalUser -Name $username
        return -not $user.PasswordRequired
    }
    catch {
        Write-Error "Error checking password status for ${username}: $_"
        return $false
    }
}

# Function to collect WinRM details
function CollectWinRMDetails {
    try {
        $username = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $computername = $env:COMPUTERNAME
        $domain = $env:USERDOMAIN
        $ipAddresses = (Get-NetIPAddress | Where-Object { $_.AddressFamily -eq 'IPv4' }).IPAddress
        $macAddresses = Get-NetAdapter | Select-Object -ExpandProperty MacAddress

        $passwordStatus = if (HasPassword($username.Split('\')[-1])) { 'has a password set' } else { 'does not have a password set' }

        return @{
            Username       = $username
            Domain         = $domain
            ComputerName   = $computername
            IPAddresses    = $ipAddresses
            MACAddresses   = $macAddresses
            PasswordStatus = "$username $passwordStatus"
        }
    }
    catch {
        Write-Error "Error collecting WinRM details: $_"
        return $null
    }
}

# Function to print WinRM details
function PrintWinRMDetails {
    param (
        [hashtable]$details,
        [string]$filePath
    )

    if ($details) {
        $content = @"
WinRM Details:
Username: $($details.Username)
Domain: $($details.Domain)
Computer Name: $($details.ComputerName)
IP Addresses: $($details.IPAddresses -join ', ')
MAC Addresses: $($details.MACAddresses -join ', ')
Password Status: $($details.PasswordStatus)
"@

        Write-Output $content

        if ($filePath) {
            try {
                $content | Out-File -FilePath $filePath
                Write-Output "`nDetails have been saved to $filePath. Keep this information safe!"
            }
            catch {
                Write-Error "Error saving details to file: $_"
            }
        }
    }
    else {
        Write-Error 'No details to print.'
    }
}

# Function to add a trusted device
function AddTrustedDevice {
    param (
        [string]$trustedDevice
    )

    try {
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value $trustedDevice -Concatenate
        Write-Output "Added $trustedDevice to the list of trusted devices."
    }
    catch {
        Write-Error "Error adding trusted device ${trustedDevice}: $_"
    }
}

# Function to setup WinRM
function SetupWinRM {
    try {
        # Enable WinRM and configure firewall rules
        Enable-PSRemoting -SkipNetworkProfileCheck -Force

        # Set all network adapters to private
        Get-NetConnectionProfile | Set-NetConnectionProfile -NetworkCategory Private

        $username = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.Split('\')[-1]
        if (-not (HasPassword $username)) {
            # Allow WinRM to authenticate without requiring a password
            Set-Item WSMan:\localhost\Service\AllowUnencrypted -Value $true
            Write-Output 'WARNING: Passwordless authentication is enabled. This is a security risk. Ensure your network is secure.'
        }

        # Temporarily disable VM adapters
        DisableVMAdapters

        # Optional: Restart WinRM service to apply changes
        Restart-Service WinRM

        # Optionally, ask the user to add a trusted device
        $trustedDevice = Read-Host 'Enter the IP address or hostname of a trusted device (or press Enter to skip)'
        if ($trustedDevice) {
            AddTrustedDevice -trustedDevice $trustedDevice
        }

        # Wait for user input to re-enable VM adapters
        Write-Output "`nPress Enter to enable VM adapters and exit script."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

        # Enable VM adapters before exiting
        EnableVMAdapters
    }
    catch {
        Write-Error "Error during WinRM setup: $_"
    }
}

# Function to prompt user for action after script completion
function PromptUserForAction {
    do {
        Write-Output "`nChoose an action:"
        Write-Output '1. Exit script'
        Write-Output '2. Go back to previous menu'
        $actionChoice = Read-Host 'Enter your choice (1 or 2)'
    } while ($actionChoice -ne '1' -and $actionChoice -ne '2')

    return $actionChoice
}

# Main script logic
function Main {
    do {
        ClearScreen
        Write-Output 'Choose an option:'
        Write-Output '1. Setup WinRM'
        Write-Output '2. Preview WinRM details'
        Write-Output '3. Both setup WinRM and preview details'
        $choice = Read-Host 'Enter your choice (1, 2, or 3)'

        if ($choice -eq '1' -or $choice -eq '3') {
            ClearScreen
            SetupWinRM
            if ($?) {
                Write-Output 'WinRM setup completed successfully.'
            }
            else {
                Write-Output 'Error occurred during WinRM setup.'
            }
            $null = Read-Host 'Press Enter to continue...'
        }

        if ($choice -eq '2' -or $choice -eq '3') {
            ClearScreen
            try {
                # Collect and print WinRM details
                $details = CollectWinRMDetails

                # Output details to the console
                PrintWinRMDetails -details $details -filePath $null

                # Ask user if they want to save details to a text file
                $saveToFile = Read-Host 'Would you like to save these details to a text file on your desktop? (yes/no)'
                if ($saveToFile -eq 'yes') {
                    $desktopPath = [System.Environment]::GetFolderPath('Desktop')
                    $filePath = Join-Path -Path $desktopPath -ChildPath 'WinRM_Details.txt'
                    PrintWinRMDetails -details $details -filePath $filePath
                }
                $null = Read-Host 'Press Enter to continue...'
            }
            catch {
                Write-Error "Error during details collection or printing: $_"
                $null = Read-Host 'Press Enter to continue...'
            }
        }

        $actionChoice = PromptUserForAction
    } while ($actionChoice -eq '2')

    Write-Output 'Exiting script.'
}

# Start the main script
Main
