# Create a log folder and place logs inside it
New-Item -ItemType Directory -Path ".\log" -Force

# Start transcript to capture the output
Start-Transcript -Path .\log\UserImport.log

# Import the Active Directory module
Import-Module ActiveDirectory

# Get the domain controller details
$domain = (Get-WmiObject Win32_ComputerSystem).Domain
$dController = nltest /dclist:$domain
Write-Host "Domain: $domain `nDC: $dController"

# Specify the path to the CSV file
$csvPath = ".\usersDelete.csv"

# Test if the file exists at the specified path
if (Test-Path $csvPath) {
    Write-Host "File exists at $csvPath"
} else {
    Write-Host "File does not exist at $csvPath"
    exit
}

# Specify the OU path where the user is to be created - Fill in the OU path, then the domain path
$ouPath = "OU=<OU #2>,OU=<OU #1>,DC=< >,DC=< >,DC=< >"
Write-Host  "Target OU path: $ouPath"

# Test if OU path exists
try {
    $ou = Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$ouPath'"
    if ($ou) {
        Write-Host "OU exists at $ouPath"
    } else {
        Write-Host "OU does not exist at $ouPath"
        exit
    }
} catch {
    Write-Host "OU path cannot be found. An error occurred: $($_.Exception.Message)"
    exit
}

# Read user details from the CSV file
$userList = Import-Csv $csvPath

# Function to check if a user exists
function Test-UserExists {
    param (
        [string]$username
    )
    $user = Get-ADUser -Filter { SamAccountName -eq $username } -ErrorAction SilentlyContinue
    if ($user) {
        return $true
    } else {
        return $false
    }
}

# Function to log messages
function Write-Message {
    param (
        [string]$message
    )
    Add-Content -Path .\log\UserImport.log -Value $message
}

# Loop through each user in the CSV - username should be a single column
foreach ($user in $userList) {
    $username = $user.Username
    
    # Check if the user already exists
    if (Check-UserExists -username $username) {
        $message = "User exists: $username"
        Log-Message $message
        Write-Output $message
        # Delete the user account
        Remove-ADUser -Identity "$username@domain.com" -Confirm:$false

        $message = "Deleted user: $username"
        Log-Message $message
        Write-Output $message
        # Pause for 50 milliseconds before processing the next user
        Start-Sleep -Milliseconds 50
    } else {
        $message = "User does not exist: $username"
        Log-Message $message
        Write-Output $message
    }
}

# Stop transcript
Stop-Transcript
