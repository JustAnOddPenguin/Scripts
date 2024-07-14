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
$csvPath = ".\usersImport.csv"

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

foreach ($user in $userList) {
    # Set user details from CSV - Enable fields required - This is column sensitive. If the .CSV is changed, it will affect the script
    $firstName = $user.FirstName            # First Name
    $lastName = $user.LastName              # Last Name
    $username = $user.Username              # UPN
    $password = ConvertTo-SecureString $user.Password -AsPlainText -Force  # Password
    $displayName = $user.DisplayName        # Display Name
    $email = $user.Email                    # Email field
    $company = $user.Company                # Company field
    $office = $user.Office                  # Office field
    $street = $user.Street                  # Street field
    $city = $user.City                      # City field
    $state = $user.State                    # State field
    $zip = $user.ZipPostCode                # ZIP field
    $countryRegion = $user.CountryRegion    # Country field
    $jobTitle = $user.JobTitle              # Job title field
    $department = $user.Department          # Department field

    # Log the details in JSON format
    $userDetails = @{
        "FirstName"     = $firstName
        "LastName"      = $lastName
        "Username"      = $username
        "Password"      = $password
        "DisplayName"   = $displayName
        "Email"         = $email
        "Company"       = $company
        "Office"        = $office
        "Street"        = $street
        "City"          = $city
        "State"         = $state
        "ZIP"           = $zip
        "Country"       = $countryRegion
        "JobTitle"      = $jobTitle
        "Department"    = $department
    }

    $userDetailsJson = $userDetails | ConvertTo-Json
    Write-Host $userDetailsJson

    # Check if the samAccountName or userPrincipalName already exists
    $existingUser = Get-ADUser -Filter {samAccountName -eq $username -or userPrincipalName -eq "$username@domain.com"}
    if ($existingUser) {
        Write-Host "A user with samAccountName $username or userPrincipalName $username@domain.com already exists. Skipping user creation.`n"
        continue
    }

    try {
        # Create the user account - REPLACE DOMAIN.COM
        New-ADUser -Name "$company - $firstName $lastName" `
            -GivenName $firstName `
            -Surname $lastName `
            -UserPrincipalName "$username@domain.com" `
            -SamAccountName $username `
            -AccountPassword $password `
            -DisplayName $displayName `
            -Company $company `
            -Office $office `
            -StreetAddress $street `
            -City $city `
            -State $state `
            -PostalCode $zip `
            -Country $countryRegion `
            -Title $jobTitle `
            -Department $department `
            -EmailAddress $email `
            -Path $ouPath `
            -Enabled $true `
            -PasswordNeverExpires $false `
            -ChangePasswordAtLogon $false

        Write-Host "User $username created successfully in $ouPath.`n"
    } catch {
        Write-Host "An error occurred while creating user '$username': $($_.Exception.Message)`n"
    }
}

# Stop transcript
Stop-Transcript
