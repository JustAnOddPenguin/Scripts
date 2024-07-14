# Define the path to the CSV file and log directory
$csvFilePath = ".\devices.csv"
$logDirectory = ".\logs"

# Create the log directory if it doesn't exist
if (-not (Test-Path -Path $logDirectory)) {
    New-Item -Path $logDirectory -ItemType Directory
}

# Get the current date and time for the log file name
$currentDateTime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = "$logDirectory\$currentDateTime Log.log"

# Log the start of the process
"Starting ping process at $currentDateTime" | Out-File -FilePath $logFilePath -Append

# Check internet access by pinging Google's public DNS server
$internetTest = Test-Connection -ComputerName 8.8.8.8 -Count 1 -ErrorAction SilentlyContinue

if ($internetTest) {
    "Internet access: Yes" | Out-File -FilePath $logFilePath -Append
} else {
    "Internet access: No" | Out-File -FilePath $logFilePath -Append
}

# Read the CSV file
try {
    $devices = Import-Csv -Path $csvFilePath
    if (-not $devices) {
        throw "CSV file is empty or not properly formatted."
    }
} catch {
    "Failed to read CSV file: $_" | Out-File -FilePath $logFilePath -Append
    exit
}

# Ping each device and log the results
foreach ($device in $devices) {
    $name = $device.Name
    $ip = $device.IP

    if (-not [string]::IsNullOrWhiteSpace($ip)) {
        $pingResult = Test-Connection -ComputerName $ip -Count 1 -ErrorAction SilentlyContinue

        if ($pingResult) {
            "$name ($ip) is online" | Out-File -FilePath $logFilePath -Append
        } else {
            "$name ($ip) is offline" | Out-File -FilePath $logFilePath -Append
        }
    } else {
        "Skipping $name as IP address is null or empty." | Out-File -FilePath $logFilePath -Append
    }
}

# Log the end of the process
"Ping process completed at $(Get-Date -Format 'yyyyMMdd_HHmmss')" | Out-File -FilePath $logFilePath -Append
