<#
.SYNOPSIS
    ProLawDocumentValidator - A script to validate document events in ProLaw SQL Database and check corresponding file existence in the file system.

.DESCRIPTION
    This script performs the following operations:
    1. Validates the existence of the SqlServer PowerShell module and installs it if necessary.
    2. Connects to a SQL Server instance and a specific database.
    3. Fetches a root directory path from the SQL database (if available).
    4. Executes a SQL query to gather document events.
    5. Checks for each document event whether the corresponding file exists in the file system.
    6. Exports the result to a CSV file, including an 'EXISTS' column to indicate file existence.

.PARAMETERS
    - None, all inputs are interactive.

.EXAMPLE
    PS> .\ProLawDocumentValidator.ps1

.INPUTS
    - SQL Server Instance
    - SQL Server Database
    - Authentication choice
    - User ID and Password (if not using Windows Authentication)

.OUTPUTS
    - CSV file containing the document events and an 'EXISTS' column.
    - Log file with information on script execution and errors.

.NOTES
    - This script assumes that the SQL Server PowerShell module is available or can be installed.
    - It also assumes that the user has the necessary permissions to connect to the SQL Server and database.

.AUTHOR
    - Darrington Bevins
    - Date: October 11, 2023

#>


# Check if SqlServer module is installed
if (-Not (Get-Module -ListAvailable -Name SqlServer)) {
    $userChoice = Read-Host "The SqlServer module is not installed. Would you like to install it now? (Y/N)"
    if ($userChoice -eq 'Y') {
        Install-Module -Name SqlServer -Scope CurrentUser -Force -SkipPublisherCheck
    } else {
        Write-Host "The SqlServer module is required to run this script. Exiting."
        Exit
    }
}

# Import SqlServer module
Import-Module SqlServer

# Initialize paths and log file
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$logFile = Join-Path $scriptDir "execution.log"
if (Test-Path $logFile) {
    Remove-Item $logFile -Force
}

try {
     # Ask user for SQL Server Instance and Database
     $ServerInstance = Read-Host "Enter SQL Server Instance (leave blank for LOCALHOST)"
     if ([string]::IsNullOrWhiteSpace($ServerInstance)) {
         $ServerInstance = "LOCALHOST"
     }
     $Database = Read-Host "Enter SQL Server Database (Leave blank for PROLAW)"
     if ([string]::IsNullOrWhiteSpace($Database)) {
         $Database = "PROLAW"
     }
 
     # Ask user for authentication type
     $authChoice = Read-Host "Use Windows Authentication? (Y/N, default is Y)"
     if ($authChoice -eq 'N') {
         $UserId = Read-Host "Enter SQL Server User ID"
         $Password = Read-Host "Enter SQL Server Password" -AsSecureString
         $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
         $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
         $connectionString = "Server=$ServerInstance;Database=$Database;User Id=$UserId;Password=$PlainPassword;TrustServerCertificate=true;"
     } else {
         $connectionString = "Server=$ServerInstance;Database=$Database;Integrated Security=True;TrustServerCertificate=true;"
     }

    # Connect to SQL
    Write-Host "Connected to SQL...."
    
    # Fetch the root directory from SQL
    Write-Host "Gathering Root DOCDIR...."
    $rootDirQuery = "SELECT value FROM ProLawINI WHERE ident = 'DocDir'"
    $rootDir = (Invoke-Sqlcmd -ConnectionString $connectionString -Query $rootDirQuery).value

    # Log the root directory
    Add-Content $logFile -Value "RootDir: $rootDir"

    # Check if $rootDir is null or empty and log it
    if ([string]::IsNullOrWhiteSpace($rootDir)) {
        Add-Content $logFile -Value "Root directory from database is null or empty. Proceeding with individual DocDirs."
    } else {
        # Validate root directory
        Write-Host "Validating Root Dir...."
        if (-Not (Test-Path $rootDir)) {
            throw "Root directory $rootDir does not exist."
        }
        Write-Host "Success!"

        # Fetch all files from root directory
        Write-Host "Gathering list of root directory children..."
        $allFiles = Get-ChildItem -Path $rootDir -Recurse | Select-Object -ExpandProperty FullName
        Write-Host "Success!"
        Write-Host "There are $($allFiles.Count) files in $rootDir"
    }

    # Main SQL Query
    $query = @"
    SELECT e.events, et.eventtypes, m.Matters, c.CONTACTS
    , e.rtf, e.ShortNote, e.DocDir
    , m.matterID, M.ClientSort
    , c.Company, C.FullName
    , COALESCE(m.Matters,c.Contacts) MCPKID
    , COALESCE(m.MatterID, c.FullName, c.COMPANY) Name
    , CASE WHEN M.MATTERS is NULL then 'Contacts' ELSE 'Matters' END Entitiy
    FROM EVENTS E 
    left outer join EVENTTYPES ET on E.EVENTTYPES = ET.EVENTTYPES
    left outer join EventMatters EM on EM.Events = e.Events
    left outer join Matters M on M.Matters = EM.Matters
    left outer join EventsContacts EC on EC.Events = E.Events
    left outer join Contacts C on C.CONTACTS = EC.Contacts
    where e.EventKind = 'o'
"@

    # Execute SQL Query
    Write-Host "Gathering Document Events..."
    $sqlResults = Invoke-Sqlcmd -ConnectionString $connectionString -Query $query
    Write-Host "Success!"
    Write-Host "There are $($sqlResults.Count) Document Events in SQL"

    # Initialize ArrayList for performance
    $results = New-Object System.Collections.ArrayList

    # Log total number of events
    Add-Content $logFile -Value "Total Events: $($sqlResults.Count)"

 # Check each event for file existence
 Write-Host "working on matching each event. this may take a while..."
foreach ($row in $sqlResults) {

    $fileExists = "N"  # Default to 'N'
    
    # Only check if DocDir is not null or empty
    if (-Not [string]::IsNullOrWhiteSpace($row.DocDir)) {
        $fileExists = if (Test-Path $row.DocDir) { "Y" } else { "N" }
    }
    
    # Create a PSObject for each row with an additional property 'EXISTS'
    $outputRow = $row | Select-Object *, @{Name='EXISTS'; Expression={$fileExists}}
    
    # Add the PSObject to ArrayList
    [void]$results.Add($outputRow)
    
    # Logging
    Add-Content $logFile -Value "Event: $($row.events), File Exists: $fileExists"
}

    # Export ArrayList to CSV
    Write-Host "Exporting to CSV..."
    $csvPath = Join-Path $scriptDir "output.csv"
    $results | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "Success!"

    Add-Content $logFile -Value "Script executed successfully."
} catch {
    # Log any errors
    Add-Content $logFile -Value "Error: $_"
}