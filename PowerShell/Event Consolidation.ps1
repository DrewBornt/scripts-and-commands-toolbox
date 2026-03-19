<#
.SYNOPSIS
    Converts Windows Event Log (.evtx) files to Excel workbook with full message details.

.DESCRIPTION
    This script allows users to select either a single .evtx file or a folder (recursive search)
    and exports all event log entries with full messages to an Excel workbook
    with each .evtx file as a separate worksheet tab with formatted Excel tables.
    Empty .evtx files are automatically skipped.

.NOTES
   
    Auto-installs ImportExcel module if not present
#>

# Add necessary assemblies for folder browser dialog
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to check, install, and import ImportExcel module
function Initialize-ImportExcelModule {
    Write-Host "`n=== Checking ImportExcel Module ===" -ForegroundColor Cyan
   
    # Check if module is already installed
    $module = Get-Module -ListAvailable -Name ImportExcel
   
    if ($null -eq $module) {
        Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow
       
        try {
            # Check if running as administrator for AllUsers scope
            $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
           
            if ($isAdmin) {
                Install-Module -Name ImportExcel -Scope AllUsers -Force -AllowClobber -ErrorAction Stop
                Write-Host "ImportExcel module installed successfully (AllUsers scope)." -ForegroundColor Green
            }
            else {
                Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-Host "ImportExcel module installed successfully (CurrentUser scope)." -ForegroundColor Green
            }
        }
        catch {
            Write-Host "ERROR: Failed to install ImportExcel module" -ForegroundColor Red
            Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "`nPlease run PowerShell as Administrator or manually install with:" -ForegroundColor Yellow
            Write-Host "Install-Module -Name ImportExcel -Scope CurrentUser -Force" -ForegroundColor Yellow
            throw "Module installation failed"
        }
    }
    else {
        Write-Host "ImportExcel module found." -ForegroundColor Green
    }
   
    # Import the module
    try {
        Import-Module ImportExcel -ErrorAction Stop
        Write-Host "ImportExcel module imported successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to import ImportExcel module" -ForegroundColor Red
        Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
        throw "Module import failed"
    }
}

# Function to prompt user for file or folder selection
function Get-SelectionChoice {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Select Input Type'
    $form.Size = New-Object System.Drawing.Size(400, 200)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
   
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(350, 40)
    $label.Text = 'Would you like to select a single .evtx file or a folder (with recursive search)?'
    $form.Controls.Add($label)
   
    $btnFile = New-Object System.Windows.Forms.Button
    $btnFile.Location = New-Object System.Drawing.Point(50, 80)
    $btnFile.Size = New-Object System.Drawing.Size(120, 40)
    $btnFile.Text = 'Single File'
    $btnFile.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $form.Controls.Add($btnFile)
   
    $btnFolder = New-Object System.Windows.Forms.Button
    $btnFolder.Location = New-Object System.Drawing.Point(220, 80)
    $btnFolder.Size = New-Object System.Drawing.Size(120, 40)
    $btnFolder.Text = 'Folder (Recursive)'
    $btnFolder.DialogResult = [System.Windows.Forms.DialogResult]::No
    $form.Controls.Add($btnFolder)
   
    $form.AcceptButton = $btnFolder
   
    $result = $form.ShowDialog()
   
    return $result
}

# Function to select a single file
function Select-File {
    param (
        [string]$Title = "Select an .evtx log file"
    )
   
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = $Title
    $openFileDialog.Filter = "Event Log Files (*.evtx)|*.evtx|All Files (*.*)|*.*"
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
   
    $result = $openFileDialog.ShowDialog()
   
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName
    }
    else {
        Write-Host "No file selected. Exiting script." -ForegroundColor Yellow
        exit
    }
}

# Function to select folder using File Explorer
function Select-Folder {
    param (
        [string]$Description = "Select the folder containing .evtx log files"
    )
   
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $Description
    $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::MyComputer
    $folderBrowser.ShowNewFolderButton = $false
   
    $result = $folderBrowser.ShowDialog()
   
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    }
    else {
        Write-Host "No folder selected. Exiting script." -ForegroundColor Yellow
        exit
    }
}

# Function to sanitize worksheet names (Excel has restrictions)
function Get-SafeWorksheetName {
    param (
        [string]$Name,
        [int]$MaxLength = 31
    )
   
    # Remove invalid characters for Excel worksheet names: \ / ? * [ ]
    $safeName = $Name -replace '[\\\/\?\*\[\]:]', '_'
   
    # Trim to max length
    if ($safeName.Length -gt $MaxLength) {
        $safeName = $safeName.Substring(0, $MaxLength)
    }
   
    return $safeName
}

# Function to sanitize table names (Excel has different restrictions)
function Get-SafeTableName {
    param (
        [string]$Name
    )
   
    # Excel table names:
    # - Must start with a letter or underscore
    # - Can only contain letters, numbers, and underscores
    # - Max 255 characters
    # - Cannot contain spaces
   
    # Replace invalid characters with underscores
    $safeName = $Name -replace '[^a-zA-Z0-9_]', '_'
   
    # Ensure it starts with a letter or underscore
    if ($safeName -match '^[0-9]') {
        $safeName = "T_$safeName"
    }
   
    # Trim to max length (255, but we'll use 100 for practicality)
    if ($safeName.Length -gt 100) {
        $safeName = $safeName.Substring(0, 100)
    }
   
    return $safeName
}

# Function to convert .evtx files to Excel with separate tabs and tables
function Convert-EvtxToExcel {
    param (
        [System.IO.FileInfo[]]$EvtxFiles,
        [string]$OutputPath
    )
   
    Write-Host "`n=== Event Log to Excel Converter ===" -ForegroundColor Cyan
    Write-Host "Processing $($EvtxFiles.Count) .evtx file(s)" -ForegroundColor Green
   
    # Track worksheet and table names to ensure uniqueness
    $worksheetNames = @{}
    $tableNames = @{}
    $worksheetCounter = 1
    $tableCounter = 1
    $processedCount = 0
    $skippedCount = 0
   
    # Process each .evtx file
    foreach ($file in $EvtxFiles) {
        Write-Host "`nProcessing: $($file.FullName)" -ForegroundColor Yellow
       
        try {
            # Get events from the .evtx file
            $events = Get-WinEvent -Path $file.FullName -ErrorAction Stop
           
            Write-Host "  Found $($events.Count) event(s)" -ForegroundColor Cyan
           
            # Skip files with no events
            if ($events.Count -eq 0) {
                Write-Host "  No events in this file. Skipping..." -ForegroundColor Yellow
                $skippedCount++
                continue
            }
           
            # Create a unique worksheet name
            $baseWorksheetName = Get-SafeWorksheetName -Name $file.BaseName
            $worksheetName = $baseWorksheetName
           
            # Ensure unique worksheet name
            if ($worksheetNames.ContainsKey($worksheetName)) {
                $worksheetName = Get-SafeWorksheetName -Name "$baseWorksheetName`_$worksheetCounter"
                $worksheetCounter++
            }
            $worksheetNames[$worksheetName] = $true
           
            # Create a unique table name
            $baseTableName = Get-SafeTableName -Name "Table_$($file.BaseName)"
            $tableName = $baseTableName
           
            # Ensure unique table name
            if ($tableNames.ContainsKey($tableName)) {
                $tableName = Get-SafeTableName -Name "$baseTableName`_$tableCounter"
                $tableCounter++
            }
            $tableNames[$tableName] = $true
           
            Write-Host "  Worksheet name: $worksheetName" -ForegroundColor Cyan
            Write-Host "  Table name: $tableName" -ForegroundColor Cyan
           
            # Collection to hold events for this file
            $fileEvents = @()
           
            # Extract detailed information from each event
            foreach ($event in $events) {
                $eventObject = [PSCustomObject]@{
                    'TimeCreated'        = $event.TimeCreated
                    'Id'                 = $event.Id
                    'Level'              = $event.LevelDisplayName
                    'LogName'            = $event.LogName
                    'ProviderName'       = $event.ProviderName
                    'Message'            = $event.Message
                    'MachineName'        = $event.MachineName
                    'UserId'             = $event.UserId
                    'ProcessId'          = $event.ProcessId
                    'ThreadId'           = $event.ThreadId
                    'TaskDisplayName'    = $event.TaskDisplayName
                    'OpcodeName'         = $event.OpcodeName
                    'KeywordsDisplayNames' = ($event.KeywordsDisplayNames -join '; ')
                    'RecordId'           = $event.RecordId
                    'ActivityId'         = $event.ActivityId
                    'RelatedActivityId'  = $event.RelatedActivityId
                    'Source_File'        = $file.Name
                    'Source_Path'        = $file.DirectoryName
                }
               
                $fileEvents += $eventObject
            }
           
            # Export to Excel with this file as a separate worksheet with a table
            Write-Host "  Exporting to worksheet with table: $tableName" -ForegroundColor Green
           
            $fileEvents | Export-Excel -Path $OutputPath `
                -WorksheetName $worksheetName `
                -TableName $tableName `
                -TableStyle Medium2 `
                -AutoSize `
                -AutoFilter `
                -FreezeTopRow `
                -BoldTopRow `
                -Append
           
            Write-Host "  Successfully exported $($fileEvents.Count) events as a formatted table" -ForegroundColor Green
            $processedCount++
        }
        catch {
            Write-Host "  ERROR: Failed to process $($file.Name)" -ForegroundColor Red
            Write-Host "  Details: $($_.Exception.Message)" -ForegroundColor Red
           
            # Create an error entry worksheet for this file
            $errorWorksheetName = Get-SafeWorksheetName -Name "ERROR_$($file.BaseName)"
            if ($worksheetNames.ContainsKey($errorWorksheetName)) {
                $errorWorksheetName = Get-SafeWorksheetName -Name "ERROR_$worksheetCounter"
                $worksheetCounter++
            }
            $worksheetNames[$errorWorksheetName] = $true
           
            # Create unique error table name
            $errorTableName = Get-SafeTableName -Name "ErrorTable_$tableCounter"
            $tableCounter++
            $tableNames[$errorTableName] = $true
           
            $errorObject = [PSCustomObject]@{
                'Source_File'    = $file.Name
                'Source_Path'    = $file.DirectoryName
                'Error'          = $_.Exception.Message
                'Error_Time'     = Get-Date
            }
           
            try {
                $errorObject | Export-Excel -Path $OutputPath `
                    -WorksheetName $errorWorksheetName `
                    -TableName $errorTableName `
                    -TableStyle Light1 `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -Append
            }
            catch {
                Write-Host "  Failed to write error worksheet" -ForegroundColor Red
            }
        }
    }
   
    Write-Host "`n=== Export Complete ===" -ForegroundColor Cyan
    Write-Host "Excel workbook created: $OutputPath" -ForegroundColor Green
    Write-Host "Files processed: $processedCount" -ForegroundColor Green
    Write-Host "Files skipped (no data): $skippedCount" -ForegroundColor Yellow
    Write-Host "Total worksheets: $($worksheetNames.Count)" -ForegroundColor Green
    Write-Host "Total tables: $($tableNames.Count)" -ForegroundColor Green
}

# Main script execution
try {
    # Initialize ImportExcel module (check, install if needed, import)
    Initialize-ImportExcelModule
   
    # Ask user to choose between single file or folder
    Write-Host "`n=== Select Input Type ===" -ForegroundColor Cyan
    $choice = Get-SelectionChoice
   
    $evtxFiles = @()
   
    if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Single file selection
        Write-Host "Single file mode selected" -ForegroundColor Green
        $filePath = Select-File -Title "Select an .evtx log file"
       
        if ([string]::IsNullOrWhiteSpace($filePath)) {
            Write-Host "Invalid file path. Exiting." -ForegroundColor Red
            exit
        }
       
        $evtxFiles = @(Get-Item -Path $filePath)
        Write-Host "Selected file: $filePath" -ForegroundColor Green
    }
    else {
        # Folder selection with recursive search
        Write-Host "Folder mode selected (recursive search)" -ForegroundColor Green
        $folderPath = Select-Folder -Description "Select the folder containing .evtx log files (will search recursively)"
       
        if ([string]::IsNullOrWhiteSpace($folderPath)) {
            Write-Host "Invalid folder path. Exiting." -ForegroundColor Red
            exit
        }
       
        Write-Host "Searching for .evtx files in: $folderPath" -ForegroundColor Green
        Write-Host "Recursively scanning subdirectories..." -ForegroundColor Green
       
        # Recursively find all .evtx files
        $evtxFiles = Get-ChildItem -Path $folderPath -Filter "*.evtx" -Recurse -ErrorAction SilentlyContinue
       
        if ($evtxFiles.Count -eq 0) {
            Write-Host "`nNo .evtx files found in the selected directory or its subdirectories." -ForegroundColor Red
            Write-Host "Press any key to exit..." -ForegroundColor Gray
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit
        }
       
        Write-Host "Found $($evtxFiles.Count) .evtx file(s)" -ForegroundColor Green
    }
   
    # Generate output file path
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outputFileName = "EventLogs_Export_$timestamp.xlsx"
   
    # Ask user where to save the output
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    $saveFileDialog.Title = "Save Excel File As"
    $saveFileDialog.FileName = $outputFileName
    $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
   
    $result = $saveFileDialog.ShowDialog()
   
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $outputPath = $saveFileDialog.FileName
    }
    else {
        # Default to Desktop if user cancels
        $outputPath = Join-Path ([Environment]::GetFolderPath("Desktop")) $outputFileName
        Write-Host "Using default output path: $outputPath" -ForegroundColor Yellow
    }
   
    # Convert .evtx files to Excel
    Convert-EvtxToExcel -EvtxFiles $evtxFiles -OutputPath $outputPath
   
    Write-Host "`n=== Process Complete ===" -ForegroundColor Cyan
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
catch {
    Write-Host "`nFATAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}