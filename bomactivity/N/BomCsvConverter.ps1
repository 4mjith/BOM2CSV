param(
    [Parameter(Mandatory=$true)]
    [string[]]$FilePaths
)

# Add Windows Forms assembly for MessageBox and progress bar
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create log directory if it doesn't exist
$logDirectory = "C:\bomactivity"
if (-not (Test-Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
}

# Define log file path with timestamp
$timestamp = (Get-Date -Format "yyyyMMdd_HHmmss")
$logFile = Join-Path $logDirectory "bom_conversion_log_$timestamp.txt"

# Create progress form
function Create-ProgressForm {
    param(
        [int]$TotalFiles
    )
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Converting CSV Files"
    $form.Size = New-Object System.Drawing.Size(500, 200)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    # Current file label
    $currentFileLabel = New-Object System.Windows.Forms.Label
    $currentFileLabel.Location = New-Object System.Drawing.Point(10, 20)
    $currentFileLabel.Size = New-Object System.Drawing.Size(460, 20)
    $currentFileLabel.Text = "Processing files..."
    $form.Controls.Add($currentFileLabel)

    # Overall progress label
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Location = New-Object System.Drawing.Point(10, 50)
    $progressLabel.Size = New-Object System.Drawing.Size(460, 20)
    $progressLabel.Text = "Overall Progress: 0/$TotalFiles"
    $form.Controls.Add($progressLabel)

    # Progress Bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(10, 80)
    $progressBar.Size = New-Object System.Drawing.Size(460, 23)
    $progressBar.Minimum = 0
    $progressBar.Maximum = $TotalFiles
    $progressBar.Value = 0
    $form.Controls.Add($progressBar)

    # Status label for additional information
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(10, 120)
    $statusLabel.Size = New-Object System.Drawing.Size(460, 40)
    $statusLabel.Text = "Starting conversion..."
    $form.Controls.Add($statusLabel)

    return @{
        Form = $form
        CurrentFileLabel = $currentFileLabel
        ProgressLabel = $progressLabel
        ProgressBar = $progressBar
        StatusLabel = $statusLabel
    }
}

# Function to process a single file
function Process-File {
    param(
        [string]$FilePath,
        [hashtable]$ProgressForm
    )

    try {
        # Get file information
        $directory = Split-Path -Parent $FilePath
        $fileName = Split-Path -Leaf $FilePath
        
        # Update status
        $ProgressForm.CurrentFileLabel.Text = "Processing: $fileName"
        $ProgressForm.StatusLabel.Text = "Reading file..."
        [System.Windows.Forms.Application]::DoEvents()

        # Check if filename already starts with BOM_
        if ($fileName -like "BOM_*") {
            Write-Host "File already has BOM_ prefix: $fileName"
            return "Already processed: BOM_ prefix exists"
        }

        # Read the file content
        $content = [System.IO.File]::ReadAllBytes($FilePath)
        
        # Check if file already has BOM
        if ($content.Length -ge 3 -and $content[0] -eq 0xEF -and $content[1] -eq 0xBB -and $content[2] -eq 0xBF) {
            $newFileName = "BOM_" + $fileName
            $newPath = Join-Path $directory $newFileName
            Rename-Item -Path $FilePath -NewName $newFileName -Force
            Write-Host "BOM already exists, renamed: $newFileName"
            return "Already has BOM, renamed file"
        }

        # Add BOM and write content
        $ProgressForm.StatusLabel.Text = "Adding BOM..."
        [System.Windows.Forms.Application]::DoEvents()
        
        $BOM = [System.Text.Encoding]::UTF8.GetPreamble()
        $newContent = $BOM + $content

        # Create the new filename with BOM_ prefix
        $newFileName = "BOM_" + $fileName
        $newPath = Join-Path $directory $newFileName

        # Delete existing file if it exists
        if (Test-Path $newPath) {
            Remove-Item $newPath -Force
        }

        # Write the content and rename
        [System.IO.File]::WriteAllBytes($FilePath, $newContent)
        Rename-Item -Path $FilePath -NewName $newFileName -Force
        
        Write-Host "BOM added and renamed: $newFileName"
        return "Successfully added BOM"
    }
    catch {
        Write-Host "Error processing $FilePath : $_"
        return "Error: $_"
    }
}

# Start recording output
Start-Transcript -Path $logFile -Append

try {
    $totalFiles = $FilePaths.Count
    $processedFiles = 0
    $results = @()

    # Create and show progress form
    $progressForm = Create-ProgressForm -TotalFiles $totalFiles
    $progressForm.Form.Show()
    [System.Windows.Forms.Application]::DoEvents()

    foreach ($file in $FilePaths) {
        if (Test-Path $file) {
            $result = Process-File -FilePath $file -ProgressForm $progressForm
            $results += "$file : $result"
        }
        else {
            $results += "$file : File not found"
        }

        $processedFiles++
        $progressForm.ProgressBar.Value = $processedFiles
        $progressForm.ProgressLabel.Text = "Overall Progress: $processedFiles/$totalFiles"
        [System.Windows.Forms.Application]::DoEvents()
    }

    # Close progress form
    $progressForm.Form.Close()
    $progressForm.Form.Dispose()

    # Show summary
    $summary = "Processing Complete:`n`n" + ($results -join "`n")
    [System.Windows.Forms.MessageBox]::Show($summary, "Conversion Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Error during processing: $_`nCheck log file: $logFile", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
finally {
    Stop-Transcript
}
