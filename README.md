# CSV BOM Converter Context Menu Tool

A PowerShell-based tool that adds a context menu option to Windows Explorer for adding BOM (Byte Order Mark) to CSV files. This tool supports both single and multiple file conversions with a user-friendly progress interface.

## Features

- Adds "Add BOM to CSV" option to Windows context menu
- Supports multiple file selection and batch processing
- Shows progress bar for large file operations
- Creates detailed logs of all operations
- Renames processed files with "BOM_" prefix
- Detects existing BOM to prevent duplicate processing
- User-friendly status messages and completion summaries

## Installation

### Prerequisites
- Windows Operating System
- PowerShell 5.1 or higher
- Administrator privileges

### Important: PowerShell Execution Policy

If you encounter permission or security-related errors, you need to adjust PowerShell's execution policy. Here's how:

1. Open PowerShell as Administrator
2. Run the following command:
```powershell
Set-ExecutionPolicy RemoteSigned
```
3. Type 'Y' when prompted to confirm the change

This allows locally created scripts to run while still requiring downloaded scripts to be signed by a trusted publisher.

> **Note**: If you still get security warnings, you might need to unblock the files:
> 1. Right-click the .ps1 file
> 2. Click Properties
> 3. Check the "Unblock" box near the bottom
> 4. Click Apply and OK

### Directory Structure Created
```
C:\bomactivity\
├── N\
│   └── BomCsvConverter.ps1
└── bom_conversion_log_[timestamp].txt
```

## Usage

1. Select one or more CSV files in Windows Explorer
2. Right-click on the selected file(s)
3. Click "Add BOM to CSV" from the context menu
4. Wait for the process to complete
5. Review the summary message

### Notes
- Files will be renamed with "BOM_" prefix after processing
- Logs are stored in C:\bomactivity with timestamps
- Already processed files (with "BOM_" prefix) will be skipped
- Files that already have BOM will be renamed but not modified

# Step-by-Step Installation Guide for CSV BOM Converter

## Step 1: Copy Files
1. Copy the `\bomactivity\` folder to your C: drive
   - The final path should be `C:\bomactivity\`
   - Ensure the `N` folder and scripts are inside this directory

## Step 2: Set PowerShell Execution Policy
1. Right-click on PowerShell
2. Select "Run as administrator"
3. Enter this command:
```powershell
Set-ExecutionPolicy RemoteSigned
```
4. Type 'Y' and press Enter when prompted

## Step 3: Run Setup Script
1. Open PowerShell as administrator (if not already open)
2. Navigate to the script location:
```powershell
cd C:\bomactivity\N
```
3. Run the setup script:
```powershell
.\SetupContextMenu.ps1
```
4. Wait for the success message

## Step 4: Restart Session
1. Sign out of your Windows account
2. Sign back in
   - This step is necessary for the context menu changes to take effect

## Step 5: Use the Tool
1. Find any CSV file
2. Right-click on the file
3. In the context menu, look for and select "Add BOM to CSV"
4. Wait for the conversion process to complete
5. A success message will appear when done

## Verification
- The processed file will be renamed with "BOM_" prefix
- Check C:\bomactivity for the log file
- The original file location will remain the same, just with the new name

## Expected Directory Structure
```
C:\bomactivity\
├── N\
│   └── BomCsvConverter.ps1
│   └── SetupContextMenu.ps1
└── bom_conversion_log_[timestamp].txt
```

## Troubleshooting
If the context menu option doesn't appear:
1. Ensure you signed out and back in
2. Verify the scripts are in the correct location
3. Check that SetupContextMenu.ps1 ran successfully
4. Make sure you ran PowerShell as administrator

If you get permission errors:
1. Verify you ran PowerShell as administrator
2. Check that execution policy is set correctly
3. Ensure you have write access to C:\bomactivity
