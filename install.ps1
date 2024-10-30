# Set the current directory to the script's directory
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $scriptPath

# Get the MSI file for Crystal Reports
$CRRFile = Get-ChildItem -Path "MyTMC" -Filter "CRRuntime_64bit*.msi" | Where-Object { $_.Name -match '^CRRuntime_64bit' }
if (-not $CRRFile) {
    Write-Host "Crystal Reports MSI not found."
    exit 1
}

# Define the target directory for NetRun files
$netRunPath = "C:\Program Files\NetRun"

# Step 1: Install Crystal Reports
Write-Host "Starting Crystal Reports installation..."
$process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$($CRRFile.FullName)`" /quiet" -Wait -PassThru

if ($process.ExitCode -eq 0) {
    Write-Host "Crystal Reports installed successfully."
    
    # Define paths for other files within the MyTMC folder
    $netRunExePath = Join-Path "MyTMC" "NetRun.exe"
    $netRunInitRegPath = Join-Path "MyTMC" "NetRunInit.reg"
    $netRunConfigPath = Join-Path "MyTMC" "NetRun.exe.config"
    $mytmcInstallerPath = Join-Path "MyTMC" "MyTMC64.msi"

    # Create the NetRun directory if it does not exist
    if (-not (Test-Path -Path $netRunPath)) {
        New-Item -ItemType Directory -Path $netRunPath | Out-Null
    }

    # Step 2: Move files to Program Files\NetRun directory
    Write-Host "Moving files to Program Files\NetRun..."
    try {
        Copy-Item -Path $netRunExePath -Destination $netRunPath -Force
        Copy-Item -Path $netRunInitRegPath -Destination $netRunPath -Force
        Copy-Item -Path $netRunConfigPath -Destination $netRunPath -Force
        Write-Host "Files moved successfully."
    } catch {
        Write-Host "Error moving files: $_"
        exit 1
    }

    # Step 3: Execute NetRun.exe
    Write-Host "Executing NetRun.exe..."
    Start-Process -FilePath "$netRunPath\NetRun.exe" -Wait

    # Step 4: Import NetRunInit.reg into registry
    Write-Host "Importing NetRunInit.reg into registry..."
    Start-Process -FilePath "regedit.exe" -ArgumentList "import `"$netRunPath\NetRunInit.reg`"" -Wait -NoNewWindow

    # Step 5: Install MyTMC64.msi
    Write-Host "Installing MyTMC64.msi..."
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$mytmcInstallerPath`" /quiet" -Wait

    Write-Host "All operations completed successfully."
} else {
    Write-Host "Crystal Reports installation failed with exit code: $($process.ExitCode)"
}
