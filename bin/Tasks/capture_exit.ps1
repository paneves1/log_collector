##### Captures exit code of silent EXE #####

### path may need to be changed 

$p = Start-Process '.\silent_new_version.exe' -Wait -PassThru
Write-Host "Exit code is $($p.ExitCode)"
