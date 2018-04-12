Invoke-Expression -Command "diskmgmt.msc"

$PrereqPath = "C:\Build\Pre-reqs"

If (Test-Path "$PrereqPath\UcmaRuntimeSetup.exe") {
    $UCMAInstall = {
        $Arguments = "-q"
        Start-Process -FilePath "$PrereqPath\UcmaRuntimeSetup.exe" -NoNewWindow -Wait -ArgumentList $Arguments
    }
    Write-Host "Installing UCMA 4.0" -ForegroundColor Cyan
    $UCMAInstallJob = Invoke-Command -ScriptBlock $UCMAInstall 
    Write-Host "UCMA 4.0 installed" -ForegroundColor Green
} else {
    Write-Warning "Executable not found at $PrereqPath\UcmaRuntimeSetup.exe; skipping UCMA install"
}

If (Test-Path "$PrereqPath\FilterPack64bit.exe") {
    $FilterpackInstall = {
        $Arguments = "/passive","/quiet","/log:$PrereqPath\FilterpackLog.txt"
        Start-Process -FilePath "$PrereqPath\FilterPack64bit.exe" -NoNewWindow -Wait -ArgumentList $Arguments
    }
    Write-Host "Installing FilterPack" -ForegroundColor Cyan
    $FilterpackInstallJob = Invoke-Command -ScriptBlock $FilterpackInstall 
    Write-Host "FilterPack installed" -ForegroundColor Green
} else {
    Write-Warning "Executable not found at $PrereqPath\FilterPack64bit.exe; skipping Office Filter Pack install"
}
 
If (Test-Path "$PrereqPath\NDP462-KB3151800-x86-x64-AllOS-ENU.exe") {
    $DotNet462Install = {
        $Arguments = "/quiet","/norestart","/log:$PrereqPath\DotNet462Log.txt"
        Start-Process -FilePath "$PrereqPath\NDP462-KB3151800-x86-x64-AllOS-ENU.exe" -NoNewWindow -Wait -ArgumentList $Arguments
    }
    Write-Host "Installing .Net 4.6.2" -ForegroundColor Cyan
    Invoke-Command -ScriptBlock $DotNet462Install
    Write-Host ".Net 4.6.2 installed" -ForegroundColor Green
} else {
    Write-Warning "Executable not found at $PrereqPath\NDP462-KB3151800-x86-x64-AllOS-ENU.exe; skipping .Net 4.6.2 install"
}
 
If (Test-Path "$PrereqPath\NDP471-KB4033342-x86-x64-AllOS-ENU.exe") {
    $DotNet471Install = {
        $Arguments = "/quiet","/norestart","/log:$PrereqPath\DotNet471Log.txt"
        Start-Process -FilePath "$PrereqPath\NDP471-KB4033342-x86-x64-AllOS-ENU.exe" -NoNewWindow -Wait -ArgumentList $Arguments
    }
    Write-Host "Installing .Net 4.7.1" -ForegroundColor Cyan
    Invoke-Command -ScriptBlock $DotNet471Install
    Write-Host ".Net 4.7.1 installed" -ForegroundColor Green
} else {
    Write-Warning "Executable not found at $PrereqPath\NDP471-KB4033342-x86-x64-AllOS-ENU.exe; skipping .Net 4.7.1 install"
}
 
$Response = Read-Host -Prompt "You must reboot the server to proceed. Do you wish to do this now? Type Yes to reboot"
 
If ($Response -eq "Yes") {
    Restart-Computer -Force 
} else {
    Write-Warning "You must restart the computer before continuing"
}

