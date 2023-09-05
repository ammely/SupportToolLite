#Forces powershell to run as an admin
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{ Start-Process powershell.exe "-NoProfile -Windowstyle Hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

#Imports Windowsforms and Drawing from system
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

#Allows the use of wshell for confirmation popups
$wshell = New-Object -ComObject Wscript.Shell
$PSScriptRoot

#Links functions to selected option in the dropdown list, activates on button click
#Outputbox.clear() Erases text output from the outputbox before continuing with the script.
Function selectedscript {

    if ($DropDownBox.Selecteditem -eq "Remove PCEye5 Bundle") {
        $Outputbox.Clear()
        UninstallPCEye5Bundle
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove all ET SW") {
        $Outputbox.Clear()
        UninstallTobiiDeviceDriversForWindows
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove WC&GP Bundle") {
        $Outputbox.Clear()
        UninstallWCGP
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove PCEye Package") {
        $Outputbox.Clear()
        UninstallPCeyePackage
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove Communicator") {
        $Outputbox.Clear()
        UninstallCommunicator
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove Compass") {
        $Outputbox.Clear()
        UninstallCompass
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove TGIS only") {
        $Outputbox.Clear()
        UninstallTGIS
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove TGIS profile calibrations") {
        $Outputbox.Clear()
        TGISProfilesremove
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove all users C5") {
        $Outputbox.Clear()
        DeleteC5User
    }
    elseif ($DropDownBox.Selecteditem -eq "Reset TETC") {
        $Outputbox.Clear()
        ResetTETC
    }
    elseif ($DropDownBox.Selecteditem -eq "Backup Gaze Interaction") {
        $Outputbox.Clear()
        BackupGazeInteraction
    }
    elseif ($DropDownBox.Selecteditem -eq "Copy License") {
        $Outputbox.Clear()
        Copylicenses
    }
    else {
        $Outputbox.AppendText( "" )
        $OutputBox.AppendText( "No option selected. `r`n" )
        Return
    }
}

#A1 Uninstalls PCEye5 Bundle
Function UninstallPCEye5Bundle {

    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will remove all software included in PCeye5Bundle.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress..`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove PCEye5Bundle`r`n" )
        Return
    }
	
    $RegPath = "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig"
    $TempPath = "$ENV:USERPROFILE\AppData\Local\Temp\EyeXConfig.reg"
    if ((Test-Path -Path $RegPath) -and (!(Test-Path -path $TempPath))) {
        $Outputbox.Appendtext("Backup profiles in %temp%\EyeXConfig.reg`r`n")
        Invoke-Command { reg export "HKLM\SOFTWARE\WOW6432Node\Tobii\EyeXConfig" $TempPath }
    }

    $GetProcess = stop-process -Name "*TobiiDynavox*" -Force
    if ($GetProcess) {
        $Outputbox.appendtext("Stopping $GetProcess `r`n" )
    }

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { 
        ($_.Displayname -Match "Tobii Dynavox Computer Control") -or
        ($_.Displayname -Match "Tobii Dynavox Control") -or
        ($_.Displayname -Match "Dynavox Computer Control Updater Service") -or
        ($_.Displayname -Match "Tobii Dynavox Update Notifier") -or
        ($_.Displayname -Match "Tobii Dynavox Eye Tracking") -or
        ($_.Displayname -Eq "Tobii Device Drivers For Windows (PCEye5)") -or
        ($_.Displayname -Eq "Tobii Experience Software For Windows (PCEye5)") -or
        ($_.Displayname -eq "Tobii Dynavox Control ") -or
        ($_.Displayname -Match "Tobii Dynavox Control Updater Service") } | Select-Object Displayname, UninstallString
    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $uninst = $ver.UninstallString -replace "msiexec.exe", "" -Replace "/I", "" -Replace "/X", ""
        $uninst = $uninst.Trim()
        $Outputbox.Appendtext( "Uninstalling - " + "$Uninstname`r`n" )
        start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
    }

    $DeleteServices = Get-Service -Name '*TobiiIS*' , '*TobiiG*' | Stop-Service -Force -passthru -ErrorAction ignore
    foreach ($Service in $DeleteServices) {
        $outputbox.appendtext(" Deleating - " + "$Service `r`n" )
        sc.exe delete $Service
    }

    $TobiiVer = Get-WindowsDriver -Online -All | Where-Object { $_.ProviderName -eq "Tobii AB" } | Select-Object Driver
    ForEach ($ver in $TobiiVer) {
        $outputBox.appendtext( "Removing Drivers - " + "$TobiiVer`r`n" )
        pnputil /delete-driver $ver.Driver /force /uninstall
    }

    #Removes WC related folders
    $paths = (
        "C:\Program Files (x86)\Tobii Dynavox\Eye Tracking Settings",	
        "C:\Program Files (x86)\Tobii Dynavox\Eye Assist",
        "C:\Program Files\Tobii\Tobii EyeX",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\EyeAssist",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\App Switcher",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Computer Control",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Computer Control Bundle",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Update Notifier",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Eye Tracking",
        "$ENV:ProgramData\Tobii Dynavox\EyeAssist",
        "$ENV:ProgramData\Tobii Dynavox\Computer Control",
        "$ENV:ProgramData\HelloDMFT" )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.appendtext( "Removing - " + "$path`r`n" )
            Remove-Item $path -Recurse -Force -ErrorAction Ignore
        }
    }
    $Keys = (
        "HKCU:\Software\Tobii\EyeAssist",
        "HKCU:\Software\Tobii\Update Notifier",
        "HKCU:\Software\Tobii Dynavox\Computer Control",
        "HKLM:\SOFTWARE\WOW6432Node\Tobii Dynavox\Computer Control Updater Service" )

    foreach ($Key in $Keys) {
        if (test-path $Key) {
            $Outputbox.appendtext( "Removing - " + "$Key`r`n" )
            Remove-item $Key -Recurse -ErrorAction Ignore
        }
    }
        
    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { 
        ($_.Displayname -Match "Tobii Dynavox Computer Control") -or
        ($_.Displayname -Match "Dynavox Computer Control Updater Service") -or
        ($_.Displayname -Match "Tobii Dynavox Update Notifier") -or
        ($_.Displayname -Match "Tobii Dynavox Eye Tracking") -or
        ($_.Displayname -Eq "Tobii Device Drivers For Windows (PCEye5)") -or
        ($_.Displayname -Eq "Tobii Experience Software For Windows (PCEye5)") } | Select-Object Displayname
    if ($TobiiVer) {
        $outputBox.appendtext( "$TobiiVer couldn't be uninstalled. Reboot your device and try again.`r`n" )
    }
    $Outputbox.Appendtext( "Done!`r`n" )
}

#A2 Uninstalls ALL Tobii Device Drivers For Windows Bundle
Function UninstallTobiiDeviceDriversForWindows {

    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will remove all software included in Tobii Device Drivers For Windows Bundles.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )

    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove Tobii Device Drivers For Windows`r`n" )
        Return
    }
    
   	$RegPath = "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig"
    $TempPath = "$ENV:USERPROFILE\AppData\Local\Temp\EyeXConfig.reg"
    if ((Test-Path -Path $RegPath) -and (!(Test-Path -path $TempPath))) {
        $Outputbox.Appendtext("Backup profiles in %temp%\EyeXConfig.reg`r`n" )
        Invoke-Command { reg export "HKLM\SOFTWARE\WOW6432Node\Tobii\EyeXConfig" $TempPath }
    }

     $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    Set-Location $fpath
    try { 
        $erroractionpreference = "Stop"
        $Firmware = .\FWUpgrade32.exe --auto --info-only 
    }
    catch [System.Management.Automation.RemoteException] {
        $outputbox.appendtext("PDK is not installed`r`n")
    }
    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Tobii Device Drivers For Windows")  } | Select-Object Displayname, DisplayVersion, UninstallString

    if ($Firmware -match "IS5_Gibbon_Gaze" -and $TobiiVer.DisplayVersion -eq "4.49.0.4000" ) { 
        $outputBox.appendtext( "Running BeforeUninstall.bat script.`r`n" )
        Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
        $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "BeforeUninstall.bat" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
        Set-Location $fpath
        $Installer = cmd /c "BeforeUninstall.bat"
        $Outputbox.appendtext($Installer)
        $outputbox.appendtext("`r`n")
        $Outputbox.appendtext( "Done! `r`n" )
        $outputbox.appendtext("`r`n")
     } 
    else { $outputbox.appendtext( "No need to run the script")}

    $GetProcess = stop-process -Name "*TobiiDynavox*" -Force
    if ($GetProcess) {
        $Outputbox.appendtext("Stopping $GetProcess `r`n" )
    }

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { 
        ($_.Displayname -Match "Tobii Device Drivers For Windows") -or
        ($_.Displayname -Match "Tobii Experience Software") -or
        ($_.Displayname -Match "Tobii Dynavox Eye Tracking Driver") -or
        ($_.Displayname -Match "Tobii Eye Tracking For Windows") -or
        ($_.Displayname -Match "Tobii Dynavox Eye Tracking") -or
        ($_.Displayname -Match "Tobii Eye Tracking") } | Select-Object Displayname, UninstallString
    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $uninst = $ver.UninstallString
        $Outputbox.Appendtext( "Uninstalling - " + "$Uninstname`r`n" )
        $uninst = $ver.UninstallString -replace "msiexec.exe", "" -Replace "/I", "" -Replace "/X", "" -replace "/uninstall", ""
        $uninst = $uninst.Trim()
        if ($uninst -match "ProgramData") {
            try {
                cmd /c $uninst /uninstall /quiet
            }
            catch { 
                Write-Output "not"
            }
        }
        else {
            start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
        }
    }
    
    if (Get-AppxPackage *TobiiAB.TobiiEyeTrackingPortal*) {
        $outputBox.appendtext( "Removing Tobii Experiecne software.`r`n" )
        Get-AppxPackage *TobiiAB.TobiiEyeTrackingPortal* | Remove-AppxPackage
    }

    $DeleteServices = Get-Service -Name '*TobiiIS*' , '*TobiiG*' | Stop-Service -Force -passthru -ErrorAction ignore
    foreach ($Service in $DeleteServices) {
        $outputbox.appendtext(" Deleating - " + "$Service `r`n" )
        sc.exe delete $Service
    }
        
    $TobiiVer = Get-WindowsDriver -Online -All | Where-Object { $_.ProviderName -eq "Tobii AB" } | Select-Object Driver
    ForEach ($ver in $TobiiVer) {
        $outputBox.appendtext( "Removing Drivers - " + "$TobiiVer`r`n" )
        pnputil /delete-driver $ver.Driver /force /uninstall
    }

    #Removes Tobii related folders
    $paths = ( 
        "C:\Program Files\Tobii\Tobii EyeX",
        "$ENV:ProgramData\TetServer",
        "$ENV:ProgramData\Tobii\HelloDMFT",
        "$ENV:ProgramData\Tobii\Statistics",
        "$ENV:ProgramData\Tobii\Tobii Interaction",
        "$ENV:ProgramData\Tobii\Tobii Stream Engine"
    )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.appendtext( "Removing - " + "$path`r`n" )
            Remove-Item $path -Recurse -Force -ErrorAction Ignore
        }
    }
    #Deleting registry keys related to WC
    $Keys = ( 
        "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeX",
        "HKCU:\Software\Tobii\EyeAssist",
        "HKCU:\Software\Tobii\EyeX",
        "HKCU:\Software\Tobii\Vouchers"
    )

    foreach ($Key in $Keys) {
        if (test-path $Key) {
            $Outputbox.appendtext( "Removing - " + "$Key`r`n" )
            Remove-item $Key -Recurse -ErrorAction Ignore
        }
    }
    
    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { 
        ($_.Displayname -Match "Tobii Device Drivers For Windows") -or
        ($_.Displayname -Match "Tobii Experience Software") -or
        ($_.Displayname -Match "Tobii Dynavox Eye Tracking Driver") -or
        ($_.Displayname -Match "Tobii Eye Tracking For Windows") -or
        ($_.Displayname -Match "Tobii Dynavox Eye Tracking") -or
        ($_.Displayname -Match "Tobii Eye Tracking") } | Select-Object Displayname
    $TobiiVer = $TobiiVer.DisplayName
    if ($TobiiVer) {
        $outputBox.appendtext( "$TobiiVer couldn't be uninstalled. Reboot your device and try again.`r`n" )
    }
    $outputbox.appendtext("`r`n")
    $Outputbox.appendtext( "Done!`r`n" )
}

#A3 Uninstalls WC Bundle
Function UninstallWCGP {

    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will remove all software included in Windows Control & Gaze Point Bundles.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )

    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove WC&GP`r`n" )
        Return
    }

    #If second answer equals yes or no
    $answer2 = $wshell.Popup("Do you want to save your licenses on your computer before continuing?", 0, "Caution", 48 + 4)
    if ($answer2 -eq 6) { CopyLicenses }

    elseif ($answer2 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Copy Licenses`r`n" )
    }


    $RegPath = "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig"
    $TempPath = "$ENV:USERPROFILE\AppData\Local\Temp\EyeXConfig.reg"
    if ((Test-Path -Path $RegPath) -and (!(Test-Path -path $TempPath))) {
       	$Outputbox.Appendtext("Backup profiles in %temp%\EyeXConfig.reg`r`n" )
        Invoke-Command { reg export "HKLM\SOFTWARE\WOW6432Node\Tobii\EyeXConfig" $TempPath }
    }


    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Windows Control") -or
        ($_.Displayname -Match "Virtual Remote") -or
        ($_.Displayname -Match "Update Notifier") -or
        ($_.Displayname -Match "Tobii Eye Tracking") -or
        ($_.Displayname -Match "GazeSelection") -or
        ($_.Displayname -Match "Tobii Dynavox Gaze Point") -or
        ($_.Displayname -Match "Tobii Dynavox Gaze Point Configuration Guide") } | Select-Object Displayname, UninstallString

    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $uninst = $ver.UninstallString
        $Outputbox.Appendtext( "Removing - " + "$Uninstname`r`n" )
        & cmd /c $uninst /quiet /norestart
    }

    #Removes WC related folders
    $paths = ( 
        "$Env:USERPROFILE\AppData\Roaming\Tobii\Tobii Interaction\",
        "$Env:USERPROFILE\AppData\Roaming\Tobii\Tobii Interaction Statistics\",
        "$Env:USERPROFILE\AppData\Roaming\Tobii Dynavox\EyeAssist",
        "$Env:USERPROFILE\AppData\Roaming\Tobii Dynavox\Gaze Selection",
        "$Env:USERPROFILE\AppData\Roaming\Tobii Dynavox\Windows Control Bundle",
        "$Env:USERPROFILE\AppData\Roaming\Tobii Dynavox\Gaze Point Bundle",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Update Notifier\",
        "$Env:USERPROFILE\AppData\Local\Tobii\Tobii Interaction\",
        "C:\Program Files (x86)\Tobii Dynavox\Windows Control Configuration Guide",
        "C:\Program Files (x86)\Tobii Dynavox\Gaze Point Configuration Guide",
        "C:\Program Files (x86)\Tobii Dynavox\Update Notifier",
        "C:\Program Files (x86)\Tobii\Service\Plugins",
        "$ENV:ProgramData\Tobii Dynavox\Tobii Interaction\ScreenPlanes\",
        "$ENV:ProgramData\Tobii Dynavox\Update Notifier\",
        "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Windows Control\",
        "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Point\",
        "$ENV:ProgramData\Tobii Dynavox\Windows Control Configuration Guide\",
        "$ENV:ProgramData\Tobii\Statistics\",
        "$ENV:ProgramData\Tobii\Tobii Interaction\",
        "$ENV:ProgramData\Tobii\Tobii Stream Engine\",
        "$ENV:ProgramData\TetServer" )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.appendtext( "Removing - " + "$path`r`n" )
            Remove-Item $path -Recurse -Force -ErrorAction Ignore
        }
    }
	
    #Deleting registry keys related to WC
    $Keys = ( 
        "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeX",
        "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig\",
        "HKLM:\SOFTWARE\Wow6432Node\Tobii\TobiiUpdater\",
        "HKLM:\SOFTWARE\Wow6432Node\Tobii\Update Notifier\",
        "HKLM:\SOFTWARE\Wow6432Node\Tobii\EyeXOverview",
        "HKCU:\Software\Tobii\ExternalNotifications",
        "HKCU:\Software\Tobii\Eye Control Suite",
        "HKCU:\Software\Tobii\EyeX",
        "HKCU:\Software\Tobii\Statistics",
        "HKCU:\Software\Tobii\Vouchers"
    )

    foreach ($Key in $Keys) {
        if (test-path $Key) {
            $Outputbox.appendtext( "Removing - " + "$Key`r`n" )
            Remove-item $Key -Recurse -ErrorAction Ignore
        }
    }

    $Outputbox.Appendtext( "Done!`r`n" )
}

#A4 Uninstall PCEye Package
Function UninstallPCEyePackage {
    #Implement functionality. (PCEye package & TGIS on i-series, start with PCEye package
    $answer1 = $wshell.Popup("This will remove all software included in PCEye Package`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove PCEye Package`r`n" )
        Return
    }

    #If second answer equals yes or no
    $answer2 = $wshell.Popup("Do you want to save your licenses on your computer before continuing?", 0, "Caution", 48 + 4)
    if ($answer2 -eq 6) { CopyLicenses }

    elseif ($answer2 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Copy Licenses`r`n" )
    }

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Tobii Dynavox Gaze Interaction Software") -or
        ($_.Displayname -Match "Tobii Dynavox PCEye Update Notifier") -or
        ($_.Displayname -Match "Tobii Dynavox Gaze Selection Language Packs") -or
        ($_.Displayname -Match "Tobii IS3 Eye Tracker Driver") -or
        ($_.Displayname -Match "Tobii IS4 Eye Tracker Driver") -or
        ($_.Displayname -Match "Tobii Eye Tracker Browser") -or
        ($_.Displayname -Match "Tobii Dynavox PCEye Configuration Guide") -or
        ($_.Displayname -Match "Tobii Dynavox Gaze HID") } | Select-Object Displayname, UninstallString

    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $uninst = $ver.UninstallString
        $Outputbox.Appendtext( "Removing - " + "$Uninstname`r`n" )
        & cmd /c $uninst /quiet /norestart
    }

    $UninstallService = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -match "Tobii Service" }

    ForEach ($Software in $UninstallService) {
        $Uninstname2 = $Software.Name
        $Outputbox.Appendtext( "Removing - " + "$Uninstname2`r`n")
        $Software.Uninstall()
    }

    $paths = ( 
        "$ENV:AppData\Tobii Dynavox\PCEye Configuration Guide",
        "$ENV:AppData\Tobii Dynavox\PCEye Update Notifier\",
        "$ENV:ProgramData\Tobii Dynavox\PCEye Configuration Guide",
        "$ENV:ProgramData\Tobii Dynavox\Gaze Interaction\Server",
        "$ENV:ProgramData\Tobii Dynavox\PCEye Update Notifier",
        "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Interaction",
        "$ENV:ProgramData\Tobii Dynavox\Tobii Interaction",
        "$ENV:ProgramData\Tobii Dynavox\Gaze Selection",
        "$ENV:ProgramData\Tobii\Statistics\",
        "$ENV:ProgramData\Tobii\Tobii Interaction",
        "$ENV:ProgramData\Tobii\Tobii Stream Engine\odin",
        "$ENV:ProgramData\TetServer",
        "$ENV:ProgramData\Tobii Dynavox\Gaze Interaction"
    )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.appendtext( "Removing - " + "$path`r`n" )
            Remove-Item $path -Recurse -Force -ErrorAction Ignore
        }
    }

    $Key = (
        "HKLM:\SOFTWARE\WOW6432Node\Tobii\ProductInformation",
        "HKCU:\SOFTWARE\Tobii\PCEye\Update Notifier",
        "HKCU:\SOFTWARE\Tobii\PCEye")
		
		
    foreach ($key in $Key) {
        if (test-path $key) {
            $Outputbox.appendtext( "Removing - " + "$key`r`n" )
            Remove-Item $key -Force -ErrorAction ignore
        }
    }

    $Outputbox.appendtext( "Done!`r`n" )
}

#A5 Uninstall Communicator
Function UninstallCommunicator {
    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will uninstall Communicator. Are you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove Communicator`r`n" )
        Return
    }

    #If second answer equals yes or no - if "Yes" then it will call the function CopyLicenses and then continue.
    $answer2 = $wshell.Popup("Do you want to save your licenses on your computer before continuing?", 0, "Caution", 48 + 4)
    if ($answer2 -eq 6) { CopyLicenses }

    elseif ($answer2 -ne 6) { $Outputbox.Appendtext( "Action canceled: Copy Licenses`r`n" ) }

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { $_.Displayname -match "Tobii Dynavox Communicator" } | Select-Object Publisher, Displayname, UninstallString

    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $Outputbox.Appendtext( "Removing - " + "$Uninstname`r`n" )
        $uninst = $ver.UninstallString
        & cmd /c $uninst /quiet /norestart
    }

    $paths = ( "$Env:USERPROFILE\AppData\Roaming\Tobii Dynavox\Communicator",
        "$ENV:ProgramData\Tobii Dynavox\Communicator" )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.AppendText( "Removing - " + "$path`r`n")
            Remove-Item $path -Recurse -Force -ErrorAction Ignore
        }
    }

    $Keys = (
        "HKLM:\SOFTWARE\WOW6432Node\Tobii\MyTobii\MPA\VS Communicator 4",
        "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 4",
        "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 5" )

    foreach ($Key in $Keys) {
        if (test-path $Key) {
            $Outputbox.appendtext( "Removing - " + "$Key`r`n" )
            Remove-item $Key -Recurse -ErrorAction Ignore
        }
    }

    $Outputbox.Appendtext( "Done!`r`n" )
}

#A6 Uninstalls only Compass
Function UninstallCompass {
    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will uninstall Compass. Are you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove Compass`r`n" )
        Return
    }

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Tobii Dynavox Compass") } | Select-Object Displayname, UninstallString

    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $Outputbox.Appendtext( "Removing - " + "$Uninstname" )
        $uninst = $ver.UninstallString
        & cmd /c $uninst /quiet /norestart
    }

    $Keys = ( "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Compass" )
    foreach ($Key in $Keys) {
        if (test-path $Key) {
            $Outputbox.appendtext( "Removing - " + "$Key`r`n" )
            Remove-item $Key -Recurse -ErrorAction Ignore
        }
    }

    $Outputbox.appendtext( "Done!`r`n" )
}

#A7 Uninstall TGIS
Function UninstallTGIS {
    $answer1 = $wshell.Popup("This will ONLY remove Tobii Gaze Interaction Software.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove PCEye5Bundle`r`n" )
        Return
    }

    #If second answer equals yes or no
    $answer2 = $wshell.Popup("Do you want to save your licenses on your computer before continuing?", 0, "Caution", 48 + 4)
    if ($answer2 -eq 6) { CopyLicenses }
    elseif ($answer2 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Copy Licenses`r`n" )
    }

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Tobii Dynavox Gaze Interaction Software") } | Select-Object Displayname, UninstallString
    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $uninst = $ver.UninstallString
        $Outputbox.Appendtext( "Removing - " + "$Uninstname`r`n" )
        & cmd /c $uninst /quiet /norestart
    }

    $paths = (
        "$env:ProgramData\Tobii Dynavox\Gaze Interaction\",
        "$ENV:ProgramData\Tobii Dynavox\Gaze Selection\Word Prediction\Language Packs\")

    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.appendtext( "Removing - " + "$path`r`n" )
            Remove-Item $path -Recurse -Force -ErrorAction Ignore
        }
    }

    $Outputbox.appendtext( "Done!`r`n" )
}

#A8 Function for the option "Remove TGIS calibration profiles #Tobii service is stopped
Function TGISProfilesremove {

    $answer1 = $wshell.Popup("This will remove ONLY calibrations for every profile, it will NOT remove the actual profiles. The Gaze Interaction software will close and tobii service will restart.`r`nContinue?", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.appendtext( "Shutting down TGIS software...`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $outputBox.appendtext( "Action canceled: Remove calibration profiles." )
    }	

    $Processkills = get-process "Tobii.Service", "TobiiEyeControlOptions", "TobiiEyeControlServer", "Notifier" | Stop-process -force -Passthru -erroraction ignore | Select Processname |
    Format-table -Hidetableheaders | Out-string
    foreach ($Processkill in $Processkills) {
        if ($Processkill) {
            $Outputbox.Appendtext( "Stopping: " + "$Processkill`r`n" )
        }
    }
    
    $paths = ( "$ENV:ProgramData\Tobii Dynavox\Gaze Interaction\Server\Calibration\*" )
    foreach ($path in $paths) {
        if (Test-Path $path) {
            remove-Item $path -Recurse -Force -ErrorAction Ignore
            $Outputbox.appendtext("Calibrations found! - Removing...`r`n" )
        }
        else {
            $Outputbox.Appendtext( "No calibration profiles were found!`r`n" )
        }
    }
    try {
        Start-Service -Name "Tobii Service" -ErrorAction Stop
        sleep 1
        $Outputbox.Appendtext( "Tobii Service started! `r`n")
    }
    Catch {
        $Outputbox.Appendtext( "Tobii Service failed to start!`r`n" )
    }

    $outputbox.appendtext( "Done!`r`n" )
}

#A9
Function DeleteC5User {
    $outputBox.clear()
    $outputBox.appendtext( "Deleting C5 users.`r`n" )
    $paths = ( 
        "$env:USERPROFILE\Documents\Communicator 5",
        "$env:USERPROFILE\AppData\Local\VirtualStore\Program Files (x86)\Tobii Dynavox\Communicator 5",
        "$env:USERPROFILE\AppData\Roaming\Tobii Dynavox\Communicator",
        "$env:ProgramData\Tobii Dynavox\Communicator")
    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.appendtext( "Removing - " + "$path`r`n" )
            Remove-Item $path -Recurse -Force -ErrorAction Ignore
        }
    }
    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("Done `r`n")
}

#A10 Resets and restart TETC Configuration
Function ResetTETC {
    $outputBox.clear()

    #Question if you want to start do this action
    $answer1 = $wshell.Popup("NOTE: This is an option for Windows Control!`r`nThis will close TETC, remove all calibration profiles and saved screenplanes to reset it to a clean state.`r`nContinue?", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $outputBox.AppendText( "Starting...`r`n" )

    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Reset TETC.`r`n" )
        return
    }

    try {
        $Processkills = get-process  -Name '*Tobii.Service*', '*Tobii.EyeX.Engine*', '*Tobii.EyeX.Interaction*', '*Tobii.EyeX.Tray*' | Stop-process -force -Passthru -erroraction ignore | Select Processname | Format-table -Hidetableheaders | Out-string
    }
    catch { 
        Write-Host ( "No processes were found`r`n" )
    }
    foreach ($Processkill in $Processkills) {
        if ($Processkill) {
            $Outputbox.Appendtext( "Stopping: " + "$Processkill`r`n" )
        }
    }
 
    $keys = ( "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig\" )
    $Keys2 = ( "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig\*" )

    Foreach ($Key in $Keys) {
        if (test-path $keys) {
            $outputBox.appendtext( "Configuration files found! - Removing...`r`n" )
            Remove-itemProperty $Keys -Name "DefaultEyeTracker" -ErrorAction Ignore
            Remove-item $Keys2 -Recurse -Force -ErrorAction Ignore
        }
        else {
            $outputBox.Appendtext("No TETC configuration files were found!`r`n")
        }
    }

    try {
        $Outputbox.Appendtext( "Attempting to start Tobii Service...`r`n" )
        Start-Service -Name "Tobii Service" -ErrorAction Stop
        $Outputbox.Appendtext( "Done!`r`n")
    }
    Catch {
        $Outputbox.Appendtext( "Tobii Service failed to start!`r`n" )
    }

    try {
        $OutputBox.AppendText( "Attempting to start TETC...`r`n" )
        Start-process "C:\Program Files (x86)\Tobii\Tobii EyeX Interaction\Tobii.EyeX.Tray.exe" -ErrorAction Stop
    }
    Catch {
        $outputBox.AppendText( "TETC failed to start!`r`n" )
    }
    $Outputbox.Appendtext( "Finished!`r`n" )
}

#A11
Function BackupGazeInteraction {
    $outputBox.clear()
    $path = ( "C:\ProgramData\Tobii Dynavox\Old Gaze Interaction" )

    $outputbox.Appendtext( "Attempting to backup folder...`r`n" )
    if (Test-path $path) {
        $outputBox.appendtext( "Backup folder already exist in: C:\ProgramData\Tobii Dynavox\Old Gaze Interaction, please move it to another location or remove it before trying to backup again." )
    }
    else {
        try {
            Copy-item "C:\ProgramData\Tobii Dynavox\Gaze Interaction\" "C:\ProgramData\Tobii Dynavox\Old Gaze Interaction\" -Recurse -Erroraction Stop
            $outputBox.appendtext( "Copying Gaze Interaction folder to 'Old Gaze Interaction' and placing it in C:\ProgramData\Tobii Dynavox\`r`n" )
            $outputBox.appendtext( "Finished!`r`n" )
        }
        Catch {
            $outputBox.appendtext( "Failed - No Gaze Interaction folder could be found!`r`n" )
        }
    }
}

#A12 Copy licenses function. If any path to $Licensepaths exists, it will make a folder "Tobii Licenses", copy the licensefolders to the new folder(Does not contain the keys.xml, it is only the folder)
Function Copylicenses {
    $outputBox.clear()
    $licensepaths = ( "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Windows Control",
        "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Interaction",
        "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 5",
        "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 4",
        "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Viewer" )

    $outputBox.appendtext( "Looking for licenses to copy...`r`n" )
    ForEach ($Path in $licensepaths) {
        if (test-path $path) {
            mkdir "C:\Tobii Licenses" -erroraction ignore
            copy-item $path "C:\Tobii Licenses" -erroraction ignore
            $outputBox.appendtext( "" )
        }
        elseif ((test-path "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\*") -eq $False) {
            $outputBox.appendtext( "No licenses found.`r`n" )
            Return
        }
    }

    $outputBox.AppendText( "Copying licenses to C:\Tobii Licenses...`r`n" )

    #Retrieves the content from keys.xml
    #Filters the content to only get the string between the activationkey words
    #Creates txt files for licenses
    if (test-path "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Windows Control\*") {
        $GetcontentWC = get-content "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Windows Control\keys.xml"
        $Outputbox.appendtext( "-- Window Control license copied.`r`n" )
        $LicenseWC = [regex]::Matches($getcontentWC, '(?<=\<ActivationKey\>).+(?=\</ActivationKey\>)', "singleline").Value.trim()
        $LicenseWC | Out-file "C:\Tobii Licenses\Windows Control\Windows Control License.txt" -erroraction ignore
    }

    if (test-path "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Interaction\*") {
        $GetcontentTGIS = get-content "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Interaction\keys.xml"
        $Outputbox.appendtext( "-- Gaze Interaction license copied.`r`n" )
        $LicenseTGIS = [regex]::Matches($getcontentTGIS, '(?<=\<ActivationKey\>).+(?=\</ActivationKey\>)', "singleline").Value.trim()
        $LicenseTGIS | Out-file "C:\Tobii Licenses\Gaze Interaction\Gaze Interaction License.txt" -erroraction ignore
    }

    if (test-path "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 5\*") {
        $GetcontentTC5 = get-content "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 5\keys.xml"
        $Outputbox.appendtext( "-- Communicator 5 license copied.`r`n" )
        $LicenseTC5 = [regex]::Matches($getcontentTC5, '(?<=\<ActivationKey\>).+(?=\</ActivationKey\>)', "singleline").Value.trim()
        $LicenseTC5 | Out-file "C:\Tobii Licenses\Communicator 5\Communicator 5 License.txt" -erroraction ignore
    }

    if (test-path "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 4\*") {
        $GetcontentTC4 = get-content "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 4\keys.xml"
        $Outputbox.appendtext( "-- Communicator 4 license copied.`r`n" )
        $LicenseTC4 = [regex]::Matches($getcontentTC4, '(?<=\<ActivationKey\>).+(?=\</ActivationKey\>)', "singleline").Value.trim()
        $LicenseTC4 | Out-file "C:\Tobii Licenses\Communicator 4\Communicator 4 License.txt" -erroraction ignore
    } #Add compass to the list.

    $outputBox.AppendText( "Done!`r`n" )
    Return
}

#B1 Function listapps - outputs all installed apps with the publisher Tobii
Function Listapps {
    $Outputbox.clear()
    $Outputbox.Appendtext( "Listing installed Tobii software... (If empty, no software found) `r`n" )
    $Listapps = Get-ChildItem -Recurse -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\,
    HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\,
    HKLM:\Software\WOW6432Node\Tobii\ |
    Get-ItemProperty | Where-Object { $_.Publisher -like '*Tobii*' } | Select Displayname, Displayversion | Sort Displayname | format-table -HideTableHeaders | out-string
    $Listwindowsapp = Get-AppxPackage | Where-Object { ($_.Publisher -like '*Tobii*') -or
        ($_.Name -like '*Snap*') } | Select name | format-table -HideTableHeaders | out-string
    $outputBox.AppendText( "TOBII INSTALLED SOFTWARE:$Listapps`n" )
    $outputBox.AppendText( "TOBII WINDOWS STORE APPS:$Listwindowsapp" )
}

#B2 Lists currently active tobii processes & services
Function GetProcesses {
    $outputBox.clear()
    $outputBox.appendtext( "Listing active Tobii processes. (If empty - no processes were found) `r`n" )
    $GetProcess = get-process "*GazeSelection*", "*Tobii*" | Select Processname | Format-table -hidetableheaders | Out-string
    $GetServices = Get-Service -Name '*Tobii*' | Select Name, Status | Format-table -hidetableheaders | Out-string
    if ($GetProcess) {
        $outputbox.appendtext("ACTIVE PROCESSES:$GetProcess")
        $Outputbox.Appendtext( "`r`n" )
    }
    else {
        $outputbox.appendtext("NO ACTIVE PROCESSES")
        $Outputbox.Appendtext( "`r`n" )
    }
    if ($GetServices) {
        $outputbox.appendtext("ACTIVE Services:$GetServices")
        $Outputbox.Appendtext( "`r`n" )
    }
    else {
        $outputbox.appendtext("NO ACTIVE Services")
        $Outputbox.Appendtext( "`r`n" )
    }
}

#B3
Function IS5PID {
    $outputBox.clear()
    $outputBox.appendtext( "Checking IS5 PID...`r`n" )
    $getdeviceid = $null
    $getdeviceid = gwmi Win32_USBControllerDevice | % { [wmi]($_.Dependent) } | Where-Object DeviceID -Like "*Tobii*" | Select-object DeviceID
    $outputbox.appendtext("$getdeviceid `r`n")								   
    $getdeviceid2 = Get-CimInstance Win32_PnPSignedDriver | Where-Object Description -Like "*WinUSB Device*" | Select-Object DeviceID
    Start-Sleep -s 5
    $outputbox.appendtext("$getdeviceid2`r`n")								   
    if (!$getdeviceid -or !$getdeviceid2) {
        $outputbox.appendtext("the tracker is not connected")
    }
    # gwmi Win32_USBControllerDevice |%{[wmi]($_.Dependent)} | Sort Manufacturer,Description,DeviceID | Ft -GroupBy Manufacturer Description,Service,DeviceID | out-file c:\VidPid.txt
    $outputbox.appendtext("Done`r`n")
    $outputbox.appendtext("`r`n")
}

#B4
Function ListDrivers {
    $outputBox.clear()
    $outputBox.appendtext( "Listing all drivers in c:/tobii.txt and here...`r`n" )
    pnputil /enum-drivers >c:\tobii.txt
    $TobiiDrivers = Get-WindowsDriver -Online -All | Where-Object { $_.ProviderName -eq "Tobii AB" } | Select-Object Driver , OriginalFileName
    ForEach ($drivers in $TobiiDrivers) {
        $inf = $drivers.Driver 
        $List = $drivers.OriginalFileName
        $List = $List.Replace("C:\Windows\System32\DriverStore\FileRepository\", "")
        $outputbox.appendtext("`r`n$inf : $List `r`n")
    }
    $outputbox.appendtext("`r`nDone")
}

#B5
Function ETfw {
    $outputBox.clear()
    $outputBox.appendtext( "Checking Eye tracker Firmware...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    Set-Location $fpath
    try { 
        $erroractionpreference = "Stop"
        $Firmware = .\FWUpgrade32.exe --auto --info-only 
    }
    Catch [System.Management.Automation.RemoteException] {
        $outputbox.appendtext("No Eye Tracker Connected`r`n")
    }
    $outputbox.appendtext($Firmware) 
    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("Done `r`n")
}

#B7 Stops all currently active tobii processes
Function RestartProcesses {
    $outputBox.clear()
    $Outputbox.Appendtext( "Restart Services...`r`n")
    $StopServices = Get-Service -Name '*Tobii*' | Stop-Service -force -Passthru -erroraction ignore | Select Name, Status | Format-table -hidetableheaders | Out-string
    $Outputbox.Appendtext( "Stopping following Services:$StopServices")

    Start-Sleep -s 3
    $Processkill = get-process "GazeSelection" , "*TobiiDynavox*", "*Tobii.EyeX*", "Notifier" | Stop-process -force -Passthru -erroraction ignore | Select Processname | Format-table -Hidetableheaders | Out-string
    $Outputbox.Appendtext( "Stopping following processes:$Processkill")

    #start all processes and services
    Start-Sleep -s 3
    try {
        $StartServices = Start-Service -Name '*Tobii*' -ErrorAction Stop
        $Outputbox.Appendtext( "Attempting to start following Services:$StartServices `r`n" )
    }
    Catch {
        $Outputbox.Appendtext( "Tobii Service failed to start!`r`n" )
    }
    Start-Sleep -s 3
    try {
        $StartProcesses = Start-process "C:\Program Files (x86)\Tobii Dynavox\Eye Assist\TobiiDynavox.EyeAssist.Engine.exe"
        $Outputbox.Appendtext( "Attempting to start Eyeassist:$StartProcesses `r`n" )
    }
    Catch {
        $outputBox.Appendtext( "EyeAssist failed to start!`r`n" )
    }
    $outputBox.Appendtext( "Done!" )
}

#B11
Function FWUpgrade {
    $outputBox.clear()
    $outputBox.appendtext( "Upgrade IS4 ET FW...`r`n" )
    $path = "C:\Program Files (x86)\Tobii\Service"
    if (Test-Path $path) {
        Set-Location -path $path
        $ETInfo = .\FWUpgrade32.exe --auto --info-only
        $outputbox.appendtext("Connected ET is: $ETInfo")
        $outputbox.appendtext("`r`n")
    }

    else {
        $outputbox.appendtext("No Eye Tracker Connected`r`n")
    }
    if ($ETInfo -match "PCE1M") {
        #PCEye Mini: tobii-ttp://PCE1M-010106010685
        $outputbox.appendtext("`r`n")
        $outputbox.appendtext("Upgrading PCEye mini FW..")
        $PCEyeMini = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii\Tobii Firmware\is4pceyemini_firmware_2.27.0-4014648.tobiipkg" --no-version-check
        $outputbox.appendtext($PCEyeMini)
        $outputbox.appendtext("`r`n")
        $outputbox.appendtext("Done `r`n")
    }
    elseif ($ETInfo -match "IS4_Large_102") {
        $outputbox.appendtext("`r`n")
        $outputbox.appendtext("Upgrading PCEye Plus FW..")
        $PCEyePlus = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii\Tobii Firmware\is4large102_firmware_2.27.0-4014648.tobiipkg" --no-version-check
        $outputbox.appendtext($PCEyePlus)
        $outputbox.appendtext("`r`n")
        $outputbox.appendtext("Done `r`n")
    }
    elseif ($ETInfo -match "IS4_Large_Peripheral") {
        $outputbox.appendtext("`r`n")
        $outputbox.appendtext("Upgrading 4C FW..")
        $4C = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii\Tobii Firmware\is4largetobiiperipheral_firmware_2.27.0-4014648.tobiipkg" --no-version-check
        $outputbox.appendtext($4C)
        $outputbox.appendtext("`r`n")
        $outputbox.appendtext("Done `r`n")
    }
    elseif ($ETInfo -match "IS4_Base_I-series") {
        $outputbox.appendtext("`r`n")
        $outputbox.appendtext("Upgrading I-Series+ FW..")
        $ISeries = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii Dynavox\Gaze Interaction\Eye Tracker Firmware Releases\IS4B1\is4iseriesb_firmware_2.9.0.tobiipkg" --no-version-check
        $outputbox.appendtext($ISeries)
        $outputbox.appendtext("`r`n")
        $outputbox.appendtext("Done. Restart ET through Control Center `r`n")
    }
    elseif ($ETInfo -match "tet-tcp") {
        #Tobii Firmware Upgrade Tool Automatically selected eye tracker tet-tcp://172.28.195.1 Failed to open file
        $outputbox.appendtext("ET model is IS20. Use ET Browser to upgrade. Make sure that Bonjure is installed.")
    }
    else {
        $outputbox.appendtext("No ET connected or ET not supported")
    }
}

#B12
Function BeforeUninstallGG {
    $outputBox.clear()
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    Set-Location $fpath
    try { 
        $erroractionpreference = "Stop"
        $Firmware = .\FWUpgrade32.exe --auto --info-only 
    }
    catch [System.Management.Automation.RemoteException] {
        $outputbox.appendtext("PDK is not installed`r`n")
    }
    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Tobii Device Drivers For Windows")  } | Select-Object Displayname, DisplayVersion, UninstallString

    if ($Firmware -match "IS5_Gibbon_Gaze" -and $TobiiVer.DisplayVersion -eq "4.49.0.4000" ) { 
        $outputBox.appendtext( "Running BeforeUninstall.bat script.`r`n" )
        Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
        $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "BeforeUninstall.bat" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
        Set-Location $fpath
        $Installer = cmd /c "BeforeUninstall.bat"
        $Outputbox.appendtext($Installer)
        $outputbox.appendtext("`r`n")
        $Outputbox.appendtext( "Done! `r`n" )
        $outputbox.appendtext("`r`n")
     } 
    else { $outputbox.appendtext( "No need to run the script")}
}

#B17
Function WCF {
    $outputBox.clear()
    $outputBox.appendtext( "Checking WCF Endpoint Blocking Software...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "handle.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    Set-Location $fpath
    Start-Process cmd "/c  `"handle.exe net.pipe & pause `""
    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("Done `r`n")
}

#B18
Function SMBios {
    $outputBox.clear()
    #[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    #$title = 'SMBios tool'
    #$msg = "Press`r`n1 to run getSMBIOSvalues.cmd `r`n " # 2 setName.cmd, `r`n 3 setSerialNumber.cmd, `r`n 4 setVendor.cmd"
    #$b = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "getSMBIOSvalues.cmd" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    Set-Location $fpath

    #if ($b -match "1") { 
        $getvaluses = Start-Process -FilePath .\getSMBIOSvalues.cmd
    #}
    #elseif ($b -match "2") {
    #    $getvaluses = Start-Process -FilePath .\setName.cmd
    #}
    #elseif ($b -match "3") { 
    #    $getvaluses = Start-Process -FilePath .\setSerialNumber.cmd
    #}
    #elseif ($b -match "4") { 
    #    $getvaluses = Start-Process -FilePath .\setVendor.cmd
    #}
    #else { $outputbox.appendtext("N/A") }

    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("Done `r`n")
}

#B23
function GetFrameworkVersionsAndHandleOperation() {
    $outputBox.clear()
    $installedFrameworks = @()
    if (IsKeyPresent "HKLM:\Software\Microsoft\.NETFramework\Policy\v1.0" "3705") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 1.0`r`n") }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v1.1.4322" "Install") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 1.1`r`n") }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v2.0.50727" "Install") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 2.0`r`n") }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v3.0\Setup" "InstallSuccess") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 3.0`r`n") }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v3.5" "Install") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 3.5`r`n" ) }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Client" "Install") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 4.0c`r`n" ) }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full" "Install") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 4.0`r`n" ) }   

    $result = -1
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Client" "Install" -or IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full" "Install") {
        # .net 4.0 is installed
        $result = 0
        $version = GetFrameworkValue "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full" "Release"
        
        if ($version -ge 528040 -Or $version -ge 528372 -Or $version -ge 528049) {
            # .net 4.8
            $outputbox.appendtext( "Installed .Net Framework 4.8")
            $result = 10
        }
        elseif ($version -ge 461808 -Or $version -ge 461814) {
            # .net 4.7.2
            $outputbox.appendtext("Installed .Net Framework 4.7.2")
            $result = 9
        }
        elseif ($version -ge 461308 -Or $version -ge 461310) {
            # .net 4.7.1
            $outputbox.appendtext( "Installed .Net Framework 4.7.1")
            $result = 8
        }
        elseif ($version -ge 460798 -Or $version -ge 460805) {
            # .net 4.7
            $outputbox.appendtext( "Installed .Net Framework 4.7")
            $result = 7
        }
        elseif ($version -ge 394802 -Or $version -ge 394806) {
            # .net 4.6.2
            $outputbox.appendtext( "Installed .Net Framework 4.6.2")
            $result = 6
        }
        elseif ($version -ge 394254 -Or $version -ge 394271) {
            # .net 4.6.1
            $outputbox.appendtext( "Installed .Net Framework 4.6.1")
            $result = 5
        }
        elseif ($version -ge 393295 -Or $version -ge 393297) {
            # .net 4.6
            $outputbox.appendtext( "Installed .Net Framework 4.6")
            $result = 4
        }
        elseif ($version -ge 379893) {
            # .net 4.5.2
            $outputbox.appendtext( "Installed .Net Framework 4.5.2")
            $result = 3
        }
        elseif ($version -ge 378675) {
            # .net 4.5.1
            $outputbox.appendtext( "Installed .Net Framework 4.5.1")
            $result = 2
        }
        elseif ($version -ge 378389) {
            # .net 4.5
            $outputbox.appendtext( "Installed .Net Framework 4.5")
            $result = 1
        }   
    }
    else {
        # .net framework 4 family isn't installed
        $result = -1
    }
    
    return $result    
    #$version = GetFramework40FamilyVersion;
    return $installedFrameworks
    
    if ($version -ge 1) { 
    }
    else { }
}

function IsKeyPresent([string]$path, [string]$key) {
    if (!(Test-Path $path)) { return $false }
    if ((Get-ItemProperty $path).$key -eq $null) { return $false }
    return $true
}
function GetFrameworkValue([string]$path, [string]$key) {
    if (!(Test-Path $path)) { return "-1" }
    return (Get-ItemProperty $path).$key  
}

#B25
Function TrackStatus {
    $outputBox.clear()
    $outputBox.appendtext( "Showing EA Track Status...`r`n" )
    $testpath = "C:\Program Files (x86)\Tobii Dynavox\Eye Assist"
    if (!(Test-path $testpath)) {
        $outputbox.appendtext("EA may not been installed. Make sure that EA is installed and try again.")
    }
    else {
        Set-Location $testpath
        $value = Get-Process | Where-Object { $_.MainWindowTitle -like "track status" } | Select-Object MainWindowTitle
        if ($value) {
            .\TobiiDynavox.EyeAssist.Smorgasbord.exe --hidetrackstatus
        }
        elseif (!($value)) {
            .\TobiiDynavox.EyeAssist.Smorgasbord.exe --showtrackstatus
        }
    }
    $outputbox.appendtext("Done `r`n")
}

#Windows forms
$Optionlist = @("Remove PCEye5 Bundle", "Remove all ET SW", "Remove WC&GP Bundle", "Remove PCEye Package", "Remove Communicator", "Remove Compass", "Remove TGIS only", "Remove TGIS profile calibrations", "Remove all users C5", "Reset TETC", "Backup Gaze Interaction", "Copy License")
$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(600, 550)
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $False

#Informationtext above the dropdown list.
$DropDownLabel = new-object System.Windows.Forms.Label
$DropDownLabel.Location = new-object System.Drawing.Size(10, 10)
$DropDownLabel.size = new-object System.Drawing.Size(160, 20)
$DropDownLabel.Text = "Select an option"
$Form.Controls.Add($DropDownLabel)

#Dropdown list with options
$DropDownBox = New-Object System.Windows.Forms.ComboBox
$DropDownBox.Location = New-Object System.Drawing.Size(10, 30)
$DropDownBox.Size = New-Object System.Drawing.Size(220, 20)
$DropDownBox.DropDownHeight = 230
$Form.Controls.Add($DropDownBox)

#For each arrayitem in optionlist, add it to $dropdownbox items.
foreach ($option in $optionlist) {
    $DropDownBox.Items.Add($option)
}

#Outputbox
$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Location = New-Object System.Drawing.Size(10, 150)
$outputBox.Size = New-Object System.Drawing.Size(400, 340)
$outputBox.MultiLine = $True
$outputBox.ScrollBars = "Vertical"
$Form.Controls.Add($outputBox)
$outputBox.font = New-Object System.Drawing.Font ("Consolas" , 8, [System.Drawing.FontStyle]::Regular)

#Button "Start"
$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(10, 80)
$Button.Size = New-Object System.Drawing.Size(180, 50)
$Button.Text = "Start"
$Button.Font = New-Object System.Drawing.Font ("" , 12, [System.Drawing.FontStyle]::Regular)
$Form.Controls.Add($Button)
$Button.Add_Click{ selectedscript }

#B1 Button1 "List Tobii Software"
$Button1 = New-Object System.Windows.Forms.Button
$Button1.Location = New-Object System.Drawing.Size(420, 10)
$Button1.Size = New-Object System.Drawing.Size(150, 30)
$Button1.Text = "Tobii Software"
$Button1.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button1)
$Button1.Add_Click{ ListApps }

#B2 Button2 "List active Tobii processes"
$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Size(420, 40)
$Button2.Size = New-Object System.Drawing.Size(150, 30)
$Button2.Text = "Active Process+Service"
$Button2.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button2)
$Button2.Add_Click{ GetProcesses }

#B3 Button3 "Check IS5 PID"
$Button3 = New-Object System.Windows.Forms.Button
$Button3.Location = New-Object System.Drawing.Size(420, 70)
$Button3.Size = New-Object System.Drawing.Size(150, 30)
$Button3.Text = "List IS5 PID"
$Button3.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button3)
$Button3.Add_Click{ IS5PID }

#B4 Button4 "List Tobii Drivers"
$Button4 = New-Object System.Windows.Forms.Button
$Button4.Location = New-Object System.Drawing.Size(420, 100)
$Button4.Size = New-Object System.Drawing.Size(150, 30)
$Button4.Text = "Tobii Drivers"
$Button4.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button4)
$Button4.Add_Click{ ListDrivers }

#B5 Button5 "ET fw"
$Button5 = New-Object System.Windows.Forms.Button
$Button5.Location = New-Object System.Drawing.Size(420, 130)
$Button5.Size = New-Object System.Drawing.Size(150, 30)
$Button5.Text = "ET firmware"
$Button5.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button5)
$Button5.Add_Click{ ETfw }

#B7 Button7 Restart Services
$Button7 = New-Object System.Windows.Forms.Button
$Button7.Location = New-Object System.Drawing.Size(420, 160)
$Button7.Size = New-Object System.Drawing.Size(150, 30)
$Button7.Text = "Restart Services"
$Button7.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button7)
$Button7.Add_Click{ RestartProcesses }

#B11 Button11 "FW Upgrade"
$Button11 = New-Object System.Windows.Forms.Button
$Button11.Location = New-Object System.Drawing.Size(420, 190)
$Button11.Size = New-Object System.Drawing.Size(150, 30)
$Button11.Text = "FW Upgrade"
$Button11.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button11)
$Button11.Add_Click{ FWUpgrade }

#B12 Button12 "Before Uninstall GG"
$Button12 = New-Object System.Windows.Forms.Button
$Button12.Location = New-Object System.Drawing.Size(420, 220)
$Button12.Size = New-Object System.Drawing.Size(150, 30)
$Button12.Text = "Before Uninstall GG"
$Button12.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button12)
$Button12.Add_Click{ BeforeUninstallGG }

#B17 Button17 "WCF"
$Button17 = New-Object System.Windows.Forms.Button
$Button17.Location = New-Object System.Drawing.Size(420, 250)
$Button17.Size = New-Object System.Drawing.Size(150, 30)
$Button17.Text = "WCF"
$Button17.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button17)
$Button17.Add_Click{ WCF }

#B18 Button18 "SMBios"
$Button18 = New-Object System.Windows.Forms.Button
$Button18.Location = New-Object System.Drawing.Size(420, 280)
$Button18.Size = New-Object System.Drawing.Size(150, 30)
$Button18.Text = "SMBIOS"
$Button18.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button18)
$Button18.Add_Click{ SMBios }

#B23 Button23 ".NET version"
$Button23 = New-Object System.Windows.Forms.Button
$Button23.Location = New-Object System.Drawing.Size(420, 310)
$Button23.Size = New-Object System.Drawing.Size(150, 30)
$Button23.Text = ".NET v."
$Button23.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button23)
$Button23.Add_Click{ GetFrameworkVersionsAndHandleOperation }

#B25 Button25 "Show Track status"
$Button25 = New-Object System.Windows.Forms.Button
$Button25.Location = New-Object System.Drawing.Size(420, 340)
$Button25.Size = New-Object System.Drawing.Size(150, 30)
$Button25.Text = "Show/hide Track S."
$Button25.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button25)
$Button25.Add_Click{ TrackStatus }

#Form name + activate form.
$Form.Text = "Support Tool 2.0.1"
$Form.Add_Shown( { $Form.Activate() })
$Form.ShowDialog()