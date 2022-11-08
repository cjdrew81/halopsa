$FolderTest = Test-Path 'c:\scripts'
If (!$FolderTest) {
    New-Item -Path c:\ -Name Scripts -ItemType Directory -Force
}


$SpaceReport = @()
$FreespaceBefore = [math]::Round(((Get-WmiObject win32_logicaldisk -filter "DeviceID='C:'" | select Freespace).FreeSpace / 1GB), 2)
$SpaceReport += "Free Space at start of script - $($FreeSpaceBefore) GB"

# This will empty the recycle bin

$Date = Get-Date -Format 'dd-MM-yy'
$objShell = New-Object -ComObject Shell.Application
$objFolder = $objShell.Namespace(0xA)
$objFolder.items() | % { remove-item $_.path -Recurse -Confirm:$true }

$FreeSpace = [math]::Round(((Get-WmiObject win32_logicaldisk -filter "DeviceID='C:'" | select Freespace).FreeSpace / 1GB), 2)
$SpaceReport += "Free Space after Empty Recycle Bin - $($FreeSpace) GB" 

$Hiber = test-path 'C:\hiber.sys'
If ($Hiber) {
    powercfg /H off
    $FreeSpace = [math]::Round(((Get-WmiObject win32_logicaldisk -filter "DeviceID='C:'" | select Freespace).FreeSpace / 1GB), 2)
    $SpaceReport += "Free Space after disabling hibernation - $($FreeSpace) GB"
}

if (test-path 'C:\Config.Msi') {
    remove-item -Path 'C:\Config.Msi' -force -recurse
}
if (test-path 'c:\Intel') {
    remove-item -Path 'c:\Intel' -force -recurse
}
if (test-path 'c:\PerfLogs') {
    remove-item -Path 'c:\PerfLogs' -force -recurse
}
if (test-path 'c:\swsetup') {
    remove-item -Path 'c:\swsetup' -force -recurse
}
if (test-path '$env:windir\memory.dmp') {
    remove-item '$env:windir\memory.dmp' -force
}

# Deleting Windows Error Reporting files
if (test-path 'C:\ProgramData\Microsoft\Windows\WER') { Get-ChildItem -Path C:\ProgramData\Microsoft\Windows\WER -Recurse | Remove-Item -force -recurse }



Write-host "Removing System and User Temp Files" -foreground yellow
Remove-Item -Path "$env:windir\Temp\*" -Force -Recurse
Remove-Item -Path "$env:windir\minidump\*" -Force -Recurse
Remove-Item -Path "$env:windir\Prefetch\*" -Force -Recurse
Remove-Item -Path "C:\Users\*\AppData\Local\Temp\*" -Force -Recurse
Remove-Item -Path "C:\Users\*\AppData\Local\Microsoft\Windows\WER\*" -Force -Recurse
Remove-Item -Path "C:\Users\*\AppData\Local\Microsoft\Windows\Temporary Internet Files\*" -Force -Recurse
Remove-Item -Path "C:\Users\*\AppData\Local\Microsoft\Windows\IECompatCache\*" -Force -Recurse
Remove-Item -Path "C:\Users\*\AppData\Local\Microsoft\Windows\IECompatUaCache\*" -Force -Recurse
Remove-Item -Path "C:\Users\*\AppData\Local\Microsoft\Windows\IEDownloadHistory\*" -Force -Recurse
Remove-Item -Path "C:\Users\*\AppData\Local\Microsoft\Windows\INetCache\*" -Force -Recurse
Remove-Item -Path "C:\Users\*\AppData\Local\Microsoft\Windows\INetCookies\*" -Force -Recurse
Remove-Item -Path "C:\Users\*\AppData\Local\Microsoft\Terminal Server Client\Cache\*" -Force -Recurse
$FreeSpace = [math]::Round(((Get-WmiObject win32_logicaldisk -filter "DeviceID='C:'" | select Freespace).FreeSpace / 1GB), 2)
$SpaceReport += "Free space after removing system and user temp files - $($FreeSpace) GB"


Write-host "Removing Windows Updates Downloads" -foreground yellow
Stop-Service wuauserv -Force -Verbose
Stop-Service TrustedInstaller -Force -Verbose
Remove-Item -Path "$env:windir\SoftwareDistribution\*" -Force -Recurse
Remove-Item $env:windir\Logs\CBS\* -force -recurse
Start-Service wuauserv -Verbose
Start-Service TrustedInstaller -Verbose
$FreeSpace = [math]::Round(((Get-WmiObject win32_logicaldisk -filter "DeviceID='C:'" | select Freespace).FreeSpace / 1GB), 2)
$SpaceReport += "Free Space after removing WIndows Updates cache - $($FreeSpace) GB"


Write-host "Checkif Windows Cleanup exists" -foreground yellow
#Mainly for 2008 servers
if (!(Test-Path c:\windows\System32\cleanmgr.exe)) {
Write-host "Windows Cleanup NOT installed now installing" -foreground yellow
copy-item $env:windir\winsxs\amd64_microsoft-windows-cleanmgr_31bf3856ad364e35_6.1.7600.16385_none_c9392808773cd7da\cleanmgr.exe $env:windir\System32
copy-item $env:windir\winsxs\amd64_microsoft-windows-cleanmgr.resources_31bf3856ad364e35_6.1.7600.16385_en-us_b9cb6194b257cc63\cleanmgr.exe.mui $env:windir\System32\en-US
}


Write-host "Running Windows System Cleanup" -foreground yellow
#Set StateFlags setting for each item in Windows disk cleanup utility
$StateFlags = 'StateFlags0013'
$StateRun = $StateFlags.Substring($StateFlags.get_Length()-2)
$StateRun = '/sagerun:' + $StateRun
if  (-not (get-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Active Setup Temp Folders' -name $StateFlags)) {
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Active Setup Temp Folders' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\BranchCache' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Downloaded Program Files' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Internet Cache Files' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Offline Pages Files' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Old ChkDsk Files' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Previous Installations' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Memory Dump Files' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Recycle Bin' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Service Pack Cleanup' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Setup Log Files' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\System error memory dump files' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\System error minidump files' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Files' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Setup Files' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Thumbnail Cache' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Update Cleanup' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Upgrade Discarded Files' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\User file versions' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Defender' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting Archive Files' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting Queue Files' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting System Archive Files' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting System Queue Files' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting Temp Files' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows ESD installation files' -name $StateFlags -type DWORD -Value 2
    set-itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Upgrade Log Files' -name $StateFlags -type DWORD -Value 2
}

Start-Process -FilePath CleanMgr.exe -ArgumentList $StateRun  -WindowStyle Hidden -Wait


$FreeSpace = [math]::Round(((Get-WmiObject win32_logicaldisk -filter "DeviceID='C:'" | select Freespace).FreeSpace / 1GB), 2)
$SPaceReport += "Free space after running disk cleanup tool - $($FreeSpace) GB"

wevtutil el | Foreach-Object { Write-Host "Clearing $_"; wevtutil cl "$_" }

$OSTList = @()
$users = Get-ChildItem 'c:\users'
foreach ($u in $users) {
    $folder = 'C:\users\' + $u.name + "\appdata\local\microsoft"
    $folderpath = test-path -Path $folder
    if ($folderpath) {
        $OSTList += (Get-ChildItem $Folder -filter "*.ost" -Recurse -Force | where-object { ($_.LastWriteTime -lt (Get-Date).AddDays(-60)) }).fullname
    }
}
If ($OSTList) {
    Foreach ($OST in $OSTList) {
        if ($ost -like "*.ost*") { 
            Remove-Item -Path $OST -Force
            $SpaceReport += "Deleted $($OST)"
        }
    }
}
else { $SpaceReport += "No OST files found to remove" }

$FreeSpace = [math]::Round(((Get-WmiObject win32_logicaldisk -filter "DeviceID='C:'" | select Freespace).FreeSpace / 1GB), 2)
$SpaceReport += "Free space after removing old OST files - $($Freespace) GB"

$FileSize = '104857600'
$fileslimit = 20
$filesLocation = 'C:\'
$largeSizefiles = get-ChildItem -path $filesLocation -recurse -ErrorAction "SilentlyContinue" | ? { $_.GetType().Name -eq "FileInfo" } | where-Object { $_.Length -gt $fileSize } | sort-Object -property length -Descending | Select-Object Name, @{Name = "Size In MB"; Expression = { "{0:N0}" -f ($_.Length / 1MB) } }, @{Name = "LastWriteTime"; Expression = { $_.LastWriteTime } }, @{Name = "Path"; Expression = { $_.directory } } -first $filesLimit
$Report = $largeSizefiles | convertto-html -fragment | out-file 'c:\scripts\filelist.html'

$FreespaceReport = $SpaceReport | select @{L = "Task"; E = { ($_.split("-"))[0] } }, @{L = "Free Space" ; E = { ($_.split("-"))[1] } } | convertto-html -Fragment
$FreespaceReport | out-file 'c:\scripts\freespacereport.html'

$DiskInfo = Get-PhysicalDisk | Select mediatype, friendlyname, operationalstatus, healthstatus, @{L = "Size"; E = { [math]::round($_.size / 1GB, 2) } } | convertto-html | out-file 'c:\scripts\drives.html'
FreeSpaceReport


