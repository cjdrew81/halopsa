#Requires -RunAsAdministrator

powercfg /H off

$SageSet = "StateFlags0099"
$Base = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\"
$Locations= @(
    "Active Setup Temp Folders"
    "BranchCache"
    "Downloaded Program Files"
    "GameNewsFiles"
    "GameStatisticsFiles"
    "GameUpdateFiles"
    "Internet Cache Files"
    "Memory Dump Files"
    "Offline Pages Files"
    "Old ChkDsk Files"
    "Previous Installations"
    "Recycle Bin"
    "Service Pack Cleanup"
    "Setup Log Files"
    "System error memory dump files"
    "System error minidump files"
    "Temporary Files"
    "Temporary Setup Files"
    "Temporary Sync Files"
    "Thumbnail Cache"
    "Update Cleanup"
    "Upgrade Discarded Files"
    "User file versions"
    "Windows Defender"
    "Windows Error Reporting Archive Files"
    "Windows Error Reporting Queue Files"
    "Windows Error Reporting System Archive Files"
    "Windows Error Reporting System Queue Files"
    "Windows ESD installation files"
    "Windows Upgrade Log Files"
)

ForEach($Location in $Locations) {
    Set-ItemProperty -Path $($Base+$Location) -Name $SageSet -Type DWORD -Value 2 -ea silentlycontinue | Out-Null
}

# do the cleanup . have to convert the SageSet number
$Args = "/sagerun:$([string]([int]$SageSet.Substring($SageSet.Length-4)))"
Start-Process -Wait "$env:SystemRoot\System32\cleanmgr.exe" -ArgumentList $Args -WindowStyle Hidden

# Removw the Stateflags
ForEach($Location in $Locations)
{
    Remove-ItemProperty -Path $($Base+$Location) -Name $SageSet -Force -ea silentlycontinue | Out-Null
}

$old = (Get-Date).adddays(-60)

$Fileset = @(Get-ChildItem -Path "C:\ProgramData\Microsoft\Windows\WER\ReportQueue\*" -Include *.dmp, *.hdmp, *.mdmp -Recurse)
Foreach ($File in $Fileset){If ($File.lastwritetime -lt $Old){Remove-Item $File -Verbose}
Else {"$File is newer than deletion range of $old"}
}

$WindowsFileset = @(Get-ChildItem -Path "C:\Windows\*", "C:\Windows\MiniDump\*" -Include *.dmp, *.hdmp, *.mdmp )
Foreach ($File in $WindowsFileset){If ($File.lastwritetime -lt $Old){Remove-Item $File -Verbose}
Else {"$File is newer than deletion range of $old"}
}

$users = Get-ChildItem 'c:\users'
foreach ($u in $users){
$folder = 'C:\users\' + $u.name +"\AppData\Local\Microsoft\Outlook"
$folderpath = test-path -Path $folder
$folderpath
if($folderpath)
{
Get-ChildItem $folder -filter *.ost | where-object {($_.LastWriteTime -lt (Get-Date).AddDays(-60)) } | remove-item
$Log += Write-Output "Deleted OST file for $user `n"
}
else{
$Log += Write-Output "OST file does not exist or meet criteria for $user `n"
}
}

Send-MailMessage -Body $log -From 'automate@cloud10.it' -SmtpServer 'cloud10-it.mail.protection.outlook.com' -Port 25 -Subject 'Low Disk Space Report - List of OST files deleted' -To 'support@cloud10.i
