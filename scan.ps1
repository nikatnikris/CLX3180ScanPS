 [CmdletBinding()]  
 param (
    [switch]$CreatePath = $false
 )  


#---------------------------------------------------------------------------------------------------------------
# script is hardcoded to my local config, but this could be spun out into a config file if flexibility is needed
$filePath = "C:\Shared\Scanner"
$deviceID = "{6BDD1FC6-810F-11D0-BEC7-08002BE2092F}\0000"
#---------------------------------------------------------------------------------------------------------------


function Test-Administrator {  
    $user = [Security.Principal.WindowsIdentity]::GetCurrent();
    (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)  
}


if(-not (Test-Path $filePath)){
    if($CreatePath){
        if(-not (Test-Administrator)){
            throw "You must run powershell as an administrator to create shares"
        }
        $xp = Get-ExecutionPolicy
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force
        New-Item -Path $filePath -ItemType Directory -Force
        if (@(Get-SmbShare | Select Name | Where {$_.Name -eq "Scanner"}).Count -ne 1){
            New-SmbShare -Name "Scanner" -Path $filePath -FullAccess "Everyone"
        }
        Set-ExecutionPolicy $xp -Force
    } else {
        throw "Folder/share '$filePath' does not exist.  Use '-CreatePath' switch to create this"
        exit
    }
}

# connect to device
$deviceManager = new-object -ComObject WIA.DeviceManager
$device = ($deviceManager.DeviceInfos | where {$_.DeviceID -eq $deviceID}).Connect()    


$wiaFormatJPG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
foreach ($item in $device.Items) { 
    foreach($prop in $item.Properties){
        if(($prop.PropertyID -eq 6147) -or ($prop.PropertyID -eq 6148)){ $prop.Value = 300 }
    }
    $image = $item.Transfer($wiaFormatJPG) 
}    

if($image.FormatID -ne $wiaFormatPNG){
    $imageProcess = new-object -ComObject WIA.ImageProcess
    $imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Convert").FilterID)
    $imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = $wiaFormatJPG
    $imageProcess.Filters.Item(1).Properties.Item("Quality").Value = 90
    $image = $imageProcess.Apply($image)
}

$fileName = "C:\Shared\Scanner\$((Get-Date).ToString("yyyy-MM-dd-HHmmss")).jpg"
$image.SaveFile("$filename")
