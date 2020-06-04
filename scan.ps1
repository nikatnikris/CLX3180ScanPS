 [CmdletBinding()]  
 param (
    [Switch]$CreatePath = $false,
    [ValidateSet(“PNG", "GIF", "JPEG", "TIFF")][String]$ImageType = "JPEG",
    [ValidateSet("75", “150", "300", "600", "1200")][String]$Resolution = "300",
    [ValidateRange(1,100)][String]$JPEGQuality = "90"
 )  


#---------------------------------------------------------------------------------------------------------------
# script is hardcoded to my local config, but this could be spun out into a config file if flexibility is needed
$filePath = "C:\Shared\Scanner"
$deviceID = "{6BDD1FC6-810F-11D0-BEC7-08002BE2092F}\0000"
#---------------------------------------------------------------------------------------------------------------


# check for admin
function Test-Administrator {  
    $user = [Security.Principal.WindowsIdentity]::GetCurrent();
    (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)  
}


# if requested, create path/share
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
$deviceManager = New-Object -ComObject WIA.DeviceManager
$device = ($deviceManager.DeviceInfos | where {$_.DeviceID -eq $deviceID}).Connect()    


# set guid from image type param
$wiaFormatPNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatGIF = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatJPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatTIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
if ($ImageType -eq "PNG"){$wiaFormat = $wiaFormatPNG}
if ($ImageType -eq "GIF"){$wiaFormat = $wiaFormatGIF}
if ($ImageType -eq "JPEG"){$wiaFormat = $wiaFormatJPEG}
if ($ImageType -eq "TIFF"){$wiaFormat = $wiaFormatTIFF}


if ($Resolution -eq "1200"){ Write-Output "! Scanning at 1200 DPI can take a few minutes !" }
Write-Output "scanning image, please wait..."

#set properties and scan
foreach ($item in $device.Items) { 
    foreach($prop in $item.Properties){
        if(($prop.PropertyID -eq 6147) -or ($prop.PropertyID -eq 6148)){ $prop.Value = $Resolution }
    }
    $image = $item.Transfer($wiaFormat) 
}    


#process image
$imageProcess = New-Object -ComObject WIA.ImageProcess
$imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Convert").FilterID)
$imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = $wiaFormat
$imageProcess.Filters.Item(1).Properties.Item("Quality").Value = $JPEGQuality
$image = $imageProcess.Apply($image)


# figure out filename extension
if ($ImageType -eq "PNG"){$fileExt = "png"}
if ($ImageType -eq "GIF"){$fileExt = "gif"}
if ($ImageType -eq "JPEG"){$fileExt = "jpg"}
if ($ImageType -eq "TIFF"){$fileExt = "tiff"}


#save
$fileName = "$((Get-Date).ToString("yyyy-MM-dd-HHmmss")).$fileExt"
$image.SaveFile((Join-Path $filePath $filename))
Write-Output "'$filename' saved to '$filePath'"
