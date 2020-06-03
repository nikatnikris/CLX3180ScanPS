if(-not (Test-Path "C:\Shared\Scanner")){
    New-Item -Path "C:\Shared\Scanner" -ItemType Directory -Force
    NET SHARE Scanner=C:\Shared\Scanner /GRANT:everyone,FULL
}

$deviceManager = new-object -ComObject WIA.DeviceManager
$device = $deviceManager.DeviceInfos.Item(1).Connect()    

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
