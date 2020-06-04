# Powershell control for Samsung CLX3180 scanner feature
The Samsung CLX-3180 all-in-one printer/scanner has a built in flat bed scanner feature.  This feature normally requires specific software to control it which requires a GUI.  This project allows the scanner to capture images and place the output files in a share so it can be used on a server via PS remoting.  I.e. the script can be placed on a server or workstation where the printer is attached, and a shortcut can be added to other computers that can trigger a scan an make the image available in a share.

The script is currently focussed around this specific printer and my home server configuration, but could be easily extended to other devices, and to write the output to an alternative network location such as a NAS share.

The script was written for Powershell 5 on Windows and hasn't been tested anywhere else, but should be adaptable for other versions / platforms.



## Usage

Clone the repo to a local path:

```powershell
git clone https://github.com/nikatnikris/CLX3180ScanPS.git
```



Run the script

```powershell
./scan.ps1 -CreatePath -ImageType JPEG -Resolution 300 -JPEGQuality 99
```

Parameter meanings:

`CreatePath` - optional.  If this is included, the image destination folder and SMB share are automatically created if they are missing.  Requires **administrator**

`ImageType` - default JPEG.  Other options are PNG, GIF and TIFF

`Resolution` - default 300.  DPI scanning quality.  Options are 75, 150, 300, 600, 1200

`JPEGQuality` - default 90.  JPEG image compression quality.  Value can be between 1 and 100



## Adapting for your scanner / server

These lines can be updated in the script:

`$filePath = "C:\Shared\Scanner"` - you can set this to whatever works for you.  UNC paths should be OK too

`$deviceID = "{6BDD1FC6-810F-11D0-BEC7-08002BE2092F}\0000"` - you can update this to point to any WIA (Windows Image Acquisition) devices connected to your computer.  These can be found by:

```powershell
$deviceManager = New-Object -ComObject WIA.DeviceManager
$deviceManager.DeviceInfos
```



