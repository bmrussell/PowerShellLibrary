# PowerShellLibrary
Useful odds and ends


## Docker-*
Misclenaneous docker shortcuts

## Publish-FreeBusy
Saves Outlook Free/Busy to an HTML file

## Send-eMail
Simple wrapper for sending email

## Get-CredentialFromFile
Store/ retrieve PSCredential from file. Don't use for anything too important :)
## Set-WallpaperFromUnsplash
Sets the wallpaper from Unsplash, displaying the image information with [Sysinternals BGInfo](https://docs.microsoft.com/en-us/sysinternals/downloads/bginfo)

Dependencies: 
* `Get-CredentialFromFile.ps1`
* (optional)[Sysinternals BGInfo](https://docs.microsoft.com/en-us/sysinternals/downloads/bginfo)
* Developer access key by [registering](https://unsplash.com/join) for a developer account at Unsplash. Add this when prompted for the password

Run it on from Task scheduler with
```
pwsh.exe -WindowStyle Hidden -File "Set-WallpaperFromUnsplash.ps1"
```