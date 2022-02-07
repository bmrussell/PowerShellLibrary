param([string]$Collections = "437035,3652377,8362253", [string]$Style='Fill')


# Get the developer access key by registering for a developer account at Unsplash https://unsplash.com/join
# Supply developer access key when prompted to cache in encrypted file

# Run hidden from Task scheduler with
# pwsh.exe -WindowStyle Hidden -File "Set-WallpaperFromUnsplash.ps1"


Function Set-WallPaper($Image, [string]$Style='Fit') {

    # Set the style of how the wallpaper should be fitted to the desktop resolution
    $WallpaperStyle = Switch ($Style) { 
        "Fill" {"10"}
        "Fit" {"6"}
        "Stretch" {"2"}
        "Tile" {"0"}
        "Center" {"0"}
        "Span" {"22"}    
    }

    If($Style -eq "Tile") {
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -PropertyType String -Value $WallpaperStyle -Force | Out-Null
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -PropertyType String -Value 1 -Force  | Out-Null
    }
    Else {
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -PropertyType String -Value $WallpaperStyle -Force | Out-Null
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -PropertyType String -Value 0 -Force | Out-Null
    }
    if (-not ([System.Management.Automation.PSTypeName]'User32Functions').Type)
    {    
        Add-Type -IgnoreWarnings -TypeDefinition @" 
            using System; 
            using System.Runtime.InteropServices;        
            public class User32Functions
            { 
                [DllImport("User32.dll",CharSet=CharSet.Unicode)] 
                public static extern int SystemParametersInfo (Int32 uAction, Int32 uParam, String lpvParam, Int32 fuWinIni);
            }
"@
    }
    $SPI_SETDESKWALLPAPER = 0x0014
    $updateIni = 0x01
    $fireChangeEvent = 0x02
    $winIniFlags = $updateIni -bor $fireChangeEvent
    [User32Functions]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $Image, $winIniFlags) | Out-Null
}


$creds = (Get-CredentialFromFile.ps1 -File "$($env:USERPROFILE)/Documents/Unsplash.cr")
$accessKey = $creds.GetNetworkCredential().password

# request parameters

$baseUrl = "https://api.unsplash.com"
$randomPhotoUrl = "$($baseUrl)/photos/random"
$headers = @{ "Accept-Version" = "v1"; Authorization = "Client-ID $($accessKey)" }
$params = @{ collections = $collections; orientation = "landscape" }

$content = Invoke-WebRequest $randomPhotoUrl -Method Get -Headers $headers -Body $params | ConvertFrom-Json

Add-Type -AssemblyName System.Windows.Forms
$screenWidth = [System.Windows.Forms.Screen]::AllScreens[0].Bounds.Width

Invoke-WebRequest $content.urls.raw -Headers $headers -Body @{ fm = "jpg"; w = "$($screenWidth)"; q = "80" } -OutFile "$($env:TEMP)/unsplash.jpg"
#Invoke-WebRequest "$baseUrl/photos/$($content.id)/download" -Headers $headers


$sel = $content | Select-Object   @{n = "Name"; e = { $_.user.name } }, @{n = "Location"; e = { $_.location.title } }, @{n = "Description"; e = { $_.description } }

($sel | Format-List | Out-String).Replace('[32;1m', '').Replace('[0m','') | Out-File -FilePath "$($env:TEMP)/unsplash.txt"

Set-WallPaper "$($env:TEMP)/unsplash.jpg" $Style

& bginfo /timer:0