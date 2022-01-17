param([string]$Collections = "437035,3652377,8362253")

# Run hidden from Task scheduler with
# pwsh.exe -WindowStyle Hidden -File "Set-WallpaperFromUnsplash.ps1"
# Supply developer access key when prompted to cache in encrypted file

Function Set-WallPaper($Image) {
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

$sel | Format-List | Out-File -FilePath "$($env:TEMP)/unsplash.txt"

Set-WallPaper "$($env:TEMP)/unsplash.jpg"