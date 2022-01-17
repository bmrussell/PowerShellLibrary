param([string]$Week = 'next', [string]$SaveAs = ($Env:TEMP + "\availability.html"), [switch]$Silent, [switch]$Help)

if ($Help -eq $true) {
    Write-Host "Cal-GetFreeBusy.ps1 [WeekDay] [SaveAs] [Silent]"
    Write-Host "Cal-GetFreeBusy.ps1 [WeekDay] [SaveAs] [Silent]"
    Write-Host "  Week:   Day in the week that we want to get the free busy for, or 'this' or 'next'. Default is 'this'"
    Write-Host "  SaveAs: File to save output to"
    Write-Host "  Silent: Don't Open file in browser when done"
    exit
}

if ($Week -eq "this") {
    $day = Get-Date
} elseif ($Week -eq "next") {
    $day = (Get-Date).AddDays(7)
} else {
    $day = [Datetime]::ParseExact($Week, 'dd/MM/yyyy', $null)
}

# Work out the previous Monday
$monday = $day.AddDays(-[Int]($day).DayOfWeek+1)
# Work out the following Friday
$friday = $monday.AddDays(4)


try {
    $original_pwd = (Get-Location).Path
    $interop_assemply_location = ''
    try {
        Set-Location 'C:\windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Outlook\15.0.0.0__71e9bce111e9429c\'
        $interop_assemply_location = (Get-ChildItem -Recurse  Microsoft.Office.Interop.Outlook.dll).Directory
        if ($interop_assemply_location -eq "") {
            Set-Location 'C:\Windows\assembly\'
            $interop_assemply_location = (Get-ChildItem -Recurse  Microsoft.Office.Interop.Outlook.dll).Directory    
        }
    }
    catch {        
        Set-Location 'C:\Windows\assembly\'
        $interop_assemply_location = (Get-ChildItem -Recurse  Microsoft.Office.Interop.Outlook.dll).Directory
    }
    Set-Location $interop_assemply_location     
    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
    Set-Location "$original_pwd"
    $Outlook = New-Object -comobject Outlook.Application
}
catch {
    write-host "Couldn't get Outlook."
    exit
}

$namespace = $Outlook.GetNameSpace("MAPI")
$user = $namespace.CreateRecipient($Env:USERNAME)
$indicators = $user.FreeBusy($monday, 30)

$currentDate = $monday
$indStart = 18
$day = 0
$matrix = New-Object 'string[,]' 5,17

do {
    $matrix[$day, 0] = $currentDate.ToShortDateString()

    $dayIndicators = $indicators.Substring($indStart, 16)
    for ($hour = 0; $hour -lt 16; $hour++) {
        if ($dayIndicators.Substring($hour, 1) -eq "0") {
            $status = "Free"
        }
        else {
            $status = "Busy"   
        }
        $matrix[$day, ($hour+1)] = $status
    }
    $indStart = $indStart + 48
    $day++
    $currentDate = $currentDate.AddDays(1)
} until ($currentDate -gt $friday)

# Dump to a Markdown array
$markdown = ""
$line = "|"
for($day = 0; $day -lt 5; $day++) {
	$line = $line +  "|" + $matrix[$day, 0]
}
$line = $line + "|"
$markdown =  $line + "`r`n"

$markdown =  $markdown  + "|--|--|--|--|--|--|`r`n"

$time = [Datetime]::ParseExact('09:00', 'HH:mm', $null) 
for ($hour = 1; $hour -lt 16; $hour++) {
	$line = "|" + $time.ToShortTimeString()
	
	for ($day = 0; $day -lt 5; $day++) {
		$line = $line +  "|" + $matrix[$day, $hour]
    }
	$line = $line + "|"
	$markdown =  $markdown + $line + "`r`n"
    $time = $time.AddMinutes(30)
}
#Write-Host $SaveAs
$markdown | Out-File ($Env:TEMP + "\availability.md")
$md = ConvertFrom-Markdown -Path ($Env:TEMP + "\availability.md")
$md.Html | Out-File -Encoding utf8 $SaveAs
if ($Silent -eq $false) {
    . ($SaveAs)
}
