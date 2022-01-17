param ([string] $name)

# Docker-Add-ContainerIpToHosts
$ip = docker inspect --format '{{ .NetworkSettings.Networks.nat.IPAddress }}' $name
$newEntry = "$ip  $name  #added by d2h# `r`n"
$path = 'C:\Windows\System32\drivers\etc\hosts'
$newEntry + (Get-Content $path -Raw) | Set-Content $path

