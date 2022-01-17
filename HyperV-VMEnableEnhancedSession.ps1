param([string]$VmName = "")

Set-VM -VMName $VmName -EnhancedSessionTransportType HvSocket
