CLS

$DNSServers = "IP1","IP2"
$computers = "Server1","Server2"

foreach ($computer in $computers) {
	Get-WmiObject -Class Win32_NetworkAdapterConfiguration -computername $computer | Foreach-Object{ $_.SetWINSServer('','') | Out-Null } 
    
    $NICs = Get-WMIObject Win32_NetworkAdapterConfiguration -computername $computer 
    $NICs.SetDNSServerSearchOrder($DNSServers) | Out-Null
    $NICs.SetDynamicDNSRegistration(“TRUE”) | Out-Null
    $NICs.SetWINSServer( "IP1","IP2") | Out-Null

    Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE -ComputerName $computer | select PSComputerName, IPAddress, IPSubnet, DefaultIPGateway, DNSServerSearchOrder

    Get-WmiObject –ComputerName $Computer –Class Win32_Volume | Select * 
                        ft –auto DriveLetter,
                        Label,
                        @{Label=”Free(GB)”;Expression={“{0:N0}” –F ($_.FreeSpace/1GB)}},
                        @{Label=”%Free”;Expression={“{0:P0}” –F ($_.FreeSpace/$_.Capacity)}}
}
