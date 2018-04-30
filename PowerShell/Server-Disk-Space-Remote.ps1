
$Computers = Get-ADComputer -Filter * | Select -Expand Name
foreach ($Computer in $Computers)
{
    if ([string]$Computer -eq "Server1" -or
        [string]$Computer -eq "Server2" ) 
    {
        echo "Name: $Computer"
        echo "========================================================="
        Get-WmiObject –ComputerName $Computer –Class Win32_Volume | ft –auto DriveLetter,
                        Label,
                        @{Label=”Free(GB)”;Expression={“{0:N0}” –F ($_.FreeSpace/1GB)}},
                        @{Label=”%Free”;Expression={“{0:P0}” –F ($_.FreeSpace/$_.Capacity)}}
    }
}