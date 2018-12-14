

$server = Get-Content -path C:\ServerList

$server | foreach { (Get-Service -computername $_) | Where-Object {$_.Status -eq "Running"}|
 Select-Object Status, Name, DisplayName | 
ConvertTo-HTML | Out-File "C:\test.htm"}