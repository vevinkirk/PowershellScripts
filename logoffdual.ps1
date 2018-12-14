
    
    


if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit } 


function Get-LoggedOnUsers ($server) {
  
if($server -eq $null){
    $server = "localhost"
}
  
$users = @()
# Query using quser, 2>$null to hide "No users exists...", then skip to the next server
$quser = quser /server:$server 2>$null
if(!($quser)){
    Continue
}
 
#Remove column headers
$quser = $quser[1..$($quser.Count)]
foreach($user in $quser){
    $usersObj = [PSCustomObject]@{Server=$null;Username=$null;SessionName=$null;SessionId=$Null;SessionState=$null;LogonTime=$null;IdleTime=$null}
    $quserData = $user -split "\s+"
  
    #We have to splice the array if the session is disconnected (as the SESSIONNAME column quserData[2] is empty)
    if(($user | select-string "Disc") -ne $null){
        #User is disconnected
        $quserData = ($quserData[0..1],"null",$quserData[2..($quserData.Length -1)]) -split "\s+"
    }
 
    # Server
    $usersObj.Server = $server
    # Username
    $usersObj.Username = $quserData[1]
    # SessionName
    $usersObj.SessionName = $quserData[2]
    # SessionID
    $usersObj.SessionID = $quserData[3]
    # SessionState
    $usersObj.SessionState = $quserData[4]
    # IdleTime
    $quserData[5] = $quserData[5] -replace "\+",":" -replace "\.","0:0" -replace "Disc","0:0"
    if($quserData[5] -like "*:*"){
        $usersObj.IdleTime = [timespan]"$($quserData[5])"
    }elseif($quserData[5] -eq "." -or $quserData[5] -eq "none"){
        $usersObj.idleTime = [timespan]"0:0"
    }else{
        $usersObj.IdleTime = [timespan]"0:$($quserData[5])"
    }
    # LogonTime
    $usersObj.LogonTime = (Get-Date "$($quserData[6]) $($quserData[7]) $($quserData[8] )")
     
    $users += $usersObj

  
}
  
return $users
  }

function LogOutUser {
    $test = Get-LoggedOnUsers
    $count = 0
    $nameArray = @()
    $sessionArray = @()
    $logonTimeArray = @()
    foreach($line in $test){
        $field1 = $line.Username
        $field2 = $line.SessionID
        $field3 = $line.LogonTime
        $nameArray += $field1
        $sessionArray += $field2
        $logonTimeArray += $field3
        $count+=1
    }
    $session1 = [INT]$sessionArray[0]
    $session2 = [INT]$sessionArray[1]
    Write-Host $nameArray
    Write-Host $sessionArray
    Write-Host $logonTimeArray
    Write-Host $count
    Write-Host $session1
    Write-Host $session2
    
    
    if(($count -gt 1) -and ($session2 -gt $session1 )){
        logoff $session1
        
    }

}


Get-LoggedOnUsers
LogOutUser