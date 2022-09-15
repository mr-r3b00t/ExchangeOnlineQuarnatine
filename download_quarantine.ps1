# Crappy POC to show how to you dowload eml files from quarantine
# mRr3b00t
Install-Module -Name ExchangeOnlineManagement

Set-ExecutionPolicy RemoteSigned

Import-Module ExchangeOnlineManagement

Connect-ExchangeOnline

$Quaratine = Get-QuarantineMessage
$c = 0
foreach($item in $Quaratine){
$item.Identity
$item.MessageId
$c = $c + 1

$path = pwd

$file = "$path\$c" + ".eml"

$name = $item.Subject
$data = Export-QuarantineMessage -Identity $item.Identity

$bytes = [Convert]::FromBase64String($data.eml)
[IO.File]::WriteAllBytes($file, $bytes)


}
