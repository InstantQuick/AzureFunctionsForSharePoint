$scriptdir = $PSScriptRoot
[Reflection.Assembly]::LoadFrom("$scriptdir\AzureFunctionsForSharePoint.Core\bin\Debug\AzureFunctionsForSharePoint.Core.dll")
$config =  New-Object AzureFunctionsForSharePoint.Core.ClientConfiguration

#Set your values here

$json = ConvertTo-Json -InputObject $config -Depth 4
$json
ConvertTo-Json -InputObject $config -Depth 4 -Compress | clip