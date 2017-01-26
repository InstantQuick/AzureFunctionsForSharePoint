$scriptdir = $PSScriptRoot
[Reflection.Assembly]::LoadFrom("$scriptdir\AzureFunctionsForSharePoint.Core\bin\Debug\AzureFunctionsForSharePoint.Core.dll")
$config =  New-Object AzureFunctionsForSharePoint.Core.ClientConfiguration

#Pretty print output to the PowerShell host window
ConvertTo-Json -InputObject $config -Depth 4

#Send to clipboard
ConvertTo-Json -InputObject $config -Depth 4 -Compress | clip