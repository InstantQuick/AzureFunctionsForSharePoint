$baseFileUrl = "site/wwwroot"

# Requires Azure PowerShell https://github.com/Azure/azure-powershell/releases

# Note - You might have to Alt-Tab to get to the login window if using interactive login
 
$deployConfig = (Get-Content "deploy.config.json" -Raw) | ConvertFrom-Json

$subscriptionId = $deployConfig.subscriptionid
$storageCnn = $deployConfig.storageconnection
$resourceGroupName = $deployConfig.resourcegroupname
$functionAppName = $deployConfig.functionappname

$storage = New-AzureStorageContext -ConnectionString $storageCnn

$storageShare = Get-AzureStorageShare -Context $storage -Name $functionAppName

$baseFileUrl = "site/wwwroot"
Login-AzureRmAccount -SubscriptionId $subscriptionId
Invoke-AzureRmResourceAction -ResourceGroupName $resourceGroupName -ResourceType Microsoft.Web/sites -ResourceName $functionAppName -Action stop -ApiVersion 2015-08-01 -Force

$functionNames = ls function.json -Recurse -File | % {$_.Directory.Name}
$functionFolders = ls function.json -Recurse -File | % {$_.Directory}
if($functionFolders.Length -gt 0)
{
    $currentFolder = (Get-Item .).FullName
    $binFolderName = ""
    $binFolderName = ($functionFolders[0].Parent.FullName + "/bin")
    if($binFolderName -ne "")
    {
        #Do all the function.json files
        For ($i=0; $i -lt $functionFolders.Length; $i++) {
            $functionName = $functionNames[$i]
            $fnjson = (Get-ChildItem -Path $functionFolders[$i] function.json)[0]
            New-AzureStorageDirectory -ShareName $functionAppName -Path "$baseFileUrl/$functionName" -Context $storage -ErrorAction SilentlyContinue
            $destination = $baseFileUrl + "/" + $functionName + "/function.json"
            Set-AzureStorageFileContent -ShareName $functionAppName -Source $fnjson.FullName -Context $storage -Path $destination -Force
            "Uploaded $destination"
        }
        
        #Only do files modified in the last n minutes. YMMV on the best time
        #$binFiles = ls $binFolderName -Recurse -File | ? { $_.LastWriteTime -gt (Get-Date).AddMinutes(-5) }
    
	    #Do all files
	    $binFiles = ls $binFolderName -Recurse -File
        New-AzureStorageDirectory -ShareName $functionAppName -Path ($baseFileUrl + "/bin") -Context $storage -ErrorAction SilentlyContinue
	
	    $lastFolder = ""
        $binFiles | % {
            $destination = $baseFileUrl + "/bin/" + $_.FullName.Substring($binFolderName.Length+1).Replace("\","/")
            $destinationFolder = $destination.Replace("/" + $_.Name, "")
            if($lastFolder -ne $destinationFolder) 
            {
                New-AzureStorageDirectory -ShareName $functionAppName -Path $destinationFolder -Context $storage -ErrorAction SilentlyContinue
                $lastFolder = $destinationFolder
            }
            Set-AzureStorageFileContent -ShareName $functionAppName -Source $_.FullName -Context $storage -Path $destination -Force
            "Uploaded $destination"
        }
    }
}

Invoke-AzureRmResourceAction -ResourceGroupName $resourceGroupName -ResourceType Microsoft.Web/sites -ResourceName $functionAppName -Action start -ApiVersion 2015-08-01 -Force

Start-Process ("https://portal.azure.com/#resource/subscriptions/" + $subscriptionId + "/resourcegroups/" + $resourceGroupName + "/providers/Microsoft.Web/sites/" + $functionAppName + "/appServices")
