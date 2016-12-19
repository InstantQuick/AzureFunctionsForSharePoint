# Requires Azure PowerShell https://github.com/Azure/azure-powershell/releases

# Note - You might have to Alt-Tab to get to the login window if using interactive login
 
$subscriptionId = "[YOUR_SUBSCRIPTION_ID]"
$storageCnn = "[YOUR_CONNECTION_STRING]"
$resourceGroupName = "[YOUR_RESOURCEGROUP_NAME]"

$storage = New-AzureStorageContext -ConnectionString $storageCnn

$functionAppName = "iqapp"
$storageShare = Get-AzureStorageShare -Context $storage -Name $functionAppName

$baseFileUrl = "site/wwwroot"
Login-AzureRmAccount -SubscriptionId $subscriptionId
Invoke-AzureRmResourceAction -ResourceGroupName $resourceGroupName -ResourceType Microsoft.Web/sites -ResourceName $functionAppName -Action stop -ApiVersion 2015-08-01 -Force

$functionNames = ls Function -Recurse -Directory | % {$_.Parent.Name}
$currentFolder = (Get-Item .).FullName

$functionNames | % {
    $functionName = $_
    
	#Only do files modified in the last n minutes. YMMV on the best time
	$functionFiles = ls $functionName/Function/*.* | ? { $_.LastWriteTime -gt (Get-Date).AddMinutes(-5) }
    New-AzureStorageDirectory -ShareName $functionAppName -Path "$baseFileUrl/$functionName" -Context $storage -ErrorAction SilentlyContinue
    New-AzureStorageDirectory -ShareName $functionAppName -Path "$baseFileUrl/$functionName/bin" -Context $storage -ErrorAction SilentlyContinue
    $functionFiles | % { 
        $destination = $baseFileUrl + "/" + $functionName + "/" + $_.Name
        Set-AzureStorageFileContent -ShareName $functionAppName -Source $_.FullName -Context $storage -Path $destination -Force
        "Uploaded $destination"
    }
    $binFolder = "$currentFolder/$functionName/bin/debug"
	
	#Only do files modified in the last n minutes. YMMV on the best time
    $binFiles = ls $binFolder -Recurse -File | ? { $_.LastWriteTime -gt (Get-Date).AddMinutes(-5) }
    $lastFolder = ""
    $binFiles | % {
        $destination = $baseFileUrl + "/" + $functionName + "/bin/" + $_.FullName.Substring($binFolder.Length+1).Replace("\","/")
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

Invoke-AzureRmResourceAction -ResourceGroupName $resourceGroupName -ResourceType Microsoft.Web/sites -ResourceName $functionAppName -Action start -ApiVersion 2015-08-01 -Force

Start-Process ("https://portal.azure.com/#resource/subscriptions/" + $subscriptionId + "/resourcegroups/" + $resourceGroupName + "/providers/Microsoft.Web/sites/" + $functionAppName + "/appServices")