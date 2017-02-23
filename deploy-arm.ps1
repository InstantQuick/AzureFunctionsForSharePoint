
# Requires Azure PowerShell https://github.com/Azure/azure-powershell/releases
                                                                                                                                               
Write-Host "`n`n*******************************" -ForegroundColor Blue
Write-Host "`n`n`t Azure Functions for SharePoint " -ForegroundColor Blue
Write-Host "`t Azure ARM Resources Setup " -ForegroundColor Blue
Write-Host "`n`n*******************************" -ForegroundColor Blue                                                                                                                                               
                                                                                
# Import developer configuration from json configuration file 
$devConfig = (Get-Content "arm.developer.json" -Raw) | ConvertFrom-Json

$acctName = $devConfig.azureaccount
$subscriptionId = $devConfig.subscriptionid
$resourceGroupName = $devConfig.resourcegroupname
$resourceGroupLocation = $devConfig.resourcegrouplocation
$deploymentName = $devConfig.deploymentname

#There is a bug with Login-AzureRmAccount that the account must be a Azure AD account
#not an outlook.com or hotmail.com account.  If an account name is not provided, prompt for the user 
#to login. See http://stackoverflow.com/questions/39657391/login-azurermaccount-cant-login-to-azure-using-pscredential
if ($devConfig.azureaccount)
{
    $acctPwd = Read-Host -Prompt "Enter password" -AsSecureString
    $psCred = New-Object System.Management.Automation.PSCredential($acctName, $acctPwd)

    Write-Host "Using $($psCred.UserName) to login to Azure..."

    $profile = Login-AzureRmAccount -Credential $psCred  -SubscriptionId $subscriptionId
} else {
    $profile = Login-AzureRmAccount 
}

Get-AzureSubscription
Select-AzureRmSubscription -SubscriptionId $devConfig.subscriptionid
$confirmation = Read-Host "Confirm using subscription above to create $($devConfig.resourcegroupname)? (y/n)"
if ($confirmation -eq 'y') {
    #create the Resource Group which will contain all resources 
    New-AzureRmResourceGroup -Name $resourceGroupName -Location $resourceGroupLocation 

    #use the resource group, tempalte file, and template parameters to create the resources 
    New-AzureRmResourceGroupDeployment -Name $deploymentName -ResourceGroupName $resourceGroupName -TemplateFile "arm-template.json" -TemplateParameterFile "arm-parameters.json" | ConvertTo-Json | Out-File "arm.outputs.json" 	

    $armOutput = (Get-Content "arm.outputs.json" -Raw) | ConvertFrom-Json

    $armOutput.outputs.ConfigurationStorageAccountKey.value

    @{subscriptionid=$subscriptionId;
        storageconnection=$armOutput.outputs.ConfigurationStorageAccountKey.value;
        resourcegroupname=$resourceGroupName;
        functionappname=$armOutput.outputs.FunctionAppName.value;} | ConvertTo-Json | Out-File "deploy.config.json" 

} else {
    Write-Host "`n`n*******************************."
    Write-Host "`n`n`t`tSkipping creation of resource group."
    Write-Host "`n`n*******************************."
}

