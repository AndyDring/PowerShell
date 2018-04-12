Login-AzureRmAccount

Get-AzureRmSubscription |Select-AzureRmSubscription

$Location = "uksouth"
$RGName = "NUCUG-Demo-RG"

If (Get-AzureRmResourceGroup -Name $RGName -EA SilentlyContinue) {
    $RG = Get-AzureRmResourceGroup $RGName
} else {
    $RG = New-AzureRmResourceGroup -Name $RGName -Location $Location
}

$rg |Get-AzureRmResourceGroup 

New-AzureRmResourceGroupDeployment -Name TestTemplateDeployment -ResourceGroupName $rg.ResourceGroupName `
    -TemplateFile 'D:\Users\AndyD\OneDrive\Documents\PowerShell\GitHub\PowerShell\AzureIaaS\Deploy-VM.json' `
    -adminUsername "ADTestAdmin" `
    -adminPassword (ConvertTo-SecureString -AsPlainText -Force "P@ssw0rdP@ssw0rd") `
    -dnsLabelPrefix "testdc1" `
    -vmName "TestDC1" 

New-AzureRmResourceGroupDeployment -Name TestTemplateDeployment -ResourceGroupName $rg.ResourceGroupName `
    -TemplateFile 'D:\Users\AndyD\OneDrive\Documents\PowerShell\GitHub\PowerShell\AzureIaaS\Deploy-VM.json' `
    -adminUsername "ADTestAdmin" `
    -adminPassword (ConvertTo-SecureString -AsPlainText -Force "P@ssw0rdP@ssw0rd") `
    -dnsLabelPrefix "testex1" `
    -vmName "Testex1" -windowsOSVersion "2012-R2-Datacenter"


New-AzureRmResourceGroupDeployment -Name TestTemplateDeployment -ResourceGroupName $rg.ResourceGroupName `
    -TemplateFile 'D:\Users\AndyD\OneDrive\Documents\PowerShell\GitHub\PowerShell\AzureIaaS\Deploy-VM.json' `
    -adminUsername "ADTestAdmin" `
    -adminPassword (ConvertTo-SecureString -AsPlainText -Force "P@ssw0rdP@ssw0rd") `
    -dnsLabelPrefix "testex2" `
    -vmName "Testex2" -windowsOSVersion "2012-R2-Datacenter"

New-AzureRmResourceGroupDeployment -Name TestTemplateDeployment -ResourceGroupName $rg.ResourceGroupName `
    -TemplateFile 'D:\Users\AndyD\OneDrive\Documents\PowerShell\GitHub\PowerShell\AzureIaaS\Deploy-VM.json' `
    -adminUsername "ADTestAdmin" `
    -adminPassword (ConvertTo-SecureString -AsPlainText -Force "P@ssw0rdP@ssw0rd") `
    -dnsLabelPrefix "testedge1" `
    -vmName "testedge1" -windowsOSVersion "2012-R2-Datacenter"
<#
Different template


$RootLocation = 'D:\Users\andyd\OneDrive\Documents\PowerShell\WorkScripts\Azure\ARM\active-directory-new-domain-ha-2-dc-zones'

$Location = "uksouth"

If (Get-AzureRmResourceGroup -Name TestTemplateRG) {
    $RG = Get-AzureRmResourceGroup TestTemplateRG
} else {
    $RG = New-AzureRmResourceGroup -Name TestTemplateRG -Location $Location
}

New-AzureRmResourceGroupDeployment -TemplateFile "$RootLocation\azuredeploy.json" `
    -TemplateParameterFile "$RootLocation\azuredeploy.parameters.json" `
    -adminPassword (ConvertTo-SecureString "P@ssw0rdP@ssw0rd" -AsPlainText -Force) `
    -location $Location `
    -dnsPrefix "adt201804" `
    -domainName "adtdom201804.local" `
    -ResourceGroupName $RG.ResourceGroupName
#>    