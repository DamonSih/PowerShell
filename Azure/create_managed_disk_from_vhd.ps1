<#
.SYNOPSIS
	Create a managed disk from a VHD file
.DESCRIPTION
	This sample demonstrates how to create a Managed Disk from a VHD file. 
	Create Managed Disks from VHD files in following scenarios:
	1. Create a Managed OS Disk from a specialized VHD file. A specialized VHD is a copy of VHD from an exisitng VM that maintains the user accounts, applications and other state data from your original VM. 
	   Attach this Managed Disk as OS disk to create a new virtual machine.
	2. Create a Managed data Disk from a VHD file. Attach the Managed Disk to an existing VM or attach it as data disk to create a new virtual machine.
.PARAMETER subscriptionId
	Specifies the subscription Id 
.PARAMETER resourceGroupName
	Specifies the name of your resource group 
.PARAMETER diskName
	Specifies the name of managed disk
.PARAMETER diskSize
	Specifies the size of managed disk
.PARAMETER vhdUri
	Specifies the URI of the VHD file that will be used to create managed disk
.PARAMETER storageAccountId
	Specifies the resource Id of the storage account where VHD file is stored
.PARAMETER sku
	Specifies the storage type for the Managed Disk. PremiumLRS or StandardLRS
.PARAMETER location
	Specifies the Azure location (e.g. westus) where Managed Disk will be located
.PARAMETER OSType
	Specifies the OS type. Can be 'None', 'Windows' or 'Linux'. 
.PARAMETER HyperVGeneration
	Specifies the HyperV generation for the OS. Can be 'None', 'V1' or 'V2'. 
.EXAMPLE
	PS> ./create-managed-disk.ps1 -subscriptionId 48d31e3b-17b4-431c-8201-794dc197ea95 -resourceGroupName USE2-NPD-EATOOL-RG -diskName azruse2eatool1-OSdisk-00 -diskSize 100 -vhdUri '<HTTPS link here>' -sku PremiumLRS -location westeurope -osType None -hyperVGeneration None
.NOTES
	Author: Damon Sih Boon Kiat
#>

param(
		[string]$subscriptionId = "",
		[string]$resourceGroupName = "",
		[string]$diskName = "",
		[int]$diskSize = "",
		[string]$vhdUri = "",
		[string]$storageAccountId = "",
		[string]$sku = "",
		[string]$location = "",
		[string]$oSType = "",
		[string]$hyperVGeneration = ""
)

Set-AzContext -Subscription $subscriptionId

if ($OSType -ieq "None"){
	$diskConfig = New-AzDiskConfig -SkuName $sku -Location $location -DiskSizeGB $diskSize -SourceUri $vhdUri -StorageAccountId $storageAccountId -CreateOption Import -Zone 1
	#Create Managed disk
	New-AzDisk -DiskName $diskName -Disk $diskConfig -ResourceGroupName $resourceGroupName -Verbose
}else if ($OSType -ieq "Windows" -or $OSType -ieq "Linux"){
	$diskConfig = New-AzDiskConfig -SkuName $sku -Location $location -DiskSizeGB $diskSize -SourceUri $vhdUri -StorageAccountId $storageAccountId -CreateOption Import -Zone 1 -OSType $osType -hyperVGeneration V2
	#Create Managed disk
	New-AzDisk -DiskName $diskName -Disk $diskConfig -ResourceGroupName $resourceGroupName -Verbose
}else {
	echo "Not doing anything. Exitting..."
}
