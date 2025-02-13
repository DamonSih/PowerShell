<#
.SYNOPSIS
	Disk export to Blob 
.DESCRIPTION
	This PowerShell script export disk to blob storage for the given executable. Azure PowerShell required.
.PARAMETER subscriptionId
	Specifies the subscription Id of the subscription where managed disk is created
.PARAMETER resourceGroupName
	Specifies the name of your resource group where managed is created
.PARAMETER diskName
	Specifies the managed disk name to be export
.PARAMETER sasExpiryDuration
	Specifies the Shared Access Signature (SAS) expiry duration in seconds (default 3600 seconds)
.PARAMETER storageAccountName
	Specifies the storage account name where you want to copy the underlying VHD of the managed disk
.PARAMETER storageContainerName
	Specifies the storage container where the downloaded VHD will be stored
.PARAMETER storageAccountKey
	Provide the key of the storage account where you want to copy the VHD of the managed disk
.PARAMETER destinationVHDFileName
	Provide the name of the destination VHD file to which the VHD of the managed disk will be copied
.EXAMPLE
	PS> ./disk-export-to-blob.ps1 -subscriptionId 48d31e3b-17b4-431c-8201-794dc197ea95 -resourceGroupName USE2-NPD-AVLCRETA-RG -diskName AZRUSE2AVLCRDB3_DataDisk_2 -sasExpiryDuration 7200 -storageAccountName migratersa592175211 -storageContainerName export-vhd -storageAccountKey '<YOUR KEY HERE>' -destinationVHDFileName AZRUSE2AVLCRDB3_DataDisk_2.vhd
.NOTES
	Author: Damon Sih Boon Kiat | License: CC0
#>

param(
		[string]$subscriptionId = "",
		[string]$resourceGroupName = "",
		[string]$diskName = "",
		[int]$sasExpiryDuration = "3600",
		[string]$storageAccountName = "",
		[string]$storageContainerName = "",
		[string]$storageAccountKey = "",
		[string]$destinationVHDFileName = ""		
)

#Set the context to the subscription Id where managed disk is created
Select-AzSubscription -SubscriptionId $subscriptionId

#Generate the SAS for the managed disk 
$sas = Grant-AzDiskAccess -ResourceGroupName $resourceGroupName -DiskName $diskName -DurationInSecond $sasExpiryDuration -Access Read

#Create the context of the storage account where the underlying VHD of the managed disk will be copied
$destinationContext = New-AzStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageAccountKey

$copyjob = Start-AzStorageBlobCopy -AbsoluteUri $sas.AccessSAS -DestContainer $storageContainerName -DestContext $destinationContext -DestBlob $destinationVHDFileName

#check copy job status
$copyjob | Get-AzStorageBlobCopyState
