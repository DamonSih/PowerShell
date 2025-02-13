<#
.SYNOPSIS
	Create SQL virtual machine 
.DESCRIPTION
	This PowerShell script create SQL virtual machine for the given executable. Azure PowerShell required.
.PARAMETER VMName
	Specifies the target VM name
.PARAMETER RGName
	Specifies the target Resource Sroup name
.EXAMPLE
	PS> ./create-sql-vm.ps1 -VMName AZRWE1AVLCRDB3 -RGName WE1-EPD-AVLCRETA-RG -License PAYG
.LINK
	https://github.com/themazdarati/Azure
.NOTES
	Author: Damon Sih Boon Kiat | License: CC0
#>

param(
		[string]$VMName = "",
		[string]$RGName = "",
		[string]$License = ""
)

# Get the existing Compute VM
$vm = Get-AzVM -Name $VMName -ResourceGroupName $RGName

New-AzSqlVM -Name $vm.Name -ResourceGroupName $vm.ResourceGroupName -Location $vm.Location -LicenseType $License -SqlManagementType Full
