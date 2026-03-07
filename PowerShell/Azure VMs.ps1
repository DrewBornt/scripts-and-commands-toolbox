# Below creates a new VM in a specific resource group
New-AzVm -ResourceGroupName {ResourceName} -Name "NAME HERE" -Credential (Get-Credential) -Location "East US" -Image UbuntuLTS -OpenPorts 22 -PublicIpAddressName "testvm-01"

# Below assigns vm to a variable
$vm = (Get-AzVM -Name "NAME HERE" -ResourceGroupName {ResourceName})

# Below shows virtual machine sizes available for the vm
$vm | Get-AzVMSize

Get-AzPublicIpAddress -ResourceGroupName {ResourceName} -Name "NAME HERE"


# Steps to shutdown the VM and delete it. 
Stop-AzVM -Name $vm.Name -ResourceGroupName $vm.ResourceGroupName

# Remove/Deletes the VM
Remove-AzVM -Name $vm.Name -ResourceGroupName $vm.ResourceGroupName

# Show the remaining resources in the resource group of the VM
Get-AzResource -ResourceGroupName $vm.ResourceGroupName | Format-Table

# Only the VM itself is deleted with above
# Remove Network Interface
$vm | Remove-AzNetworkInterface –Force

# Delete OS disk
Get-AzDisk -ResourceGroupName $vm.ResourceGroupName -DiskName $vm.StorageProfile.OSDisk.Name | Remove-AzDisk -Force

# Delete network security group
Get-AzNetworkSecurityGroup -ResourceGroupName $vm.ResourceGroupName | Remove-AzNetworkSecurityGroup -Force

# Delete the public IP address
Get-AzPublicIpAddress -ResourceGroupName $vm.ResourceGroupName | Remove-AzPublicIpAddress -Force