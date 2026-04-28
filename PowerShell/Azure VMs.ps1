# ============================================================
# Azure VM – Quick-Reference Commands
# ============================================================
# Replace the variables below before running any section.

$ResourceGroupName = "YOUR-RESOURCE-GROUP"
$VMName            = "YOUR-VM-NAME"
$Location          = "East US"
$PublicIPName      = "$VMName-pip"


# ── Create a new VM ─────────────────────────────────────────
# UbuntuLTS alias is deprecated; use a specific URN instead.
# Run: Get-AzVMImageSku -Location $Location -PublisherName Canonical -Offer 0001-com-ubuntu-server-jammy
New-AzVm `
    -ResourceGroupName $ResourceGroupName `
    -Name              $VMName `
    -Credential        (Get-Credential) `
    -Location          $Location `
    -Image             "Canonical:0001-com-ubuntu-server-jammy:22_04-lts:latest" `
    -OpenPorts         22 `
    -PublicIpAddressName $PublicIPName


# ── Get VM object (required before the sections below) ──────
$vm = Get-AzVM -Name $VMName -ResourceGroupName $ResourceGroupName
if (-not $vm) { throw "VM '$VMName' not found in resource group '$ResourceGroupName'." }


# ── List available VM sizes for this VM's region ────────────
$vm | Get-AzVMSize


# ── Get the VM's public IP address ──────────────────────────
Get-AzPublicIpAddress -ResourceGroupName $ResourceGroupName -Name $PublicIPName


# ── Shut down and fully delete a VM (and its dependent resources) ──
#
# CAUTION: Run each Remove-Az* command deliberately.
# Consider adding -WhatIf first to preview what will be deleted.

# 1. Stop the VM gracefully
Stop-AzVM -Name $vm.Name -ResourceGroupName $vm.ResourceGroupName -Force

# 2. Delete the VM compute resource
Remove-AzVM -Name $vm.Name -ResourceGroupName $vm.ResourceGroupName -Force

# 3. Show remaining resources in the group (sanity check before deleting more)
Get-AzResource -ResourceGroupName $vm.ResourceGroupName | Format-Table

# 4. Delete each Network Interface attached to the VM
foreach ($nicRef in $vm.NetworkProfile.NetworkInterfaces) {
    $nic = Get-AzNetworkInterface -ResourceId $nicRef.Id
    Remove-AzNetworkInterface -Name $nic.Name -ResourceGroupName $nic.ResourceGroupName -Force
}

# 5. Delete the OS disk
Get-AzDisk -ResourceGroupName $vm.ResourceGroupName -DiskName $vm.StorageProfile.OSDisk.Name |
    Remove-AzDisk -Force

# 6. Delete the VM's Network Security Group
#    WARNING: filters by the VM name prefix – adjust if your NSG is named differently.
Get-AzNetworkSecurityGroup -ResourceGroupName $vm.ResourceGroupName |
    Where-Object { $_.Name -like "*$($vm.Name)*" } |
    Remove-AzNetworkSecurityGroup -Force

# 7. Delete the VM's public IP address
Get-AzPublicIpAddress -ResourceGroupName $vm.ResourceGroupName -Name $PublicIPName |
    Remove-AzPublicIpAddress -Force
