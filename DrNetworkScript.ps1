# VM name & network name export script
# Syntax: ./DrNetworkScript.ps1 -Task (import/export) -CsvPath "C:\vmnetworklist.csv"
# Import will change the network assignments to VMs in the CSV, Export will write the VM and network to the CSV

# Retrieve our arguments. If the task is omitted, it will default to exporting the list. CsvPath is Required regardless.
Param(
    [Parameter(Mandatory=$False)][String]$Task,
    [Parameter(Mandatory=$True)][String]$CsvPath
    )

# Get the Start Time
$StartTime = (Get-Date)

###########################################################
# Check for VMware PowerCLI module installation           #
###########################################################

# Get the PowerCLI Version
$PowerCLIVersion = Get-Module -Name VMware.PowerCLI -ListAvailable | Select-Object -Property Version

# If the PowerCLI Version is not v10 or higher, recommend that the user install PowerCLI 10 or higher
If ($PowerCLIVersion.Version.Major -ge "10") {
    Write-Host "PowerCLI version 10 or higher present, " -NoNewLine
    Write-Host "proceeding" -ForegroundColor Green 
} else {
    Write-Host "PowerCLI version could not be determined or is less than version 10" -Foregroundcolor Red
    Write-Host "Please install PowerCLI 10 or higher and rerun this script" -Foregroundcolor Yellow
    Write-Host " "
    exit
}

#########################################################################################################################
# Check to see if we're connected to vCenter. If not, prompt for the vCenter and for Credentials to connect to it.      #
#########################################################################################################################

If ($Global:DefaultVIServer) {
    Write-Host "Connected to " -NoNewline 
    Write-Host $Global:DefaultVIServer -ForegroundColor Green
} else {
    # If not connected to vCenter Server make a connection
    Write-Host "Not connected to vCenter Server" -ForegroundColor Red
    $VIFQDN = Read-Host "Please enter the vCenter Server FQDN"  
    # Prompt for credentials using the native PowerShell Get-Credential cmdlet
    $VICredentials = Get-Credential -Message "Enter credentials for vCenter Server" 
    try {
        # Attempt to connect to the vCenter Server 
        Connect-VIServer -Server $VIFQDN -Credential $VICredentials -ErrorAction Stop | Out-Null
        Write-Host "Connected to $VIFQDN" -ForegroundColor Green 
        # Note that we connected to vCenter so we can disconnect upon termination
        $ConnectVc = $True
    }
    catch {
        # If we could not connect to vCenter report that and exit the script
        Write-Host "Failed to connect to $VIFQDN" -BackgroundColor Red
        Write-Host $Error
        Write-Host "Terminating the script " -BackgroundColor Red
        # Note that we did not connect to vCenter Server
        $ConnectVc = $False
        return
    }
}

########################################
# Preform the import or export process #
########################################

Switch ($Task) {
    # If the import process is specified, changes will be performed to VMs
    "import" { 
        Write-Host "Import Process"

        # Read the data from the CSV file into the $Data array
        $Data = Import-Csv -Path $CsvPath

        # Loop through all lines in the CSV
        # VM's with multiple NICs will go through this process for every NIC
        Foreach ($item in $data) {
            # Get the VM object for the VM specified in this record
            $VM = Get-VM -name $($item.VMname)
            # Get the specific NIC object for the VM named
            $NIC = $VM | Get-NetworkAdapter | Where-Object {$_.Name -eq $item.NIC}
            # Get the Network Object with the name specified in this record 
            $Network = Get-VirtualPortGroup -Name $item.NetworkName        
            Write-Host "Updating VM: $($VM.Name), Adapter: $($NIC), Network: $($Network)"
            # Update the VM NIC to the specified Network 
            $NIC | Set-NetworkAdapter -Portgroup $Network -Confirm:$False | Out-Null
        }
    }
    # If anything other than import is specified, only export the results
    default {
        $VmList = # use -Name '*NAME*' to only get a subset of VMs
        # Could also use Get-Cluster -Name "Cluster1" | Get-VM only to return VM's in a given cluster

        # Create an array to put all of our data in. This will help with the CSV export process
        $Data = @()

        # Loop through the VM's included in the $VmList array
        foreach ($Vm in $VmList){
            Write-Host "Exporting Network info for VM: $($VM.Name)"
            # Get the VM view of the current VM
            $VMGuest = Get-View $VM.Id  
            # Get each NIC for the VM (this takes into account VMs with multiple NICs)
            $NICs = $VM | Get-NetworkAdapter
            # Loop through all of the NICs for the current VM
            foreach ($NIC in $NICs) {
                Write-Host "    $($NIC.Name)"
                # Create a PSObject to convert VM, NIC, and Network into a single record
                $into = New-Object PSObject
                Add-Member -InputObject $into -MemberType NoteProperty -Name VMname $VM.Name
                Add-Member -InputObject $into -MemberType NoteProperty -Name NIC $NIC.Name
                Add-Member -InputObject $into -MemberType NoteProperty -Name NetworkName $NIC.NetworkName
                # Add the current record to the Data array
                $Data += $into   
            }
        }
        # When all VM's have been looped through, export that data to the CSV file 
        $Data | Export-Csv -Path $CsvPath -NoTypeInformation
    }
}

#####################################################################
# Disconnect from vCenter if it was necessary to connect to vCenter #
#####################################################################
# If we had to connect to vCenter, disconnect.
If ($ConnectVc -eq $True) {
    Disconnect-VIserver * -Confirm:$False
}
$EndTime = (Get-Date)
Write-Host "This script took $($EndTime - $StartTime) to run"
