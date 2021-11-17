# VM name & network name export script
# Syntax: ./DrNetworkScriptParallel.ps1 -Task (import/export) -CsvPath "C:\vmnetworklist.csv"
# Import will change the network assignments to VMs in the CSV, Export will write the VM and network to the CSV

# Retrieve our arguments. If the task is omitted, it will default to exporting the list. CsvPath is Required regardless.
Param(
    [Parameter(Mandatory=$False)][String]$Task,
    [Parameter(Mandatory=$True)][String]$CsvPath,
    [Parameter(Mandatory=$False)][Int]$Count
    )

# Get the Start Time
$StartTime = (Get-Date)

###########################################################
# Check for PowerShell Core                               #
###########################################################
# Get the PowerShell Version
$PowerShellVersion = $PSVersionTable

# If the PowerCLI Version is not v10 or higher, recommend that the user install PowerShell 7 Core or higher
If ($PowerShellVersion.PSVersion.Major -ge "7" ) {
    Write-Host "PowerShell version 7 or higher present, " -NoNewLine
    Write-Host "proceeding" -ForegroundColor Green 
} else {
    Write-Host "PowerShell version could not be determined or is less than version 7" -Foregroundcolor Red
    Write-Host "Please install PowerCLI 7 or higher and rerun this script" -Foregroundcolor Yellow
    Write-Host " "
    exit
}

###########################################################
# Check for VMware PowerCLI module installation           #
###########################################################

# Get the PowerCLI Version
$PowerCLIVersion = Get-Module -Name VMware.PowerCLI -ListAvailable | Select-Object -Property Version

# If the PowerCLI Version is not v10 or higher, recommend that the user install PowerCLI 10 or higher
If ($PowerCLIVersion.Version.Major -ge "12" -and $PowerCLIVersion.Version.Minor -ge "2") {
    Write-Host "PowerCLI version 12.2 or higher present, " -NoNewLine
    Write-Host "proceeding" -ForegroundColor Green 
} else {
    Write-Host "PowerCLI version could not be determined or is less than version 12.2" -Foregroundcolor Red
    Write-Host "Please install PowerCLI 12.2 or higher and rerun this script" -Foregroundcolor Yellow
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

# Get the PowerCLI Context so it can be passed to the Parallel Foreach Loop
$PowerCLIContext = Get-PowerCLIContext

# If the throttle limit hasn't been specified, default to 10
If (-Not $Count) { $Count = 10 } else { $Count = $Count }

Switch ($Task) {
    # If the import process is specified, changes will be performed to VMs
    "import" { 
        Write-Host "Import Process"

        # Read the data from the CSV file into the $Data array
        $Data = Import-Csv -Path $CsvPath

        # Loop through all lines in the CSV
        # VM's with multiple NICs will go through this process for every NIC
        $Data | ForEach-Object -ThrottleLimit $Count -Parallel {

            # Set the PowerCLI Context so we can use the PowerCLI environment
            Use-PowerCLIContext -PowerCLIContext $using:PowerCLIContext -SkipImportModuleChecks

            # Get the VM object for the VM specified in this record
            $VM = Get-VM -name $($_.VMname)
            # Get the specific NIC object for the VM named
            $NIC = $VM | Get-NetworkAdapter -Name $_.NIC

            # Get the Network Object with the name specified in this record 
            $Network = Get-VirtualPortGroup -Name $_.NetworkName        
            Write-Host "Updating VM: $($VM.Name), Adapter: $($NIC), Network: $($Network)"
            # Update the VM NIC to the specified Network 
            $NIC | Set-NetworkAdapter -Portgroup $Network -Confirm:$False | Out-Null
        }
    }
    # If anything other than import is specified, only export the results
    default {
        $VmList = Get-VM # use -Name '*NAME*' to only get a subset of VMs
        # Could also use Get-Cluster -Name "Cluster1" | Get-VM only to return VM's in a given cluster

        # Create a concurrent bag to put all of our data in. This is "Parallel Safe"
        # This will help with the CSV export process
        $Data = [System.Collections.Concurrent.ConcurrentBag[object]]::new()

        # Perform the export in parallel - Note this will result in the CSV not necessarily being sorted by VM
        $VmList | ForEach-Object -ThrottleLimit $Count -Parallel {

            # Set the PowerCLI Context so we can use the PowerCLI environment
            Use-PowerCLIContext -PowerCLIContext $using:PowerCLIContext -SkipImportModuleChecks

            Write-Host "Exporting Network info for VM: $($_.Name)"
            # Get the VM view of the current VM
            $VMGuest = Get-View $_.Id  
            # Get each NIC for the VM (this takes into account VMs with multiple NICs)
            $NICs = $_ | Get-NetworkAdapter
            # Loop through all of the NICs for the current VM
            foreach ($NIC in $NICs) {
                # Write-Host "    $($NIC.Name)"
                # Create a PSObject to convert VM, NIC, and Network into a single record
                $into = New-Object PSObject
                Add-Member -InputObject $into -MemberType NoteProperty -Name VMname $_.Name
                Add-Member -InputObject $into -MemberType NoteProperty -Name NIC $NIC.Name
                Add-Member -InputObject $into -MemberType NoteProperty -Name NetworkName $NIC.NetworkName
                # Add the current record to the Data array
                $array = $using:data
                $array.Add($into)
            }
        }
        # Sort the list of VMs by VM name and then Network Adapter
        $Data = $Data | Sort-Object VMname,NIC 
        # Write the Data out to the CSV file specified
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
