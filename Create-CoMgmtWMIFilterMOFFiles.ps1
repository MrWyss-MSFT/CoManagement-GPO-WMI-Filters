<#
  .SYNOPSIS
  Creates WMI Filter MOF Files for specific Co-Management workloads and to what they are set.
  .DESCRIPTION
  Creates two WMI Filter MOF files per workload in a out folder of the current directory
  or in the directory specified with the -OutDir Param.
  One WMI Filter for when the workload is on Intune and one if it's on ConfigMgr.
#>

[CmdletBinding()]
Param (
    # Output Directory of the MOF Files, excluding trailing "\""
    [Parameter(Mandatory = $false)]
    #[ValidateScript({if(Test-Path $_ -PathType Container) {$true} else {throw "Path $_ is not valid"}})]
    [string]
    $OutDir = ".\out",
    # For which workload config the MOF files should be created Intune or ConfigMgr, if not specified both
    [Parameter(Mandatory = $false)]
    [ValidateSet("Intune", "ConfigMgr")]
    [string[]]$WorkloadConfigs = @('Intune', 'ConfigMgr')
)

Begin {
    #region Functions
    Function New-WMIFilterMOF {
        [CmdletBinding()]
        param (
            # Name of the WMI Filter
            [Parameter(Mandatory = $true)]
            [string]
            $Name,
            # WQL Query
            [Parameter(Mandatory = $true)]
            [string]
            $WQLQuery,
            # Creation and Change Date
            [Parameter(Mandatory = $false)]
            [string]
            $Date = [Management.ManagementDateTimeConverter]::ToDmtfDateTime($(get-date)),
            # GUID 
            [Parameter(Mandatory = $false)]
            [string]
            $Guid = [string](New-Guid)
        )
    
        #region WMIFilter MOF Template
        $wmiFilterMOF = @"

instance of MSFT_SomFilter
{
	Author = "https://github.com/mrwyss-msft";
	ChangeDate = $Date;
	CreationDate = $Date;
	Domain = "domain.tld";
	ID = "{D8D7BCE7-71A4-4BB9-9F4C-6322B132BF43}";
	Name = "$Name";
	Rules = {
instance of MSFT_Rule
{
	Query = "$WQLQuery";
	QueryLanguage = "WQL";
	TargetNameSpace = "root\\ccm\\InvAgt";
}};
};
"@
        #endregion 
        return $wmiFilterMOF
    }
    #endregion
    #region CoManagementFlag Enum
    [Flags()] enum CoManagementFlag {
        CompliancePolicies = 2
        ResourceAccessPolicies = 4
        DeviceConfiguration = 8
        WindowsUpdatesPolicies = 16
        ClientApps = 64
        OfficeClickToRunApps = 128
        EndpointProtection = 4128
        CoManagementConfigured = 8193
    }
    #endregion
    
    if (-not (Test-Path $OutDir)) {
        Write-Warning "Cannot find OutDir: $OutDir"
        exit (3)
    }
    
    Write-Host "Configured Workload Configs $($WorkloadConfigs -Join " and ")"
}

Process {
    Write-Host "Generating MOF Files" -ForegroundColor Green
    # Loop all Workloads except CoManagementConfigured
    foreach ( $Workload in [CoManagementFlag].GetEnumNames() | Where-Object { $_ -ne [CoManagementFlag]::CoManagementConfigured } ) {
        
        # Loop all Workload Configs (intune or ConfigMgr)
        Write-Host "Workload $Workload" -ForegroundColor Magenta
        foreach ($WorkloadConfig in $WorkloadConfigs) {
            
            # Build WQL Query and set the CheckBit for the binary representation
            if ($WorkloadConfig -eq "Intune") {
                $CheckBit = "1"
                $WQLWhereQuery1 = "= "
                $WQLWhereQuery2 = ") or `n (ComgmtWorkloads = "
            }
            if ($WorkloadConfig -eq "ConfigMgr") { 
                $CheckBit = "0"
                $WQLWhereQuery1 = "!= "
                $WQLWhereQuery2 = ") and `n (ComgmtWorkloads != "
            }
            $WQLQuery = "Select ComgmtWorkloads from CCM_System Where `n (ComgmtWorkloads "

            # Get all keys from the hashtable except for current Workload
            $keys = [CoManagementFlag].GetEnumNames() | Where-Object { $_ -ne $Workload }
        
            # Create an empty array list to store the sums of each combination
            $combinations = [System.Collections.ArrayList]@()
        
            # Loop over all possible combinations using binary representation of numbers
            for ($i = 0; $i -lt [Math]::Pow(2, $keys.Count); $i++) {
            
                # Convert the current number to binary and pad it with zeros on the left to match the number of keys
                $binaryStr = [Convert]::ToString($i, 2).PadLeft($keys.Count, '0')
            
                # Initialize the sum with the value of the current Workload
                $sum = [CoManagementFlag]::$Workload.value__
            
                # Loop over each character in the binary string and check if it's '1'
                for ($j = 0; $j -lt $binaryStr.Length; $j++) {

                    if ($binaryStr[$j] -eq $CheckBit) {
                        # If it's the CheckBit value 1 or 0, add the corresponding value from the hashtable to the sum
                        $sum += [CoManagementFlag]::$($keys[$j]).value__
                    }
                }
            
                # Add a new property to the array list with a unique name and set its value to the calculated sum 
                $combinations.add($sum) | Out-Null
            }
        
            # Build query for each combination
            $WQLQuery += $WQLWhereQuery1 + ($combinations -Join $WQLWhereQuery2) + ")"
            
            Write-Verbose "$WorkloadConfig Query: `n $WQLQuery"
            
            # Generates the WMI Filer MOF Files
            $filename = "{0}\Co-Mgmt_{1}_{2}.mof" -f $OutDir, $Workload, $WorkloadConfig
            $mof = New-WMIFilterMOF -WQLQuery $WQLQuery -Name "Co-Mgmt $Workload $WorkloadConfig"
            $mof | Out-File -FilePath $filename -Encoding Unicode
            Write-Host "  Created $filename" -ForegroundColor DarkGreen
        }   
    }
}

End {
    Write-Host "All done!!!" -ForegroundColor Green
}
