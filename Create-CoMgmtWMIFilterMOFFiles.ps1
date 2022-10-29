<#
  .SYNOPSIS
  Creates WMI Filter MOF Files for specific Co-Management workloads and to what they are set.
  .DESCRIPTION
  Creates two WMI Filter MOF files per workload in the current directory subfolder out
  or in the directory specified with the -OutDir Param.
  One WMI Filter for when the workload is on Intune and one if it's on ConfigMgr.
#>

[CmdletBinding()]
Param (
    # Output Directory of the MOF Files, excluding trailing "\""
    [Parameter()]
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
        CoManagementConfigured = 1
        CompliancePolicies = 3
        ResourceAccessPolicies = 5
        DeviceConfiguration = 9
        WindowsUpdatesPolicies = 17
        EndpointProtection = 33
        ClientApps = 65
        OfficeClickToRunApps = 129
    }
    #endregion
    write-host "Configured Workload Configs $($WorkloadConfigs -Join " and ")"

    if (-not (Test-Path $OutDir)) {
        Write-Warning "cannot find OutDir $OutDir"
        exit (3)
    }
}

Process {
    Write-Host "Generating MOF Files" -ForegroundColor Green
    # Loop all Workloads except CoManagementConfigured
    foreach ( $Workload in [CoManagementFlag].GetEnumNames() | Where-Object{$_ -ne [CoManagementFlag]::CoManagementConfigured} ) {

        Write-Host "Workload $Workload" -ForegroundColor Magenta
        foreach ($WorkloadConfig in $WorkloadConfigs) {

            # Gets all the Decimals where the current workload is flagged
            $WorkloadNumbers = (1..255).where({ (($_ -band [CoManagementFlag]::$Workload) -eq [CoManagementFlag]::$Workload) })
            
            if ($WorkloadConfig -eq "ConfigMgr") {
                # Builds the wql where part, (NOT IN), Filter would match if workload is set to ConfigMgr
                $WQLQuery = "Select ComgmtWorkloads from CCM_System Where `n (ComgmtWorkloads != " `
                    + ($WorkloadNumbers -Join (") and `n (ComgmtWorkloads != "))  `
                    + ")"
            }
            if ($WorkloadConfig -eq "Intune") {
                # Builds the wql where part, (IN), Filter would match if workload is set to Intune
                $WQLQuery = "Select ComgmtWorkloads from CCM_System Where `n (ComgmtWorkloads = " `
                    + ($WorkloadNumbers -Join (") or `n (ComgmtWorkloads = "))  `
                    + ")"
            }
            
            # Generates the WMI Filer MOF Files
            $filename = "{0}\Co-Mgmt_{1}_{2}.mof" -f $OutDir, $Workload, $WorkloadConfig
            $mof = New-WMIFilterMOF -WQLQuery $WQLQuery -Name "Co-Mgmt $Workload $WorkloadConfig"
            $mof | out-file -FilePath $filename -Encoding Unicode
            Write-Host "  Created $filename" -ForegroundColor DarkGreen
        }   
    }
}

End {
    Write-Host "All Done!!!" -ForegroundColor Green
}
