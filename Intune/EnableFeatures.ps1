



Enable-WindowsOptionalFeature -FeatureName Microsoft-Hyper-V,Microsoft-Windows-Subsystem-Linux,TelnetClient -Online -All -ErrorVariable ErrorEnabled
IF ($ErrorEnabled)
    {
    Write-Warning "Error during deploiement"
    }
else 
    {
    New-Item "C:\IVQ\Logs\EnableFeaturesOK.txt" -Type file | Out-Null
    }