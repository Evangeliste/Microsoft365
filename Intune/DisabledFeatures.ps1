



Disable-WindowsOptionalFeature -FeatureName Microsoft-Hyper-V,Microsoft-Windows-Subsystem-Linux,TelnetClient -Online -ErrorVariable ErrorDisabled
IF ($ErrorDisabled)
    {
    Write-Warning "Error during disabling features"
    }
else 
    {
    New-Item "C:\IVQ\Logs\DisabledFeaturesOK.txt" -Type file | Out-Null
    }