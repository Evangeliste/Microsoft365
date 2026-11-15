#########################################################################################################################
#
#		Script for : Install PowerShell module
#
#		Script author : Alexis Charrière / Evangeliste
#		mail 	: alexis@inventiq.fr
#		Create 	: 02 Mai 2024
#
#		Version	: 1.0
#		UPDate 	:
#
#########################################################################################################################
<#
.DESCRIPTION

    This script install PowerShell Module for Entra ID / M365 Administrations

    [pscustomobject]@{Module='Microsoft365DSC'                                  ; Description='Microsoft 365 DSC PowerShell Module'}
    [pscustomobject]@{Module='ORCAPreview'                                      ; Description='Microsoft Defender for Office 365 PowerShell Module'}

    [pscustomobject]@{Module='Microsoft.Graph.Teams'                            ; Description='Microsoft Graph Teams PowerShell Module'} ### Installed with Microsoft.Graph
    [pscustomobject]@{Module='Microsoft.Graph.Devices.ServiceAnnouncement'      ; Description= 'Microsoft Graph Device Management PowerShell Module'} ### Installed with Microsoft.Graph
    [pscustomobject]@{Module='Microsoft.Graph.Intune'                           ; Description='Microsoft Graph Intune PowerShell Module'} ### Installed with Microsoft.Graph

    [pscustomobject]@{Module='WindowsAutoPilotIntune'                           ; Description='Intune Windows Autopilot PowerShell Module'}

    [pscustomobject]@{Module='AIPService'                                       ; Description='Azure Information Protection PowerShell Module'}
    [pscustomobject]@{Module='UnifiedLabelingSupportTool'                       ; Description='Unified Labeling Support Tool PowerShell Module'}

    [pscustomobject]@{Module='ZeroTrustAssessment'                              ; Description='Perform a Zero Trust Assessment of your Microsoft 365 environment'}
    [pscustomobject]@{Module='AzureADAssessment'                                ; Description='Perform a Azure AD Assessment'}
    [pscustomobject]@{Module='MSCloudIdAssessment'                              ; Description='To be added soon'}
#>

#########################################################################################################################
#
#   Gestion des variables et logs
#
#########################################################################################################################

$MyDate = Get-Date -Format yyyy-MM-dd
$MyDateTime = "$($((Get-Date).ToString('yyyy-MM-dd_HH-mm')))"

$WorkPath = Test-Path -Path C:\_Admin
    IF ($WorkPath -eq $False)
        {
        New-Item -Path "C:\" -Name "_Admin" -ItemType Directory -Confirm:$False
        }

$WorkDirectory = "C:\_Admin"
$FolderBuild = $WorkDirectory+"\LOG\"+$MyDate

$TranscriptLog = $FolderBuild+"\"+$MyDateTime+"_InstallStatus_TraceRun.log"
Start-Transcript -Path $TranscriptLog

$ExecutionContext.SessionState.LanguageMode = 'FullLanguage'

#[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
#Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -Confirm:$False

New-Item $FolderBuild -Type Directory -Confirm:$False -ErrorAction SilentlyContinue
$LogFile = $WorkDirectory+"\LOG\"+$MyDate+"\"+$MyDateTime+"_InstallStatus_Modules.log"

"Date;Module;Version;Status" | Tee-Object -FilePath $LogFile -Append

#########################################################################################################################
#
#   Start
#
#########################################################################################################################

#
# Install / UPDate Package Providers
#

Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$False
Install-Module -Name PowerShellGet -Force -AllowClobber -Confirm:$False

#
# Add module in list bellow for deploying module
#

$Modules = @(
    [pscustomobject]@{Module='Microsoft.Graph'                          ; Description='Microsoft Graph PowerShell Module'}
    [pscustomobject]@{Module='Microsoft.Graph.Beta'                     ; Description='Microsoft Graph Beta PowerShell Module'}
#    [pscustomobject]@{Module='Msonline'                                 ; Description='Legacy Office 365 Management (MSOnline)'}    ### Soon Deprecated ###
#    [pscustomobject]@{Module='AzureAD'                                  ; Description='Legacy Azure AD PowerShell Module'}          ### Soon Deprecated ###
    [pscustomobject]@{Module='MicrosoftTeams'                           ; Description='Microsoft Teams PowerShell Module'}
    [pscustomobject]@{Module='ExchangeOnlineManagement'                 ; Description='Exchange Online PowerShell Module'}
#    [pscustomobject]@{Module='Az'                                       ; Description='Microsoft Azure PowerShell Module'}
    [pscustomobject]@{Module='Microsoft.Online.SharePoint.PowerShell'   ; Description='Microsoft SharePoint Online Services Module for Windows PowerShell'}

    [pscustomobject]@{Module='Microsoft365DSC'                          ; Description='Microsoft 365 DSC PowerShell Module'}
    [pscustomobject]@{Module='ORCAPreview'                              ; Description='Microsoft Defender for Office 365 PowerShell Module'}

    [pscustomobject]@{Module='Microsoft.Graph.Teams'                            ; Description='Microsoft Graph Teams PowerShell Module'} ### Installed with Microsoft.Graph
    [pscustomobject]@{Module='Microsoft.Graph.Devices.ServiceAnnouncement'      ; Description= 'Microsoft Graph Device Management PowerShell Module'} ### Installed with Microsoft.Graph
    [pscustomobject]@{Module='Microsoft.Graph.Intune'                           ; Description='Microsoft Graph Intune PowerShell Module'} ### Installed with Microsoft.Graph

    [pscustomobject]@{Module='WindowsAutoPilotIntune'                   ; Description='Intune Windows Autopilot PowerShell Module'}

    [pscustomobject]@{Module='AIPService'                               ; Description='Azure Information Protection PowerShell Module'}
    [pscustomobject]@{Module='UnifiedLabelingSupportTool'               ; Description='Unified Labeling Support Tool PowerShell Module'}

    [pscustomobject]@{Module='ZeroTrustAssessment'                      ; Description='Perform a Zero Trust Assessment of your Microsoft 365 environment'}
    [pscustomobject]@{Module='AzureADAssessment'                        ; Description='Perform a Azure AD Assessment'}
    [pscustomobject]@{Module='MSCloudIdAssessment'                      ; Description='To be added soon'}
)

#
# Clean PowerShellModules Load
#

$ModulesLoad = Get-Module | Where-Object {($_.Name -ne "Microsoft.PowerShell.Management") -and ($_.Name -ne "PSReadLine")}
Foreach ($ModuleLoad in $ModulesLoad.Name) 
    {
    Remove-Module -Name $ModuleLoad -Force -Confirm:$False
    }

Get-Module | Write-Host -ForegroundColor Green

foreach ($Module in $Modules.Module)
    {
    $MaximumFunctionCount = 16320
    $MaximumVariableCount = 16320
    #Check action for module (UPDate or Install)
    $TestModule = Get-InstalledModule -Name $Module -ErrorAction SilentlyContinue
    IF ($null -eq $TestModule)
        {
        Write-Host $Module "Installation InProgress" -ForegroundColor Magenta
        Install-Module -Name $Module -AllowClobber -Scope AllUsers -Force -Confirm:$False -ErrorAction SilentlyContinue -ErrorVariable InstallFailed
        IF ($InstallFailed)
            {
            Write-Warning "Failed to install Powershell Module" $Module
            "$($MyDate);" + $Module + ";" + $Module.version + ";" +"Installation Failed" | Tee-Object -FilePath $LogFile -Append
            }
        }
    ELSE
        {
        Write-Host $Module "UPGrade InProgress" -ForegroundColor Cyan
        Update-Module -Name $Module -Force -Confirm:$False -ErrorAction SilentlyContinue -ErrorVariable UPDateFailed
        Remove-Module -Name $Module -Confirm:$false -Force -ErrorAction SilentlyContinue
        IF ($UPDateFailed)
            {
            Write-Warning "Failed to UPDate Powershell Module" $Module
            "$($MyDate);" + $Module + ";" + $Module.version + ";" +"UPDate Failed" | Tee-Object -FilePath $LogFile -Append
            }
        }

    #Check install
    $Check = Get-InstalledModule -Name $Module -ErrorAction SilentlyContinue
    IF ($null -ne $Check)
        {
        Write-Host $Check.Name, $Check.version "Correctly Installed" -ForegroundColor Green
        "$($MyDate);" + $Check.Name + ";" + $Check.Version + ";" + "Installation Success" | Tee-Object -FilePath $LogFile -Append
        }
    }

Get-InstalledModule | Select-Object Name,Version,Type,Repository,InstalledLocation | ft -AutoSize

Stop-Transcript