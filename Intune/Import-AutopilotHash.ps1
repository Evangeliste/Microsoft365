#########################################################################################################################
#
#		Script for : Import-AutoPilotHash.ps1
#
#		Script author : Alexis Charrière / Evangeliste
#		mail 	: alexis@inventiq.fr
#		Create 	: 18 Avril 2024
#
#		Version	: 1.1
#		UPDate 	: 1.1 : 2024-07-10 Add variable for GroupTAG
#
#########################################################################################################################

$env:Path += ";C:\Program Files\WindowsPowerShell\Scripts"
$TenantID = "ff3615ac-20fe-4391-a318-72fe3e3cc023"
$AppID = "0818960b-1532-4389-a373-903913652d21"
$SecretAppID = "avJ8Q~40oAw1hB4oEZD8oNOX_YVUqd5L5C.2IcR~"
$GroupTAG = 'IVQ_AutoP_Common'

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$False
Install-Module -Name PowerShellGet -Force -AllowClobber -Confirm:$False

Install-Script -Name Get-WindowsAutopilotInfo -Force

Get-WindowsAutoPilotInfo.ps1 -Online -TenantId $TenantID -AppId $AppID -AppSecret $SecretAppID -GroupTag $GroupTAG