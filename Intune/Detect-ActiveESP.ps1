#########################################################################################################################
#
#		Script for : Detect-ActiveESP.ps1
#
#		Script author : Alexis Charrière / Evangeliste
#		mail 	: alexis@inventiq.fr
#		Create 	: 18 Avril 2024
#
#		Version	: 1.1
#		UPDate 	: 2024-07-09 - Change detection rules to Defender System Tray
#
#########################################################################################################################


#########################################################################################################################
#
#   Start
#
#########################################################################################################################

#$proc = Get-Process -Name WWAHost -ErrorAction SilentlyContinue
$proc = Get-Process -Name SecurityHealthSystray -ErrorAction SilentlyContinue

# Process not found, so ESP is not running
if ($proc -eq $null) 
    {
    Write-Output "ESP-active"
    exit 0
    } 
else 
    {
    Write-Output "ESP-not-active"
    exit 0
    }
