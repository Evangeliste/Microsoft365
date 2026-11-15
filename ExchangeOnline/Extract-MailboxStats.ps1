#########################################################################################################################
#
#   Extract Mailbox Statistics
#
#########################################################################################################################

#Param/Config
#########################################################################################################################
[CmdletBinding(SupportsShouldProcess = $true)]
param 
    (
    [String]$PFtoApply,
    [String]$LogName
    )

#Variables
#########################################################################################################################
$MyDateTime = "$($((Get-Date).ToString('yyyy-MM-dd_HH-mm')))"
$CurrentDirectory = Get-Location
$TranscriptLog = $CurrentDirectory.Path+"\"+$MyDateTime+"_Trace_"+$LogName+".log"

#Run
#########################################################################################################################
$MBX = Get-Mailbox -ResultSize Unlimited

IF ($MBX)
    {$counter #++
    $Inc = 0
	Foreach ($MB in $MBX.userprincipalname)
		{
		##Progress
		$counter = $counter + 1
		Write-Progress -Activity 'Extracting mailbox statistics...' -CurrentOperation $MB -PercentComplete (($counter / $MBX.count) * 100)
		$Inc = $Inc + 1
		Write-Host "   Working on the mailbox" $Inc "out of" $MBX.count // $MB -ForegroundColor Magenta
		
        #Write-Host "Working on" $MB "Mailbox" -ForegroundColor Green
		Get-Mailbox -identity $MB -ResultSize Unlimited | Get-MailboxStatistics | Select-Object DisplayName,LastLogonTime,StorageLimitStatus,@{name="TotalItemSize_MB"; expression={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}},ItemCount,LegacyDN | Export-CSV "D:\_Talan\All_RATP_Mailboxes.csv" -NoTypeInformation -Encoding UTF8 -Append
		}
	}