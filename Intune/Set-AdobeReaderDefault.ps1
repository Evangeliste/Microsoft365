<#

https://helpx.adobe.com/acrobat/kb/disable-prompt-to-make-acrobat-my-default.html
https://community.adobe.com/t5/acrobat-reader/pop-up-asking-to-make-adobe-reader-my-default-pdf-application/td-p/11339052?page=1

[HKEY_CURRENT_USER\SOFTWARE\Adobe\Acrobat Reader\DC\AVAlert\cCheckbox]
"iAppDoNotTakePDFOwnershipAtLaunchWin10"=dword:00000001


#>

$RegPath = "HKCU:\Software\Adobe\Acrobat Reader\DC\AVAlert\cCheckBox"
$RegPath2 = "HKCU:\Software\Adobe\Acrobat Reader\DC"

#Accent Color Menu Key
$AdobeDefaultKey = @{
	Key   = 'iAppDoNotTakePDFOwnershipAtLaunchWin10';
	Type  = "DWORD";
	Value = '1'
}

If ($Null -eq (Get-ItemProperty -Path $RegPath -Name $AdobeDefaultKey.Key -ErrorAction SilentlyContinue))
    {
    New-Item -Path $RegPath2 -Name AVAlert
    New-Item -Path $RegPath2\AVAlert -Name cCheckBox
    New-ItemProperty -Path $RegPath -Name $AdobeDefaultKey.Key -Value $AdobeDefaultKey.Value -PropertyType $AdobeDefaultKey.Type -Force
    }
Else
    {
    Set-ItemProperty -Path $RegPath -Name $AdobeDefaultKey.Key -Value $AdobeDefaultKey.Value -Force
    }          
