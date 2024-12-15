<#PSScriptInfo .VERSION 1.0.1#>

using namespace System.IO
using namespace System.Runtime.InteropServices
[CmdletBinding()]
param ()

& {
  Import-Module "$PSScriptRoot\tools"
  Format-ProjectCode @('*.js','*.ps*1','.gitignore'| ForEach-Object { "$PSScriptRoot\$_" })
  Set-ProjectVersion $PSScriptRoot
  Remove-Module tools
  $shell = New-Object -ComObject 'WScript.Shell'
  $shell.CreateShortcut("$PSScriptRoot\MsgBox.lnk") | ForEach-Object {
    $_.TargetPath = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
    $_.Arguments = '-File "{0}"' -f ([Path]::Combine($PSScriptRoot, 'rsc\messageBox.ps1'))
    $_.IconLocation = [Path]::Combine($PSScriptRoot, 'rsc\menu.ico')
    $_.Save()
    [void][Marshal]::FinalReleaseComObject($_)
    $_ = $null
  }
  [void][Marshal]::FinalReleaseComObject($shell)
  $shell = $null
  [GC]::Collect()
}