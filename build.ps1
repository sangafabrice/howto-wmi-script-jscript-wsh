<#PSScriptInfo .VERSION 1.0.2#>

[CmdletBinding()]
param ()

& {
  Import-Module "$PSScriptRoot\tools"
  Format-ProjectCode @('*.js','*.ps*1','.gitignore'| ForEach-Object { "$PSScriptRoot\$_" })
  Set-ProjectVersion $PSScriptRoot
  Remove-Module tools
}