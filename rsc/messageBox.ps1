<#PSScriptInfo .VERSION 0.0.2#>

using namespace System.Windows

Add-Type -AssemblyName PresentationFramework
[MessageBox]::Show($args[0].Remove($args[0].Length - 1).Trim(), 'Convert to HTML', $args[1], $args[2])