<#PSScriptInfo .VERSION 0.0.1#>

using assembly PresentationFramework
using namespace System.Windows

[MessageBox]::Show($args[0].Substring(1).Remove($args[0].Length - 2).Trim(), 'Convert to HTML', $args[1], $args[2])