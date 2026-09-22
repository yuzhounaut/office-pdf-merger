$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$wpf = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\WPF'
$renderer = Join-Path $PSScriptRoot 'RenderIcon.exe'
& $compiler /nologo /target:exe "/out:$renderer" "/reference:$wpf\PresentationCore.dll" "/reference:$wpf\WindowsBase.dll" /reference:System.Xaml.dll (Join-Path $PSScriptRoot 'RenderIcon.cs')
if ($LASTEXITCODE -ne 0) { throw 'Icon renderer compilation failed' }
& $renderer (Join-Path $root 'assets\OfficePDF.svg') (Join-Path $root 'assets\OfficePDF-master.png') 1024
if ($LASTEXITCODE -ne 0) { throw 'SVG rendering failed' }
& 'C:\ProgramData\anaconda3\python.exe' -X utf8 (Join-Path $PSScriptRoot 'make_icon.py')
if ($LASTEXITCODE -ne 0) { throw 'ICO generation failed' }
