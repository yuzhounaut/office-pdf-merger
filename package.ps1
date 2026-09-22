param([switch]$Force)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$target = Join-Path $root '成品'
if (Test-Path -LiteralPath $target) {
    if ($Force) { Remove-Item -LiteralPath $target -Recurse -Force }
    else { throw 'Product directory already exists; inspect first or use -Force.' }
}
New-Item -ItemType Directory -Path $target,(Join-Path $target '图标素材') | Out-Null
foreach ($name in @('Office资料一键合并PDF_v1.4.3.exe','OfficePdfWorker.exe','qpdf','THIRD_PARTY_LICENSES')) {
    Copy-Item -LiteralPath (Join-Path $root $name) -Destination $target -Recurse
}
Copy-Item -LiteralPath (Join-Path $root 'docs\使用说明.md') -Destination (Join-Path $target 'README_使用说明.md')
Copy-Item -LiteralPath (Join-Path $root 'docs\验证报告_v1.4.3.md') -Destination (Join-Path $target '验证结果.md')
Copy-Item -LiteralPath (Join-Path $root 'tools\收集诊断信息.ps1') -Destination (Join-Path $target '收集诊断信息.ps1')
foreach ($name in @('OfficePDF.svg','OfficePDF.ico','OfficePDF.png')) { Copy-Item -LiteralPath (Join-Path $root ('assets\'+$name)) -Destination (Join-Path $target '图标素材') }
$hashes = @(Get-ChildItem -LiteralPath $target -File -Recurse | Sort-Object FullName | ForEach-Object {
    [pscustomobject]@{File=$_.FullName.Substring($target.Length+1);Bytes=$_.Length;SHA256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()}
})
$hashes | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $target 'SHA256.json') -Encoding UTF8
$archive = Join-Path (Split-Path $root -Parent) 'OfficePDF_v1.4.3_Windows11_x64.zip'
if (Test-Path -LiteralPath $archive) {
    if ($Force) { Remove-Item -LiteralPath $archive -Force }
    else { throw 'Archive already exists; inspect first or use -Force.' }
}
Add-Type -AssemblyName System.IO.Compression.FileSystem
[IO.Compression.ZipFile]::CreateFromDirectory($target,$archive,[IO.Compression.CompressionLevel]::Optimal,$false)
Get-FileHash -LiteralPath $archive -Algorithm SHA256 | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root '交付包SHA256.json') -Encoding UTF8
Get-Item -LiteralPath $archive | Select-Object FullName,Length
