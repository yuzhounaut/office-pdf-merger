$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot | Split-Path -Parent
$worker = Join-Path $root 'OfficePdfWorker.exe'
$testOutput = Join-Path $root 'tests\output_test'
$officeData = Join-Path $root 'OfficePdfData'

if (Test-Path $testOutput) { Remove-Item $testOutput -Recurse -Force }
New-Item -ItemType Directory -Path $testOutput | Out-Null

$fixturesDir = Join-Path $PSScriptRoot 'fixtures'
$img1 = Join-Path $fixturesDir '03 横图.jpg'
$img2 = Join-Path $fixturesDir '04 透明图片.png'
$pdf1 = Join-Path $fixturesDir '06 原始.pdf'

Write-Host "=== TEST 1: OutputName with dots and no extension ==="
$jobDir1 = Join-Path $officeData ('jobs\test1_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $jobDir1 | Out-Null

$job1 = [pscustomobject]@{
    Items = @(
        [pscustomobject]@{ Path = $img1; Group = '散文件'; State = '待处理' },
        [pscustomobject]@{ Path = $img2; Group = '散文件'; State = '待处理' },
        [pscustomobject]@{ Path = $pdf1; Group = '散文件'; State = '待处理' }
    )
    OutputDirectory = $testOutput
    OutputName = '18.1-杨培茵（CV、GCP、执业证书）'
    PerGroup = $false
    ContinueOnError = $false
    TimeoutSeconds = 60
}
$job1 | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $jobDir1 'job.json') -Encoding UTF8

$proc1 = Start-Process -FilePath $worker -ArgumentList ('--worker ' + [char]34 + $jobDir1 + [char]34) -PassThru -Wait
if ($proc1.ExitCode -ne 0) {
    if (Test-Path (Join-Path $jobDir1 'worker-error.log')) {
        Get-Content (Join-Path $jobDir1 'worker-error.log')
    }
    throw "Worker exited with code $($proc1.ExitCode)"
}

$expectedPdf1 = Join-Path $testOutput '18.1-杨培茵（CV、GCP、执业证书）.pdf'
if (!(Test-Path -LiteralPath $expectedPdf1)) {
    throw "FAIL: Expected PDF not found: $expectedPdf1. Actual files in output: $(Get-ChildItem $testOutput | Select-Object -ExpandProperty Name)"
}
Write-Host "PASS: Generated PDF exists with full name: $expectedPdf1"

# Check output directory has NO manifest or log
$outputFiles = Get-ChildItem -LiteralPath $testOutput
foreach ($f in $outputFiles) {
    if ($f.Name -like '*.清单*' -or $f.Name -like '*.日志*') {
        throw "FAIL: Manifest or log found in output directory: $($f.FullName)"
    }
}
Write-Host "PASS: Output directory contains only output PDF, NO manifest/log."

# Check OfficePdfData has manifest and log
$expectedManifest1 = Join-Path $officeData '18.1-杨培茵（CV、GCP、执业证书）.清单.csv'
$expectedLog1 = Join-Path $officeData '18.1-杨培茵（CV、GCP、执业证书）.日志.txt'

if (!(Test-Path -LiteralPath $expectedManifest1)) {
    throw "FAIL: Expected manifest not found in OfficePdfData: $expectedManifest1"
}
Write-Host "PASS: Manifest found in OfficePdfData: $expectedManifest1"

if (!(Test-Path -LiteralPath $expectedLog1)) {
    throw "FAIL: Expected log not found in OfficePdfData: $expectedLog1"
}
Write-Host "PASS: Log found in OfficePdfData: $expectedLog1"

Write-Host "`n=== TEST 2: Conflict resolution with same name ==="
$jobDir2 = Join-Path $officeData ('jobs\test2_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $jobDir2 | Out-Null

$job2 = [pscustomobject]@{
    Items = @(
        [pscustomobject]@{ Path = $img1; Group = '散文件'; State = '待处理' }
    )
    OutputDirectory = $testOutput
    OutputName = '18.1-杨培茵（CV、GCP、执业证书）.pdf'
    PerGroup = $false
    ContinueOnError = $false
    TimeoutSeconds = 60
}
$job2 | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $jobDir2 'job.json') -Encoding UTF8

$proc2 = Start-Process -FilePath $worker -ArgumentList ('--worker ' + [char]34 + $jobDir2 + [char]34) -PassThru -Wait
if ($proc2.ExitCode -ne 0) {
    throw "Worker exited with code $($proc2.ExitCode)"
}

$expectedPdf2 = Join-Path $testOutput '18.1-杨培茵（CV、GCP、执业证书） (2).pdf'
if (!(Test-Path -LiteralPath $expectedPdf2)) {
    throw "FAIL: Conflict PDF not found: $expectedPdf2. Actual files in output: $(Get-ChildItem $testOutput | Select-Object -ExpandProperty Name)"
}
Write-Host "PASS: Duplicate PDF named cleanly: $expectedPdf2"

$expectedManifest2 = Join-Path $officeData '18.1-杨培茵（CV、GCP、执业证书） (2).清单.csv'
$expectedLog2 = Join-Path $officeData '18.1-杨培茵（CV、GCP、执业证书） (2).日志.txt'

if (!(Test-Path -LiteralPath $expectedManifest2)) {
    throw "FAIL: Expected conflict manifest not found: $expectedManifest2"
}
Write-Host "PASS: Conflict manifest found in OfficePdfData: $expectedManifest2"

if (!(Test-Path -LiteralPath $expectedLog2)) {
    throw "FAIL: Expected conflict log not found: $expectedLog2"
}
Write-Host "PASS: Conflict log found in OfficePdfData: $expectedLog2"

# Clean up test output
Remove-Item $testOutput -Recurse -Force
Write-Host "`nALL TESTS PASSED SUCCESSFULLY!"
