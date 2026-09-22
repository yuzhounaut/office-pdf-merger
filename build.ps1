$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$archive = Join-Path $root 'vendor\qpdf-12.4.1-mingw64.zip'
$expectedHash = '6a47eeddc8ff712a6e003314daae25569402c6e904ba82b7b9181d7b0301b689'

if (!(Test-Path -LiteralPath $archive)) {
    $vendorDir = Join-Path $root 'vendor'
    if (!(Test-Path -LiteralPath $vendorDir)) { New-Item -ItemType Directory -Path $vendorDir | Out-Null }
    Write-Host "Downloading official qpdf 12.4.1 mingw64 archive from GitHub releases..."
    $url = 'https://github.com/qpdf/qpdf/releases/download/v12.4.1/qpdf-12.4.1-mingw64.zip'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $url -OutFile $archive -UseBasicParsing
}

if ((Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash -ne $expectedHash) { throw 'qpdf official archive SHA256 mismatch' }
Add-Type -AssemblyName System.IO.Compression.FileSystem
$source = [IO.Compression.ZipFile]::OpenRead($archive)
try {
    foreach ($entry in $source.Entries) {
        $relative = ($entry.FullName -split '/', 2)[1]
        if ($entry.Name -and ($relative -match '^bin/(qpdf\.exe|.+\.dll)$' -or $relative -match '(?i)(license|notice|copying|copyright)' -or $relative -match '^share/doc/qpdf/README')) {
            $target = [IO.Path]::GetFullPath((Join-Path (Join-Path $root 'qpdf') $relative))
            if (!$target.StartsWith((Join-Path $root 'qpdf') + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe archive entry' }
            [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target)) | Out-Null
            [IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $target, $true)
        }
    }
} finally { $source.Dispose() }
$common = @('Models.cs','Conversion.cs','Engine.cs','RuntimeFiles.cs','AssemblyInfo.cs') | ForEach-Object { Join-Path $root ('src\' + $_) }
$references = @('/reference:System.dll','/reference:System.Core.dll','/reference:System.Management.dll','/reference:System.Drawing.dll','/reference:System.Web.Extensions.dll','/reference:Microsoft.CSharp.dll')
$worker = Join-Path $root 'OfficePdfWorker.exe'
& $compiler /nologo /target:winexe /platform:x64 /optimize+ /main:OfficePdf.WorkerProgram "/out:$worker" "/win32icon:$root\assets\OfficePDF.ico" "/win32manifest:$root\src\app.manifest" @references @common (Join-Path $root 'src\WorkerProgram.cs')
if ($LASTEXITCODE -ne 0) { throw 'Worker build failed' }
$ui = @('Program.cs','Scheduler.cs','ModernWindow.cs','FolderPicker.cs') | ForEach-Object { Join-Path $root ('src\' + $_) }
$exe = Join-Path $root 'Office资料一键合并PDF_v1.4.3.exe'
$wpf = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\WPF'
& $compiler /nologo /target:winexe /platform:x64 /optimize+ /main:OfficePdf.Program "/out:$exe" "/win32icon:$root\assets\OfficePDF.ico" "/win32manifest:$root\src\app.manifest" "/resource:$root\assets\OfficePDF.png,AppIcon.png" "/resource:$root\src\MainWindow.xaml,MainWindow.xaml" "/reference:$wpf\PresentationFramework.dll" "/reference:$wpf\PresentationCore.dll" "/reference:$wpf\WindowsBase.dll" /reference:System.Xaml.dll @references @common @ui
if ($LASTEXITCODE -ne 0) { throw 'GUI build failed' }
Get-FileHash -LiteralPath $exe,$worker -Algorithm SHA256 | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'build-hashes.json') -Encoding UTF8
Get-Item -LiteralPath $exe,$worker | Select-Object FullName,Length
