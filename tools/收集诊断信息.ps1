param([Parameter(Mandatory=$true)][string]$ProgramDirectory, [string]$ReportDirectory = [Environment]::GetFolderPath('Desktop'))
$ErrorActionPreference = 'Stop'
$OutputEncoding = [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$directory = [IO.Path]::GetFullPath($ProgramDirectory).TrimEnd('\')
if (!(Test-Path -LiteralPath $directory -PathType Container)) { throw '指定的软件目录不存在。无需恢复已被隔离的文件。' }
$prefix = $directory + '\'
$report = [ordered]@{CapturedAt=[DateTime]::Now.ToString('o');ProgramDirectory=$directory;Notes='只读诊断；不执行目标文件、不修改安全设置、不上传。'}
try { $report.Defender = Get-MpComputerStatus | Select-Object AMServiceEnabled,AntivirusEnabled,RealTimeProtectionEnabled,BehaviorMonitorEnabled,AntivirusSignatureVersion,AntivirusSignatureLastUpdated,AMEngineVersion,AMProductVersion } catch { $report.DefenderError=$_.Exception.Message }
$report.Files = @(Get-ChildItem -LiteralPath $directory -File | Where-Object { $_.Extension -eq '.exe' } | ForEach-Object {
    try { [pscustomobject]@{Name=$_.Name;Length=$_.Length;SHA256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash;SignatureStatus=[string](Get-AuthenticodeSignature -LiteralPath $_.FullName).Status} }
    catch { [pscustomobject]@{Name=$_.Name;Error=$_.Exception.Message} }
})
try { $report.Detections = @(Get-MpThreatDetection | Where-Object { $resources=$_.Resources -join ' '; $resources.IndexOf($directory,[StringComparison]::OrdinalIgnoreCase) -ge 0 } | Select-Object InitialDetectionTime,LastThreatStatusChangeTime,ThreatID,ActionSuccess,Resources,ProcessName) } catch { $report.DetectionError=$_.Exception.Message }
try { $report.RelatedTasks = @(Get-ScheduledTask | Where-Object { $_.TaskName -like 'OfficePdfMerger-*' -and @($_.Actions | Where-Object { $_.Execute -and $_.Execute.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase) }).Count -gt 0 } | Select-Object TaskName,State,@{N='Actions';E={@($_.Actions | Select-Object Execute,Arguments)}}) } catch { $report.TaskError=$_.Exception.Message }
try { $report.RelatedProcesses = @(Get-CimInstance Win32_Process | Where-Object { $_.ExecutablePath -and $_.ExecutablePath.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase) } | Select-Object Name,ProcessId,ParentProcessId,ExecutablePath,CommandLine) } catch { $report.ProcessError=$_.Exception.Message }
$destination = Join-Path $ReportDirectory ('OfficePDF诊断_' + [DateTime]::Now.ToString('yyyyMMdd_HHmmss') + '.json')
$report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $destination -Encoding UTF8
Write-Host ('报告已保存：' + $destination)
Write-Host '报告可能含本机路径、账户名及文件名；分享前请检查。'
