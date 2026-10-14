$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$taskCompilerPath = 'C:\Program Files\AutoHotkey\Compiler\Ahk2Exe.exe'
function Build-SapArtifact($sourceName, $outputName, $baseName, $machine) {
    $taskArguments = '/in "' + $sourceName + '" /out "' + $outputName + '" /base "C:\Program Files\AutoHotkey\v2\' + $baseName + '" /silent verbose'
    if ($sourceName -eq 'Main.ahk') { $taskArguments += ' /icon "ui\assets\icon.ico"' }
    $taskProcess = Start-Process -FilePath $taskCompilerPath -ArgumentList $taskArguments -WindowStyle Hidden -PassThru -RedirectStandardOutput 'release-build.log' -RedirectStandardError 'release-build-error.log'
    if (!$taskProcess.WaitForExit(60000)) { throw "Build timed out: $sourceName" }
    $taskProcess.Refresh()
    if ($taskProcess.ExitCode -ne 0) { throw "Build failed: $sourceName (see release-build-error.log)" }
    $taskBytes = [IO.File]::ReadAllBytes((Join-Path $PSScriptRoot $outputName))
    $taskPeOffset = [BitConverter]::ToInt32($taskBytes, 60)
    if ([BitConverter]::ToUInt16($taskBytes, $taskPeOffset + 4) -ne $machine) { throw "Incorrect architecture: $outputName" }
}
if (Get-Process SAPRFC -ErrorAction SilentlyContinue) { throw 'Close active SAP RFC queries before building.' }
Build-SapArtifact 'SAPRFC.ahk' 'SAPRFC.build.exe' 'AutoHotkey32.exe' 0x14c
Copy-Item -LiteralPath 'SAPRFC.build.exe' -Destination 'SAPRFC.exe' -Force
$taskRequest = Join-Path $env:TEMP ('sap-release-version-' + [guid]::NewGuid().ToString('N') + '.json')
$taskResponse = $taskRequest + '.result'
try {
    [IO.File]::WriteAllText($taskRequest, '{"protocol":1,"command":"Version"}', [Text.UTF8Encoding]::new($false))
    $taskProcess = Start-Process -FilePath (Join-Path $PSScriptRoot 'SAPRFC.exe') -ArgumentList ('"' + $taskRequest + '" "' + $taskResponse + '"') -WindowStyle Hidden -PassThru
    if (!$taskProcess.WaitForExit(10000)) { throw 'RFC version check timed out.' }
    $taskProcess.Refresh()
    $taskResult = Get-Content -LiteralPath $taskResponse -Raw | ConvertFrom-Json
    if ($taskProcess.ExitCode -ne 0 -or !$taskResult.ok -or $taskResult.protocol -ne 1) { throw 'RFC version check failed.' }
} finally {
    foreach ($taskPath in @($taskRequest, $taskResponse)) {
        if (Test-Path -LiteralPath $taskPath) { Remove-Item -LiteralPath $taskPath }
    }
}
Build-SapArtifact 'Main.ahk' 'Main.release.exe' 'AutoHotkey64.exe' 0x8664
$taskFolder = Join-Path $env:TEMP ('sap-package-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $taskFolder | Out-Null
Copy-Item -LiteralPath 'Main.release.exe' -Destination (Join-Path $taskFolder 'Main.exe')
$taskProcess = Start-Process -FilePath (Join-Path $taskFolder 'Main.exe') -ArgumentList '--verify-sap-package' -WorkingDirectory $taskFolder -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $taskFolder 'package.log') -RedirectStandardError (Join-Path $taskFolder 'package-error.log')
if (!$taskProcess.WaitForExit(20000)) { throw "Package test timed out: $taskFolder" }
$taskProcess.Refresh()
if ($taskProcess.ExitCode -ne 0 -or (Get-FileHash -LiteralPath 'SAPRFC.exe').Hash -ne (Get-FileHash -LiteralPath (Join-Path $taskFolder 'SAPRFC.exe')).Hash) { throw "Package extraction failed: $taskFolder" }
Copy-Item -LiteralPath 'Main.release.exe' -Destination 'Main.exe' -Force
if ((Get-FileHash -LiteralPath 'Main.release.exe').Hash -ne (Get-FileHash -LiteralPath 'Main.exe').Hash) { throw 'Main.exe copy verification failed.' }
Remove-Item -LiteralPath 'SAPRFC.build.exe', 'Main.release.exe'
Write-Output "Built and verified: Main.exe (x64) + embedded SAPRFC.exe (x86), release $($taskResult.version)."
