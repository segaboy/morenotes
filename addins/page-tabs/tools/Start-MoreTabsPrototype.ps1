[CmdletBinding()]
param(
    [ValidateSet('Build', 'Measure', 'Placement', 'Com', 'Bind', 'Run', 'Verify', 'HandsOn', 'SelfTest')]
    [string]$Mode = 'Measure',
    [long]$WindowHandle
)
$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $PSScriptRoot
$source = Join-Path $project 'prototype'
$output = Join-Path $project 'artifacts\prototype'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$pia = Join-Path $env:WINDIR 'assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.OneNote.dll'
$office = Join-Path $env:WINDIR 'assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll'
if (-not (Test-Path -LiteralPath $pia)) { throw 'The OneNote 15 desktop interop assembly is required for this prototype.' }
$sources = Get-ChildItem -LiteralPath $source -Filter '*.cs' | Sort-Object Name | Select-Object -ExpandProperty FullName
$build = New-Item -ItemType Directory -Path (Join-Path $output ([Guid]::NewGuid().ToString('N'))) -Force
$exe = Join-Path $build.FullName 'MoreTabsPrototype.exe'
try {
    & $compiler /nologo /target:exe /platform:x64 /warnaserror+ "/out:$exe" "/reference:$pia" "/reference:$office" /reference:System.Drawing.dll /reference:System.Windows.Forms.dll /reference:System.Xml.Linq.dll $sources
    if ($LASTEXITCODE -ne 0) { throw "MoreTabs build failed with exit code $LASTEXITCODE" }
    if ($Mode -eq 'Build') {
        $pageMenu = & (Join-Path $PSScriptRoot 'Set-MoreTabsPageMenu.ps1') -Mode Build
        $binaryName = 'MoreTabsPrototype.exe'
        try {
            [IO.File]::Copy($exe, (Join-Path $output $binaryName), $true)
        } catch [IO.IOException] {
            if (($_.Exception.HResult -band 0xffff) -ne 32) { throw }
            # Let the running prototype release its OneNote row when the user closes it.
            $binaryName = 'MoreTabsPrototype-next.exe'
            [IO.File]::Copy($exe, (Join-Path $output $binaryName), $true)
        }
        @('@echo off', ('"%~dp0{0}" --register' -f $pageMenu.Substring($output.Length + 1)),
            'if errorlevel 1 echo MoreTabs could not register the OneNote page menu for this user.',
            ('"%~dp0{0}" Run' -f $binaryName), 'if errorlevel 1 pause') |
            Set-Content -LiteralPath (Join-Path $output 'Run-MoreTabs.cmd') -Encoding Ascii
        Write-Output (Join-Path $output $binaryName)
        return
    }
    if ($Mode -eq 'Run') {
        try { & (Join-Path $PSScriptRoot 'Set-MoreTabsPageMenu.ps1') -Mode Enable }
        catch { Write-Warning "The tab strip will start without the native page menu. $($_.Exception.Message)" }
    }
    if ($PSBoundParameters.ContainsKey('WindowHandle')) {
        if ($Mode -notin 'Placement', 'Bind', 'Run', 'Verify', 'HandsOn') { throw 'WindowHandle applies to Placement, Bind, Run, Verify and HandsOn. Use a frame number reported by Measure.' }
        if ($WindowHandle -le 0) { throw 'WindowHandle must be a positive frame number reported by Measure.' }
        & $exe $Mode $WindowHandle
    } else {
        & $exe $Mode
    }
    if ($LASTEXITCODE -ne 0) { throw "MoreTabs $Mode failed with exit code $LASTEXITCODE" }
} finally {
    if (Test-Path -LiteralPath $exe) { Remove-Item -LiteralPath $exe }
    [IO.Directory]::Delete($build.FullName, $false)
}
