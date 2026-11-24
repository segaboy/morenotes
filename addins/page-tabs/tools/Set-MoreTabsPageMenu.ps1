[CmdletBinding()]
param([ValidateSet('Build','Enable','Remove','Status')][string]$Mode = 'Enable')
$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $PSScriptRoot
$source = Join-Path $project 'page-menu\NativePageMenu.cs'
$setup = Join-Path $project 'page-menu\PageMenuSetup.cs'
$protocol = Join-Path $project 'prototype\PageMenuCommand.cs'

# Keep a running COM server available until OneNote releases it.
$version = ((@($source, $protocol, $setup) | ForEach-Object { (Get-FileHash $_).Hash.Substring(0,16) }) -join '')
$folder = Join-Path $project "artifacts\prototype\page-menu\$version"
$exe = Join-Path $folder 'MoreTabs.PageMenu.Setup.exe'
if (-not [IO.File]::Exists($exe)) {
    $staging = Join-Path (Split-Path $folder) ([Guid]::NewGuid().ToString('N'))
    [IO.Directory]::CreateDirectory($staging) | Out-Null
    $dll = Join-Path $staging 'MoreTabs.PageMenu.dll'
    $setupExe = Join-Path $staging 'MoreTabs.PageMenu.Setup.exe'
    $gac = Join-Path $env:WINDIR 'assembly'
    $office = Join-Path $gac 'GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll'
    $one = Join-Path $gac 'GAC_MSIL\Microsoft.Office.Interop.OneNote\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.OneNote.dll'
    $ext = Join-Path $gac 'GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\Extensibility.dll'
    try {
        $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
        & $csc /nologo /warnaserror+ /target:library /platform:x64 "/out:$dll" "/reference:$office" "/reference:$one" "/reference:$ext" /reference:System.Xml.Linq.dll /reference:System.Windows.Forms.dll $source $protocol
        if ($LASTEXITCODE -ne 0) { throw 'MoreTabs page menu build failed.' }
        & $csc /nologo /warnaserror+ /target:exe /platform:x64 "/out:$setupExe" "/reference:$dll" "/reference:$office" "/reference:$ext" $setup
        if ($LASTEXITCODE -ne 0) { throw 'MoreTabs page menu setup build failed.' }
        try { [IO.Directory]::Move($staging, $folder) }
        catch [IO.IOException] { if (-not [IO.File]::Exists($exe)) { throw } }
    } finally {
        if ([IO.Directory]::Exists($staging)) {
            [IO.File]::Delete($dll)
            [IO.File]::Delete($setupExe)
            [IO.Directory]::Delete($staging)
        }
    }
}
if ($Mode -eq 'Build') { Write-Output $exe; return }
$operation = @{ Remove = '--unregister'; Status = '--status'; Enable = '--register' }[$Mode]
& $exe $operation | Out-Host
if ($LASTEXITCODE -ne 0) { throw "Could not $operation the MoreTabs page menu for the current user." }
