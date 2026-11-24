[CmdletBinding()]
param([ValidateSet('Build','Enable','Remove')][string]$Mode = 'Enable')
$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $PSScriptRoot
$source = Join-Path $project 'page-menu\NativePageMenu.cs'
$protocol = Join-Path $project 'prototype\PageMenuCommand.cs'

# Keep a running COM server available until OneNote releases it.
$version = (Get-FileHash $source).Hash.Substring(0,16) + (Get-FileHash $protocol).Hash.Substring(0,16)
$folder = Join-Path $project "artifacts\prototype\page-menu\$version"
$exe = Join-Path $folder 'MoreTabs.PageMenu.exe'
if (-not [IO.File]::Exists($exe)) {
    [IO.Directory]::CreateDirectory($folder) | Out-Null
    $temporary = Join-Path $folder ([Guid]::NewGuid().ToString('N') + '.exe')
    $gac = Join-Path $env:WINDIR 'assembly'
    $office = Join-Path $gac 'GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll'
    $one = Join-Path $gac 'GAC_MSIL\Microsoft.Office.Interop.OneNote\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.OneNote.dll'
    $ext = Join-Path $gac 'GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\Extensibility.dll'
    try {
        & (Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe') /nologo /warnaserror+ /target:winexe /platform:x64 "/out:$temporary" "/reference:$office" "/reference:$one" "/reference:$ext" /reference:System.Xml.Linq.dll /reference:System.Windows.Forms.dll $source $protocol
        if ($LASTEXITCODE -ne 0) { throw 'MoreTabs page menu build failed.' }
        try { [IO.File]::Move($temporary, $exe) }
        catch [IO.IOException] { if (-not [IO.File]::Exists($exe)) { throw } }
    } finally { if ([IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) } }
}
if ($Mode -eq 'Build') { Write-Output $exe; return }
$operation = if ($Mode -eq 'Remove') { '--unregister' } else { '--register' }
& $exe $operation | Out-Host
if ($LASTEXITCODE -ne 0) { throw "Could not $operation the MoreTabs page menu for the current user." }
