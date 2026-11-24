# Run in 64-bit Windows PowerShell. This probe reads local installation metadata only.
[CmdletBinding()]
param()

$config = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
$appPathKey = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\ONENOTE.EXE'
$candidates = @(
    (Get-ItemPropertyValue -LiteralPath $appPathKey -Name '(default)' -ErrorAction SilentlyContinue)
    foreach ($root in @($env:ProgramFiles, ${env:ProgramFiles(x86)}) | Where-Object { $_ }) {
        Join-Path $root 'Microsoft Office\root\Office16\ONENOTE.EXE'
    }
)
$executable = $candidates | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1
$version = if ($executable) { (Get-Item -LiteralPath $executable).VersionInfo.ProductVersion }
$clsid = Get-ItemPropertyValue -LiteralPath 'Registry::HKEY_CLASSES_ROOT\OneNote.Application\CLSID' -Name '(default)' -ErrorAction SilentlyContinue

[pscustomobject]@{
    OneNoteExecutable   = $executable
    OneNoteVersion      = $version
    ClickToRunPlatform  = $config.Platform
    DesktopComClassId   = $clsid
    ProbeProcessBits    = [IntPtr]::Size * 8
}
