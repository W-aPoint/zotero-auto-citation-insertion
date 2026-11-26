[CmdletBinding()]
param([ValidateSet('Auto','x86','x64')][string]$WordArchitecture='Auto')
$ErrorActionPreference = 'Stop'
$dist = Join-Path $PSScriptRoot 'dist'
$dll = Join-Path $dist 'WordZoteroScan.dll'
$xpi = Join-Path $dist 'WordZoteroScan.xpi'
if (-not (Test-Path -LiteralPath $dll) -or -not (Test-Path -LiteralPath $xpi)) { & (Join-Path $PSScriptRoot 'Build.ps1') }
$installRoot = Join-Path $env:LOCALAPPDATA 'WordZoteroScan'
$build = Join-Path $installRoot ((Get-FileHash -LiteralPath $dll).Hash.Substring(0,16))
[void][IO.Directory]::CreateDirectory($build)
$installed = Join-Path $build 'WordZoteroScan.dll'
if (-not [IO.File]::Exists($installed)) { [IO.File]::Copy($dll,$installed) }
if ((Get-FileHash -LiteralPath $installed).Hash -ne (Get-FileHash -LiteralPath $dll).Hash) { throw 'Installed Word assembly differs from build.' }
$token = Join-Path $installRoot 'connection.key'
if (-not [IO.File]::Exists($token)) {
    $bytes = [byte[]]::new(32)
    $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($bytes) } finally { $rng.Dispose() }
    $hex = [BitConverter]::ToString($bytes).Replace('-','').ToLowerInvariant()
    $fs = [IO.File]::Open($token,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
    try { $buf=[Text.Encoding]::ASCII.GetBytes($hex); $fs.Write($buf,0,$buf.Length) } finally { $fs.Dispose() }
}
$prog = 'WordZoteroScan.Addin'
$guid = '{2077C602-306B-4F7A-9FA3-EB45C7F1AF34}'
$class = 'WordZoteroScan.Addin'
$asm = [Reflection.AssemblyName]::GetAssemblyName($installed)
$configuration = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
if ($WordArchitecture -eq 'Auto') {
    if (-not [Environment]::Is64BitOperatingSystem) { $WordArchitecture = 'x86' }
    elseif ($configuration.Platform -in 'x86','x64') { $WordArchitecture = $configuration.Platform }
    else { throw 'Word architecture could not be detected. Check File > Account > About Word and rerun Install.ps1 -WordArchitecture x86 or x64.' }
}
$view = if ($WordArchitecture -eq 'x86') { [Microsoft.Win32.RegistryView]::Registry32 } else { [Microsoft.Win32.RegistryView]::Registry64 }
$userHive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
function Set-Registration([string]$path, [hashtable]$values) {
    $reg = $userHive.CreateSubKey($path)
    try { foreach ($name in $values.Keys) {
        $kind = if ($values[$name] -is [int]) { [Microsoft.Win32.RegistryValueKind]::DWord } else { [Microsoft.Win32.RegistryValueKind]::String }
        $reg.SetValue($name,$values[$name],$kind)
    } } finally { $reg.Dispose() }
}
try {
    $clsid = "Software\Classes\CLSID\$guid"
    Set-Registration "Software\Classes\$prog" @{''=$prog}
    Set-Registration "Software\Classes\$prog\CLSID" @{''=$guid}
    Set-Registration $clsid @{''=$class}
    Set-Registration "$clsid\ProgID" @{''=$prog}
    $settings=@{Class=$class;Assembly=$asm.FullName;RuntimeVersion='v4.0.30319';CodeBase=([Uri]$installed).AbsoluteUri}
    Set-Registration "$clsid\InprocServer32\$($asm.Version)" $settings
    $settings['']='mscoree.dll'
    $settings['ThreadingModel']='Both'
    Set-Registration "$clsid\InprocServer32" $settings
    Set-Registration "Software\Microsoft\Office\Word\Addins\$prog" @{FriendlyName='Word Zotero Scan';Description='Scan citation markers and insert the native Zotero bibliography at the cursor';LoadBehavior=3}
} finally { $userHive.Dispose() }
Write-Output "Word add-in registered for the current user: $installed"
Write-Output "Zotero companion package: $xpi"
