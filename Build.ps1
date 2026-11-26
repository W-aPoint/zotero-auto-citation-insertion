$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$dist = Join-Path $here 'dist'
[void][IO.Directory]::CreateDirectory($dist)
$assemblies = @{
    word   = (Join-Path $env:WINDIR 'assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll')
    office = (Join-Path $env:WINDIR 'assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll')
    extensibility = (Join-Path $env:WINDIR 'assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\extensibility.dll')
}
foreach ($assembly in $assemblies.Values) {
    if (-not (Test-Path -LiteralPath $assembly)) { throw "Office interop assembly not found: $assembly" }
}
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $compiler)) { $compiler=Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
$dll = Join-Path $dist 'WordZoteroScan.dll'
$src = @(Get-ChildItem -LiteralPath (Join-Path $here 'WordAddin') -Filter '*.cs' | ForEach-Object FullName)
$opts = @('/nologo','/target:library','/platform:anycpu','/codepage:65001','/optimize+',"/out:$dll",
    '/r:System.dll','/r:System.Core.dll','/r:System.Drawing.dll','/r:System.Windows.Forms.dll','/r:System.Web.Extensions.dll')
$opts += @($assemblies.Values | ForEach-Object { "/r:$_" })
if (-not [IO.File]::Exists($dll) -or @($src | Where-Object { [IO.File]::GetLastWriteTimeUtc($_) -gt [IO.File]::GetLastWriteTimeUtc($dll) }).Count) {
    & $compiler @opts @src
    if ($LASTEXITCODE) { throw 'Word add-in build failed.' }
}
Add-Type -AssemblyName System.IO.Compression.FileSystem
$xpi = Join-Path $dist 'WordZoteroScan.xpi'
$staged = Join-Path $dist ([guid]::NewGuid().ToString('N') + '.zip')
[IO.Compression.ZipFile]::CreateFromDirectory((Join-Path $here 'zotero'), $staged)
if ([IO.File]::Exists($xpi)) { [IO.File]::Replace($staged,$xpi,[System.Management.Automation.Language.NullString]::Value) }
else { [IO.File]::Move($staged,$xpi) }
Get-Item -LiteralPath $dll,$xpi | Select-Object Name,Length
