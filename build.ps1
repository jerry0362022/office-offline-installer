param([switch]$CheckOnly, [string]$Out = 'Office离线安装器.exe')

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$outPath = Join-Path $root $Out
$tmp = Join-Path $env:TEMP ('o365build_' + [Guid]::NewGuid().ToString('N') + '.exe')

$main = [IO.File]::ReadAllText((Join-Path $root 'src\Installer.cs'), [Text.Encoding]::UTF8)

# 其余源文件只保留类型定义，using 已在主文件中
$extra = ''
foreach ($f in @('src\Lang.cs', 'src\Theme.cs', 'src\Lang.en.cs', 'src\Lang.ja.cs', 'src\Lang.ko.cs', 'src\Products.cs', 'src\Common.cs', 'src\Consent.cs', 'src\Uninstall.cs', 'src\License.cs')) {
    $t = [IO.File]::ReadAllText((Join-Path $root $f), [Text.Encoding]::UTF8)
    $extra += ((($t -split "`r?`n") | Where-Object { $_ -notmatch '^\s*using\s+[\w\.]+\s*;\s*$' }) -join "`r`n") + "`r`n"
}
$code = $main + "`r`n" + $extra

$refs = @(
    'System.dll',
    'System.Core.dll',
    'System.Drawing.dll',
    'System.Windows.Forms.dll'
)

Add-Type -TypeDefinition $code -ReferencedAssemblies $refs -OutputAssembly $tmp -OutputType WindowsApplication -Language CSharp

if (-not (Test-Path $tmp)) {
    Write-Output 'COMPILE FAILED'
    exit 1
}

if ($CheckOnly) {
    [IO.File]::Delete($tmp)
    Write-Output 'COMPILE OK (exe 未改动)'
    exit 0
}

[IO.File]::Copy($tmp, $outPath, $true)
[IO.File]::Delete($tmp)
$fi = Get-Item $outPath
Write-Output ("BUILD OK " + [math]::Round($fi.Length / 1KB, 1) + " KB")
