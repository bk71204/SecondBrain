$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $compiler)) {
    $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
}
if (-not (Test-Path -LiteralPath $compiler)) {
    throw 'No se encontró el compilador C# de .NET Framework.'
}
$outputDirectory = Join-Path $root 'dist'
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
$arguments = @(
    '/nologo'
    '/optimize+'
    '/target:winexe'
    ('/out:' + (Join-Path $outputDirectory 'ExportarAObsidian.exe'))
    ('/resource:' + (Join-Path $root 'ExportarAObsidian.ps1') + ',ExportarAObsidian.ps1')
    (Join-Path $root 'Launcher.cs')
    '/reference:System.Windows.Forms.dll'
)
& $compiler $arguments
if ($LASTEXITCODE -ne 0) {
    throw ('Falló la compilación (código {0}).' -f $LASTEXITCODE)
}
Write-Host ('Ejecutable creado: ' + (Join-Path $outputDirectory 'ExportarAObsidian.exe'))
