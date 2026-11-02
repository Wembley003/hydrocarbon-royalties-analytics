$ErrorActionPreference = "SilentlyContinue"
$ProjectRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = (Get-Location).Path
}
$VbaModPath = Join-Path $PSScriptRoot "Constructor_VBA_Nativo.bas"
$TempBootstrapDb = Join-Path $env:TEMP ("Hidrocarburos_Init_" + [guid]::NewGuid().ToString("N") + ".accdb")

Write-Host "Iniciando motor nativo de Microsoft Access (Modo Universal)..." -ForegroundColor Cyan
$accRunner = New-Object -ComObject Access.Application
$accRunner.AutomationSecurity = 1
$accRunner.Visible = $false
$accRunner.NewCurrentDatabase($TempBootstrapDb)
$accRunner.LoadFromText(5, "ModInstaladorUniversal", $VbaModPath)
$accRunner.Run("ConstruirTodoDesdeCero", $ProjectRoot)
$accRunner.CloseCurrentDatabase()
$accRunner.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($accRunner) | Out-Null
if (Test-Path $TempBootstrapDb) {
    Remove-Item -Path $TempBootstrapDb -Force -ErrorAction SilentlyContinue
}
