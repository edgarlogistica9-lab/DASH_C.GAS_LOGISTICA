param(
  [string]$WorkbookPath = (Join-Path $PSScriptRoot "ENVIOS ARCHIVO GENERAL.xlsm"),
  [string]$DashboardPath = (Join-Path $PSScriptRoot "dashboard.html"),
  [string]$SheetName = "Enero2026"
)

$ErrorActionPreference = "Stop"

function Invoke-DashboardUpdate {
  powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "actualizar_dashboard.ps1") -WorkbookPath $WorkbookPath -SheetName $SheetName -OutputPath $DashboardPath
  if ($LASTEXITCODE -ne 0) {
    throw "No se pudo actualizar dashboard.html."
  }
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
  throw "No encontre ENVIOS ARCHIVO GENERAL.xlsm en esta carpeta: $PSScriptRoot"
}

Write-Host "Preparando dashboard inicial..."
Invoke-DashboardUpdate

Write-Host "Abriendo dashboard..."
Start-Process -FilePath $DashboardPath

$lastWrite = (Get-Item -LiteralPath $WorkbookPath).LastWriteTimeUtc
Write-Host ""
Write-Host "Estado: activo"
Write-Host "Fuente: ENVIOS ARCHIVO GENERAL.xlsm"
Write-Host "Hoja:   $SheetName"
Write-Host "Modo:   portable, sin rutas fijas"
Write-Host "Vigilando cambios guardados. Presiona Ctrl+C para detener."
Write-Host ""

$watchDir = Split-Path -Parent $WorkbookPath
$watchFile = Split-Path -Leaf $WorkbookPath
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $watchDir
$watcher.Filter = $watchFile
$watcher.NotifyFilter = [System.IO.NotifyFilters]'FileName, LastWrite, Size'
$watcher.EnableRaisingEvents = $true

try {
  Register-ObjectEvent -InputObject $watcher -EventName Changed -SourceIdentifier DashboardExcelChanged | Out-Null
  Register-ObjectEvent -InputObject $watcher -EventName Created -SourceIdentifier DashboardExcelCreated | Out-Null
  Register-ObjectEvent -InputObject $watcher -EventName Renamed -SourceIdentifier DashboardExcelRenamed | Out-Null

  while ($true) {
    $event = Wait-Event
    Remove-Event -EventIdentifier $event.EventIdentifier -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Get-Event | Where-Object { $_.SourceIdentifier -in @('DashboardExcelChanged','DashboardExcelCreated','DashboardExcelRenamed') } | Remove-Event
    
    if (-not (Test-Path -LiteralPath $WorkbookPath)) {
      Write-Host "Aviso: no encuentro el Excel. Esperando el siguiente cambio..."
      continue
    }
    
    $currentWrite = (Get-Item -LiteralPath $WorkbookPath).LastWriteTimeUtc
    if ($currentWrite -eq $lastWrite) { continue }
    $lastWrite = $currentWrite
    
    try {
      Write-Host ("[{0}] Cambio guardado detectado. Regenerando dashboard..." -f (Get-Date -Format "HH:mm:ss"))
      Invoke-DashboardUpdate
      Write-Host ("[{0}] Dashboard actualizado." -f (Get-Date -Format "HH:mm:ss"))
      Write-Host "Si ya lo tienes abierto, actualiza la pagina cuando quieras ver el cambio."
    } catch {
      Write-Host ("Aviso: {0}" -f $_.Exception.Message)
    }
  }
} finally {
  Get-EventSubscriber | Where-Object { $_.SourceIdentifier -in @('DashboardExcelChanged','DashboardExcelCreated','DashboardExcelRenamed') } | Unregister-Event -ErrorAction SilentlyContinue
  $watcher.Dispose()
}
