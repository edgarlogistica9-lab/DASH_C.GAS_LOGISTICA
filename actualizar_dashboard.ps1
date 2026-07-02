param(
  [string]$WorkbookPath = (Join-Path $PSScriptRoot "ENVIOS ARCHIVO GENERAL.xlsm"),
  [string]$SheetName = "Enero2026",
  [string]$OutputPath = (Join-Path $PSScriptRoot "dashboard.html")
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
  throw "No encontre el archivo ENVIOS ARCHIVO GENERAL.xlsm en: $WorkbookPath"
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipText {
  param($Zip, [string]$EntryName)
  $entry = $Zip.GetEntry($EntryName)
  if (-not $entry) { throw "No encontre $EntryName dentro del Excel." }
  $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
  try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-ColIndex {
  param([string]$CellRef)
  $letters = ([regex]::Match($CellRef, "^[A-Z]+")).Value
  $n = 0
  foreach ($ch in $letters.ToCharArray()) {
    $n = ($n * 26) + ([int][char]$ch - [int][char]'A' + 1)
  }
  return $n
}

function Convert-ExcelSerialDate {
  param($Value)
  try {
    if ($null -eq $Value -or "$Value" -eq "") { return $null }
    return ([datetime]"1899-12-30").AddDays([double]$Value)
  } catch {
    try { return [datetime]$Value } catch { return $null }
  }
}

function Normalize-Text {
  param($Value)
  if ($null -eq $Value) { return "" }
  return (($Value.ToString()).Trim() -replace "\s+", " ").ToUpperInvariant()
}

function Get-OpenXmlText {
  param($Node)
  if ($null -eq $Node) { return "" }
  $textNodes = @($Node.SelectNodes("./*[local-name()='t'] | ./*[local-name()='r']/*[local-name()='t']"))
  if ($textNodes.Count -gt 0) {
    return (($textNodes | ForEach-Object { $_.InnerText }) -join "")
  }
  if ($Node.InnerText) { return [string]$Node.InnerText }
  return ""
}

function Get-DisplayText {
  param($Value)
  $text = Normalize-Text $Value
  $compact = $text -replace "[\s\./\\_-]+", ""
  if ([string]::IsNullOrWhiteSpace($text) -or $text -eq "SYSTEM.XML.XMLELEMENT" -or $compact -eq "SN") { return "SIN ESPECIFICAR" }
  return $text
}

function Test-AllowedOperator {
  param([string]$Operador)
  if ($Operador -eq "SIN ESPECIFICAR") { return $true }
  return -not [string]::IsNullOrWhiteSpace($Operador)
}

function Get-CellValue {
  param($Cell, $SharedStrings)
  $type = $Cell.t
  if ($type -eq "s") {
    $idx = [int]$Cell.v
    if ($idx -ge 0 -and $idx -lt $SharedStrings.Count) { return $SharedStrings[$idx] }
    return ""
  }
  if ($type -eq "inlineStr") {
    return Get-OpenXmlText $Cell.is
  }
  return $Cell.v
}

function Get-StablePoint {
  param([string]$Localidad)
  $known = @{
    "SIN ESPECIFICAR" = @(50, 50); "TOLUCA" = @(33, 55); "METEPEC" = @(36, 58); "LERMA" = @(40, 54); "ZINACANTEPEC" = @(27, 57);
    "NAUCALPAN" = @(50, 48); "TLALNEPANTLA" = @(56, 44); "TLANEPANTLA" = @(56, 44); "ATIZAPAN" = @(51, 38);
    "CUAUTITLAN" = @(58, 34); "CUAUTITLAN IZCALLI" = @(55, 34); "TULTITLAN" = @(60, 36); "COACALCO" = @(63, 38);
    "TECÁMAC" = @(68, 34); "TECAMAC" = @(68, 34); "ZUMPANGO" = @(64, 24); "ACOLMAN" = @(70, 40);
    "ECATEPEC" = @(65, 43); "ECATEPEC DE MORELOS" = @(65, 43); "SAN CRISTOBAL" = @(64, 41);
    "TEXCOCO" = @(76, 52); "CHIMALHUACAN" = @(68, 55); "NEZAHUALCOYOTL" = @(61, 57);
    "LOS REYES" = @(62, 63); "LA PAZ" = @(62, 63); "IXTAPALUCA" = @(66, 68); "CHALCO" = @(62, 75);
    "VALLE DE BRAVO" = @(15, 58); "ATLACOMULCO" = @(38, 24); "IXTLAHUACA" = @(35, 35);
    "CD AZTECA" = @(66, 45); "CIUDAD AZTECA" = @(66, 45); "VALLE DE ARAGON" = @(63, 52);
    "PLAZAS DE ARAGON" = @(64, 53); "SAN AGUSTIN" = @(63, 45); "XALOSTOC" = @(61, 47);
    "JARDINES DE MORELOS" = @(69, 42); "CASAS ALEMAN" = @(61, 50); "JARDINES DE SANTA CLARA" = @(63, 44);
    "EJIDOS DE SAN CRISTOBAL" = @(66, 40); "SAGITARIO" = @(67, 44); "JARDINES DEL TEPEYAC" = @(59, 50);
    "TULPETLAC" = @(62, 40); "TABLAS DEL POZO" = @(69, 39); "CHAMIZAL" = @(63, 46);
    "CHAMIZALITO" = @(63, 46); "EL SALADO XALOSTOC" = @(61, 49); "JORGE JIMENEZ CANTU" = @(70, 44);
    "VERGEL DE GUADALUPE" = @(64, 44); "GRANJAS VALLE DE GUADALUPE" = @(64, 45);
    "GRANJAS VALLE" = @(64, 45); "SAN PEDRO XALOSTOC" = @(61, 48);
    "LAZARO CARDENAS" = @(60, 44); "LA PRESA" = @(61, 39); "SAN JUAN IXHUATEPEC" = @(58, 47);
    "SAN JUAN IXHUASTEPEC" = @(58, 47); "URBANA IXHUATEPEC" = @(58, 48); "HURBANA IXHUASTEPEC" = @(58, 48);
    "NUEVA ATZACOALCO" = @(59, 52); "NVA ATZACOALCO" = @(59, 52); "SAN FELIPE DE JESUS" = @(59, 54);
    "JUAN GONZALEZ ROMERO" = @(63, 45); "PRIZO" = @(67, 42); "CODICE MENDOCINO" = @(67, 43);
    "SANTA CLARA" = @(63, 44); "CERRO GORDO" = @(64, 41); "JARDINES DE CERRO GORDO" = @(64, 41)
  }
  if ($known.ContainsKey($Localidad)) {
    return @{ x = $known[$Localidad][0]; y = $known[$Localidad][1]; approx = $false }
  }
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($Localidad)
  $hash = 0
  foreach ($b in $bytes) { $hash = (($hash * 31) + $b) % 100000 }
  return @{ x = 14 + ($hash % 72); y = 16 + (($hash / 97) % 70); approx = $true }
}

$tempFile = Join-Path ([System.IO.Path]::GetTempPath()) ("ENVIOS_DASH_{0}.xlsm" -f ([guid]::NewGuid().ToString("N")))
Copy-Item -LiteralPath $WorkbookPath -Destination $tempFile -Force

try {
  $zip = [System.IO.Compression.ZipFile]::OpenRead($tempFile)
  try {
    $sharedStrings = New-Object System.Collections.Generic.List[string]
    $sharedEntry = $zip.GetEntry("xl/sharedStrings.xml")
    if ($sharedEntry) {
      [xml]$sharedXml = Read-ZipText $zip "xl/sharedStrings.xml"
      foreach ($si in $sharedXml.sst.si) {
        [void]$sharedStrings.Add((Get-OpenXmlText $si))
      }
    }

    [xml]$workbookXml = Read-ZipText $zip "xl/workbook.xml"
    [xml]$relsXml = Read-ZipText $zip "xl/_rels/workbook.xml.rels"
    $sheetNode = $workbookXml.workbook.sheets.sheet | Where-Object { $_.name -eq $SheetName } | Select-Object -First 1
    if (-not $sheetNode) { throw "No encontre la hoja '$SheetName'." }
    $rid = $sheetNode.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $rel = $relsXml.Relationships.Relationship | Where-Object { $_.Id -eq $rid } | Select-Object -First 1
    if (-not $rel) { throw "No encontre la relacion interna de la hoja '$SheetName'." }
    $target = $rel.Target -replace "^/", ""
    if ($target -notlike "xl/*") { $target = "xl/$target" }
    $target = $target -replace "\\", "/"

    [xml]$sheetXml = Read-ZipText $zip $target
    $rows = @()
    foreach ($row in $sheetXml.worksheet.sheetData.row) {
      $obj = @{}
      foreach ($cell in $row.c) {
        $col = Get-ColIndex $cell.r
        $obj[$col] = Get-CellValue $cell $sharedStrings
      }
      $rows += ,$obj
    }
  } finally {
    $zip.Dispose()
  }
} finally {
  Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
}

if ($rows.Count -lt 2) { throw "La hoja '$SheetName' no tiene datos suficientes." }

$headers = @{}
foreach ($key in $rows[0].Keys) { $headers[(Normalize-Text $rows[0][$key])] = [int]$key }

$colOperador = $headers["OPERADOR"]
$colCliente = $headers["CLIENTE"]
$colLocalidad = $headers["LOCALIDAD"]
$colTicketFactura = $headers["T/ FACT"]
$colImporte = $headers["IMPORTE"]
$colFecha = $headers["FECHA ENVIO"]
if (-not $colOperador -or -not $colLocalidad -or -not $colImporte -or -not $colFecha) {
  throw "Faltan columnas requeridas: OPERADOR, LOCALIDAD, IMPORTE o FECHA ENVIO."
}

$monthNames = @("", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
$orders = New-Object System.Collections.Generic.List[object]

for ($i = 1; $i -lt $rows.Count; $i++) {
  $r = $rows[$i]
  $operador = Get-DisplayText $r[$colOperador]
  $localidad = Get-DisplayText $r[$colLocalidad]
  $cliente = Get-DisplayText $r[$colCliente]
  $ticketFactura = Get-DisplayText $r[$colTicketFactura]
  if (-not (Test-AllowedOperator $operador)) { continue }
  $fecha = Convert-ExcelSerialDate $r[$colFecha]
  if ($null -eq $fecha) { continue }
  $importe = 0.0
  $importeText = ""
  if ($null -ne $r[$colImporte]) { $importeText = ($r[$colImporte]).ToString() }
  [void][double]::TryParse($importeText, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$importe)
  $point = Get-StablePoint $localidad
  [void]$orders.Add([pscustomobject]@{
    operador = $operador
    cliente = $cliente
    localidad = $localidad
    ticketFactura = $ticketFactura
    importe = [math]::Round($importe, 2)
    fecha = $fecha.ToString("yyyy-MM-dd")
    mes = $monthNames[[int]$fecha.Month]
    mesNum = [int]$fecha.Month
    x = [math]::Round([double]$point.x, 2)
    y = [math]::Round([double]$point.y, 2)
    approx = [bool]$point.approx
  })
}

$generatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
$data = [pscustomobject]@{
  generatedAt = $generatedAt
  sourceFile = [System.IO.Path]::GetFileName($WorkbookPath)
  sheet = $SheetName
  orders = $orders
}
$json = $data | ConvertTo-Json -Depth 8 -Compress

$html = @'
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Dashboard Envios</title>
  <style>
    :root {
      --bg: #f5f6f2;
      --panel: #ffffff;
      --panel-soft: #fbfcfa;
      --ink: #17202c;
      --muted: #647084;
      --line: #d9dfd6;
      --line-strong: #c4cdbf;
      --blue: #2454a6;
      --green: #166b54;
      --amber: #b57614;
      --red: #b4473f;
      --teal: #0f766e;
      --shadow: 0 12px 28px rgba(23, 32, 44, .08);
      --radius: 8px;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: Segoe UI, Arial, sans-serif;
      background: var(--bg);
      color: var(--ink);
      letter-spacing: 0;
    }
    header {
      background: #ffffff;
      color: var(--ink);
      padding: 18px 24px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
      border-bottom: 1px solid var(--line);
      box-shadow: 0 1px 0 rgba(23, 32, 44, .04);
    }
    h1 { margin: 0; font-size: 24px; font-weight: 700; letter-spacing: 0; }
    .stamp {
      color: #3f4b5e;
      font-size: 12px;
      text-align: right;
      line-height: 1.45;
      background: #f7f8f5;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 9px 11px;
      max-width: 420px;
    }
    main { padding: 18px; max-width: 1500px; margin: 0 auto; }
    .filters {
      display: grid;
      grid-template-columns: repeat(7, minmax(130px, 1fr));
      gap: 12px;
      margin-bottom: 16px;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: var(--radius);
      box-shadow: 0 4px 14px rgba(23, 32, 44, .05);
      padding: 13px;
    }
    label { font-size: 12px; color: var(--muted); display: grid; gap: 6px; font-weight: 600; }
    select, input {
      height: 38px;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: white;
      padding: 0 10px;
      color: var(--ink);
      min-width: 0;
      outline: none;
      transition: border-color .15s ease, box-shadow .15s ease, background .15s ease;
    }
    select:focus, input:focus {
      border-color: var(--blue);
      box-shadow: 0 0 0 3px rgba(36, 84, 166, .12);
      background: #fff;
    }
    .kpis {
      display: grid;
      grid-template-columns: repeat(5, minmax(150px, 1fr));
      gap: 12px;
      margin-bottom: 16px;
    }
    .kpi, .panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: var(--radius);
      box-shadow: 0 5px 18px rgba(23, 32, 44, .06);
    }
    .kpi {
      padding: 14px 15px;
      border-top: 3px solid var(--blue);
      min-height: 86px;
    }
    .kpi:nth-child(2) { border-top-color: var(--green); }
    .kpi:nth-child(3) { border-top-color: var(--amber); }
    .kpi:nth-child(4) { border-top-color: var(--teal); }
    .kpi:nth-child(5) { border-top-color: var(--red); }
    .kpi span { display: block; color: var(--muted); font-size: 12px; margin-bottom: 8px; font-weight: 600; }
    .kpi strong { display: block; font-size: 25px; line-height: 1.08; color: var(--ink); }
    .grid {
      display: grid;
      grid-template-columns: 1.15fr .85fr;
      gap: 16px;
      align-items: start;
    }
    .panel { padding: 15px; min-width: 0; }
    .panel h2 {
      font-size: 15px;
      margin: 0 0 14px;
      display: flex;
      justify-content: space-between;
      gap: 10px;
      align-items: baseline;
      color: #1f2937;
      padding-bottom: 9px;
      border-bottom: 1px solid #eef1eb;
    }
    .panel h2 small { color: var(--muted); font-weight: 400; }
    canvas { display: block; width: 100%; height: 320px; }
    #map { width: 100%; height: 430px; border: 1px solid var(--line); border-radius: 6px; background: #eef3ee; }
    table { width: 100%; border-collapse: collapse; font-size: 12px; }
    th, td { border-bottom: 1px solid #edf0e9; padding: 9px 8px; text-align: left; }
    th { color: #4d5968; font-weight: 700; background: #f7f8f5; position: sticky; top: 0; z-index: 1; }
    tbody tr:nth-child(even) { background: #fbfcfa; }
    tbody tr:hover { background: #eef5f1; }
    td.num, th.num { text-align: right; }
    .table-wrap { max-height: 360px; overflow: auto; border: 1px solid var(--line); border-radius: 6px; background: #fff; }
    .two { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-top: 16px; }
    .three { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 16px; margin-top: 16px; }
    .legend { color: var(--muted); font-size: 12px; margin-top: 4px; }
    .bubble { cursor: pointer; stroke: white; stroke-width: .9; }
    .map-note { color: var(--muted); font-size: 11px; margin-top: 8px; }
    .swatch { display: inline-block; width: 11px; height: 11px; border-radius: 3px; border: 1px solid rgba(16,24,32,.18); margin-right: 6px; vertical-align: -1px; }
    .color-legend { display: flex; flex-wrap: wrap; gap: 7px 12px; margin-top: 9px; color: var(--muted); font-size: 11px; }
    .color-legend span { white-space: nowrap; }
    .municipality-region { stroke: #ffffff; stroke-width: .55; opacity: .72; }
    .municipality-outline { fill: none; stroke: #73846f; stroke-width: 1.7; }
    .municipality-label { fill: #334155; font-size: 3.3px; text-anchor: middle; pointer-events: none; font-weight: 600; }
    .shape { fill: none; stroke: #a6b9ab; stroke-width: 1.7; }
    .road { stroke: #c2c8bd; stroke-width: 2; stroke-dasharray: 5 5; }
    .label { fill: #526071; font-size: 11px; }
    .map-city { fill: #64748b; font-size: 4.2px; text-anchor: middle; }
    .tooltip {
      position: fixed;
      pointer-events: none;
      background: #142033;
      color: white;
      padding: 8px 10px;
      border-radius: 6px;
      font-size: 12px;
      max-width: 240px;
      display: none;
      z-index: 10;
      box-shadow: var(--shadow);
    }
    @media (max-width: 980px) {
      header { align-items: start; flex-direction: column; }
      .stamp { text-align: left; }
      .filters, .kpis, .grid, .two, .three { grid-template-columns: 1fr; }
      canvas { height: 280px; }
    }
  </style>
</head>
<body>
  <header>
    <div>
      <h1>Dashboard Envios</h1>
      <div class="legend">Pedidos, importes, choferes y concentracion por localidad</div>
    </div>
    <div class="stamp" id="stamp"></div>
  </header>

  <main>
    <section class="filters">
      <label>Mes
        <select id="monthFilter"></select>
      </label>
      <label>Chofer
        <select id="driverFilter"></select>
      </label>
      <label>Localidad
        <select id="localityFilter"></select>
      </label>
      <label>Desde
        <input id="startDateFilter" type="date">
      </label>
      <label>Hasta
        <input id="endDateFilter" type="date">
      </label>
      <label>Ticket/factura
        <input id="ticketSearch" type="search" placeholder="T/ FACT">
      </label>
      <label>Buscar cliente/localidad
        <input id="searchBox" type="search" placeholder="Buscar">
      </label>
    </section>

    <section class="kpis">
      <div class="kpi"><span>Pedidos</span><strong id="kpiOrders">0</strong></div>
      <div class="kpi"><span>Importe total</span><strong id="kpiAmount">$0</strong></div>
      <div class="kpi"><span>Choferes</span><strong id="kpiDrivers">0</strong></div>
      <div class="kpi"><span>Localidades</span><strong id="kpiLocalities">0</strong></div>
      <div class="kpi"><span>Clientes</span><strong id="kpiClients">0</strong></div>
    </section>

    <section class="grid">
      <div class="panel">
        <h2>Pedidos e importe por mes <small id="monthHint"></small></h2>
        <canvas id="monthChart"></canvas>
      </div>
      <div class="panel">
        <h2>Mapa Estado de Mexico <small>color por municipio | puntos por localidad</small></h2>
        <svg id="map" viewBox="0 0 100 100" preserveAspectRatio="none" role="img" aria-label="Mapa de concentracion del Estado de Mexico"></svg>
        <div class="map-note">Cada region de color representa un municipio aproximado; las burbujas muestran localidades y volumen de entregas.</div>
      </div>
    </section>

    <section class="three">
      <div class="panel">
        <h2>Pedidos por chofer <small id="driverHint"></small></h2>
        <canvas id="driverChart"></canvas>
        <div class="color-legend" id="driverColorLegend"></div>
      </div>
      <div class="panel">
        <h2>Top localidades <small id="localityHint"></small></h2>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Localidad</th><th class="num">Pedidos</th><th class="num">Importe</th></tr></thead>
            <tbody id="localityRows"></tbody>
          </table>
        </div>
      </div>
      <div class="panel">
        <h2>Pedidos por dia <small id="dayHint"></small></h2>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Dia</th><th class="num">Pedidos</th><th class="num">Importe acumulado</th></tr></thead>
            <tbody id="dayRows"></tbody>
          </table>
        </div>
      </div>
    </section>

    <section class="panel" style="margin-top:14px">
      <h2>Pedidos diarios por localidad <small id="dailyLocalityHint"></small></h2>
      <div class="table-wrap">
        <table>
          <thead><tr><th>Fecha</th><th>Localidad</th><th class="num">Pedidos</th><th class="num">Importe</th></tr></thead>
          <tbody id="dailyLocalityRows"></tbody>
        </table>
      </div>
    </section>

    <section class="panel" style="margin-top:14px">
      <h2>Detalle de pedidos <small id="detailHint"></small></h2>
      <div class="table-wrap">
        <table>
          <thead><tr><th>Fecha</th><th>Ticket/factura</th><th>Chofer</th><th>Cliente</th><th>Localidad</th><th class="num">Importe</th></tr></thead>
          <tbody id="detailRows"></tbody>
        </table>
      </div>
    </section>
  </main>
  <div class="tooltip" id="tooltip"></div>

  <script>
    const DATA = __DATA_JSON__;
    const money = new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', maximumFractionDigits: 0 });
    const number = new Intl.NumberFormat('es-MX');
    const monthOrder = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
    const $ = (id) => document.getElementById(id);

    function groupBy(rows, key) {
      const out = new Map();
      rows.forEach(row => {
        const k = row[key] || 'SIN DATO';
        const item = out.get(k) || { name: k, pedidos: 0, importe: 0, x: row.x, y: row.y, approx: row.approx };
        item.pedidos += 1;
        item.importe += Number(row.importe || 0);
        if (row.x != null) { item.x = row.x; item.y = row.y; item.approx = row.approx; }
        out.set(k, item);
      });
      return Array.from(out.values());
    }

    function unique(rows, key) {
      return Array.from(new Set(rows.map(r => r[key]).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'es'));
    }

    function fillSelect(id, values, label) {
      const sel = $(id);
      const current = sel.value;
      sel.innerHTML = `<option value="">${label}</option>` + values.map(v => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
      sel.value = values.includes(current) ? current : '';
    }

    function clamp(value, min, max) {
      return Math.min(max, Math.max(min, value));
    }

    function heatColor(value, min, max) {
      const t = max <= min ? 1 : clamp((value - min) / (max - min), 0, 1);
      const mid = t < .5 ? t / .5 : (t - .5) / .5;
      const a = t < .5 ? [31, 157, 85] : [240, 180, 41];
      const b = t < .5 ? [240, 180, 41] : [214, 69, 69];
      const rgb = a.map((start, i) => Math.round(start + (b[i] - start) * mid));
      return `rgb(${rgb[0]}, ${rgb[1]}, ${rgb[2]})`;
    }

    function localityColor(name) {
      const palette = ['#2563eb','#dc2626','#16a34a','#f59e0b','#7c3aed','#0891b2','#db2777','#65a30d','#ea580c','#475569','#0f766e','#9333ea','#be123c','#1d4ed8','#a16207','#15803d'];
      let hash = 0;
      String(name).split('').forEach(ch => { hash = ((hash * 31) + ch.charCodeAt(0)) >>> 0; });
      return palette[hash % palette.length];
    }

    function driverColor(name) {
      const palette = ['#1d4ed8','#dc2626','#15803d','#c2410c','#7e22ce','#0e7490','#be123c','#4d7c0f','#b45309','#334155','#047857','#6d28d9','#0369a1','#a21caf','#ca8a04','#166534'];
      let hash = 0;
      String(name).split('').forEach(ch => { hash = ((hash * 33) + ch.charCodeAt(0)) >>> 0; });
      return palette[hash % palette.length];
    }

    function monthColor(name) {
      const colors = {
        Enero: '#2563eb',
        Febrero: '#dc2626',
        Marzo: '#16a34a',
        Abril: '#f59e0b',
        Mayo: '#7c3aed',
        Junio: '#0891b2',
        Julio: '#db2777',
        Agosto: '#65a30d',
        Septiembre: '#ea580c',
        Octubre: '#475569',
        Noviembre: '#0f766e',
        Diciembre: '#9333ea'
      };
      return colors[name] || '#1f6feb';
    }

    function escapeHtml(value) {
      return String(value).replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
    }

    function filteredRows() {
      const month = $('monthFilter').value;
      const driver = $('driverFilter').value;
      const locality = $('localityFilter').value;
      const startDate = $('startDateFilter').value;
      const endDate = $('endDateFilter').value;
      const ticket = $('ticketSearch').value.trim().toUpperCase();
      const q = $('searchBox').value.trim().toUpperCase();
      return DATA.orders.filter(r =>
        (!month || r.mes === month) &&
        (!driver || r.operador === driver) &&
        (!locality || r.localidad === locality) &&
        (!startDate || r.fecha >= startDate) &&
        (!endDate || r.fecha <= endDate) &&
        (!ticket || String(r.ticketFactura || '').toUpperCase().includes(ticket)) &&
        (!q || `${r.cliente} ${r.localidad} ${r.operador}`.toUpperCase().includes(q))
      );
    }

    function drawBarLine(canvas, labels, barValues, lineValues, opts = {}) {
      const ctx = canvas.getContext('2d');
      const dpr = window.devicePixelRatio || 1;
      const rect = canvas.getBoundingClientRect();
      canvas.width = Math.max(320, rect.width * dpr);
      canvas.height = Math.max(240, rect.height * dpr);
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      const w = rect.width, h = rect.height;
      ctx.clearRect(0, 0, w, h);
      const pad = { l: 46, r: 16, t: 18, b: 54 };
      const chartW = w - pad.l - pad.r, chartH = h - pad.t - pad.b;
      const maxBar = Math.max(1, ...barValues);
      const maxLine = Math.max(1, ...lineValues);
      ctx.strokeStyle = '#d9dee7';
      ctx.lineWidth = 1;
      ctx.beginPath();
      for (let i = 0; i <= 4; i++) {
        const y = pad.t + chartH * i / 4;
        ctx.moveTo(pad.l, y); ctx.lineTo(w - pad.r, y);
      }
      ctx.stroke();
      const slot = chartW / Math.max(1, labels.length);
      const barW = Math.max(12, Math.min(44, slot * .48));
      labels.forEach((label, i) => {
        const bh = chartH * (barValues[i] / maxBar);
        const x = pad.l + i * slot + (slot - barW) / 2;
        const y = pad.t + chartH - bh;
        ctx.fillStyle = opts.barColors ? opts.barColors[i] : (opts.bar || '#1f6feb');
        ctx.fillRect(x, y, barW, bh);
        ctx.save();
        ctx.translate(x + barW / 2, h - 38);
        ctx.rotate(labels.length > 6 ? -Math.PI / 5 : 0);
        ctx.fillStyle = '#526071';
        ctx.font = '11px Segoe UI, Arial';
        ctx.textAlign = labels.length > 6 ? 'right' : 'center';
        ctx.fillText(label.length > 15 ? label.slice(0, 14) + '.' : label, 0, 0);
        ctx.restore();
      });
      ctx.strokeStyle = opts.line || '#168a5b';
      ctx.lineWidth = 2;
      ctx.beginPath();
      labels.forEach((label, i) => {
        const x = pad.l + i * slot + slot / 2;
        const y = pad.t + chartH - chartH * (lineValues[i] / maxLine);
        if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
      });
      ctx.stroke();
      labels.forEach((label, i) => {
        const x = pad.l + i * slot + slot / 2;
        const y = pad.t + chartH - chartH * (lineValues[i] / maxLine);
        ctx.fillStyle = opts.pointColors ? opts.pointColors[i] : (opts.line || '#168a5b');
        ctx.beginPath(); ctx.arc(x, y, 3.5, 0, Math.PI * 2); ctx.fill();
      });
      ctx.fillStyle = '#526071';
      ctx.font = '11px Segoe UI, Arial';
      ctx.textAlign = 'right';
      ctx.fillText(number.format(maxBar), pad.l - 8, pad.t + 4);
      ctx.fillText('0', pad.l - 8, pad.t + chartH);
    }

    function renderMap(rows) {
      const map = $('map');
      const tip = $('tooltip');
      const locs = groupBy(rows, 'localidad').sort((a, b) => b.pedidos - a.pedidos);
      const max = Math.max(1, ...locs.map(l => l.pedidos));
      map.innerHTML = `
        <defs>
          <clipPath id="edomexClip">
            <path d="M44,8 L60,13 L74,25 L88,36 L91,52 L82,66 L74,84 L55,92 L39,86 L25,91 L13,77 L10,60 L17,43 L13,28 L27,17 Z"></path>
          </clipPath>
        </defs>
        <g clip-path="url(#edomexClip)">
          <polygon class="municipality-region" fill="#c7d2fe" points="13,28 27,17 44,8 48,25 38,34 22,38"></polygon>
          <polygon class="municipality-region" fill="#bbf7d0" points="48,25 60,13 74,25 68,36 56,35"></polygon>
          <polygon class="municipality-region" fill="#fde68a" points="68,36 74,25 88,36 91,52 78,52"></polygon>
          <polygon class="municipality-region" fill="#fecaca" points="56,35 68,36 66,47 56,48 50,42"></polygon>
          <polygon class="municipality-region" fill="#bae6fd" points="66,47 78,52 82,66 70,66 61,58"></polygon>
          <polygon class="municipality-region" fill="#ddd6fe" points="50,42 56,48 61,58 48,60 40,52"></polygon>
          <polygon class="municipality-region" fill="#fed7aa" points="38,34 50,42 40,52 29,53 22,38"></polygon>
          <polygon class="municipality-region" fill="#a7f3d0" points="17,43 29,53 28,70 14,72 10,60"></polygon>
          <polygon class="municipality-region" fill="#fbcfe8" points="29,53 40,52 48,60 42,74 28,70"></polygon>
          <polygon class="municipality-region" fill="#bfdbfe" points="48,60 61,58 70,66 62,80 47,76"></polygon>
          <polygon class="municipality-region" fill="#d9f99d" points="70,66 82,66 74,84 62,80"></polygon>
          <polygon class="municipality-region" fill="#fef3c7" points="14,72 28,70 42,74 39,86 25,91 13,77"></polygon>
          <polygon class="municipality-region" fill="#e9d5ff" points="42,74 47,76 62,80 55,92 39,86"></polygon>
        </g>
        <path class="municipality-outline" d="M44,8 L60,13 L74,25 L88,36 L91,52 L82,66 L74,84 L55,92 L39,86 L25,91 L13,77 L10,60 L17,43 L13,28 L27,17 Z"></path>
        <path class="road" d="M33,55 C44,50 53,47 65,43"></path>
        <path class="road" d="M50,48 C58,52 66,58 76,52"></path>
        <path class="road" d="M65,43 C68,50 66,60 62,75"></path>
        <text class="label" x="50" y="7" text-anchor="middle">Estado de Mexico</text>
        <text class="municipality-label" x="32" y="27">Atlacomulco</text>
        <text class="municipality-label" x="60" y="26">Zumpango</text>
        <text class="municipality-label" x="80" y="39">Texcoco</text>
        <text class="municipality-label" x="61" y="42">Ecatepec</text>
        <text class="municipality-label" x="70" y="57">Chimalhuacan</text>
        <text class="municipality-label" x="49" y="53">Tlalnepantla</text>
        <text class="municipality-label" x="33" y="44">Naucalpan</text>
        <text class="municipality-label" x="21" y="62">Valle Bravo</text>
        <text class="municipality-label" x="36" y="64">Toluca</text>
        <text class="municipality-label" x="58" y="70">Ixtapaluca</text>
        <text class="municipality-label" x="72" y="75">Chalco</text>
        <text class="municipality-label" x="28" y="80">Zinacantepec</text>
        <text class="municipality-label" x="50" y="84">Metepec</text>
      `;
      locs.slice(0, 80).forEach(l => {
        const r = 1.8 + Math.sqrt(l.pedidos / max) * 4.2;
        const c = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
        c.setAttribute('class', 'bubble');
        c.setAttribute('cx', l.x);
        c.setAttribute('cy', l.y);
        c.setAttribute('r', r.toFixed(2));
        c.setAttribute('fill', localityColor(l.name));
        c.setAttribute('fill-opacity', l.approx ? '.72' : '.88');
        c.addEventListener('mousemove', ev => {
          tip.style.display = 'block';
          tip.style.left = (ev.clientX + 12) + 'px';
          tip.style.top = (ev.clientY + 12) + 'px';
          tip.innerHTML = `<strong>${escapeHtml(l.name)}</strong><br>${number.format(l.pedidos)} pedidos<br>${money.format(l.importe)}<br><span style="color:${localityColor(l.name)}">Color de localidad</span>`;
        });
        c.addEventListener('mouseleave', () => tip.style.display = 'none');
        c.addEventListener('click', () => { $('localityFilter').value = l.name; render(); });
        map.appendChild(c);
      });
    }

    function renderDailyTable(rows) {
      const days = groupBy(rows, 'fecha').sort((a, b) => a.name.localeCompare(b.name));
      let accumulated = 0;
      $('dayRows').innerHTML = days.map(d => {
        accumulated += d.importe;
        return `<tr><td>${d.name}</td><td class="num">${number.format(d.pedidos)}</td><td class="num">${money.format(accumulated)}</td></tr>`;
      }).join('');
      $('dayHint').textContent = `${number.format(days.length)} dias | acumulado ${money.format(accumulated)}`;
    }

    function renderDailyLocalityTable(rows) {
      const out = new Map();
      rows.forEach(r => {
        const key = `${r.fecha}||${r.localidad}`;
        const item = out.get(key) || { fecha: r.fecha, localidad: r.localidad, pedidos: 0, importe: 0 };
        item.pedidos += 1;
        item.importe += Number(r.importe || 0);
        out.set(key, item);
      });
      const items = Array.from(out.values()).sort((a, b) => b.fecha.localeCompare(a.fecha) || b.pedidos - a.pedidos || a.localidad.localeCompare(b.localidad, 'es'));
      $('dailyLocalityRows').innerHTML = items.slice(0, 500).map(item =>
        `<tr><td>${item.fecha}</td><td><span class="swatch" style="background:${localityColor(item.localidad)}"></span>${escapeHtml(item.localidad)}</td><td class="num">${number.format(item.pedidos)}</td><td class="num">${money.format(item.importe)}</td></tr>`
      ).join('');
      $('dailyLocalityHint').textContent = `${number.format(Math.min(items.length, 500))} de ${number.format(items.length)} combinaciones`;
    }

    function renderTables(rows) {
      const locs = groupBy(rows, 'localidad').sort((a, b) => b.pedidos - a.pedidos || b.importe - a.importe);
      $('localityRows').innerHTML = locs.slice(0, 30).map(l =>
        `<tr><td><span class="swatch" style="background:${localityColor(l.name)}"></span>${escapeHtml(l.name)}</td><td class="num">${number.format(l.pedidos)}</td><td class="num">${money.format(l.importe)}</td></tr>`
      ).join('');
      $('detailRows').innerHTML = rows.slice().sort((a, b) => b.fecha.localeCompare(a.fecha)).slice(0, 300).map(r =>
        `<tr><td>${r.fecha}</td><td>${escapeHtml(r.ticketFactura)}</td><td>${escapeHtml(r.operador)}</td><td>${escapeHtml(r.cliente)}</td><td>${escapeHtml(r.localidad)}</td><td class="num">${money.format(r.importe)}</td></tr>`
      ).join('');
    }

    function render() {
      const rows = filteredRows();
      const total = rows.reduce((s, r) => s + Number(r.importe || 0), 0);
      $('kpiOrders').textContent = number.format(rows.length);
      $('kpiAmount').textContent = money.format(total);
      $('kpiDrivers').textContent = number.format(unique(rows, 'operador').length);
      $('kpiLocalities').textContent = number.format(unique(rows, 'localidad').length);
      $('kpiClients').textContent = number.format(unique(rows, 'cliente').length);

      const monthGroups = groupBy(rows, 'mes').sort((a, b) => monthOrder.indexOf(a.name) - monthOrder.indexOf(b.name));
      const monthColors = monthGroups.map(x => monthColor(x.name));
      drawBarLine($('monthChart'), monthGroups.map(x => x.name), monthGroups.map(x => x.pedidos), monthGroups.map(x => x.importe), { barColors: monthColors, pointColors: monthColors, line: '#334155' });
      $('monthHint').textContent = 'color por mes | linea: importe';

      const driverGroups = groupBy(rows, 'operador').sort((a, b) => b.pedidos - a.pedidos).slice(0, 12);
      const driverColors = driverGroups.map(x => driverColor(x.name));
      drawBarLine($('driverChart'), driverGroups.map(x => x.name), driverGroups.map(x => x.pedidos), driverGroups.map(x => x.importe), { barColors: driverColors, pointColors: driverColors, line: '#334155' });
      $('driverHint').textContent = `${number.format(driverGroups.length)} visibles`;
      $('driverColorLegend').innerHTML = driverGroups.map((x, i) => `<span><span class="swatch" style="background:${driverColors[i]}"></span>${escapeHtml(x.name)}</span>`).join('');
      $('localityHint').textContent = `${number.format(groupBy(rows, 'localidad').length)} localidades`;
      $('detailHint').textContent = `${number.format(Math.min(rows.length, 300))} de ${number.format(rows.length)} registros`;

      renderMap(rows);
      renderDailyTable(rows);
      renderDailyLocalityTable(rows);
      renderTables(rows);
    }

    function init() {
      $('stamp').textContent = `Fuente: ${DATA.sourceFile} / ${DATA.sheet} | Actualizado: ${DATA.generatedAt}`;
      fillSelect('monthFilter', monthOrder.filter(m => DATA.orders.some(r => r.mes === m)), 'Todos los meses');
      fillSelect('driverFilter', unique(DATA.orders, 'operador'), 'Todos los choferes');
      fillSelect('localityFilter', unique(DATA.orders, 'localidad'), 'Todas las localidades');
      const dates = unique(DATA.orders, 'fecha');
      if (dates.length) {
        ['startDateFilter','endDateFilter'].forEach(id => {
          $(id).min = dates[0];
          $(id).max = dates[dates.length - 1];
        });
      }
      ['monthFilter','driverFilter','localityFilter','startDateFilter','endDateFilter','ticketSearch','searchBox'].forEach(id => $(id).addEventListener('input', render));
      window.addEventListener('resize', render);
      render();
    }
    init();
  </script>
</body>
</html>
'@

$html = $html.Replace("__DATA_JSON__", $json)
[System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.Encoding]::UTF8)
Write-Host "Dashboard actualizado: $OutputPath"
