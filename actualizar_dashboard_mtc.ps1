param(
  [string]$WorkbookPath = "E:\REPORTE_MTC.xlsx",
  [string]$OutputPath = (Join-Path $PSScriptRoot "dashboard_mtc.html"),
  [switch]$Watch
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Normalize-Text {
  param($Value)
  if ($null -eq $Value) { return "" }
  $text = ($Value.ToString()).Trim() -replace "\s+", " "
  return $text.ToUpperInvariant()
}

function Get-DisplayText {
  param($Value)
  $text = Normalize-Text $Value
  if ([string]::IsNullOrWhiteSpace($text) -or $text -eq "SYSTEM.XML.XMLELEMENT") { return "" }
  return $text
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

function Read-ZipText {
  param($Zip, [string]$EntryName)
  $entry = $Zip.GetEntry($EntryName)
  if (-not $entry) { throw "No encontre $EntryName dentro del archivo Excel." }
  $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
  try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-OpenXmlText {
  param($Node)
  if ($null -eq $Node) { return "" }
  $textNodes = @($Node.SelectNodes("./*[local-name()='t'] | ./*[local-name()='r']/*[local-name()='t']"))
  if ($textNodes.Count -gt 0) { return (($textNodes | ForEach-Object { $_.InnerText }) -join "") }
  if ($Node.InnerText) { return [string]$Node.InnerText }
  return ""
}

function Get-CellValue {
  param($Cell, $SharedStrings)
  $type = $Cell.t
  if ($type -eq "s") {
    $idx = [int]$Cell.v
    if ($idx -ge 0 -and $idx -lt $SharedStrings.Count) { return $SharedStrings[$idx] }
    return ""
  }
  if ($type -eq "inlineStr") { return Get-OpenXmlText $Cell.is }
  return $Cell.v
}

function Convert-ExcelDate {
  param($Value)
  if ($null -eq $Value -or "$Value" -eq "") { return $null }
  try {
    $n = 0.0
    if ([double]::TryParse(($Value.ToString()), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) {
      return ([datetime]"1899-12-30").AddDays($n)
    }
    return [datetime]$Value
  } catch { return $null }
}

function Convert-ExcelTime {
  param($Value)
  if ($null -eq $Value -or "$Value" -eq "") { return "" }
  try {
    $n = 0.0
    if ([double]::TryParse(($Value.ToString()), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) {
      if ($n -ge 0 -and $n -lt 1) { return ([datetime]"1899-12-30").AddDays($n).ToString("HH:mm") }
    }
    $dt = [datetime]$Value
    return $dt.ToString("HH:mm")
  } catch {
    return (Get-DisplayText $Value)
  }
}

function Convert-Number {
  param($Value)
  if ($null -eq $Value -or "$Value" -eq "") { return $null }
  $n = 0.0
  if ([double]::TryParse(($Value.ToString()), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) {
    return [math]::Round($n, 2)
  }
  return $null
}

function Get-WeekLabel {
  param($Date)
  if ($null -eq $Date) { return "SIN FECHA" }
  $calendar = [System.Globalization.CultureInfo]::InvariantCulture.Calendar
  $week = $calendar.GetWeekOfYear($Date, [System.Globalization.CalendarWeekRule]::FirstFourDayWeek, [DayOfWeek]::Monday)
  return ("{0}-S{1:00}" -f $Date.Year, $week)
}

function Get-WorkbookRows {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { throw "No encontre el archivo Excel: $Path" }

  $tempFile = Join-Path ([System.IO.Path]::GetTempPath()) ("MTC_DASH_{0}.xlsx" -f ([guid]::NewGuid().ToString("N")))
  Copy-Item -LiteralPath $Path -Destination $tempFile -Force
  $records = New-Object System.Collections.Generic.List[object]

  try {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($tempFile)
    try {
      $sharedStrings = New-Object System.Collections.Generic.List[string]
      $sharedEntry = $zip.GetEntry("xl/sharedStrings.xml")
      if ($sharedEntry) {
        [xml]$sharedXml = Read-ZipText $zip "xl/sharedStrings.xml"
        foreach ($si in $sharedXml.sst.si) { [void]$sharedStrings.Add((Get-OpenXmlText $si)) }
      }

      [xml]$workbookXml = Read-ZipText $zip "xl/workbook.xml"
      [xml]$relsXml = Read-ZipText $zip "xl/_rels/workbook.xml.rels"

      foreach ($sheetNode in $workbookXml.workbook.sheets.sheet) {
        $sheetName = [string]$sheetNode.name
        $rid = $sheetNode.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        $rel = $relsXml.Relationships.Relationship | Where-Object { $_.Id -eq $rid } | Select-Object -First 1
        if (-not $rel) { continue }
        $target = $rel.Target -replace "^/", ""
        if ($target -notlike "xl/*") { $target = "xl/$target" }
        $target = $target -replace "\\", "/"
        [xml]$sheetXml = Read-ZipText $zip $target

        $sheetRows = @()
        foreach ($row in $sheetXml.worksheet.sheetData.row) {
          $obj = @{}
          foreach ($cell in $row.c) {
            $obj[(Get-ColIndex $cell.r)] = Get-CellValue $cell $sharedStrings
          }
          $sheetRows += ,$obj
        }
        if ($sheetRows.Count -lt 2) { continue }

        $headerRowIndex = -1
        for ($i = 0; $i -lt [math]::Min(8, $sheetRows.Count); $i++) {
          $headersProbe = @{}
          foreach ($key in $sheetRows[$i].Keys) { $headersProbe[(Normalize-Text $sheetRows[$i][$key])] = [int]$key }
          if ($headersProbe.ContainsKey("OPERADOR") -and $headersProbe.ContainsKey("FECHA")) {
            $headerRowIndex = $i
            break
          }
        }
        if ($headerRowIndex -lt 0) { continue }

        $headers = @{}
        foreach ($key in $sheetRows[$headerRowIndex].Keys) {
          $name = Normalize-Text $sheetRows[$headerRowIndex][$key]
          if (-not [string]::IsNullOrWhiteSpace($name) -and -not $headers.ContainsKey($name)) { $headers[$name] = [int]$key }
        }

        $colOperador = $headers["OPERADOR"]
        $colSerie = $headers["N. SERIE"]
        $colFecha = $headers["FECHA"]
        $colHoraInicio = 4
        $colVacio = $headers["VACIO"]
        $colTanque = $headers["N/TANQUE"]
        $colHoraFin = 7
        $colLleno = $headers["LLENO"]
        $colTiempo = $headers["TIEMPO OPERATIVO"]

        for ($i = $headerRowIndex + 1; $i -lt $sheetRows.Count; $i++) {
          $r = $sheetRows[$i]
          $operador = Get-DisplayText $r[$colOperador]
          $serie = Get-DisplayText $r[$colSerie]
          $fechaObj = Convert-ExcelDate $r[$colFecha]
          $tiempo = Convert-Number $r[$colTiempo]
          $vacio = Get-DisplayText $r[$colVacio]
          $lleno = Get-DisplayText $r[$colLleno]
          $tanque = Get-DisplayText $r[$colTanque]
          if ([string]::IsNullOrWhiteSpace($operador) -and [string]::IsNullOrWhiteSpace($serie) -and $null -eq $fechaObj -and $null -eq $tiempo) { continue }
          if ($null -eq $fechaObj -and $null -eq $tiempo -and [string]::IsNullOrWhiteSpace($tanque)) { continue }

          [void]$records.Add([pscustomobject]@{
            hoja = $sheetName.Trim()
            operador = $operador
            serie = $serie
            fecha = $(if ($fechaObj) { $fechaObj.ToString("yyyy-MM-dd") } else { "" })
            mes = $(if ($fechaObj) { $fechaObj.ToString("yyyy-MM") } else { "SIN FECHA" })
            semana = Get-WeekLabel $fechaObj
            horaInicio = Convert-ExcelTime $r[$colHoraInicio]
            vacio = $vacio
            tanque = $(if ([string]::IsNullOrWhiteSpace($tanque)) { "S/N" } else { $tanque })
            horaFin = Convert-ExcelTime $r[$colHoraFin]
            lleno = $lleno
            tiempoOperativo = $tiempo
          })
        }
      }
    } finally {
      $zip.Dispose()
    }
  } finally {
    Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
  }

  return $records
}

function New-Dashboard {
  $records = Get-WorkbookRows $WorkbookPath
  $generatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  $data = [pscustomobject]@{
    generatedAt = $generatedAt
    sourceFile = [System.IO.Path]::GetFileName($WorkbookPath)
    sourcePath = $WorkbookPath
    records = $records
  }
  $json = $data | ConvertTo-Json -Depth 8 -Compress

  $html = @'
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Dashboard MTC</title>
  <style>
    :root {
      --bg: #e9edf2;
      --panel: #ffffff;
      --panel-dark: #dbeafe;
      --panel-mid: #bfdbfe;
      --ink: #111827;
      --muted: #64748b;
      --line: #cbd5e1;
      --blue: #1d4ed8;
      --green: #047857;
      --amber: #b45309;
      --red: #b91c1c;
      --cyan: #0e7490;
      --radius: 8px;
    }
    * { box-sizing: border-box; }
    body { margin: 0; font-family: Segoe UI, Arial, sans-serif; background: var(--bg); color: var(--ink); letter-spacing: 0; }
    header { background: #111827; color: #fff; border-bottom: 1px solid #0f172a; padding: 16px 20px; display: grid; grid-template-columns: minmax(260px, .75fr) minmax(680px, 1.7fr); gap: 18px; align-items: start; box-shadow: 0 8px 20px rgba(15,23,42,.18); }
    h1 { margin: 0; font-size: 25px; line-height: 1.12; }
    .brand { display: grid; gap: 8px; align-content: start; }
    .brand-subtitle { color: #cbd5e1; font-size: 13px; font-weight: 600; }
    .stamp { font-size: 12px; color: #cbd5e1; line-height: 1.45; }
    main { max-width: 1480px; margin: 0 auto; padding: 18px; }
    .filters { justify-self: end; width: min(100%, 980px); display: grid; grid-template-columns: repeat(4, minmax(135px, 1fr)); gap: 10px; background: #1f2937; border: 1px solid #475569; border-radius: var(--radius); padding: 12px; }
    label { display: grid; gap: 5px; color: #dbe3ed; font-size: 11px; font-weight: 700; text-transform: uppercase; }
    select, input { height: 36px; min-width: 0; border: 1px solid #475569; border-radius: 6px; padding: 0 9px; background: #f8fafc; color: var(--ink); }
    select:focus, input:focus { outline: 3px solid rgba(96,165,250,.28); border-color: #60a5fa; }
    .kpis { display: grid; grid-template-columns: repeat(5, minmax(135px, 1fr)); gap: 12px; margin-bottom: 14px; }
    .kpi, .panel { background: var(--panel); border: 1px solid var(--line); border-radius: var(--radius); box-shadow: 0 8px 22px rgba(15,23,42,.08); overflow: hidden; }
    .kpi { min-height: 88px; padding: 14px 15px; border-left: 5px solid var(--blue); }
    .kpi:nth-child(2) { border-left-color: var(--green); }
    .kpi:nth-child(3) { border-left-color: var(--amber); }
    .kpi:nth-child(4) { border-left-color: var(--cyan); }
    .kpi:nth-child(5) { border-left-color: var(--red); }
    .kpi span { display: block; color: #475569; font-size: 12px; font-weight: 800; margin-bottom: 8px; text-transform: uppercase; }
    .kpi strong { display: block; font-size: 25px; line-height: 1.12; color: #0f172a; }
    .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; align-items: start; }
    .three { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 14px; margin-top: 14px; }
    .panel { min-width: 0; }
    .panel h2 { margin: 0; font-size: 15px; display: flex; justify-content: space-between; gap: 10px; align-items: baseline; background: linear-gradient(90deg, #dbeafe, #e0f2fe); color: #1e3a5f; padding: 11px 13px; border-bottom: 1px solid #bfdbfe; }
    .panel h2 small { color: #475569; font-weight: 600; }
    .panel canvas, .panel .table-wrap { margin: 14px; }
    canvas { display: block; width: calc(100% - 28px); height: 300px; }
    table { width: 100%; border-collapse: collapse; font-size: 12px; }
    th, td { border-bottom: 1px solid #edf1f5; padding: 8px 7px; text-align: left; vertical-align: top; }
    th { position: sticky; top: 0; background: #c7d2fe; color: #1e293b; z-index: 1; }
    td.num, th.num { text-align: right; }
    tbody tr:nth-child(even) { background: #fbfcfd; }
    .table-wrap { max-height: 360px; overflow: auto; border: 1px solid var(--line); border-radius: 6px; }
    .pill { display: inline-block; min-width: 34px; padding: 2px 7px; border-radius: 999px; font-weight: 700; font-size: 11px; text-align: center; }
    .yes { background: #dcfce7; color: #166534; }
    .no { background: #fee2e2; color: #991b1b; }
    @media (max-width: 1180px) { header { grid-template-columns: 1fr; } .filters { justify-self: stretch; width: 100%; } }
    @media (max-width: 1050px) { .filters, .kpis, .grid, .three { grid-template-columns: 1fr 1fr; } }
    @media (max-width: 680px) { header { align-items: flex-start; } .filters, .kpis, .grid, .three { grid-template-columns: 1fr; } canvas { height: 260px; } }
  </style>
</head>
<body>
  <header>
    <div class="brand">
      <h1>Dashboard MTC</h1>
      <div class="brand-subtitle">Control operativo de montacargas</div>
      <div class="stamp" id="stamp"></div>
    </div>
    <section class="filters">
      <label>Operador<select id="operatorFilter"></select></label>
      <label>Serie<select id="serieFilter"></select></label>
      <label>Mes<select id="monthFilter"></select></label>
      <label>Semana<select id="weekFilter"></select></label>
      <label>Vacio<select id="emptyFilter"></select></label>
      <label>Lleno<select id="fullFilter"></select></label>
      <label>Desde<input id="startDateFilter" type="date"></label>
      <label>Buscar<input id="searchBox" type="search" placeholder="tanque, hoja, serie"></label>
    </section>
  </header>
  <main>

    <section class="kpis">
      <div class="kpi"><span>Registros</span><strong id="kpiRecords">0</strong></div>
      <div class="kpi"><span>Operadores</span><strong id="kpiOperators">0</strong></div>
      <div class="kpi"><span>Equipos</span><strong id="kpiSeries">0</strong></div>
      <div class="kpi"><span>Tanques</span><strong id="kpiTanks">0</strong></div>
      <div class="kpi"><span>Ultimo tiempo</span><strong id="kpiLastTime">0</strong></div>
    </section>

    <section class="grid">
      <div class="panel"><h2>Registros por fecha <small id="dateHint"></small></h2><canvas id="dateChart"></canvas></div>
      <div class="panel"><h2>Tiempo operativo por operador <small id="timeHint"></small></h2><canvas id="timeChart"></canvas></div>
    </section>

    <section class="three">
      <div class="panel"><h2>Operadores <small id="operatorHint"></small></h2><canvas id="operatorChart"></canvas></div>
      <div class="panel"><h2>Tanques por operador <small id="tankHint"></small></h2><canvas id="tankChart"></canvas></div>
      <div class="panel"><h2>Resumen por equipo <small id="serieHint"></small></h2><div class="table-wrap"><table><thead><tr><th>Serie</th><th>Operador</th><th class="num">Reg.</th><th class="num">Ultimo tiempo</th></tr></thead><tbody id="serieRows"></tbody></table></div></div>
    </section>

    <section class="panel" style="margin-top:14px"><h2>Detalle <small id="detailHint"></small></h2><div class="table-wrap"><table><thead><tr><th>Fecha</th><th>Operador</th><th>Serie</th><th>Vacio</th><th>Tanque</th><th>Hora</th><th>Lleno</th><th class="num">Tiempo operativo</th><th>Hoja</th></tr></thead><tbody id="detailRows"></tbody></table></div></section>
  </main>

  <script>
    const DATA = __DATA_JSON__;
    const number = new Intl.NumberFormat('es-MX');
    const dec = new Intl.NumberFormat('es-MX', { maximumFractionDigits: 2 });
    const $ = id => document.getElementById(id);
    const colors = ['#60a5fa','#34d399','#fbbf24','#fb7185','#a78bfa','#22d3ee','#f97316','#84cc16','#f472b6','#38bdf8','#c084fc','#2dd4bf'];

    function unique(rows, field) {
      return [...new Set(rows.map(r => r[field]).filter(v => v !== null && v !== undefined && String(v).trim() !== ''))].sort((a,b) => String(a).localeCompare(String(b), 'es'));
    }
    function fillSelect(id, values, label) {
      $(id).innerHTML = `<option value="">${label}</option>` + values.map(v => `<option>${escapeHtml(v)}</option>`).join('');
    }
    function escapeHtml(value) {
      return String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
    }
    function statusPill(value) {
      const text = String(value || '').toUpperCase();
      const cls = text === 'SI' ? 'yes' : 'no';
      return `<span class="pill ${cls}">${escapeHtml(text || 'N/D')}</span>`;
    }
    function filteredRows() {
      const op = $('operatorFilter').value, serie = $('serieFilter').value, month = $('monthFilter').value;
      const week = $('weekFilter').value;
      const empty = $('emptyFilter').value, full = $('fullFilter').value, start = $('startDateFilter').value;
      const q = $('searchBox').value.trim().toUpperCase();
      return DATA.records.filter(r =>
        (!op || r.operador === op) && (!serie || r.serie === serie) && (!month || r.mes === month) &&
        (!week || r.semana === week) &&
        (!empty || r.vacio === empty) && (!full || r.lleno === full) && (!start || r.fecha >= start) &&
        (!q || `${r.hoja} ${r.operador} ${r.serie} ${r.tanque}`.toUpperCase().includes(q))
      );
    }
    function group(rows, field) {
      const map = new Map();
      rows.forEach(r => {
        const key = r[field] || 'SIN DATO';
        const item = map.get(key) || { name: key, registros: 0, tanques: new Set(), maxTiempo: null, minTiempo: null };
        item.registros += 1;
        if (r.tanque && r.tanque !== 'S/N') item.tanques.add(r.tanque);
        if (typeof r.tiempoOperativo === 'number') {
          item.maxTiempo = item.maxTiempo === null ? r.tiempoOperativo : Math.max(item.maxTiempo, r.tiempoOperativo);
          item.minTiempo = item.minTiempo === null ? r.tiempoOperativo : Math.min(item.minTiempo, r.tiempoOperativo);
        }
        map.set(key, item);
      });
      return [...map.values()];
    }
    function byDate(rows) {
      return group(rows.filter(r => r.fecha), 'fecha').sort((a,b) => a.name.localeCompare(b.name));
    }
    function accumulatedByOperator(rows) {
      const byEquipment = new Map();
      rows.filter(r => r.fecha && typeof r.tiempoOperativo === 'number').forEach(r => {
        const operator = r.operador || 'SIN DATO';
        const serie = r.serie || 'SIN SERIE';
        const key = `${operator}||${serie}`;
        const item = byEquipment.get(key) || { operator, serie, readings: [] };
        item.readings.push(r);
        byEquipment.set(key, item);
      });
      const operators = new Map();
      byEquipment.forEach(item => {
        item.readings.sort((a,b) => `${a.fecha} ${a.horaInicio || ''}`.localeCompare(`${b.fecha} ${b.horaInicio || ''}`));
        let accumulated = 0;
        for (let i = 1; i < item.readings.length; i++) {
          const diff = Number(item.readings[i].tiempoOperativo || 0) - Number(item.readings[i - 1].tiempoOperativo || 0);
          if (diff > 0) accumulated += diff;
        }
        const operatorItem = operators.get(item.operator) || { name: item.operator, readings: 0, acumulado: 0 };
        operatorItem.readings += item.readings.length;
        operatorItem.acumulado += accumulated;
        operators.set(item.operator, operatorItem);
      });
      return [...operators.values()]
        .map(item => ({ name: item.name, readings: item.readings, acumulado: Math.round(item.acumulado * 100) / 100 }))
        .sort((a,b) => b.acumulado - a.acumulado || b.readings - a.readings);
    }
    function drawBars(canvas, labels, values, opts = {}) {
      const ctx = canvas.getContext('2d'), dpr = window.devicePixelRatio || 1, rect = canvas.getBoundingClientRect();
      canvas.width = Math.max(320, rect.width * dpr); canvas.height = Math.max(220, rect.height * dpr);
      ctx.setTransform(dpr,0,0,dpr,0,0);
      const w = rect.width, h = rect.height, pad = {l:46,r:14,t:16,b:54}, cw = w-pad.l-pad.r, ch = h-pad.t-pad.b;
      ctx.clearRect(0,0,w,h); ctx.strokeStyle = '#dce3ea'; ctx.lineWidth = 1;
      ctx.beginPath(); for (let i=0;i<=4;i++){ const y=pad.t+ch*i/4; ctx.moveTo(pad.l,y); ctx.lineTo(w-pad.r,y); } ctx.stroke();
      const max = Math.max(1, ...values), slot = cw / Math.max(1, labels.length), bw = Math.max(10, Math.min(42, slot*.5));
      labels.forEach((label,i) => {
        const bh = ch * (values[i] / max), x = pad.l + i*slot + (slot-bw)/2, y = pad.t + ch - bh;
        ctx.fillStyle = opts.colors ? opts.colors[i % opts.colors.length] : colors[i % colors.length]; ctx.fillRect(x,y,bw,bh);
        ctx.save(); ctx.translate(x+bw/2,h-36); ctx.rotate(labels.length > 7 ? -Math.PI/5 : 0); ctx.fillStyle = '#5d6b7a'; ctx.font = '11px Segoe UI, Arial'; ctx.textAlign = labels.length > 7 ? 'right' : 'center'; ctx.fillText(String(label).slice(0,16),0,0); ctx.restore();
      });
      ctx.fillStyle = '#5d6b7a'; ctx.font = '11px Segoe UI, Arial'; ctx.textAlign = 'right'; ctx.fillText(number.format(max), pad.l-7, pad.t+4); ctx.fillText('0', pad.l-7, pad.t+ch);
    }
    function drawPie(canvas, labels, values) {
      const ctx = canvas.getContext('2d'), dpr = window.devicePixelRatio || 1, rect = canvas.getBoundingClientRect();
      canvas.width = Math.max(320, rect.width * dpr); canvas.height = Math.max(220, rect.height * dpr);
      ctx.setTransform(dpr,0,0,dpr,0,0);
      const w = rect.width, h = rect.height;
      ctx.clearRect(0,0,w,h);
      const clean = labels.map((label, i) => ({ label, value: Number(values[i] || 0), color: colors[i % colors.length] })).filter(x => x.value > 0);
      const total = clean.reduce((s, x) => s + x.value, 0);
      const radius = Math.min(w * .28, h * .38, 118);
      const cx = Math.max(radius + 26, w * .33), cy = h * .48;
      if (!total) {
        ctx.fillStyle = '#64748b'; ctx.font = '13px Segoe UI, Arial'; ctx.textAlign = 'center'; ctx.fillText('Sin lecturas para graficar', w / 2, h / 2);
        return;
      }
      let start = -Math.PI / 2;
      clean.forEach(item => {
        const slice = (item.value / total) * Math.PI * 2;
        ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, radius, start, start + slice); ctx.closePath();
        ctx.fillStyle = item.color; ctx.fill();
        ctx.strokeStyle = '#ffffff'; ctx.lineWidth = 2; ctx.stroke();
        start += slice;
      });
      ctx.beginPath(); ctx.arc(cx, cy, radius * .48, 0, Math.PI * 2); ctx.fillStyle = '#ffffff'; ctx.fill();
      ctx.fillStyle = '#0f172a'; ctx.font = '700 18px Segoe UI, Arial'; ctx.textAlign = 'center'; ctx.fillText(dec.format(total), cx, cy - 2);
      ctx.fillStyle = '#64748b'; ctx.font = '11px Segoe UI, Arial'; ctx.fillText('total', cx, cy + 15);
      const lx = Math.min(w - 220, cx + radius + 34), ly = Math.max(28, cy - radius + 8);
      ctx.textAlign = 'left'; ctx.font = '12px Segoe UI, Arial';
      clean.slice(0, 7).forEach((item, i) => {
        const y = ly + i * 26;
        ctx.fillStyle = item.color; ctx.fillRect(lx, y - 10, 12, 12);
        ctx.fillStyle = '#334155'; ctx.fillText(String(item.label).slice(0, 18), lx + 20, y);
        ctx.fillStyle = '#64748b'; ctx.textAlign = 'right'; ctx.fillText(`${Math.round(item.value / total * 100)}%`, w - 20, y);
        ctx.textAlign = 'left';
      });
    }
    function render() {
      const rows = filteredRows();
      const dated = rows.filter(r => r.fecha).sort((a,b) => `${a.fecha} ${a.horaInicio || ''}`.localeCompare(`${b.fecha} ${b.horaInicio || ''}`));
      const times = dated.filter(r => typeof r.tiempoOperativo === 'number');
      const last = times.length ? times[times.length - 1].tiempoOperativo : 0;
      $('kpiRecords').textContent = number.format(rows.length);
      $('kpiOperators').textContent = number.format(unique(rows, 'operador').length);
      $('kpiSeries').textContent = number.format(unique(rows, 'serie').length);
      $('kpiTanks').textContent = number.format(unique(rows, 'tanque').filter(x => x !== 'S/N').length);
      $('kpiLastTime').textContent = dec.format(last);

      const days = byDate(rows);
      drawBars($('dateChart'), days.map(x => x.name), days.map(x => x.registros));
      $('dateHint').textContent = `${number.format(days.length)} dias`;
      const ops = group(rows, 'operador').sort((a,b) => b.registros - a.registros);
      const accumulatedOps = accumulatedByOperator(rows);
      const accumulatedTotal = accumulatedOps.reduce((s, x) => s + x.acumulado, 0);
      drawPie($('timeChart'), accumulatedOps.map(x => x.name), accumulatedOps.map(x => x.acumulado));
      $('timeHint').textContent = `${dec.format(accumulatedTotal)} hrs acumuladas | ${number.format(times.length)} lecturas`;
      drawBars($('operatorChart'), ops.map(x => x.name), ops.map(x => x.registros));
      $('operatorHint').textContent = `${number.format(ops.length)} operadores`;
      drawBars($('tankChart'), ops.map(x => x.name), ops.map(x => x.tanques.size));
      $('tankHint').textContent = 'tanques con numero';

      const series = group(rows, 'serie').sort((a,b) => b.registros - a.registros);
      $('serieRows').innerHTML = series.map(s => {
        const sample = rows.find(r => r.serie === s.name) || {};
        return `<tr><td>${escapeHtml(s.name)}</td><td>${escapeHtml(sample.operador || '')}</td><td class="num">${number.format(s.registros)}</td><td class="num">${s.maxTiempo === null ? '' : dec.format(s.maxTiempo)}</td></tr>`;
      }).join('');
      $('serieHint').textContent = `${number.format(series.length)} equipos`;
      $('detailRows').innerHTML = rows.slice().sort((a,b) => String(b.fecha).localeCompare(String(a.fecha))).slice(0,500).map(r =>
        `<tr><td>${escapeHtml(r.fecha)}</td><td>${escapeHtml(r.operador)}</td><td>${escapeHtml(r.serie)}</td><td>${statusPill(r.vacio)}</td><td>${escapeHtml(r.tanque)}</td><td>${escapeHtml(r.horaInicio)} - ${escapeHtml(r.horaFin)}</td><td>${statusPill(r.lleno)}</td><td class="num">${r.tiempoOperativo === null ? '' : dec.format(r.tiempoOperativo)}</td><td>${escapeHtml(r.hoja)}</td></tr>`
      ).join('');
      $('detailHint').textContent = `${number.format(Math.min(rows.length,500))} de ${number.format(rows.length)} registros`;
    }
    function init() {
      $('stamp').textContent = `Fuente: ${DATA.sourceFile} | Actualizado: ${DATA.generatedAt}`;
      fillSelect('operatorFilter', unique(DATA.records, 'operador'), 'Todos');
      fillSelect('serieFilter', unique(DATA.records, 'serie'), 'Todas');
      fillSelect('monthFilter', unique(DATA.records, 'mes'), 'Todos');
      fillSelect('weekFilter', unique(DATA.records, 'semana'), 'Todas');
      fillSelect('emptyFilter', unique(DATA.records, 'vacio'), 'Todos');
      fillSelect('fullFilter', unique(DATA.records, 'lleno'), 'Todos');
      const dates = unique(DATA.records, 'fecha');
      if (dates.length) { $('startDateFilter').min = dates[0]; $('startDateFilter').max = dates[dates.length - 1]; }
      ['operatorFilter','serieFilter','monthFilter','weekFilter','emptyFilter','fullFilter','startDateFilter','searchBox'].forEach(id => $(id).addEventListener('input', render));
      window.addEventListener('resize', render);
      render();
    }
    init();
  </script>
</body>
</html>
'@

  $html = $html.Replace("__DATA_JSON__", $json)
  $outDir = Split-Path -Parent $OutputPath
  if (-not [string]::IsNullOrWhiteSpace($outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }
  [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.Encoding]::UTF8)
  Write-Host "Dashboard MTC actualizado: $OutputPath"
}

function Start-WatchMode {
  New-Dashboard
  $folder = Split-Path -Parent $WorkbookPath
  $file = Split-Path -Leaf $WorkbookPath
  $watcher = New-Object System.IO.FileSystemWatcher
  $watcher.Path = $folder
  $watcher.Filter = $file
  $watcher.NotifyFilter = [System.IO.NotifyFilters]'LastWrite, Size, FileName'
  $watcher.EnableRaisingEvents = $true
  Write-Host "Vigilando cambios en: $WorkbookPath"
  Write-Host "Deja esta ventana abierta. Cada guardado del Excel actualizara el dashboard."
  while ($true) {
    $event = Wait-Event -SourceIdentifier MtcDashboardChanged -Timeout 1
    if ($event) {
      Remove-Event -EventIdentifier $event.EventIdentifier
      Start-Sleep -Milliseconds 900
      try { New-Dashboard } catch { Write-Warning $_.Exception.Message }
    }
  }
}

if ($Watch) {
  $folder = Split-Path -Parent $WorkbookPath
  $file = Split-Path -Leaf $WorkbookPath
  $watcher = New-Object System.IO.FileSystemWatcher $folder, $file
  $watcher.NotifyFilter = [System.IO.NotifyFilters]'LastWrite, Size, FileName'
  Register-ObjectEvent -InputObject $watcher -EventName Changed -SourceIdentifier MtcDashboardChanged | Out-Null
  Register-ObjectEvent -InputObject $watcher -EventName Created -SourceIdentifier MtcDashboardChangedCreated | Out-Null
  $watcher.EnableRaisingEvents = $true
  New-Dashboard
  Write-Host "Vigilando cambios en: $WorkbookPath"
  Write-Host "Deja esta ventana abierta. Cada guardado del Excel actualizara el dashboard."
  while ($true) {
    $event = Wait-Event -Timeout 2
    if ($event -and ($event.SourceIdentifier -like "MtcDashboardChanged*")) {
      Remove-Event -EventIdentifier $event.EventIdentifier
      Start-Sleep -Milliseconds 900
      try { New-Dashboard } catch { Write-Warning $_.Exception.Message }
    }
  }
} else {
  New-Dashboard
}
