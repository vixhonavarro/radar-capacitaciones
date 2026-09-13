# ============================================================================
#  RADAR DE LICITACIONES DE CAPACITACION - Mercado Publico (ChileCompra)
# ----------------------------------------------------------------------------
#  Consulta la API oficial de Mercado Publico, filtra las licitaciones ACTIVAS
#  relacionadas con capacitacion y genera:
#    - salida\informe_licitaciones_capacitacion.html  (informe compartible)
#    - salida\licitaciones_capacitacion.csv           (para Excel, separador ;)
#  Cachea los detalles en data\cache_detalles.json para que las siguientes
#  ejecuciones sean rapidas (solo consulta licitaciones nuevas).
#
#  Uso:  powershell -NoProfile -ExecutionPolicy Bypass -File buscar_licitaciones.ps1
#  (o doble clic en Actualizar_Licitaciones.cmd)
# ============================================================================

param(
    # Ticket de acceso a la API. El valor por defecto es el ticket PUBLICO DE
    # PRUEBA de ChileCompra (compartido y con limite de consultas). Para uso
    # productivo pida un ticket propio gratis en https://api.mercadopublico.cl
    # y reemplacelo aqui o defina la variable de entorno MP_TICKET.
    [string]$Ticket = $(if ($env:MP_TICKET) { $env:MP_TICKET } else { 'F8537A18-6766-4DEF-9E59-426B4FEE2844' }),
    # Tope de seguridad de consultas de detalle por ejecucion
    [int]$MaxDetalles = 250
)

$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

# Compatibilidad: Windows PowerShell 5.1 (local) usa JavaScriptSerializer;
# PowerShell 7+ (GitHub Actions / Linux) usa ConvertFrom-Json -AsHashtable.
$EsCore = $PSVersionTable.PSEdition -eq 'Core'
if (-not $EsCore) {
    Add-Type -AssemblyName System.Web.Extensions
    $script:Ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $script:Ser.MaxJsonLength = [int]::MaxValue
}
function ConvertFrom-JsonCompat([string]$s) {
    if ($EsCore) { return (ConvertFrom-Json -InputObject $s -AsHashtable) }
    return $script:Ser.DeserializeObject($s)
}
function ConvertTo-JsonCompat($obj) {
    if ($EsCore) { return (ConvertTo-Json -InputObject $obj -Depth 10 -Compress) }
    return $script:Ser.Serialize($obj)
}

# --- Palabras clave (regex sobre el nombre normalizado: minusculas y sin tildes).
#     \b evita falsos positivos como "INFORMACION" (formacion) o "CONCURSO" (curso).
$PatronClaves = '\bcapacit|\bcurso|\btaller(es)? (de|en|para)\b|\bformacion|\bdiplomado|\bentrenamiento|\bcoaching|\brelator|e-?learning|\bmentoria|\bperfeccionamiento|\bseminario|\bcharla'

$Base       = Split-Path -Parent $MyInvocation.MyCommand.Path
$DirSalida  = Join-Path $Base 'salida'
$DirData    = Join-Path $Base 'data'
$RutaCache  = Join-Path $DirData 'cache_detalles.json'
$RutaHtml   = Join-Path $DirSalida 'informe_licitaciones_capacitacion.html'
$RutaCsv    = Join-Path $DirSalida 'licitaciones_capacitacion.csv'
foreach ($d in @($DirSalida, $DirData)) { if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d | Out-Null } }

$CulturaCL = [Globalization.CultureInfo]::GetCultureInfo('es-CL')
$Inv       = [Globalization.CultureInfo]::InvariantCulture

# ---------------------------------------------------------------- utilidades
function Get-ApiJson([string]$Url, [int]$Reintentos = 4) {
    $espera = 8
    for ($i = 1; $i -le $Reintentos; $i++) {
        try {
            $wc = New-Object System.Net.WebClient
            $bytes = $wc.DownloadData($Url)
            $texto = [Text.Encoding]::UTF8.GetString($bytes)
            $obj = ConvertFrom-JsonCompat $texto
            if ($obj -is [System.Collections.IDictionary] -and $obj.ContainsKey('Listado')) { return $obj }
            # La API devuelve {"Codigo":...,"Mensaje":"..."} cuando hay error/limite
            Write-Host ("   aviso API (intento {0}): {1}" -f $i, $texto.Substring(0, [Math]::Min(120, $texto.Length)))
        } catch {
            Write-Host ("   error red (intento {0}): {1}" -f $i, $_.Exception.Message)
        }
        if ($i -lt $Reintentos) { Start-Sleep -Seconds $espera; $espera = $espera * 2 }
    }
    return $null
}

function Normalizar([string]$s) {
    # minusculas + quita tildes/enie descomponiendo Unicode (FormD) y
    # eliminando las marcas diacriticas; asi el .ps1 se mantiene 100% ASCII
    if ($null -eq $s) { return '' }
    $d = $s.ToLower().Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object Text.StringBuilder
    foreach ($ch in $d.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
    }
    return $sb.ToString()
}

function EscHtml([string]$s) {
    if ($null -eq $s) { return '' }
    return $s.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('"','&quot;')
}

function ParseFecha($s) {
    if ($null -eq $s -or '' -eq ('' + $s)) { return $null }
    try { return [datetime]::Parse(('' + $s), $Inv) } catch { return $null }
}

# --- Modalidad de la capacitacion, detectada en el texto (nombre + descripcion
#     normalizados). "Online" implica que puede dictarse desde cualquier lugar.
$PatronMixta      = 'semi-?presencial|b-?learning|hibrid|modalidad mixta'
$PatronOnline     = '\bonline\b|\bon-? ?line\b|e-?learning|\ba distancia\b|\bremot[oa]s?\b|\bvirtual(es)?\b|telematic|asincronic|sincronic|autoinstruccion|autoaprendizaje|videoconferencia|\bstreaming\b|\bzoom\b|\bteams\b|plataforma (virtual|online|digital|e-?learning)'
$PatronPresencial = 'presencial|\bin situ\b|en (las |sus )?dependencias|en terreno|\ben sala\b|\ben aula\b|lugar de (la )?ejecucion|coffee break|coffe break|cofee break|arriendo (de )?(salon|local|sala)|alojamiento|en la ciudad de|habilitacion de espacios|en (diferentes|distintos) sectores'

function Detectar-Modalidad([string]$textoNorm) {
    # 'semipresencial' tambien matchea el patron presencial: mixta se evalua primero
    if ($textoNorm -match $PatronMixta) { return 'Mixta' }
    $esOnline = $textoNorm -match $PatronOnline
    $esPresencial = $textoNorm -match $PatronPresencial
    if ($esOnline -and $esPresencial) { return 'Mixta' }
    if ($esOnline) { return 'Online' }
    if ($esPresencial) { return 'Presencial' }
    return 'No especificada'
}

# --------------------------------------------------- 1) licitaciones activas
Write-Host '1) Descargando licitaciones activas...'
$activas = Get-ApiJson ("https://api.mercadopublico.cl/servicios/v1/publico/licitaciones.json?estado=activas&ticket={0}" -f $Ticket)
if ($null -eq $activas) { throw 'No fue posible obtener las licitaciones activas desde la API.' }
Write-Host ("   {0} licitaciones activas en total." -f $activas['Cantidad'])

# ------------------------------------------------ 2) filtro por palabra clave
$coincidencias = @()
foreach ($lic in $activas['Listado']) {
    $nom = Normalizar ('' + $lic['Nombre'])
    $m = [regex]::Match($nom, $PatronClaves)
    if ($m.Success) {
        $coincidencias += New-Object PSObject -Property @{
            Codigo       = '' + $lic['CodigoExterno']
            Nombre       = '' + $lic['Nombre']
            FechaCierre  = ParseFecha $lic['FechaCierre']
            PalabraClave = $m.Value.Trim()
        }
    }
}
Write-Host ("2) {0} licitaciones coinciden con las palabras clave." -f $coincidencias.Count)

# ------------------------------------------------------- 3) detalles (+cache)
$cache = @{}
if (Test-Path $RutaCache) {
    try {
        $tmp = ConvertFrom-JsonCompat ([IO.File]::ReadAllText($RutaCache, [Text.Encoding]::UTF8))
        foreach ($k in $tmp.Keys) { $cache[$k] = $tmp[$k] }
    } catch { Write-Host '   (cache ilegible, se regenera)' }
}

$pendientes = @($coincidencias | Where-Object { -not $cache.ContainsKey($_.Codigo) })
if ($pendientes.Count -gt $MaxDetalles) { $pendientes = @($pendientes | Select-Object -First $MaxDetalles) }
Write-Host ("3) Consultando detalle de {0} licitaciones nuevas ({1} ya en cache)..." -f $pendientes.Count, ($coincidencias.Count - $pendientes.Count))

function Guardar-Cache {
    [IO.File]::WriteAllText($RutaCache, (ConvertTo-JsonCompat $cache), (New-Object Text.UTF8Encoding $false))
}

$n = 0
foreach ($p in $pendientes) {
    $n++
    $det = Get-ApiJson ("https://api.mercadopublico.cl/servicios/v1/publico/licitaciones.json?codigo={0}&ticket={1}" -f $p.Codigo, $Ticket) 3
    if ($null -ne $det -and $det['Listado'] -and $det['Listado'].Count -gt 0) {
        $d = $det['Listado'][0]
        $comp = $d['Comprador']; $fechas = $d['Fechas']
        $rubro86 = $false
        if ($d.ContainsKey('Items') -and $null -ne $d['Items'] -and $null -ne $d['Items']['Listado']) {
            foreach ($it in $d['Items']['Listado']) {
                if (('' + $it['CodigoCategoria']).StartsWith('86')) { $rubro86 = $true; break }
            }
        }
        $desc = ('' + $d['Descripcion']) -replace '\s+', ' '
        if ($desc.Length -gt 700) { $desc = $desc.Substring(0, 700) + '...' }
        $cache[$p.Codigo] = @{
            Organismo        = '' + $comp['NombreOrganismo']
            Unidad           = '' + $comp['NombreUnidad']
            Region           = ('' + $comp['RegionUnidad']).Trim()
            Comuna           = '' + $comp['ComunaUnidad']
            Tipo             = '' + $d['Tipo']
            Moneda           = '' + $d['Moneda']
            MontoEstimado    = $d['MontoEstimado']
            VisibilidadMonto = $d['VisibilidadMonto']
            FechaPublicacion = '' + $fechas['FechaPublicacion']
            Descripcion      = $desc
            RubroEducacion   = $rubro86
        }
    } else {
        Write-Host ("   sin detalle para {0}" -f $p.Codigo)
    }
    if ($n % 10 -eq 0) { Write-Host ("   {0}/{1}..." -f $n, $pendientes.Count); Guardar-Cache }
    Start-Sleep -Milliseconds 2200
}
Guardar-Cache

# ------------------------------------------------------ 4) registros finales
$ahora = Get-Date
$registros = @()
foreach ($c in $coincidencias) {
    if ($null -ne $c.FechaCierre -and $c.FechaCierre -lt $ahora) { continue }  # ya cerradas
    $det = $null
    if ($cache.ContainsKey($c.Codigo)) { $det = $cache[$c.Codigo] }
    $dias = $null
    if ($null -ne $c.FechaCierre) { $dias = [int][Math]::Ceiling(($c.FechaCierre - $ahora).TotalDays) }
    $monto = $null; $moneda = 'CLP'; $montoPublicado = $false
    $org = ''; $region = ''; $comuna = ''; $tipo = ''; $descr = ''; $fpub = $null; $rubro = $false; $unidad = ''
    if ($null -ne $det) {
        $org = '' + $det['Organismo']; $region = '' + $det['Region']; $tipo = '' + $det['Tipo']
        $comuna = ('' + $det['Comuna']).Trim()
        $descr = '' + $det['Descripcion']; $fpub = ParseFecha $det['FechaPublicacion']
        $rubro = [bool]$det['RubroEducacion']; $unidad = '' + $det['Unidad']
        if ('' + $det['Moneda'] -ne '') { $moneda = '' + $det['Moneda'] }
        if (('' + $det['VisibilidadMonto']) -eq '1' -and $null -ne $det['MontoEstimado']) {
            $m = 0.0
            if ([double]::TryParse(('' + $det['MontoEstimado']), [Globalization.NumberStyles]::Any, $Inv, [ref]$m) -and $m -gt 1) {
                $monto = $m; $montoPublicado = $true
            }
        }
    }
    $modalidad = Detectar-Modalidad (Normalizar ($c.Nombre + ' ' + $descr))
    $registros += New-Object PSObject -Property @{
        Codigo = $c.Codigo; Nombre = $c.Nombre; FechaCierre = $c.FechaCierre; Dias = $dias
        PalabraClave = $c.PalabraClave; Organismo = $org; Unidad = $unidad; Region = $region
        Comuna = $comuna; Modalidad = $modalidad
        Tipo = $tipo; Monto = $monto; Moneda = $moneda; MontoPublicado = $montoPublicado
        Descripcion = $descr; FechaPublicacion = $fpub; RubroEducacion = $rubro
        Link = ('https://www.mercadopublico.cl/fichaLicitacion.html?idLicitacion=' + $c.Codigo)
    }
}
$registros = @($registros | Sort-Object @{Expression={ if ($null -eq $_.FechaCierre) { [datetime]::MaxValue } else { $_.FechaCierre } }})
Write-Host ("4) {0} licitaciones abiertas en el informe." -f $registros.Count)

# ---------------------------------------------------------------- 5) CSV
$registros | Select-Object Codigo, Nombre, Organismo, Unidad, Region, Comuna, Tipo, Modalidad,
    @{n='FechaPublicacion'; e={ if ($_.FechaPublicacion) { $_.FechaPublicacion.ToString('dd-MM-yyyy') } else { '' } }},
    @{n='FechaCierre'; e={ if ($_.FechaCierre) { $_.FechaCierre.ToString('dd-MM-yyyy HH:mm') } else { '' } }},
    @{n='DiasRestantes'; e={ $_.Dias }},
    @{n='MontoEstimado'; e={ if ($_.MontoPublicado) { [long]$_.Monto } else { '' } }},
    Moneda, PalabraClave,
    @{n='RubroEducacionONU'; e={ if ($_.RubroEducacion) { 'SI' } else { '' } }},
    Link, Descripcion |
    Export-Csv -Path $RutaCsv -NoTypeInformation -Delimiter ';' -Encoding UTF8

# ---------------------------------------------------------------- 6) HTML
$kpiTotal  = $registros.Count
$kpiSemana = @($registros | Where-Object { $null -ne $_.Dias -and $_.Dias -le 7 }).Count
$sumaClp = 0.0
$conMonto = 0
foreach ($r in $registros) { if ($r.MontoPublicado -and $r.Moneda -eq 'CLP') { $sumaClp += $r.Monto; $conMonto++ } }
$kpiMonto = '$' + $sumaClp.ToString('N0', $CulturaCL)

$kpiOnline = @($registros | Where-Object { $_.Modalidad -eq 'Online' }).Count

$regiones = @($registros | ForEach-Object { $_.Region } | Where-Object { $_ -ne '' } | Sort-Object -Unique)
$sbReg = New-Object Text.StringBuilder
foreach ($rg in $regiones) { [void]$sbReg.AppendFormat('<option value="{0}">{0}</option>', (EscHtml $rg)) }

$comunas = @($registros | ForEach-Object { $_.Comuna } | Where-Object { $_ -ne '' } | Sort-Object -Unique)
$sbCom = New-Object Text.StringBuilder
foreach ($cm in $comunas) { [void]$sbCom.AppendFormat('<option value="{0}">{0}</option>', (EscHtml $cm)) }

$sbFilas = New-Object Text.StringBuilder
foreach ($r in $registros) {
    $cierreTxt = '&mdash;'
    if ($null -ne $r.FechaCierre) { $cierreTxt = $r.FechaCierre.ToString('dd-MM-yyyy HH:mm') }
    $chipClase = 'd-ok'; $chipTxt = 'sin fecha'
    if ($null -ne $r.Dias) {
        if     ($r.Dias -le 0) { $chipTxt = 'hoy';                          $chipClase = 'd-crit' }
        elseif ($r.Dias -eq 1) { $chipTxt = 'ma&ntilde;ana';                $chipClase = 'd-crit' }
        elseif ($r.Dias -le 2) { $chipTxt = ('' + $r.Dias + ' d&iacute;as'); $chipClase = 'd-crit' }
        elseif ($r.Dias -le 7) { $chipTxt = ('' + $r.Dias + ' d&iacute;as'); $chipClase = 'd-warn' }
        else                   { $chipTxt = ('' + $r.Dias + ' d&iacute;as'); $chipClase = 'd-ok' }
    }
    $montoTxt = '<span class="sinmonto">No publicado</span>'
    if ($r.MontoPublicado) {
        if ($r.Moneda -eq 'CLP') { $montoTxt = '$' + $r.Monto.ToString('N0', $CulturaCL) }
        else { $montoTxt = $r.Monto.ToString('N0', $CulturaCL) + ' ' + (EscHtml $r.Moneda) }
    }
    $tipoTxt = 'n/d'; if ($r.Tipo -ne '') { $tipoTxt = EscHtml $r.Tipo }
    $orgTxt = '(detalle no disponible)'; if ($r.Organismo -ne '') { $orgTxt = EscHtml $r.Organismo }
    $lugarHtml = EscHtml $r.Region
    if ($r.Comuna -ne '') { $lugarHtml = (EscHtml $r.Comuna) + ' &middot; ' + (EscHtml $r.Region) }
    $modClase = 'm-nd'
    if     ($r.Modalidad -eq 'Online')     { $modClase = 'm-online' }
    elseif ($r.Modalidad -eq 'Presencial') { $modClase = 'm-pres' }
    elseif ($r.Modalidad -eq 'Mixta')      { $modClase = 'm-mixta' }
    $descHtml = ''
    if ($r.Descripcion -ne '') {
        $descHtml = '<details><summary>Ver descripci&oacute;n</summary><p>' + (EscHtml $r.Descripcion) + '</p></details>'
    }
    $rubroHtml = ''
    if ($r.RubroEducacion) { $rubroHtml = ' <span class="rubro" title="Incluye items del rubro ONU 86: Educacion y Capacitacion">rubro 86</span>' }
    $buscable = Normalizar ($r.Codigo + ' ' + $r.Nombre + ' ' + $r.Organismo + ' ' + $r.Comuna + ' ' + $r.Region + ' ' + $r.Tipo + ' ' + $r.Modalidad)
    [void]$sbFilas.AppendFormat(@'
<tr data-buscar="{0}" data-region="{1}" data-comuna="{2}" data-mod="{3}" data-dias="{4}">
<td class="c-cierre"><span class="fecha">{5}</span><span class="chip {6}">{7}</span></td>
<td class="c-lic"><a href="{8}" target="_blank" rel="noopener">{9}</a><span class="codigo">{10}</span>{11}{12}</td>
<td class="c-org">{13}<span class="region">{14}</span></td>
<td class="c-tipo"><span class="tipo">{15}</span></td>
<td class="c-mod"><span class="mod {16}">{17}</span></td>
<td class="c-monto">{18}</td>
</tr>
'@, (EscHtml $buscable), (EscHtml $r.Region), (EscHtml $r.Comuna), (EscHtml $r.Modalidad),
        $(if ($null -ne $r.Dias) { $r.Dias } else { 9999 }),
        $cierreTxt, $chipClase, $chipTxt, (EscHtml $r.Link), (EscHtml $r.Nombre), (EscHtml $r.Codigo),
        $rubroHtml, $descHtml, $orgTxt, $lugarHtml, $tipoTxt, $modClase, (EscHtml $r.Modalidad), $montoTxt)
}

$plantilla = @'
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Radar de Capacitaciones</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Archivo:wght@600;700;800&family=Public+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap">
<style>
:root{
  --ground:#f7f9fa; --surface:#ffffff; --ink:#1a2733; --ink-2:#5b6b77; --ink-3:#8195a1;
  --linea:#dde5ea; --acento:#0e6e8c; --acento-tenue:#e3eff4;
  --warn-bg:#faeede; --warn-tx:#8f4208; --crit-bg:#f9e7e5; --crit-tx:#8c231e;
  --ok-bg:#eef2f4; --ok-tx:#4d5f6b; --rubro-bg:#e7f1e9; --rubro-tx:#2b6a3f;
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    --ground:#10181e; --surface:#17222a; --ink:#e4ebf0; --ink-2:#9db0bc; --ink-3:#6d8290;
    --linea:#26343e; --acento:#5cb8d6; --acento-tenue:#143240;
    --warn-bg:#3a2712; --warn-tx:#f0b06a; --crit-bg:#3d1a17; --crit-tx:#f09a92;
    --ok-bg:#1e2b33; --ok-tx:#a7bac5; --rubro-bg:#16301f; --rubro-tx:#83c99a;
  }
}
:root[data-theme="dark"]{
  --ground:#10181e; --surface:#17222a; --ink:#e4ebf0; --ink-2:#9db0bc; --ink-3:#6d8290;
  --linea:#26343e; --acento:#5cb8d6; --acento-tenue:#143240;
  --warn-bg:#3a2712; --warn-tx:#f0b06a; --crit-bg:#3d1a17; --crit-tx:#f09a92;
  --ok-bg:#1e2b33; --ok-tx:#a7bac5; --rubro-bg:#16301f; --rubro-tx:#83c99a;
}
*{box-sizing:border-box}
body{margin:0;background:var(--ground);color:var(--ink);
  font:15px/1.55 "Public Sans",-apple-system,"Segoe UI",Roboto,sans-serif}
.wrap{max-width:1100px;margin:0 auto;padding:28px 20px 60px}
header{margin-bottom:22px}
.eyebrow{font-size:12px;font-weight:600;letter-spacing:.09em;text-transform:uppercase;color:var(--acento);margin-bottom:6px}
h1{font-family:"Archivo","Public Sans",sans-serif;font-weight:800;font-size:30px;line-height:1.15;margin:0 0 8px;text-wrap:balance}
.sub{color:var(--ink-2);margin:0;max-width:68ch}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px;margin:22px 0 18px}
.kpi{background:var(--surface);border:1px solid var(--linea);border-radius:8px;padding:14px 16px}
.kpi .num{font-family:"Archivo",sans-serif;font-weight:700;font-size:26px;font-variant-numeric:tabular-nums}
.kpi .lbl{font-size:12.5px;color:var(--ink-2);margin-top:2px}
.controles{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-bottom:12px}
.controles input,.controles select{font:inherit;color:var(--ink);background:var(--surface);
  border:1px solid var(--linea);border-radius:6px;padding:8px 10px}
.controles input{flex:1 1 240px;min-width:200px}
.controles input:focus,.controles select:focus{outline:2px solid var(--acento);outline-offset:1px}
#contador{font-size:13px;color:var(--ink-2);margin-left:auto;font-variant-numeric:tabular-nums}
.tablewrap{overflow-x:auto;background:var(--surface);border:1px solid var(--linea);border-radius:8px}
table{border-collapse:collapse;width:100%;min-width:940px}
th{font-size:11.5px;font-weight:600;letter-spacing:.07em;text-transform:uppercase;color:var(--ink-3);
  text-align:left;padding:10px 14px;border-bottom:1px solid var(--linea);background:var(--surface);position:sticky;top:0}
td{padding:12px 14px;border-bottom:1px solid var(--linea);vertical-align:top}
tr:last-child td{border-bottom:none}
.c-cierre{white-space:nowrap;width:150px}
.c-cierre .fecha{display:block;font-variant-numeric:tabular-nums;font-weight:500}
.chip{display:inline-block;margin-top:4px;font-size:11.5px;font-weight:600;padding:2px 8px;border-radius:999px}
.d-crit{background:var(--crit-bg);color:var(--crit-tx)}
.d-warn{background:var(--warn-bg);color:var(--warn-tx)}
.d-ok{background:var(--ok-bg);color:var(--ok-tx)}
.c-lic a{color:var(--acento);font-weight:600;text-decoration:none}
.c-lic a:hover,.c-lic a:focus{text-decoration:underline}
.codigo{display:block;font:12px "IBM Plex Mono",Consolas,monospace;color:var(--ink-3);margin-top:3px}
.rubro{display:inline-block;font-size:11px;font-weight:600;background:var(--rubro-bg);color:var(--rubro-tx);
  padding:1px 7px;border-radius:999px;margin-left:6px;vertical-align:1px}
details{margin-top:6px}
summary{font-size:12.5px;color:var(--ink-2);cursor:pointer}
details p{font-size:13px;color:var(--ink-2);margin:6px 0 0;max-width:70ch}
.c-org{color:var(--ink);font-size:13.5px}
.c-org .region{display:block;font-size:12px;color:var(--ink-3);margin-top:2px}
.c-tipo .tipo{font:12px "IBM Plex Mono",Consolas,monospace;background:var(--acento-tenue);color:var(--acento);
  padding:2px 7px;border-radius:4px}
.c-mod{white-space:nowrap}
.c-mod .mod{display:inline-block;font-size:11.5px;font-weight:600;padding:2px 8px;border-radius:999px}
.m-online{background:var(--acento-tenue);color:var(--acento)}
.m-pres{background:var(--ok-bg);color:var(--ok-tx)}
.m-mixta{background:var(--ok-bg);color:var(--ok-tx)}
.m-nd{background:none;color:var(--ink-3);font-weight:400;padding-left:0}
.c-monto{text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums;font-weight:600}
.sinmonto{color:var(--ink-3);font-weight:400;font-size:13px}
footer{margin-top:26px;color:var(--ink-2);font-size:13px}
footer h2{font-family:"Archivo",sans-serif;font-size:15px;color:var(--ink);margin:18px 0 6px}
footer dl{display:grid;grid-template-columns:auto 1fr;gap:4px 12px;margin:0}
footer dt{font:12px "IBM Plex Mono",Consolas,monospace;color:var(--acento)}
footer dd{margin:0}
.vacio{padding:28px;text-align:center;color:var(--ink-2)}
@media (max-width:640px){h1{font-size:24px}.wrap{padding:20px 12px 40px}}
</style>
<div class="wrap">
<header>
  <div class="eyebrow">Mercado P&uacute;blico &middot; ChileCompra</div>
  <h1>Radar de Capacitaciones</h1>
  <p class="sub">Licitaciones p&uacute;blicas <strong>abiertas</strong> relacionadas con capacitaci&oacute;n, cursos, formaci&oacute;n y relator&iacute;as, ordenadas por fecha de cierre. Actualizado el <strong>__GENERADO__</strong> &middot; Fuente: API oficial de Mercado P&uacute;blico.</p>
</header>
<section class="kpis">
  <div class="kpi"><div class="num">__KPI_TOTAL__</div><div class="lbl">licitaciones abiertas</div></div>
  <div class="kpi"><div class="num">__KPI_SEMANA__</div><div class="lbl">cierran en 7 d&iacute;as o menos</div></div>
  <div class="kpi"><div class="num">__KPI_ONLINE__</div><div class="lbl">en modalidad online (sin restricci&oacute;n geogr&aacute;fica)</div></div>
  <div class="kpi"><div class="num">__KPI_MONTO__</div><div class="lbl">monto estimado publicado (CLP, __KPI_CONMONTO__ licitaciones)</div></div>
</section>
<section class="controles">
  <input id="buscar" type="search" placeholder="Buscar por nombre, organismo o c&oacute;digo..." aria-label="Buscar licitaciones">
  <select id="region" aria-label="Filtrar por regi&oacute;n"><option value="">Todas las regiones</option>__REGIONES__</select>
  <select id="comuna" aria-label="Filtrar por comuna"><option value="">Todas las comunas</option>__COMUNAS__</select>
  <select id="modalidad" aria-label="Filtrar por modalidad">
    <option value="">Cualquier modalidad</option>
    <option value="Online">Online</option>
    <option value="Presencial">Presencial</option>
    <option value="Mixta">Mixta</option>
    <option value="No especificada">No especificada</option>
  </select>
  <select id="plazo" aria-label="Filtrar por plazo de cierre">
    <option value="">Cualquier plazo</option>
    <option value="7">Cierran en 7 d&iacute;as o menos</option>
    <option value="3">Cierran en 3 d&iacute;as o menos</option>
    <option value="14">Cierran en 14 d&iacute;as o menos</option>
  </select>
  <span id="contador"></span>
</section>
<div class="tablewrap">
<table>
<thead><tr><th>Cierre</th><th>Licitaci&oacute;n</th><th>Organismo y ubicaci&oacute;n</th><th>Tipo</th><th>Modalidad</th><th>Monto estimado</th></tr></thead>
<tbody id="cuerpo">
__FILAS__
</tbody>
</table>
<div id="vacio" class="vacio" hidden>Ninguna licitaci&oacute;n coincide con los filtros.</div>
</div>
<footer>
  <h2>Tipos de licitaci&oacute;n (tramos referenciales en UTM)</h2>
  <dl>
    <dt>L1</dt><dd>Menor a 100 UTM</dd>
    <dt>LE</dt><dd>Entre 100 y 1.000 UTM</dd>
    <dt>LP</dt><dd>Entre 1.000 y 2.000 UTM</dd>
    <dt>LQ</dt><dd>Entre 2.000 y 5.000 UTM</dd>
    <dt>LR</dt><dd>Sobre 5.000 UTM</dd>
    <dt>CO</dt><dd>Compra coordinada</dd>
  </dl>
  <h2>Sobre la ubicaci&oacute;n</h2>
  <p>La comuna y regi&oacute;n mostradas corresponden a la <strong>unidad compradora</strong>, que es donde normalmente se dicta una capacitaci&oacute;n presencial. La <strong>modalidad</strong> (Online / Presencial / Mixta) se detecta autom&aacute;ticamente en el texto de cada licitaci&oacute;n: las marcadas <em>Online</em> pueden dictarse desde cualquier lugar. El lugar exacto de ejecuci&oacute;n se confirma siempre en las bases.</p>
  <h2>C&oacute;mo postular</h2>
  <p>Cada enlace abre la ficha p&uacute;blica de la licitaci&oacute;n en mercadopublico.cl, con sus bases, plazos y anexos. Para ofertar, el proveedor debe estar inscrito en el Registro de Proveedores del Estado (proveedores.mercadopublico.cl). La etiqueta <em>rubro 86</em> indica que la licitaci&oacute;n incluye &iacute;tems clasificados en el rubro ONU &laquo;Educaci&oacute;n y Capacitaci&oacute;n&raquo;.</p>
  <p>Informe generado autom&aacute;ticamente a partir de la API p&uacute;blica de Mercado P&uacute;blico. Los montos y plazos son referenciales; verifique siempre las bases oficiales. <a href="licitaciones_capacitacion.csv" download>Descargar los datos en CSV (Excel)</a>.</p>
</footer>
</div>
<script>
(function(){
  var buscar = document.getElementById('buscar');
  var region = document.getElementById('region');
  var comuna = document.getElementById('comuna');
  var modalidad = document.getElementById('modalidad');
  var plazo  = document.getElementById('plazo');
  var filas  = Array.prototype.slice.call(document.querySelectorAll('#cuerpo tr'));
  var contador = document.getElementById('contador');
  var vacio = document.getElementById('vacio');
  function norm(s){
    s = (s || '').toLowerCase();
    try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch(e) {}
    return s;
  }
  function aplicar(){
    var q = norm(buscar.value), reg = region.value, com = comuna.value, mod = modalidad.value,
        pl = plazo.value ? parseInt(plazo.value, 10) : null;
    var visibles = 0;
    filas.forEach(function(tr){
      var ok = true;
      if (q && tr.getAttribute('data-buscar').indexOf(q) === -1) ok = false;
      if (ok && reg && tr.getAttribute('data-region') !== reg) ok = false;
      if (ok && com && tr.getAttribute('data-comuna') !== com) ok = false;
      if (ok && mod && tr.getAttribute('data-mod') !== mod) ok = false;
      if (ok && pl !== null && parseInt(tr.getAttribute('data-dias'), 10) > pl) ok = false;
      tr.hidden = !ok;
      if (ok) visibles++;
    });
    contador.textContent = visibles + ' de ' + filas.length + ' licitaciones';
    vacio.hidden = visibles !== 0;
  }
  buscar.addEventListener('input', aplicar);
  region.addEventListener('change', aplicar);
  comuna.addEventListener('change', aplicar);
  modalidad.addEventListener('change', aplicar);
  plazo.addEventListener('change', aplicar);
  aplicar();
})();
</script>
'@

$html = $plantilla.Replace('__GENERADO__', $ahora.ToString('dd-MM-yyyy HH:mm')).
    Replace('__KPI_TOTAL__', ('' + $kpiTotal)).
    Replace('__KPI_SEMANA__', ('' + $kpiSemana)).
    Replace('__KPI_MONTO__', $kpiMonto).
    Replace('__KPI_CONMONTO__', ('' + $conMonto)).
    Replace('__KPI_ONLINE__', ('' + $kpiOnline)).
    Replace('__REGIONES__', $sbReg.ToString()).
    Replace('__COMUNAS__', $sbCom.ToString()).
    Replace('__FILAS__', $sbFilas.ToString())
[IO.File]::WriteAllText($RutaHtml, $html, (New-Object Text.UTF8Encoding $false))

# Si existe la carpeta de entrega en el Escritorio, dejar alli copia fresca
# (en GitHub Actions no hay Escritorio: se omite)
$Escritorio = [Environment]::GetFolderPath('Desktop')
if ($Escritorio) {
    $DirEntrega = Join-Path $Escritorio 'mercado publico'
    if (Test-Path $DirEntrega) {
        Copy-Item $RutaHtml $DirEntrega -Force
        Copy-Item $RutaCsv  $DirEntrega -Force
        Write-Host (' Copia para enviar actualizada en: {0}' -f $DirEntrega)
    }
}

Write-Host ''
Write-Host '============================================================'
Write-Host (' Listo: {0} licitaciones de capacitacion abiertas' -f $registros.Count)
Write-Host (' Informe HTML : {0}' -f $RutaHtml)
Write-Host (' CSV (Excel)  : {0}' -f $RutaCsv)
Write-Host '============================================================'
