<#
  ============================================================================
   INFORME_PC.ps1  -  Aplicacion de diagnostico de hardware y software
  ============================================================================
   QUE HACE
     - Corre 10 secciones de pruebas sobre el equipo (hardware, rendimiento,
       procesos sospechosos, arranque, seguridad, drivers, eventos, energia,
       disco) y genera un reporte completo en HTML (para leer/imprimir) y TXT
       (para copiar y pegar).
     - Trae una pestana de MONITOREO DE PICOS: graba cada 2 s que proceso
       consume la CPU, con boton "marcar ventilador ahora" para anotar el
       momento exacto en que escuchas subir el ventilador.
     - Revisa si hay que ACTUALIZAR LOS CONTROLADORES (drivers): dispositivos
       sin controlador o con error, los controladores mas antiguos y los que
       Windows Update tiene disponibles para este equipo.
     - Se auto-eleva a Administrador (varias pruebas lo necesitan).
   NO MODIFICA NADA del sistema: solo lee. Lo unico que escribe es la carpeta
   de reportes (Escritorio\InformePC_...).
   ============================================================================
   COMO USARLA
     - Doble clic en  Ejecutar_Informe_PC.bat      -> interfaz grafica
     - Modo consola (sin ventana), reporte + resumen en pantalla:
         powershell -NoProfile -ExecutionPolicy Bypass -File .\INFORME_PC.ps1 -Auto
     - Opciones utiles:
         -MuestreoSegundos 90     muestreo de CPU mas largo (default 45)
         -TestDisco               agrega prueba de velocidad del SSD
         -ActualizacionesPendientes  consulta Windows Update: actualizaciones y drivers
         -DriversPendientes       consulta Windows Update solo para ver los drivers
                                  (equivale a la casilla de la interfaz grafica)
         -MonitorSegundos 600     modo consola: monitorea 10 minutos y graba CSV
         -OutDir "C:\Temp\rep"    carpeta de salida
         -NoElevar                no pedir elevacion (para pruebas)
  ============================================================================
#>
param(
    [switch]$Auto,
    [switch]$NoElevar,
    [int]$MuestreoSegundos = 45,
    [int]$MonitorSegundos  = 0,
    [int]$MonitorUmbral    = 20,
    [switch]$TestDisco,
    [switch]$ActualizacionesPendientes,
    [switch]$DriversPendientes,
    [string]$OutDir,
    [switch]$TestGUI
)

$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference    = 'SilentlyContinue'

$esAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# --- Auto-elevacion (solo en modo grafico; el .bat ya eleva en modo consola) ---
if (-not $esAdmin -and -not $NoElevar -and -not $Auto) {
    try {
        $arg = '-NoProfile -ExecutionPolicy Bypass -File "' + $PSCommandPath + '"'
        Start-Process -FilePath 'powershell.exe' -ArgumentList $arg -Verb RunAs | Out-Null
        exit
    } catch { }
}

if (-not $OutDir) {
    $OutDir = Join-Path ([Environment]::GetFolderPath('Desktop')) ('InformePC_' + $env:COMPUTERNAME + '_' + (Get-Date -Format 'yyyyMMdd_HHmm'))
}
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

$global:Txt     = Join-Path $OutDir ('Informe_' + $env:COMPUTERNAME + '.txt')
$global:Htm     = Join-Path $OutDir ('Informe_' + $env:COMPUTERNAME + '.html')
$global:CsvMon  = Join-Path $OutDir 'monitor_picos.csv'
$global:Sec     = [ordered]@{}
$global:UI      = $null          # callback de progreso para la GUI
$global:Bombear = $false         # true = refrescar la ventana dentro de los bucles
$global:Admin   = $esAdmin

# --- Resultados del analisis de controladores (los llena Analizar-Drivers) ---
$global:DrvTotal      = 0
$global:DrvProblemas  = @()
$global:DrvViejos     = @()
$global:DrvPendientes = @()
$global:DrvConsultado = $false
$global:UpdSoftware   = 0

function UI-Texto([string]$t) {
    if ($global:Bombear) { [System.Windows.Forms.Application]::DoEvents() }
    if ($global:UI) { & $global:UI $t }
}
function Sec-Ini([string]$n) { if (-not $global:Sec.Contains($n)) { $global:Sec[$n] = New-Object System.Collections.ArrayList } }
function Ln([string]$n, [string]$t) { Sec-Ini $n; [void]$global:Sec[$n].Add($t) }
function Seg([scriptblock]$sb) { try { return (& $sb) } catch { return $null } }
function HtmlEnc([string]$t) {
    if ($null -eq $t) { return '' }
    return ($t -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;')
}
function Fmt-Props($obj, [string[]]$props) {
    $o = New-Object System.Collections.ArrayList
    foreach ($p in $props) { $v = Seg { $obj.$p }; [void]$o.Add(('  {0,-26} {1}' -f $p, ('' + $v))) }
    return $o
}
function Uptime-Texto {
    $os = Seg { Get-CimInstance Win32_OperatingSystem }
    if (-not $os) { return 'no disponible' }
    $b = $os.LastBootUpTime
    if ($b -is [string]) { $b = Seg { [Management.ManagementDateTimeConverter]::ToDateTime($b) } }
    if (-not $b) { return 'no disponible' }
    $up = (Get-Date) - $b
    return ('{0:N1} horas ({1:N1} dias) - ultimo arranque {2}' -f $up.TotalHours, $up.TotalDays, $b)
}
function Nueva-Inst {
    $h = @{}
    foreach ($p in Get-Process) { $h[$p.Id] = [double]$p.TotalProcessorTime.TotalSeconds }
    return $h
}
function Get-MuestraCPU($prev, [int]$seg, [int]$nuc) {
    $ahora = @{}; $delta = @{}
    foreach ($p in Get-Process) {
        $t = [double]$p.TotalProcessorTime.TotalSeconds
        $ahora[$p.Id] = $t
        if ($prev.ContainsKey($p.Id)) {
            $d = $t - $prev[$p.Id]
            if ($d -gt 0) {
                if ($delta.ContainsKey($p.ProcessName)) { $delta[$p.ProcessName] += $d } else { $delta[$p.ProcessName] = $d }
            }
        }
    }
    $factor = $seg * $nuc
    if ($factor -le 0) { $factor = 1 }
    $suma = 0.0
    foreach ($v in $delta.Values) { $suma += $v }
    $cpu = [Math]::Round(($suma / $factor) * 100, 0)
    if ($cpu -gt 100) { $cpu = 100 }
    $top = (($delta.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 3 | ForEach-Object { $_.Key + ' ' + [Math]::Round(($_.Value / $factor) * 100, 0) + '%' }) -join ' | ')
    return [pscustomobject]@{ Nueva = $ahora; CPU = $cpu; Top = $top; Delta = $delta }
}

function Rec-Identidad {
    $o = New-Object System.Collections.ArrayList
    [void]$o.Add(('  {0,-30} {1}' -f 'Nombre del equipo', $env:COMPUTERNAME))
    [void]$o.Add(('  {0,-30} {1}' -f 'Usuario', $env:USERNAME))
    [void]$o.Add(('  {0,-30} {1}' -f 'Fecha y hora del informe', (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
    [void]$o.Add(('  {0,-30} {1}' -f 'Ejecutado como Administrador', $global:Admin))
    [void]$o.Add(('  {0,-30} {1}' -f 'Arranque / uptime', (Uptime-Texto)))
    [void]$o.Add(('  {0,-30} {1}' -f 'Firmware (UEFI/Legacy)', $env:firmware_type))
    $cs = Seg { Get-CimInstance Win32_ComputerSystem }
    if ($cs) {
        [void]$o.Add(('  {0,-30} {1}' -f 'Fabricante y modelo', ('' + $cs.Manufacturer + ' ' + $cs.Model)))
        [void]$o.Add(('  {0,-30} {1}' -f 'Familia y SKU', ('' + $cs.SystemFamily + ' / ' + $cs.SystemSKUNumber)))
    }
    $bi = Seg { Get-CimInstance Win32_BIOS }
    if ($bi) { [void]$o.Add(('  {0,-30} {1}' -f 'BIOS/UEFI', ('' + $bi.SMBIOSBIOSVersion + '   fecha ' + $bi.ReleaseDate))) }
    $bb = Seg { Get-CimInstance Win32_BaseBoard }
    if ($bb) { [void]$o.Add(('  {0,-30} {1}' -f 'Placa base', ('' + $bb.Manufacturer + ' ' + $bb.Product))) }
    $os = Seg { Get-CimInstance Win32_OperatingSystem }
    if ($os) {
        [void]$o.Add(('  {0,-30} {1}' -f 'Windows', ('' + $os.Caption + ' (build ' + $os.BuildNumber + ', ' + $os.OSArchitecture + ')')))
        [void]$o.Add(('  {0,-30} {1}' -f 'Windows instalado el', $os.InstallDate))
    }
    return $o
}
function Rec-Hardware {
    $o = New-Object System.Collections.ArrayList
    $cp = Seg { Get-CimInstance Win32_Processor | Select-Object -First 1 }
    if ($cp) {
        [void]$o.Add('  --- PROCESADOR ---')
        [void]$o.Add(('  {0,-30} {1}' -f 'Modelo', $cp.Name))
        [void]$o.Add(('  {0,-30} {1}' -f 'Nucleos / hilos', ('' + $cp.NumberOfCores + ' / ' + $cp.NumberOfLogicalProcessors)))
        [void]$o.Add(('  {0,-30} {1}' -f 'Frecuencia max / actual', ('' + $cp.MaxClockSpeed + ' MHz / ' + $cp.CurrentClockSpeed + ' MHz')))
        [void]$o.Add(('  {0,-30} {1}' -f 'Cache L2 / L3', ('' + $cp.L2CacheSize + ' KB / ' + $cp.L3CacheSize + ' KB')))
    }
    [void]$o.Add('')
    [void]$o.Add('  --- MEMORIA RAM (clave: cuantos modulos hay) ---')
    foreach ($m in (Seg { Get-CimInstance Win32_PhysicalMemory })) {
        [void]$o.Add(('  {0,-30} {1}' -f ('Modulo ' + $m.DeviceLocator), ('{0} GB @ {1} MHz ({2}) {3}' -f [int]($m.Capacity/1GB), $m.Speed, $m.ConfiguredClockSpeed, $m.Manufacturer)))
    }
    $marr = Seg { Get-CimInstance Win32_PhysicalMemoryArray | Select-Object -First 1 }
    if ($marr) { [void]$o.Add(('  {0,-30} {1}' -f 'Ranuras de memoria', $marr.MemoryDevices)) }
    $os = Seg { Get-CimInstance Win32_OperatingSystem }
    if ($os) { [void]$o.Add(('  {0,-30} {1}' -f 'RAM libre / total', ('{0} GB libres de {1} GB' -f [int]($os.FreePhysicalMemory/1MB), [int]($os.TotalVisibleMemorySize/1MB)))) }
    [void]$o.Add('')
    [void]$o.Add('  --- VIDEO ---')
    foreach ($v in (Seg { Get-CimInstance Win32_VideoController })) {
        [void]$o.Add(('  {0,-30} {1}' -f 'Adaptador', $v.Name))
        [void]$o.Add(('  {0,-30} {1}' -f '   Driver', ('' + $v.DriverVersion + '   fecha ' + $v.DriverDate)))
        [void]$o.Add(('  {0,-30} {1}' -f '   Resolucion', ('' + $v.CurrentHorizontalResolution + 'x' + $v.CurrentVerticalResolution)))
    }
    [void]$o.Add('')
    [void]$o.Add('  --- ALMACENAMIENTO ---')
    foreach ($d in (Seg { Get-PhysicalDisk })) {
        [void]$o.Add(('  {0,-30} {1}' -f 'Disco', ('{0} - {1} GB - {2} - {3}' -f $d.FriendlyName, [int]($d.Size/1GB), $d.MediaType, $d.HealthStatus)))
    }
    foreach ($v in (Seg { Get-Volume | Where-Object DriveLetter })) {
        [void]$o.Add(('  {0,-30} {1}' -f ('Volumen ' + $v.DriveLetter + ':'), ('{0} GB libres de {1} GB ({2})' -f [int]($v.SizeRemaining/1GB), [int]($v.Size/1GB), $v.FileSystem)))
    }
    [void]$o.Add('')
    [void]$o.Add('  --- SALUD DEL DISCO (SMART) ---')
    $sm = Seg { Get-PhysicalDisk | Get-StorageReliabilityCounter | Select-Object -First 4 }
    if ($sm) {
        foreach ($s in $sm) { [void]$o.Add(('  {0,-30} {1}' -f ('Disco ' + $s.DeviceId), ('desgaste ' + $s.Wear + ' %, temp ' + $s.Temperature + ' C, horas ' + $s.PowerOnHours))) }
    } else { [void]$o.Add('  El controlador no expone SMART por WMI (usar CrystalDiskInfo).') }
    [void]$o.Add('')
    [void]$o.Add('  --- BATERIA ---')
    $ba = Seg { Get-CimInstance Win32_Battery | Select-Object -First 1 }
    if ($ba) { [void]$o.Add(('  {0,-30} {1}' -f 'Estado', ('' + $ba.EstimatedChargeRemaining + '% - ' + $ba.BatteryStatus))) } else { [void]$o.Add('  sin bateria detectada (es un equipo de escritorio?)') }
    return $o
}

function Rec-Rendimiento {
    $o = New-Object System.Collections.ArrayList
    $pt = Seg { (Get-CimInstance Win32_PerfFormattedData_PerfOS_Processor | Where-Object Name -eq '_Total').PercentProcessorTime }
    [void]$o.Add(('  {0,-30} {1}' -f 'CPU total en este instante', ('' + $pt + ' %')))
    [void]$o.Add('')
    [void]$o.Add('  --- TOP 20 PROCESOS POR CPU ACUMULADA (desde el arranque) ---')
    foreach ($p in (Seg { Get-Process | Sort-Object CPU -Descending | Select-Object -First 20 })) {
        [void]$o.Add(('  {0,10:N0} s  {1,7:N0} MB  {2,-26} {3}' -f $p.CPU, ($p.WorkingSet64/1MB), $p.ProcessName, $p.Path))
    }
    [void]$o.Add('')
    [void]$o.Add('  --- TOP 10 PROCESOS POR MEMORIA ---')
    foreach ($p in (Seg { Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 10 })) {
        [void]$o.Add(('  {0,10:N0} MB  {1,-26} {2}' -f ($p.WorkingSet64/1MB), $p.ProcessName, $p.Company))
    }
    [void]$o.Add('')
    [void]$o.Add('  --- CARGA POR NUCLEO (desbalance = un hilo saturado) ---')
    foreach ($n in (Seg { Get-CimInstance Win32_PerfFormattedData_PerfOS_Processor | Where-Object Name -ne '_Total' })) {
        [void]$o.Add(('  Nucleo {0,3}   CPU {1,4} %   usuario {2,4} %   interrupciones {3,4} %' -f $n.Name, $n.PercentProcessorTime, $n.PercentUserTime, $n.PercentInterruptTime))
    }
    return $o
}
function Rec-MuestreoCPU {
    $o = New-Object System.Collections.ArrayList
    $nuc = Seg { (Get-CimInstance Win32_ComputerSystem).NumberOfLogicalProcessors }
    if (-not $nuc) { $nuc = 1 }
    $intervalo = 3
    if ($MuestreoSegundos -lt 6) { $intervalo = 2 }
    [void]$o.Add(('  Muestreo de {0} segundos, una muestra cada {1} s. CPU en % del total del sistema.' -f $MuestreoSegundos, $intervalo))
    $prev = Nueva-Inst
    $acum = @{}; $pico = @{}
    $global:MuestreoMaxCPU = 0
    $csv = Join-Path $OutDir 'muestreo_cpu.csv'
    'Hora;CPU_Total;Top3;Candidatos' | Out-File -FilePath $csv -Encoding utf8
    $fin = (Get-Date).AddSeconds($MuestreoSegundos)
    $n = 0
    while ((Get-Date) -lt $fin) {
        Start-Sleep -Seconds $intervalo
        $mu = Get-MuestraCPU $prev $intervalo $nuc
        $prev = $mu.Nueva
        $n++
        foreach ($k in $mu.Delta.Keys) {
            if ($acum.ContainsKey($k)) { $acum[$k] += $mu.Delta[$k] } else { $acum[$k] = $mu.Delta[$k] }
            if (-not $pico.ContainsKey($k) -or $pico[$k] -lt $mu.Delta[$k]) { $pico[$k] = $mu.Delta[$k] }
        }
        $factor = $intervalo * $nuc
        $os = Seg { Get-CimInstance Win32_OperatingSystem }
        $ram = 0
        if ($os) { $ram = [Math]::Round((($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / $os.TotalVisibleMemorySize) * 100, 0) }
        $mhz = Seg { (Get-CimInstance Win32_Processor).CurrentClockSpeed }
        $gordos = (($mu.Delta.GetEnumerator() | Where-Object { ($_.Value / $factor) * 100 -ge 5 } | Sort-Object Value -Descending | ForEach-Object { $_.Key + '(' + [Math]::Round(($_.Value / $factor) * 100, 0) + '%)' }) -join ', ')
        ((Get-Date -Format 'HH:mm:ss') + ';' + $mu.CPU + ';' + $mu.Top + ';' + $gordos) | Out-File -FilePath $csv -Append -Encoding utf8
        if ($mu.CPU -gt $global:MuestreoMaxCPU) { $global:MuestreoMaxCPU = $mu.CPU }
        UI-Texto ('Muestreo ' + $n + ': CPU ' + $mu.CPU + '%  ->  ' + $mu.Top)
        [void]$o.Add(('  {0}   CPU {1,3} %   {2,5} MHz   RAM {3,3} %   top: {4}' -f (Get-Date -Format 'HH:mm:ss'), $mu.CPU, $mhz, $ram, $mu.Top))
    }
    [void]$o.Add('')
    [void]$o.Add('  --- PROCESOS QUE MAS CPU CONSUMIERON DURANTE EL MUESTREO ---')
    foreach ($e in ($acum.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 15)) { [void]$o.Add(('  {0,-30} {1,9:N1} s de CPU' -f $e.Key, $e.Value)) }
    [void]$o.Add('')
    [void]$o.Add('  --- PICOS MAS ALTOS EN UN SOLO INTERVALO (los "picos raros") ---')
    foreach ($e in ($pico.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 15)) { [void]$o.Add(('  {0,-30} {1,9:N1} s en un intervalo de {2} s' -f $e.Key, $e.Value, $intervalo)) }
    $global:MuestreoTopProc = ''
    $global:MuestreoTopPorc = 0
    $tops = $acum.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1
    if ($tops) {
        $global:MuestreoTopProc = $tops.Key
        if ($MuestreoSegundos -gt 0) { $global:MuestreoTopPorc = [Math]::Round(($tops.Value / $MuestreoSegundos) * 100, 1) }
    }
    [void]$o.Add('')
    [void]$o.Add('  CSV completo del muestreo: ' + $csv)
    return $o
}

function Rec-Sospechosos {
    $o = New-Object System.Collections.ArrayList
    [void]$o.Add('  --- PROCESOS CON FIRMA DIGITAL NO VALIDA (posible malware) ---')
    $malos = 0
    foreach ($p in (Seg { Get-Process | Where-Object Path | Sort-Object { $_.Path } -Unique })) {
        $sig = Seg { (Get-AuthenticodeSignature $p.Path).Status }
        if ($sig -and $sig -ne 'Valid') {
            $malos++
            [void]$o.Add(('  [{0}] {1}  ->  {2}' -f $sig, $p.ProcessName, $p.Path))
        }
    }
    if ($malos -eq 0) { [void]$o.Add('  ninguno: todos los procesos tienen firma digital valida') }
    [void]$o.Add('')
    [void]$o.Add('  --- PROCESOS DESDE CARPETAS DE USUARIO O TEMPORALES ---')
    $raro = 0
    foreach ($p in (Seg { Get-Process | Where-Object { $_.Path -match 'AppData|ProgramData|Users\\Public|Downloads|\\Temp\\' } })) {
        $raro++
        [void]$o.Add(('  {0,-26} {1}' -f $p.ProcessName, $p.Path))
    }
    if ($raro -eq 0) { [void]$o.Add('  ninguno: nada corriendo desde AppData / ProgramData / Temp') }
    [void]$o.Add('')
    [void]$o.Add('  --- CONEXIONES DE RED ACTIVAS CON SU PROCESO ---')
    foreach ($c in (Seg { Get-NetTCPConnection -State Established })) {
        $pn = Seg { (Get-Process -Id $c.OwningProcess).ProcessName }
        [void]$o.Add(('  {0,-26} {1}:{2}  ->  {3}:{4}' -f $pn, $c.LocalAddress, $c.LocalPort, $c.RemoteAddress, $c.RemotePort))
    }
    return $o
}
function Sugerencia-Arranque([string]$texto) {
    if ($texto -match 'javau|jusched|java auto')            { return 'DESACTIVAR: actualizador de Java, viejo e inutil' }
    if ($texto -match 'onedrivesetup')                      { return 'DESACTIVAR: resto de la instalacion de OneDrive' }
    if ($texto -match 'opera_autoupdate|opera gx')          { return 'PASAR A MANUAL: actualizador de Opera GX' }
    if ($texto -match 'wallpaper')                          { return 'DESACTIVAR: fondo animado, consume GPU y RAM' }
    if ($texto -match 'blitz')                              { return 'DESACTIVAR: se abre solo, abrilo cuando lo uses' }
    if ($texto -match 'nox|MultiPlayerManager')             { return 'DESACTIVAR: emulador Nox, es pesado' }
    if ($texto -match 'epicgames|battle\.net|ubisoft|ea desktop|eadm|riotclient|roblox') { return 'OPCIONAL: launcher de juegos, abrilo a mano' }
    if ($texto -match 'steam')                              { return 'OPCIONAL: Steam al inicio (sacalo si no lo usas siempre)' }
    if ($texto -match 'msedge|edgeautolaunch|edgeupdate')   { return 'DESACTIVAR: Edge arrancando con Windows' }
    if ($texto -match 'braveupdate|gupdate|google')         { return 'PASAR A MANUAL: actualizador del navegador' }
    if ($texto -match 'zoom')                               { return 'PASAR A MANUAL: actualizador de Zoom' }
    if ($texto -match 'adobe|creative cloud')               { return 'PASAR A MANUAL: actualizador de Adobe' }
    if ($texto -match 'vanguard')                           { return 'DESACTIVAR si no jugas Valorant: anticheat siempre activo' }
    if ($texto -match 'customcursor|blife')                 { return 'DESACTIVAR: cursor personalizado, es prescindible' }
    if ($texto -match 'discord')                            { return 'OPCIONAL: dejamelo solo si estas siempre en Discord' }
    if ($texto -match 'securityhealth|defender|winDefend|mdcoresvc') { return 'DEJAR: es la seguridad de Windows' }
    if ($texto -match 'rktaud|realtek|audio|nvidia|nvcontainer|amd|intel|lg ?hub|jbl|quantum|aorus|easytune|gcc') { return 'DEJAR: es driver o servicio del equipo' }
    return ''
}
function Rec-Arranque {
    $o = New-Object System.Collections.ArrayList
    [void]$o.Add('  --- PROGRAMAS DE INICIO (lo que ve el Administrador de tareas) ---')
    foreach ($s in (Seg { Get-CimInstance Win32_StartupCommand })) {
        $sug = Sugerencia-Arranque ('' + $s.Name + '  ' + $s.Command)
        if ($sug) { [void]$o.Add(('  {0,-32} {1}   [{2}]' -f $s.Name, $s.Command, $sug)) }
        else      { [void]$o.Add(('  {0,-32} {1}' -f $s.Name, $s.Command)) }
    }
    [void]$o.Add('')
    [void]$o.Add('  --- CLAVES RUN / RUNONCE DEL REGISTRO ---')
    $claves = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
                'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run',
                'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
                'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce',
                'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce',
                'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run')
    foreach ($k in $claves) {
        $it = Seg { Get-ItemProperty $k }
        if ($it) {
            foreach ($pp in ($it.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' })) {
                $sug = Sugerencia-Arranque ('' + $pp.Name + '  ' + $pp.Value)
                if ($sug) { [void]$o.Add(('  [{0}] {1}   [{2}]' -f $k.Split('\')[-1], $pp.Name, $sug)) }
                else      { [void]$o.Add(('  [{0}] {1}' -f $k.Split('\')[-1], $pp.Name)) }
                [void]$o.Add('        ' + $pp.Value)
            }
        }
    }
    [void]$o.Add('')
    [void]$o.Add('  --- CARPETAS DE INICIO ---')
    foreach ($f in (Seg { Get-ChildItem (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup'), (Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs\Startup') })) { [void]$o.Add('  ' + $f.FullName) }
    [void]$o.Add('')
    [void]$o.Add('  --- TAREAS PROGRAMADAS QUE NO SON DE MICROSOFT ---')
    foreach ($t in (Seg { Get-ScheduledTask | Where-Object { $_.TaskPath -notlike '\Microsoft\*' } })) {
        $acc = Seg { ($t.Actions | ForEach-Object { $_.Execute }) -join ' || ' }
        [void]$o.Add(('  [{0,-9}] {1}{2}' -f ('' + $t.State), $t.TaskPath, $t.TaskName))
        [void]$o.Add('        ' + $acc)
    }
    [void]$o.Add('')
    [void]$o.Add('  --- TAREAS QUE ARRANCAN CON EL SISTEMA O AL INICIAR SESION ---')
    foreach ($t in (Seg { Get-ScheduledTask | Where-Object { $_.TaskPath -notlike '\Microsoft\*' } | Where-Object { $_.Triggers | Where-Object { $_.CimClass.CimClassName -match 'Boot|Logon' } } })) {
        $sug = Sugerencia-Arranque ('' + $t.TaskName)
        if ($sug) { [void]$o.Add(('  [{0,-9}] {1}{2}   [{3}]' -f ('' + $t.State), $t.TaskPath, $t.TaskName, $sug)) }
        else      { [void]$o.Add(('  [{0,-9}] {1}{2}' -f ('' + $t.State), $t.TaskPath, $t.TaskName)) }
    }
    [void]$o.Add('')
    [void]$o.Add('  --- SERVICIOS AUTOMATICOS QUE NO SON DE WINDOWS ---')
    foreach ($s in (Seg { Get-CimInstance Win32_Service | Where-Object { $_.StartMode -eq 'Auto' -and $_.PathName -notmatch 'C:\\Windows\\' } })) {
        [void]$o.Add(('  [{0,-8}] {1}' -f $s.State, $s.Name))
        [void]$o.Add('        ' + $s.PathName)
    }
    [void]$o.Add('')
    [void]$o.Add('  --- SERVICIOS AUTOMATICOS QUE ESTAN DETENIDOS ---')
    foreach ($s in (Seg { Get-Service | Where-Object { $_.StartType -eq 'Automatic' -and $_.Status -ne 'Running' } })) { [void]$o.Add(('  {0,-32} {1}' -f $s.Name, $s.DisplayName)) }
    [void]$o.Add('')
    [void]$o.Add('  --- COMO USAR ESTA LISTA ---')
    [void]$o.Add('  No hace falta desinstalar nada: se desactiva y se puede volver atras cuando quieras.')
    [void]$o.Add('   - Programas de inicio: Administrador de tareas > Aplicaciones de inicio > clic derecho > Deshabilitar.')
    [void]$o.Add('   - Tareas programadas: Programador de tareas > Deshabilitar (no borrar).')
    [void]$o.Add('   - Servicios: services.msc > cambiar a Manual (solo los de marcas que no uses).')
    [void]$o.Add('   - Para ver absolutamente todo lo que arranca: Autoruns (Microsoft Sysinternals).')
    [void]$o.Add('   Regla practica: deja los drivers y la seguridad; el resto, activalo cuando lo uses.')
    return $o
}
function Rec-Seguridad {
    $o = New-Object System.Collections.ArrayList
    [void]$o.Add('  --- WINDOWS DEFENDER ---')
    $d = Seg { Get-MpComputerStatus }
    if ($d) {
        [void]$o.Add(('  {0,-30} {1}' -f 'Proteccion en tiempo real', $d.RealTimeProtectionEnabled))
        [void]$o.Add(('  {0,-30} {1}' -f 'Modulo antivirus', $d.AntivirusEnabled))
        [void]$o.Add(('  {0,-30} {1}' -f 'Firmas actualizadas el', $d.AntivirusSignatureLastUpdated))
        $qr = $d.QuickScanAge; if ($qr -gt 40000) { $qr = 'nunca' }
        $fr = $d.FullScanAge;  if ($fr -gt 40000) { $fr = 'nunca' }
        [void]$o.Add(('  {0,-30} {1}' -f 'Dias desde examen rapido', $qr))
        [void]$o.Add(('  {0,-30} {1}' -f 'Dias desde examen completo', $fr))
    } else { [void]$o.Add('  Defender no responde (puede haber otro antivirus instalado)') }
    [void]$o.Add('  Exclusiones de Defender (el malware suele agregarlas):')
    $mp = Seg { Get-MpPreference }
    if ($mp) {
        $exR = @($mp.ExclusionPath      | Where-Object { $_ -and ('' + $_).Trim().Length -gt 1 -and $_ -notlike 'N/A*' })
        $exP = @($mp.ExclusionProcess   | Where-Object { $_ -and ('' + $_).Trim().Length -gt 1 -and $_ -notlike 'N/A*' })
        $exE = @($mp.ExclusionExtension | Where-Object { $_ -and ('' + $_).Trim().Length -gt 1 -and $_ -notlike 'N/A*' })
        if ($exR.Count -gt 0) { foreach ($e in $exR) { [void]$o.Add('    RUTA: ' + $e) } } else { [void]$o.Add('    sin exclusiones de rutas') }
        if ($exP.Count -gt 0) { foreach ($e in $exP) { [void]$o.Add('    PROCESO: ' + $e) } }
        if ($exE.Count -gt 0) { foreach ($e in $exE) { [void]$o.Add('    EXTENSION: ' + $e) } }
        if ($mp.DisableRealtimeMonitoring) { [void]$o.Add('    ATENCION: proteccion en tiempo real desactivada por politica') }
    }
    [void]$o.Add('  Amenazas detectadas (historial):')
    $am = Seg { Get-MpThreatDetection | Sort-Object InitialDetectionTime -Descending | Select-Object -First 15 }
    if ($am) { foreach ($a in $am) { [void]$o.Add(('    {0}   {1}' -f $a.InitialDetectionTime, ($a.Resources -join ' , '))) } } else { [void]$o.Add('    ninguna registrada (buena senal)') }
    [void]$o.Add('')
    [void]$o.Add('  --- FIREWALL ---')
    foreach ($f in (Seg { Get-NetFirewallProfile })) { [void]$o.Add(('  {0,-12} habilitado: {1}   entrante: {2}   saliente: {3}' -f $f.Name, $f.Enabled, $f.DefaultInboundAction, $f.DefaultOutboundAction)) }
    [void]$o.Add('')
    [void]$o.Add('  --- ACCESO REMOTO Y CUENTAS ---')
    $ts = Seg { Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' }
    if ($ts) {
        $rdp = 'HABILITADO - si no lo usas, desactivalo (Configuracion > Sistema > Escritorio remoto)'
        if ($ts.fDenyTSConnections -eq 1) { $rdp = 'bloqueado (correcto)' }
        [void]$o.Add(('  {0,-30} {1}' -f 'Escritorio remoto', $rdp))
    }
    foreach ($s in (Seg { Get-CimInstance Win32_Service | Where-Object { $_.Name -in @('TermService','WinRM','RemoteRegistry','RemoteAccess','UmRdpService') } })) {
        [void]$o.Add(('  {0,-18} {1,-8} {2}' -f $s.Name, $s.State, $s.StartMode))
    }
    foreach ($u in (Seg { Get-LocalUser })) { [void]$o.Add(('  USUARIO {0,-16} habilitado: {1}   ultimo acceso: {2}' -f $u.Name, $u.Enabled, $u.LastLogon)) }
    foreach ($a in (Seg { Get-LocalGroupMember -SID 'S-1-5-32-544' })) { [void]$o.Add('  ADMINISTRADOR: ' + $a.Name) }
    [void]$o.Add('')
    [void]$o.Add('  --- SECUESTROS TIPICOS ---')
    $wl = Seg { Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' }
    if ($wl) { [void]$o.Add(('  Shell: {0}   Userinit: {1}' -f $wl.Shell, $wl.Userinit)) }
    $ai = Seg { Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows' }
    if ($ai) { [void]$o.Add(('  AppInit_DLLs: "{0}"   (debe estar vacio)' -f $ai.AppInit_DLLs)) }
    [void]$o.Add('  Archivo hosts (lineas activas):')
    $hst = Seg { Get-Content (Join-Path $env:SystemRoot 'System32\drivers\etc\hosts') | Where-Object { $_ -and $_ -notmatch '^\s*#' } }
    if ($hst) { foreach ($h in $hst) { [void]$o.Add('    ' + $h) } } else { [void]$o.Add('    vacio (correcto)') }
    [void]$o.Add('')
    [void]$o.Add('  --- DNS CONFIGURADO (un DNS raro puede ser secuestro) ---')
    foreach ($dd in (Seg { Get-DnsClientServerAddress -AddressFamily IPv4 })) { [void]$o.Add(('  {0,-26} {1}' -f $dd.InterfaceAlias, ($dd.ServerAddresses -join ', '))) }
    return $o
}
function Texto-ErrorDispositivo([int]$c) {
    # Devuelve 'texto del problema|gravedad' para el codigo de error del Administrador de dispositivos
    switch ($c) {
        1  { return 'no esta configurado correctamente|REVISAR' }
        3  { return 'el controlador esta danado o falta memoria del sistema|REVISAR' }
        10 { return 'no puede iniciarse|REVISAR' }
        12 { return 'no hay recursos libres suficientes|REVISAR' }
        14 { return 'necesita reiniciar el equipo para funcionar|ATENCION' }
        18 { return 'hay que reinstalar los controladores|REVISAR' }
        19 { return 'el registro del dispositivo esta danado|REVISAR' }
        21 { return 'Windows lo esta quitando|ATENCION' }
        22 { return 'deshabilitado (a proposito o por una aplicacion)|ATENCION' }
        24 { return 'no esta conectado en este momento|ATENCION' }
        28 { return 'NO TIENE LOS CONTROLADORES INSTALADOS|REVISAR' }
        29 { return 'el firmware del dispositivo no dio los recursos necesarios|REVISAR' }
        31 { return 'no se puede cargar el controlador|REVISAR' }
        32 { return 'tiene el controlador de inicio deshabilitado|ATENCION' }
        37 { return 'el controlador no puede inicializarse|REVISAR' }
        39 { return 'el controlador esta danado o no esta|REVISAR' }
        43 { return 'Windows lo detuvo porque dio un error|REVISAR' }
        45 { return 'no esta conectado en este momento|ATENCION' }
        48 { return 'no se pudo iniciar el software del dispositivo|REVISAR' }
        52 { return 'Windows no pudo verificar la firma del controlador|ATENCION' }
        default { return ('codigo de error ' + $c + '|REVISAR') }
    }
}

function Analizar-Drivers {
    # Analiza si hay que actualizar los controladores. Devuelve las lineas del informe
    # y deja los resultados en variables globales para el tablero de estado.
    $o = New-Object System.Collections.ArrayList
    $global:DrvTotal      = 0
    $global:DrvProblemas  = @()
    $global:DrvViejos     = @()
    $global:DrvPendientes = @()
    $global:DrvConsultado = $false
    $global:UpdPendientes = $null
    $global:UpdSoftware   = 0
    $anios = 4

    # --- A) Dispositivos sin controlador o con error (lo mismo que muestra el Administrador de dispositivos) ---
    [void]$o.Add('  --- DISPOSITIVOS SIN CONTROLADOR O CON ERROR ---')
    $prob = @()
    foreach ($d in (Seg { Get-CimInstance Win32_PnPEntity })) {
        if (-not $d.Name) { continue }
        $c = 0
        try { $c = [int]$d.ConfigManagerErrorCode } catch { $c = 0 }
        if ($c -eq 0) { continue }
        $pp = (Texto-ErrorDispositivo $c) -split '\|'
        $cl = ('' + $d.PNPClass)
        if (-not $cl) { $cl = (('' + $d.DeviceID) -split '\\')[0] }   # algunos dispositivos no informan la clase: usamos el tipo de bus
        $prob += [pscustomobject]@{
            Nombre = ('' + $d.Name)
            Codigo = $c
            Texto  = ('' + $pp[0])
            Estado = ('' + $pp[1])
            Clase  = $cl
        }
    }
    $global:DrvProblemas = $prob
    if ($prob.Count -gt 0) {
        foreach ($p in ($prob | Sort-Object Estado, Nombre)) {
            [void]$o.Add(('  [{0}] {1}' -f $p.Estado, $p.Nombre))
            [void]$o.Add(('         codigo {0}: {1}   (clase de dispositivo: {2})' -f $p.Codigo, $p.Texto, $p.Clase))
        }
        [void]$o.Add('  Codigo 28 = falta el controlador: se arregla en Administrador de dispositivos >')
        [void]$o.Add('  clic derecho en el dispositivo > Actualizar controlador > Buscar automaticamente.')
    } else {
        [void]$o.Add('  ninguno: todos los dispositivos tienen su controlador y responden bien')
    }
    # --- B) Controladores instalados y los mas antiguos ---
    [void]$o.Add('')
    [void]$o.Add('  --- CONTROLADORES INSTALADOS Y LOS MAS ANTIGUOS ---')
    $clases = 'DISPLAY|NET|MEDIA|BLUETOOTH|CAMERA|USB|HDC|FIRMWARE|MONITOR|PRINTER|IMAGE|BATTERY|MOUSE|KEYBOARD|SCANNER|SMARTCARD'
    $drv = @(Seg { Get-CimInstance Win32_PnPSignedDriver })
    $global:DrvTotal = $drv.Count
    $lim = (Get-Date).AddYears(-1 * $anios)
    $viejos = @()
    $conFecha = 0
    foreach ($d in $drv) {
        if (-not $d.DeviceName -or -not $d.DriverDate) { continue }
        if (('' + $d.DeviceClass) -notmatch $clases) { continue }
        $f = $d.DriverDate
        if ($f -is [string]) { $f = Seg { [Management.ManagementDateTimeConverter]::ToDateTime($d.DriverDate) } }
        if (-not $f -or -not ($f -is [datetime])) { continue }
        if ($f.Year -le 2000) { continue }   # muchos INF traen una fecha falsa (1968) que no sirve para comparar
        if (('' + $d.DriverProviderName) -match 'Microsoft') { continue }   # los genericos de Windows usan siempre la fecha 21/06/2006
        $conFecha++
        if ($f -lt $lim) {
            $viejos += [pscustomobject]@{
                Fecha     = $f
                Nombre    = ('' + $d.DeviceName)
                Version   = ('' + $d.DriverVersion)
                Proveedor = ('' + $d.DriverProviderName)
                Clase     = ('' + $d.DeviceClass)
            }
        }
    }
    $global:DrvViejos = $viejos
    [void]$o.Add(('  Controladores que ve Windows: ' + $global:DrvTotal + '   |   de un fabricante y con fecha util: ' + $conFecha))
    if ($viejos.Count -gt 0) {
        [void]$o.Add(('  Con mas de ' + $anios + ' anos: ' + $viejos.Count + '   (si el equipo anda bien, no hace falta tocarlos)'))
        foreach ($v in ($viejos | Sort-Object Fecha | Select-Object -First 20)) {
            [void]$o.Add(('  {0}  {1}  v{2}  ({3}, {4})' -f $v.Fecha.ToString('yyyy-MM-dd'), $v.Nombre, $v.Version, $v.Clase, $v.Proveedor))
        }
    } else {
        [void]$o.Add(('  ninguno con mas de ' + $anios + ' anos: estan razonablemente al dia'))
    }
    $gen = @($drv | Where-Object { ('' + $_.DeviceClass) -eq 'DISPLAY' -and ('' + $_.DriverProviderName) -match 'Microsoft' -and $_.DeviceName })
    if ($gen.Count -gt 0) {
        [void]$o.Add('  AVISO: hay un adaptador de video con el controlador generico de Microsoft (no el del fabricante):')
        foreach ($g in $gen) { [void]$o.Add('    ' + $g.DeviceName) }
    }
    # --- C) Consulta a Windows Update: que controladores ofrece para este equipo ---
    [void]$o.Add('')
    [void]$o.Add('  --- CONTROLADORES PARA ACTUALIZAR (Windows Update) ---')
    if ($ActualizacionesPendientes -or $DriversPendientes) {
        $global:DrvConsultado = $true
        UI-Texto '   Consultando Windows Update (actualizaciones y controladores)...'
        $res = Seg {
            $ses = New-Object -ComObject Microsoft.Update.Session
            $bus = $ses.CreateUpdateSearcher()
            $up  = $bus.Search('IsInstalled=0').Updates
            $lista = @()
            foreach ($x in $up) {
                $fec = ''
                try { if ($x.DriverVerDate) { $fec = ('{0:yyyy-MM-dd}' -f $x.DriverVerDate) } } catch { $fec = '' }
                $lista += [pscustomobject]@{
                    Titulo     = ('' + $x.Title)
                    Tipo       = [int]$x.Type          # 1 = Windows   |   2 = controlador
                    Modelo     = ('' + $x.DriverModel)
                    Fabricante = ('' + $x.DriverManufacturer)
                    Fecha      = $fec
                }
            }
            return [pscustomobject]@{ Items = $lista }
        }
        if ($null -eq $res) {
            $global:DrvConsultado = $false
            [void]$o.Add('  no se pudo consultar Windows Update (sin internet, servicio detenido o politica del equipo)')
        } else {
            $global:DrvPendientes = @($res.Items | Where-Object { $_.Tipo -eq 2 })
            $soft = @($res.Items | Where-Object { $_.Tipo -ne 2 })
            $global:UpdPendientes = @($res.Items).Count
            $global:UpdSoftware   = $soft.Count
            if ($global:DrvPendientes.Count -eq 0) {
                [void]$o.Add('  ninguno: Windows Update no tiene controladores nuevos para este equipo')
            } else {
                [void]$o.Add(('  HAY ' + $global:DrvPendientes.Count + ' CONTROLADOR(ES) PARA ACTUALIZAR:'))
                foreach ($p in $global:DrvPendientes) {
                    [void]$o.Add('  - ' + $p.Titulo)
                    if ($p.Modelo) { [void]$o.Add(('       dispositivo: ' + $p.Modelo + '   fabricante: ' + $p.Fabricante + '   fecha del driver: ' + $p.Fecha)) }
                }
            }
            if ($soft.Count -gt 0) {
                [void]$o.Add('')
                [void]$o.Add(('  Ademas hay ' + $soft.Count + ' actualizacion(es) de Windows pendientes:'))
                foreach ($p in ($soft | Select-Object -First 30)) { [void]$o.Add('  - ' + $p.Titulo) }
            }
        }
    } else {
        [void]$o.Add('  no consultado (esta prueba usa internet): marca la casilla "Consultar Windows Update"')
        [void]$o.Add('  en Opciones avanzadas, o ejecuta el script con -DriversPendientes, para saber si este')
        [void]$o.Add('  equipo tiene controladores nuevos disponibles.')
    }
    # --- D) Como actualizarlos bien ---
    [void]$o.Add('')
    [void]$o.Add('  --- COMO ACTUALIZAR LOS CONTROLADORES (en este orden) ---')
    [void]$o.Add('  1. Windows Update > Opciones avanzadas > Actualizaciones opcionales > Actualizaciones de controladores')
    [void]$o.Add('  2. Administrador de dispositivos > clic derecho en el dispositivo > Actualizar controlador')
    [void]$o.Add('     (los que dicen "(codigo 28) no tiene los controladores instalados" son los que faltan)')
    [void]$o.Add('  3. Placa de video y chipset: bajarlos del sitio oficial (NVIDIA, AMD, Intel o el del equipo)')
    [void]$o.Add('  4. NO usar "actualizadores" tipo Driver Booster / IObit: instalan controladores viejos o equivocados')
    return $o
}

function Rec-Drivers {
    $o = New-Object System.Collections.ArrayList
    [void]$o.Add('  --- DRIVERS (.sys) MODIFICADOS EN LOS ULTIMOS 90 DIAS ---')
    [void]$o.Add('  (Driver Booster y otros "actualizadores" escriben aqui: fechas recientes = revisar)')
    foreach ($f in (Seg { Get-ChildItem (Join-Path $env:SystemRoot 'System32\drivers') -Filter *.sys | Where-Object { $_.LastWriteTime -gt (Get-Date).AddDays(-90) } | Sort-Object LastWriteTime -Descending | Select-Object -First 40 })) {
        [void]$o.Add(('  {0}   {1}' -f $f.LastWriteTime.ToString('yyyy-MM-dd'), $f.Name))
    }
    [void]$o.Add('')
    [void]$o.Add('  --- RASTROS DE DRIVER BOOSTER / IObit (conviene desinstalarlo) ---')
    $sv = Seg { Get-CimInstance Win32_Service | Where-Object { $_.Name -match 'iobit|driverbooster|ascdrv|advancedsystemcare' } }
    if ($sv) { foreach ($s in $sv) { [void]$o.Add(('  SERVICIO {0} [{1}] {2}' -f $s.Name, $s.State, $s.PathName)) } } else { [void]$o.Add('  sin servicios de IObit') }
    $tw = Seg { Get-ScheduledTask | Where-Object { $_.TaskName -match 'iobit|driverbooster|driver booster|ascdrv|advancedsystemcare|advanced systemcare' } }
    if ($tw) { foreach ($t in $tw) { [void]$o.Add(('  TAREA {0} [{1}]' -f $t.TaskName, $t.State)) } } else { [void]$o.Add('  sin tareas programadas de IObit') }
    foreach ($dir in @('C:\Program Files (x86)\IObit', 'C:\Program Files\IObit')) { if (Test-Path $dir) { [void]$o.Add('  CARPETA INSTALADA: ' + $dir) } }
    [void]$o.Add('')
    [void]$o.Add('  --- PROGRAMAS INSTALADOS (los 30 mas recientes) ---')
    $rutas = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*')
    $progs = Seg { Get-ItemProperty $rutas | Where-Object DisplayName | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | Sort-Object InstallDate -Descending | Select-Object -First 30 }
    foreach ($p in $progs) { [void]$o.Add(('  {0,-12} {1,-50} {2}' -f $p.InstallDate, $p.DisplayName, $p.Publisher)) }
    [void]$o.Add('')
    [void]$o.Add('  --- ULTIMAS ACTUALIZACIONES DE WINDOWS INSTALADAS ---')
    foreach ($h in (Seg { Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 15 })) { [void]$o.Add(('  {0}   {1,-22} {2}' -f $h.HotFixID, $h.Description, $h.InstalledOn)) }
    [void]$o.Add('')
    $dl = Seg { Analizar-Drivers }
    if ($dl) { foreach ($l in $dl) { [void]$o.Add('' + $l) } } else { [void]$o.Add('  (no se pudo analizar los controladores del equipo)') }
    return $o
}
function Rec-Eventos {
    $o = New-Object System.Collections.ArrayList
    [void]$o.Add('  --- ERRORES CRITICOS DEL SISTEMA AGRUPADOS (ultimos 7 dias) ---')
    $g1 = Seg { Get-WinEvent -FilterHashtable @{LogName='System';Level=1,2;StartTime=(Get-Date).AddDays(-7)} -MaxEvents 500 -ErrorAction Stop | Group-Object Id,ProviderName | Sort-Object Count -Descending | Select-Object -First 20 }
    if ($g1) { foreach ($g in $g1) { [void]$o.Add(('  x{0,-5} {1}' -f $g.Count, $g.Name)) } } else { [void]$o.Add('  sin errores o sin permisos para leerlos') }
    [void]$o.Add('')
    [void]$o.Add('  --- DETALLE: SERVICIOS QUE FALLARON (primeros 8, ultimos 7 dias) ---')
    foreach ($e in (Seg { Get-WinEvent -FilterHashtable @{LogName='System';Id=7000,7024,7034;StartTime=(Get-Date).AddDays(-7)} -MaxEvents 8 -ErrorAction Stop })) {
        $m = ($e.Message -replace '\s+',' ')
        if ($m.Length -gt 170) { $m = $m.Substring(0,170) }
        [void]$o.Add(('  {0}  Id {1}  {2}' -f $e.TimeCreated, $e.Id, $m))
    }
    [void]$o.Add('')
    [void]$o.Add('  --- DETALLE: APLICACIONES QUE SE CERRARON CON ERROR (primeros 8, ultimos 7 dias) ---')
    foreach ($e in (Seg { Get-WinEvent -FilterHashtable @{LogName='Application';Id=1000,1002,1026;StartTime=(Get-Date).AddDays(-7)} -MaxEvents 8 -ErrorAction Stop })) {
        $m = ($e.Message -replace '\s+',' ')
        if ($m.Length -gt 170) { $m = $m.Substring(0,170) }
        [void]$o.Add(('  {0}  Id {1}  {2}' -f $e.TimeCreated, $e.Id, $m))
    }
    [void]$o.Add('')
    [void]$o.Add('  --- ARRANQUE: DURACION (Id 100) Y COMPONENTES QUE LO FRENAN (Id 101) ---')
    $ev = Seg { Get-WinEvent -LogName Microsoft-Windows-Diagnostics-Performance/Operational -MaxEvents 150 -ErrorAction Stop }
    if ($ev) {
        foreach ($e in ($ev | Where-Object Id -eq 100 | Select-Object -First 6)) {
            $m = ($e.Message -replace '\s+',' ')
            if ($m.Length -gt 150) { $m = $m.Substring(0,150) }
            [void]$o.Add(('  {0}   {1}' -f $e.TimeCreated, $m))
        }
        [void]$o.Add('  Culpables que se repiten:')
        foreach ($g in ($ev | Where-Object Id -eq 101 | Group-Object { $m = ($_.Message -replace '\s+',' '); if ($m.Length -gt 100) { $m = $m.Substring(0,100) }; $m } | Sort-Object Count -Descending | Select-Object -First 10)) {
            [void]$o.Add(('  x{0,-5} {1}' -f $g.Count, $g.Name))
        }
        [void]$o.Add('  Cierres y apagados anomalos (Id 200-203):')
        foreach ($e in ($ev | Where-Object { $_.Id -in 200,201,202,203 } | Select-Object -First 8)) {
            $m = ($e.Message -replace '\s+',' ')
            if ($m.Length -gt 120) { $m = $m.Substring(0,120) }
            [void]$o.Add(('  {0}  Id {1}  {2}' -f $e.TimeCreated, $e.Id, $m))
        }
    } else { [void]$o.Add('  sin datos (esta seccion necesita Administrador)') }
    [void]$o.Add('')
    [void]$o.Add('  --- HARDWARE, ENERGIA DEL CPU Y DISCO ---')
    foreach ($e in (Seg { Get-WinEvent -FilterHashtable @{LogName='System';ProviderName='Microsoft-Windows-WHEA-Logger';StartTime=(Get-Date).AddDays(-30)} -MaxEvents 10 -ErrorAction Stop })) { [void]$o.Add(('  ERROR DE HARDWARE (WHEA) {0}  Id {1}' -f $e.TimeCreated, $e.Id)) }
    foreach ($e in (Seg { Get-WinEvent -FilterHashtable @{LogName='System';ProviderName='Microsoft-Windows-Kernel-Processor-Power';StartTime=(Get-Date).AddDays(-7)} -MaxEvents 10 -ErrorAction Stop })) { [void]$o.Add(('  LIMITACION DE CPU O TERMICA {0}  Id {1}' -f $e.TimeCreated, $e.Id)) }
    foreach ($e in (Seg { Get-WinEvent -FilterHashtable @{LogName='System';Id=41,6008,7,51,129,153,157;StartTime=(Get-Date).AddDays(-30)} -MaxEvents 25 -ErrorAction Stop })) { [void]$o.Add(('  {0}  Id {1}  {2}' -f $e.TimeCreated, $e.Id, $e.ProviderName)) }
    [void]$o.Add('')
    [void]$o.Add('  --- ERRORES DE APLICACIONES AGRUPADOS (ultimos 7 dias) ---')
    $g2 = Seg { Get-WinEvent -FilterHashtable @{LogName='Application';Level=1,2;StartTime=(Get-Date).AddDays(-7)} -MaxEvents 400 -ErrorAction Stop | Group-Object Id,ProviderName | Sort-Object Count -Descending | Select-Object -First 10 }
    if ($g2) { foreach ($g in $g2) { [void]$o.Add(('  x{0,-5} {1}' -f $g.Count, $g.Name)) } } else { [void]$o.Add('  sin errores o sin permisos') }
    return $o
}
function Rec-Energia {
    $o = New-Object System.Collections.ArrayList
    [void]$o.Add('  --- PLAN DE ENERGIA ACTIVO ---')
    foreach ($l in (Seg { powercfg /getactivescheme })) { [void]$o.Add('  ' + $l) }
    [void]$o.Add('')
    [void]$o.Add('  --- LIMITES DEL PROCESADOR Y BOOST (afectan los picos de ventilador) ---')
    foreach ($l in (Seg { powercfg /query SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMIN })) { if ($l -match 'Índice|Index|valor actual|Current AC|Current DC') { [void]$o.Add('  MIN  : ' + ($l -replace '\s+',' ')) } }
    foreach ($l in (Seg { powercfg /query SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMAX })) { if ($l -match 'Índice|Index|valor actual|Current AC|Current DC') { [void]$o.Add('  MAX  : ' + ($l -replace '\s+',' ')) } }
    foreach ($l in (Seg { powercfg /query SCHEME_CURRENT SUB_PROCESSOR PERFBOOSTMODE })) { if ($l -match 'Índice|Index|valor actual|Current AC|Current DC') { [void]$o.Add('  BOOST: ' + ($l -replace '\s+',' ')) } }
    [void]$o.Add('')
    [void]$o.Add('  --- INICIO RAPIDO E HIBERNACION ---')
    $pw = Seg { Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' }
    if ($pw) {
        [void]$o.Add(('  {0,-30} {1}' -f 'Inicio rapido (1 = activado)', $pw.HiberbootEnabled))
        [void]$o.Add(('  {0,-30} {1}' -f 'Hibernacion', $pw.HibernateEnabled))
    }
    [void]$o.Add('')
    [void]$o.Add('  --- TEMPERATURAS ACPI (si el equipo las expone) ---')
    $tz = Seg { Get-CimInstance -Namespace root/wmi -ClassName MSAcpi_ThermalZoneTemperature }
    [void]$o.Add('  (valores orientativos: las zonas termicas ACPI no son la temperatura real del procesador)')
    if ($tz) { foreach ($t in $tz) { [void]$o.Add(('  Zona termica: {0:N1} C' -f (($t.CurrentTemperature / 10) - 273.15))) } } else { [void]$o.Add('  el equipo no expone temperaturas por ACPI: usar HWiNFO64 en modo sensores') }
    return $o
}

function Rec-Disco {
    $o = New-Object System.Collections.ArrayList
    $destino = Join-Path $OutDir 'prueba_velocidad.bin'
    [void]$o.Add('  Prueba de 256 MB escribiendo y leyendo en la carpeta del informe (el archivo se borra al terminar).')
    UI-Texto 'Prueba de velocidad del disco (256 MB)...'
    $buf = New-Object byte[] (1MB)
    (New-Object Random).NextBytes($buf)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $fs = [IO.File]::Create($destino)
    for ($i = 0; $i -lt 256; $i++) { $fs.Write($buf, 0, $buf.Length) }
    $fs.Flush($true); $fs.Close(); $sw.Stop()
    $global:DiscoEscrituraMBs = [Math]::Round((256 / $sw.Elapsed.TotalSeconds), 1)
    [void]$o.Add(('  {0,-30} {1:N1} MB/s' -f 'Escritura secuencial', $global:DiscoEscrituraMBs))
    $sw.Restart()
    $rs = [IO.File]::OpenRead($destino)
    $b2 = New-Object byte[] (1MB)
    while ($rs.Read($b2, 0, $b2.Length) -gt 0) { }
    $rs.Close(); $sw.Stop()
    $global:DiscoLecturaMBs = [Math]::Round((256 / $sw.Elapsed.TotalSeconds), 1)
    [void]$o.Add(('  {0,-30} {1:N1} MB/s' -f 'Lectura secuencial', $global:DiscoLecturaMBs))
    Remove-Item $destino -Force -ErrorAction SilentlyContinue
    [void]$o.Add('')
    [void]$o.Add('  Referencias: SSD SATA sano 300-550 MB/s, NVMe 1000-3000 MB/s.')
    [void]$o.Add('  Menos de 150 MB/s de escritura = disco degradado o saturado.')
    return $o
}

function Rec-Resumen {
    $o = New-Object System.Collections.ArrayList
    [void]$o.Add('  --- RESUMEN AUTOMATICO DE HALLAZGOS ---')
    $cs = Seg { Get-CimInstance Win32_ComputerSystem }
    $cp = Seg { Get-CimInstance Win32_Processor | Select-Object -First 1 }
    $os = Seg { Get-CimInstance Win32_OperatingSystem }
    $mods = @(Seg { Get-CimInstance Win32_PhysicalMemory })
    if ($cs) { [void]$o.Add('  Equipo: ' + $cs.Manufacturer + ' ' + $cs.Model) }
    if ($cp) { [void]$o.Add('  CPU: ' + $cp.Name) }
    [void]$o.Add('  Modulos de RAM detectados: ' + $mods.Count + $(if ($mods.Count -le 1) { '  <-- SINGLE CHANNEL (rendimiento limitado)' } else { '  (doble canal, correcto)' }))
    if ($os) {
        $libre = [Math]::Round($os.FreePhysicalMemory / 1MB, 1)
        [void]$o.Add('  RAM libre: ' + $libre + ' GB de ' + [int]($os.TotalVisibleMemorySize / 1MB) + ' GB')
        if ($libre -lt 1.5) { [void]$o.Add('  ALERTA: muy poca RAM libre, es probable que la lentitud venga de ahi.') }
    }
    $arranque = Uptime-Texto
    [void]$o.Add('  Encendida desde: ' + $arranque)
    $d = Seg { Get-MpComputerStatus }
    if ($d) {
        [void]$o.Add('  Defender activo: ' + $d.RealTimeProtectionEnabled + '   firmas: ' + $d.AntivirusSignatureLastUpdated)
        if ($d.RealTimeProtectionEnabled -eq $false) { [void]$o.Add('  ALERTA: la proteccion en tiempo real esta APAGADA.') }
    }
    $malos = 0
    foreach ($p in (Seg { Get-Process | Where-Object Path | Sort-Object { $_.Path } -Unique })) {
        if ((Seg { (Get-AuthenticodeSignature $p.Path).Status }) -notin @('Valid', $null)) { $malos++ }
    }
    [void]$o.Add('  Procesos activos sin firma digital valida: ' + $malos)
    $iob = Seg { Get-CimInstance Win32_Service | Where-Object { $_.Name -match 'iobit|driverbooster' } }
    if ($iob) { [void]$o.Add('  ATENCION: hay servicios de Driver Booster/IObit activos (recomendado desinstalar).') } else { [void]$o.Add('  Sin rastros de Driver Booster/IObit en servicios.') }
    $ev = Seg { Get-WinEvent -FilterHashtable @{LogName='System';Id=41;StartTime=(Get-Date).AddDays(-30)} -MaxEvents 1 -ErrorAction Stop }
    if ($ev) { [void]$o.Add('  ATENCION: hubo apagados inesperados (Id 41) en los ultimos 30 dias.') }
    return $o
}
function Nuevo-H($titulo, $valor, $detalle, $estado) {
    return [pscustomobject]@{ Titulo = $titulo; Valor = $valor; Detalle = $detalle; Estado = $estado }
}

function Rec-Hallazgos {
    $h = New-Object System.Collections.ArrayList

    # --- 1. Equipo y Windows ---
    $cs = Seg { Get-CimInstance Win32_ComputerSystem }
    $os = Seg { Get-CimInstance Win32_OperatingSystem }
    $nombre = $env:COMPUTERNAME
    if ($cs) { $nombre = ('' + $cs.Manufacturer + ' ' + $cs.Model) }
    $det = ''
    if ($os) { $det = ('' + $os.Caption + '  build ' + $os.BuildNumber) }
    [void]$h.Add((Nuevo-H 'Equipo y Windows' $nombre ($det + '   |   encendida hace ' + (Uptime-Texto)) 'OK'))

    # --- 2. Procesador ---
    $cp = Seg { Get-CimInstance Win32_Processor | Select-Object -First 1 }
    if ($cp) {
        $carga = [int](Seg { (Get-CimInstance Win32_PerfFormattedData_PerfOS_Processor | Where-Object Name -eq '_Total').PercentProcessorTime })
        $est = 'OK'
        if ($carga -ge 70) { $est = 'ATENCION' }
        [void]$h.Add((Nuevo-H 'Procesador' $cp.Name ('' + $cp.NumberOfCores + ' nucleos / ' + $cp.NumberOfLogicalProcessors + ' hilos   |   carga en este momento ' + $carga + ' %   |   ' + $cp.CurrentClockSpeed + ' MHz') $est))
    }

    # --- 3. Memoria RAM ---
    $mods = @(Seg { Get-CimInstance Win32_PhysicalMemory })
    $libre = 0; $total = 0
    if ($os) { $libre = [Math]::Round($os.FreePhysicalMemory / 1MB, 1); $total = [int]($os.TotalVisibleMemorySize / 1MB) }
    $canal = 'doble canal (correcto)'
    $est = 'OK'
    if ($mods.Count -le 1) { $canal = 'UN SOLO MODULO: single channel, limita mucho el rendimiento de la grafica integrada'; $est = 'ATENCION' }
    if ($libre -lt 1.5) { $canal += '   |   POCA RAM LIBRE'; $est = 'REVISAR' }
    [void]$h.Add((Nuevo-H 'Memoria RAM' ('' + $total + ' GB en ' + $mods.Count + ' modulo(s)') ('' + $libre + ' GB libres   |   ' + $canal) $est))

    # --- 4. Graficos ---
    $gpu = Seg { Get-CimInstance Win32_VideoController | Select-Object -First 1 }
    if ($gpu) {
        $est = 'OK'; $fd = ''
        if ($gpu.DriverDate) {
            $fd = ('' + $gpu.DriverDate)
            try { if (((Get-Date) - $gpu.DriverDate).TotalDays -gt 730) { $est = 'ATENCION'; $fd += ' (driver viejo)' } } catch { }
        }
        [void]$h.Add((Nuevo-H 'Graficos' $gpu.Name ('Driver ' + $gpu.DriverVersion + '   |   fecha ' + $fd) $est))
    }

    # --- 5. Almacenamiento ---
    $discos = @(Seg { Get-PhysicalDisk })
    $salud = 'Healthy'; $nombres = @()
    foreach ($d in $discos) { $nombres += $d.FriendlyName; if (('' + $d.HealthStatus) -ne 'Healthy') { $salud = ('' + $d.HealthStatus) } }
    $vol = Seg { Get-Volume | Where-Object { $_.DriveLetter -eq 'C' } | Select-Object -First 1 }
    $lib = 0; $tt = 0
    if ($vol) { $lib = [int]($vol.SizeRemaining/1GB); $tt = [int]($vol.Size/1GB) }
    $est = 'OK'; $det = ('{0} discos   |   C: {1} GB libres de {2} GB' -f $discos.Count, $lib, $tt)
    if ($salud -ne 'Healthy') { $est = 'REVISAR'; $det += '   |   SALUD: ' + $salud }
    if ($tt -gt 0 -and ($lib / $tt) -lt 0.10) { $est = 'REVISAR'; $det += '   |   POCO ESPACIO LIBRE' }
    if ($nombres.Count -gt 0) { [void]$h.Add((Nuevo-H 'Almacenamiento' ($nombres -join ', ') $det $est)) }

    # --- 6. Velocidad del disco ---
    $est = 'OK'; $val = 'no medido'; $det = 'Marca la opcion "prueba de disco" y volve a analizar para medirlo.'
    if ($global:DiscoEscrituraMBs) {
        $val = ('' + $global:DiscoEscrituraMBs + ' MB/s de escritura')
        $det = ('Lectura ' + $global:DiscoLecturaMBs + ' MB/s   |   Referencia: SSD SATA sano 300-550, NVMe 1000-3000')
        if ([double]$global:DiscoEscrituraMBs -lt 150) { $est = 'REVISAR'; $det += '   |   MUY LENTO' }
        elseif ([double]$global:DiscoEscrituraMBs -lt 300) { $est = 'ATENCION' }
    }
    [void]$h.Add((Nuevo-H 'Velocidad del disco' $val $det $est))

    # --- 7. Windows Update ---
    $hf = Seg { Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 1 }
    $val = 'no consultado'; $det = 'Marca la casilla "Consultar Windows Update" para saberlo.'; $est = 'OK'
    if ($hf) { $val = 'ultima: ' + $hf.HotFixID; $det = ('Instalada el ' + $hf.InstalledOn) }
    if ($global:UpdPendientes -ne $null) {
        $nDrv = @($global:DrvPendientes).Count
        $que = ('' + $global:UpdSoftware + ' de Windows')
        if ($nDrv -gt 0) { $que += ' + ' + $nDrv + ' de drivers' }
        if ([int]$global:UpdPendientes -gt 0) { $est = 'ATENCION'; $det += ('   |   ' + $global:UpdPendientes + ' PENDIENTES (' + $que + ')') }
        else { $det += '   |   sin actualizaciones pendientes' }
    }
    [void]$h.Add((Nuevo-H 'Windows Update' $val $det $est))

    # --- 8. Seguridad ---
    $d = Seg { Get-MpComputerStatus }
    $exR = @(Seg { (Get-MpPreference).ExclusionPath } | Where-Object { $_ -and $_ -notlike 'N/A*' })
    $amen = @(Seg { Get-MpThreatDetection })
    $est = 'OK'; $val = 'Windows Defender activo'; $det = ''
    if ($d) {
        $det = 'Proteccion en tiempo real: ' + $d.RealTimeProtectionEnabled + '   |   firmas: ' + $d.AntivirusSignatureLastUpdated
        if ($d.RealTimeProtectionEnabled -eq $false) { $est = 'REVISAR'; $det += '   |   PROTECCION APAGADA' }
    } else { $est = 'ATENCION'; $val = 'sin datos de Defender'; $det = 'Defender no responde (puede haber otro antivirus instalado)' }
    if ($exR.Count -gt 0) { $est = 'REVISAR'; $det += ('   |   ' + $exR.Count + ' exclusion(es) de rutas: revisar') }
    if ($amen.Count -gt 0) { $est = 'REVISAR'; $det += ('   |   ' + $amen.Count + ' amenaza(s) en el historial') }
    $tsr = Seg { (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server').fDenyTSConnections }
    if ($tsr -eq 0) { if ($est -eq 'OK') { $est = 'ATENCION' }; $det += '   |   Escritorio remoto HABILITADO (desactivalo si no lo usas)' }
    [void]$h.Add((Nuevo-H 'Seguridad (Defender)' $val $det $est))

    # --- 9. Programas de arranque ---
    $ini = @(Seg { Get-CimInstance Win32_StartupCommand })
    $tar = @(Seg { Get-ScheduledTask | Where-Object { $_.TaskPath -notlike '\Microsoft\*' } | Where-Object { $_.Triggers | Where-Object { $_.CimClass.CimClassName -match 'Boot|Logon' } } })
    $est = 'OK'
    if ($ini.Count -gt 15 -or $tar.Count -gt 10) { $est = 'ATENCION' }
    [void]$h.Add((Nuevo-H 'Programas de arranque' ('' + $ini.Count + ' programas de inicio') ('' + $tar.Count + ' tareas programadas corren al encender o iniciar sesion   |   menos es mas: todas consumen RAM y CPU al arrancar') $est))

    # --- 10. Procesos sospechosos ---
    $sinFirma = 0
    foreach ($p in (Seg { Get-Process | Where-Object Path | Sort-Object { $_.Path } -Unique })) {
        if ((Seg { (Get-AuthenticodeSignature $p.Path).Status }) -notin @('Valid', $null)) { $sinFirma++ }
    }
    $enUsr = @(Seg { Get-Process | Where-Object { $_.Path -match 'AppData|ProgramData|Users\\Public|Downloads' } }).Count
    $est = 'OK'
    if ($sinFirma -gt 0) { $est = 'ATENCION' }
    if ($enUsr -gt 3) { $est = 'ATENCION' }
    [void]$h.Add((Nuevo-H 'Procesos en ejecucion' ('' + $sinFirma + ' sin firma digital valida') ('' + $enUsr + ' corriendo desde carpetas de usuario o temporales   |   si no los reconoces, hay que mirarlos en el detalle') $est))

    # --- 11. Picos de CPU ---
    $val = 'sin muestreo'; $det = 'Muestreo desactivado (0 segundos).'; $est = 'OK'
    if ($MuestreoSegundos -gt 0) {
        $val = ('pico maximo ' + $global:MuestreoMaxCPU + ' % de CPU total')
        $det = ('Proceso que mas CPU acumulo: ' + $global:MuestreoTopProc + ' (' + $global:MuestreoTopPorc + ' % del tiempo)   |   muestreo de ' + $MuestreoSegundos + ' s')
        if ([int]$global:MuestreoMaxCPU -ge 90) { $est = 'ATENCION' }
    }
    [void]$h.Add((Nuevo-H 'Uso de CPU (picos)' $val $det $est))

    # --- 12. Energia y temperatura ---
    $e37 = @(Seg { Get-WinEvent -FilterHashtable @{LogName='System';ProviderName='Microsoft-Windows-Kernel-Processor-Power';Id=37;StartTime=(Get-Date).AddDays(-7)} -MaxEvents 20 -ErrorAction Stop })
    $e55 = @(Seg { Get-WinEvent -FilterHashtable @{LogName='System';ProviderName='Microsoft-Windows-Kernel-Processor-Power';Id=55;StartTime=(Get-Date).AddDays(-7)} -MaxEvents 60 -ErrorAction Stop })
    $plan = Seg { (powercfg /getactivescheme) -join ' ' }
    $est = 'OK'; $det = 'Plan de energia: ' + $plan
    if ($e37.Count -gt 0) { $est = 'ATENCION'; $det += ('   |   ' + $e37.Count + ' avisos de limitacion del procesador por firmware o temperatura (Id 37)') }
    if ($e55.Count -gt 0) { $det += ('   |   ' + $e55.Count + ' avisos informativos de energia del CPU (Id 55: es normal, uno por nucleo al cargar el perfil)') }
    [void]$h.Add((Nuevo-H 'Energia y temperatura' 'Ventilador y rendimiento' $det $est))

    # --- 13. Estabilidad ---
    $errores = @(Seg { Get-WinEvent -FilterHashtable @{LogName='System';Level=1,2;StartTime=(Get-Date).AddDays(-7)} -MaxEvents 300 -ErrorAction Stop })
    $cortes  = @(Seg { Get-WinEvent -FilterHashtable @{LogName='System';Id=41;StartTime=(Get-Date).AddDays(-30)} -MaxEvents 10 -ErrorAction Stop })
    $edisco  = @(Seg { Get-WinEvent -FilterHashtable @{LogName='System';Id=7,51,129,153,157;StartTime=(Get-Date).AddDays(-30)} -MaxEvents 40 -ErrorAction Stop } | Where-Object { ('' + $_.ProviderName) -notmatch 'Kernel-Boot' })
    $whea    = @(Seg { Get-WinEvent -FilterHashtable @{LogName='System';ProviderName='Microsoft-Windows-WHEA-Logger';StartTime=(Get-Date).AddDays(-30)} -MaxEvents 20 -ErrorAction Stop })
    $est = 'OK'
    if ($errores.Count -gt 20) { $est = 'ATENCION' }
    if ($cortes.Count -gt 0 -or $edisco.Count -gt 0) { $est = 'REVISAR' }
    if ($whea.Count -gt 0 -and $est -eq 'OK') { $est = 'ATENCION' }
    $det = ('' + $errores.Count + ' errores criticos en 7 dias   |   ' + $cortes.Count + ' apagados inesperados   |   ' + $edisco.Count + ' errores de disco reales   |   ' + $whea.Count + ' avisos de hardware WHEA (30 dias)')
    [void]$h.Add((Nuevo-H 'Estabilidad del sistema' 'Errores y apagados' $det $est))

    # --- 14. Arranque de Windows ---
    $ev = Seg { Get-WinEvent -LogName Microsoft-Windows-Diagnostics-Performance/Operational -MaxEvents 80 -ErrorAction Stop }
    if ($ev) {
        $u100 = $ev | Where-Object Id -eq 100 | Select-Object -First 1
        $c101 = @($ev | Where-Object Id -eq 101)
        $val = 'sin medicion de arranque'
        if ($u100) { $val = ('ultimo arranque medido: ' + ('{0:yyyy-MM-dd HH:mm}' -f $u100.TimeCreated)) }
        $det = ('' + $c101.Count + ' avisos de arranque lento en el registro reciente')
        if ($u100 -and $u100.TimeCreated -lt (Get-Date).AddDays(-7)) { $det += '   |   la medicion es vieja: con el Inicio rapido activado Windows deja de registrar la duracion del arranque' }
        $masRep = $c101 | Group-Object { $m = ($_.Message -replace '\s+',' '); if ($m.Length -gt 60) { $m = $m.Substring(0,60) }; $m } | Sort-Object Count -Descending | Select-Object -First 1
        if ($masRep) { $det += '   |   lo que mas lo frena: ' + $masRep.Name }
        $est = 'OK'
        if ($c101.Count -gt 5) { $est = 'ATENCION' }
        [void]$h.Add((Nuevo-H 'Arranque de Windows' $val $det $est))
    } else {
        [void]$h.Add((Nuevo-H 'Arranque de Windows' 'sin datos' 'El registro de rendimiento de arranque necesita ejecutar la app como Administrador.' 'ATENCION'))
    }

    # --- 15. Controladores ---
    $sys = @(Seg { Get-ChildItem (Join-Path $env:SystemRoot 'System32\drivers') -Filter *.sys | Where-Object { $_.LastWriteTime -gt (Get-Date).AddDays(-90) } })
    $iob = Seg { Get-CimInstance Win32_Service | Where-Object { $_.Name -match 'iobit|driverbooster|ascdrv|advancedsystemcare' } }
    $hfUlt = Seg { Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 1 }
    $prob = @($global:DrvProblemas); $viej = @($global:DrvViejos); $pend = @($global:DrvPendientes)
    $val = 'Controladores al dia'; $est = 'OK'
    $det = ('' + $global:DrvTotal + ' controladores instalados   |   ' + $sys.Count + ' archivos .sys modificados en los ultimos 90 dias')
    if ($sys.Count -gt 0 -and $hfUlt -and $hfUlt.InstalledOn) { $det += (' (coincide con la actualizacion de Windows del ' + ('{0:yyyy-MM-dd}' -f $hfUlt.InstalledOn) + ')') }
    if ($prob.Count -gt 0) {
        $est = 'REVISAR'
        $val = ('' + $prob.Count + ' dispositivo(s) con problema de controlador')
        $det += ('   |   ' + (($prob | Select-Object -First 2 | ForEach-Object { $_.Nombre + ' (codigo ' + $_.Codigo + ')' }) -join ' / '))
    } elseif ($pend.Count -gt 0) {
        $est = 'ATENCION'
        $val = ('' + $pend.Count + ' controlador(es) para actualizar')
        $det += ('   |   ' + (($pend | Select-Object -First 2 | ForEach-Object { $_.Titulo }) -join ' / '))
    } elseif ($global:DrvConsultado) {
        $val = 'Sin controladores pendientes'
        $det += '   |   Windows Update no ofrece controladores nuevos'
    }
    if ($viej.Count -gt 0) { $det += ('   |   ' + $viej.Count + ' con fecha de mas de 4 anos') }
    if ($iob) { if ($est -eq 'OK') { $est = 'ATENCION' }; $det += '   |   Driver Booster / IObit instalado: conviene desinstalarlo' }
    [void]$h.Add((Nuevo-H 'Controladores (drivers)' $val $det $est))

    # --- Estado global ---
    $global:NRevisar = @($h | Where-Object { $_.Estado -eq 'REVISAR' }).Count
    $global:NAtender = @($h | Where-Object { $_.Estado -eq 'ATENCION' }).Count
    if ($global:NRevisar -gt 0) { $global:EstadoGlobal = 'REVISAR' }
    elseif ($global:NAtender -gt 0) { $global:EstadoGlobal = 'ATENCION' }
    else { $global:EstadoGlobal = 'OK' }
    return $h
}

function Escribir-Txt {

    $nl = [Environment]::NewLine
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('==========================================================================================' + $nl)
    [void]$sb.Append('INFORME DE PC   ' + $env:COMPUTERNAME + '   ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + $nl)
    [void]$sb.Append('Generado por INFORME_PC.ps1 - solo lectura, no modifica el sistema' + $nl)
    [void]$sb.Append('==========================================================================================' + $nl)
    foreach ($k in $global:Sec.Keys) {
        [void]$sb.Append($nl + '##########################################################################################' + $nl)
        [void]$sb.Append('## ' + $k + $nl)
        [void]$sb.Append('##########################################################################################' + $nl)
        foreach ($l in $global:Sec[$k]) { [void]$sb.Append($l + $nl) }
    }
    $sb.ToString() | Out-File -FilePath $global:Txt -Encoding utf8
}

function Escribir-Html {
    $nl = [Environment]::NewLine
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">' + $nl)
    [void]$sb.Append('<title>Informe de PC - ' + (HtmlEnc $env:COMPUTERNAME) + '</title>' + $nl)
    [void]$sb.Append('<style>body{font-family:Segoe UI,Arial,sans-serif;margin:24px;background:#f6f7f9;color:#1b1f23}' + $nl)
    [void]$sb.Append('h1{font-size:22px;margin:0 0 4px 0}h2{background:#1f6feb;color:#fff;padding:8px 12px;border-radius:6px;font-size:15px;margin:26px 0 8px 0}' + $nl)
    [void]$sb.Append('pre{background:#fff;border:1px solid #d7dbe0;border-radius:6px;padding:12px;overflow-x:auto;font-family:Consolas,monospace;font-size:12px;line-height:1.45}' + $nl)
    [void]$sb.Append('.meta{background:#fff;border:1px solid #d7dbe0;border-radius:6px;padding:12px;font-size:13px}' + $nl)
    [void]$sb.Append('table.tb{width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;font-size:13px}' + $nl)
    [void]$sb.Append('table.tb th{background:#0f172a;color:#fff;text-align:left;padding:8px 10px;font-size:12px}' + $nl)
    [void]$sb.Append('table.tb td{border-bottom:1px solid #e5e7eb;padding:8px 10px;vertical-align:top}' + $nl)
    [void]$sb.Append('.badge{font-weight:700;font-size:11px;border-radius:10px;color:#fff;text-align:center;white-space:nowrap}' + $nl)
    [void]$sb.Append('.badge.ok{background:#16a34a}.badge.warn{background:#d97706}.badge.bad{background:#dc2626}.badge.info{background:#2563eb}' + $nl)
    [void]$sb.Append('</style></head><body>' + $nl)
    [void]$sb.Append('<h1>Informe de PC</h1>' + $nl)
    [void]$sb.Append('<div class="meta"><b>Equipo:</b> ' + (HtmlEnc $env:COMPUTERNAME) + ' &nbsp; <b>Fecha:</b> ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + ' &nbsp; <b>Administrador:</b> ' + $global:Admin + '<br>' + $nl)
    [void]$sb.Append('Reporte generado por INFORME_PC.ps1. Solo lectura: no se modifico ningun ajuste del sistema.</div>' + $nl)
    if ($global:Hallazgos) {
        [void]$sb.Append('<h2>Tablero de estado</h2>' + $nl)
        [void]$sb.Append('<table class="tb"><tr><th>Componente</th><th>Estado</th><th>Valor</th><th>Detalle</th></tr>' + $nl)
        foreach ($x in $global:Hallazgos) {
            $clase = 'ok'
            if ($x.Estado -eq 'ATENCION') { $clase = 'warn' }
            elseif ($x.Estado -eq 'REVISAR') { $clase = 'bad' }
            [void]$sb.Append('<tr><td><b>' + (HtmlEnc $x.Titulo) + '</b></td><td class="badge ' + $clase + '">' + (HtmlEnc $x.Estado) + '</td><td>' + (HtmlEnc $x.Valor) + '</td><td>' + (HtmlEnc $x.Detalle) + '</td></tr>' + $nl)
        }
        [void]$sb.Append('</table>' + $nl)
        [void]$sb.Append('<p><b>Resultado global:</b> ' + (HtmlEnc $global:EstadoGlobal) + '  (' + $global:NRevisar + ' para revisar, ' + $global:NAtender + ' para atender, ' + $global:Hallazgos.Count + ' componentes analizados)</p>' + $nl)
    }
    foreach ($k in $global:Sec.Keys) {
        [void]$sb.Append('<h2>' + (HtmlEnc $k) + '</h2>' + $nl)
        [void]$sb.Append('<pre>' + (HtmlEnc ($global:Sec[$k] -join $nl)) + '</pre>' + $nl)
    }
    [void]$sb.Append('</body></html>')
    $sb.ToString() | Out-File -FilePath $global:Htm -Encoding utf8
}
function Ejecutar-Diagnostico {
    $global:Sec = [ordered]@{}
    $global:Hallazgos = $null
    $global:DiscoEscrituraMBs = $null
    $global:DiscoLecturaMBs = $null
    $global:UpdPendientes = $null
    $global:UpdSoftware = 0
    $global:DrvTotal = 0
    $global:DrvProblemas = @()
    $global:DrvViejos = @()
    $global:DrvPendientes = @()
    $global:DrvConsultado = $false
    $global:MuestreoMaxCPU = 0
    $global:MuestreoTopProc = ''
    $global:MuestreoTopPorc = 0
    $global:NRevisar = 0
    $global:NAtender = 0
    $global:EstadoGlobal = 'OK'
    Sec-Ini '0. RESUMEN DE HALLAZGOS'
    $pasos = New-Object System.Collections.ArrayList
    [void]$pasos.Add(@{ T = '1. IDENTIFICACION DEL EQUIPO Y DEL SISTEMA'; F = { Rec-Identidad } })
    [void]$pasos.Add(@{ T = '2. HARDWARE: CPU, RAM, VIDEO, DISCOS Y BATERIA'; F = { Rec-Hardware } })
    [void]$pasos.Add(@{ T = '3. RENDIMIENTO Y PROCESOS'; F = { Rec-Rendimiento } })
    [void]$pasos.Add(@{ T = ('4. MUESTREO DE CPU (' + $MuestreoSegundos + ' segundos)'); F = { Rec-MuestreoCPU } })
    [void]$pasos.Add(@{ T = '5. PROCESOS SOSPECHOSOS Y CONEXIONES DE RED'; F = { Rec-Sospechosos } })
    [void]$pasos.Add(@{ T = '6. ARRANQUE AUTOMATICO, TAREAS PROGRAMADAS Y SERVICIOS'; F = { Rec-Arranque } })
    [void]$pasos.Add(@{ T = '7. SEGURIDAD: DEFENDER, FIREWALL, CUENTAS Y SECUESTROS'; F = { Rec-Seguridad } })
    [void]$pasos.Add(@{ T = '8. DRIVERS, DISPOSITIVOS SIN CONTROLADOR Y ACTUALIZACIONES'; F = { Rec-Drivers } })
    [void]$pasos.Add(@{ T = '9. EVENTOS, ERRORES Y TIEMPO DE ARRANQUE'; F = { Rec-Eventos } })
    [void]$pasos.Add(@{ T = '10. ENERGIA, TEMPERATURA E INICIO RAPIDO'; F = { Rec-Energia } })
    if ($TestDisco) { [void]$pasos.Add(@{ T = '11. PRUEBA DE VELOCIDAD DEL DISCO'; F = { Rec-Disco } }) }
    $i = 0
    foreach ($p in $pasos) {
        $i++
        UI-Texto ('[' + $i + '/' + $pasos.Count + '] ' + $p.T)
        $res = $null
        try { $res = & $p.F } catch { $res = $null }
        Sec-Ini $p.T
        if ($res) { foreach ($l in $res) { if ($null -ne $l) { [void]$global:Sec[$p.T].Add(('' + $l)) } } }
        if ($global:Sec[$p.T].Count -eq 0) { [void]$global:Sec[$p.T].Add('  (sin datos para esta seccion)') }
    }
    UI-Texto 'Calculando el tablero de estado...'
    $global:Hallazgos = Rec-Hallazgos
    $k = '0. RESUMEN DE HALLAZGOS'
    [void]$global:Sec[$k].Add('  TABLERO DE ESTADO   -   OK: sin problemas   |   ATENCION: para tener en cuenta   |   REVISAR: hay que mirarlo')
    [void]$global:Sec[$k].Add('  ------------------------------------------------------------------------------------------')
    foreach ($x in $global:Hallazgos) {
        [void]$global:Sec[$k].Add(('  [{0,-8}] {1}' -f $x.Estado, $x.Titulo))
        [void]$global:Sec[$k].Add(('             ' + $x.Valor))
        if ($x.Detalle) { [void]$global:Sec[$k].Add(('             ' + $x.Detalle)) }
        [void]$global:Sec[$k].Add('')
    }
    [void]$global:Sec[$k].Add(('  RESULTADO GLOBAL: ' + $global:EstadoGlobal + '    (' + $global:NRevisar + ' para revisar, ' + $global:NAtender + ' para atender, ' + $global:Hallazgos.Count + ' componentes analizados)'))
    [void]$global:Sec[$k].Add('')
    $res = $null
    try { $res = Rec-Resumen } catch { $res = $null }
    if ($res) { foreach ($l in $res) { [void]$global:Sec[$k].Add(('' + $l)) } }
    UI-Texto 'Escribiendo el informe en disco...'
    Escribir-Txt
    Escribir-Html
    return $true
}

function Iniciar-CsvMonitor {
    if (-not (Test-Path $global:CsvMon)) { 'Hora;CPU_Total;Proceso_top;MHz;RAM_porc;Nota' | Out-File -FilePath $global:CsvMon -Encoding utf8 }
}
function Nueva-FilaMonitor($mu, [int]$mhz, [int]$ram, [string]$nota) {
    $ncpu = 0; $top = ''
    if ($mu) { $ncpu = $mu.CPU; $top = $mu.Top }
    return ((Get-Date -Format 'HH:mm:ss') + ';' + $ncpu + ';' + $top + ';' + $mhz + ';' + $ram + ';' + $nota)
}
function Datos-Ahora {
    $mhz = Seg { (Get-CimInstance Win32_Processor).CurrentClockSpeed }
    $os  = Seg { Get-CimInstance Win32_OperatingSystem }
    $ram = 0
    if ($os) { $ram = [Math]::Round((($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / $os.TotalVisibleMemorySize) * 100, 0) }
    return [pscustomobject]@{ MHz = $mhz; RAM = $ram }
}
function Ejecutar-Monitoreo([int]$segundos, [int]$umbral, [switch]$Todo) {
    Iniciar-CsvMonitor
    $nuc = Seg { (Get-CimInstance Win32_ComputerSystem).NumberOfLogicalProcessors }
    if (-not $nuc) { $nuc = 1 }
    $intervalo = 2
    $prev = Nueva-Inst
    $fin = (Get-Date).AddSeconds($segundos)
    $picos = 0
    while ((Get-Date) -lt $fin) {
        Start-Sleep -Seconds $intervalo
        $mu = Get-MuestraCPU $prev $intervalo $nuc
        $prev = $mu.Nueva
        $d = Datos-Ahora
        if ($mu.CPU -ge $umbral) { $picos++; $nota = 'PICO' } else { $nota = '' }
        if ($mu.CPU -ge $umbral -or $Todo) {
            (Nueva-FilaMonitor $mu $d.MHz $d.RAM $nota) | Out-File -FilePath $global:CsvMon -Append -Encoding utf8
            UI-Texto ('   PICO ' + (Get-Date -Format 'HH:mm:ss') + '   CPU ' + $mu.CPU + '%   ->   ' + $mu.Top)
        } else {
            UI-Texto ('   CPU ' + $mu.CPU + '%   ' + $mu.Top)
        }
    }
    UI-Texto ('Monitoreo terminado. Picos por encima de ' + $umbral + '%: ' + $picos)
    return $picos
}
function Mostrar-GUI {
    param([switch]$Autotest)

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $C_FONDO = [System.Drawing.Color]::FromArgb(243, 246, 251)
    $C_CARD  = [System.Drawing.Color]::White
    $C_TXT   = [System.Drawing.Color]::FromArgb(15, 23, 42)
    $C_SUB   = [System.Drawing.Color]::FromArgb(90, 100, 115)
    $C_BORDE = [System.Drawing.Color]::FromArgb(222, 228, 236)
    $C_AZUL  = [System.Drawing.Color]::FromArgb(37, 99, 235)
    $C_OK    = [System.Drawing.Color]::FromArgb(22, 163, 74)
    $C_WARN  = [System.Drawing.Color]::FromArgb(217, 119, 6)
    $C_BAD   = [System.Drawing.Color]::FromArgb(220, 38, 38)
    $C_GRIS  = [System.Drawing.Color]::FromArgb(100, 116, 139)

    $F_TIT = New-Object System.Drawing.Font('Segoe UI', 20, [System.Drawing.FontStyle]::Bold)
    $F_H2  = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
    $F_TXT = New-Object System.Drawing.Font('Segoe UI', 9.5)
    $F_SUB = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $F_VAL = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
    $F_BTN = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $F_ICO = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
    $F_BIG = New-Object System.Drawing.Font('Segoe UI', 22, [System.Drawing.FontStyle]::Bold)
    $F_MONO = New-Object System.Drawing.Font('Consolas', 9)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Diagnostico del equipo'
    $form.Size = New-Object System.Drawing.Size(1120, 820)
    $form.StartPosition = 'CenterScreen'
    $form.MinimumSize = New-Object System.Drawing.Size(980, 660)
    $form.BackColor = $C_FONDO
    $form.Font = $F_TXT

    function Nuevo-Label($texto, $x, $y, $w, $h, $fuente, $color) {
        if ($null -eq $color)  { $color  = [System.Drawing.Color]::FromArgb(15, 23, 42) }
        if ($null -eq $fuente) { $fuente = [System.Drawing.SystemFonts]::DefaultFont }
        if ($null -eq $texto)  { $texto  = '' }
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $texto
        $l.Location = New-Object System.Drawing.Point($x, $y)
        if ($w -gt 0) { $l.Size = New-Object System.Drawing.Size($w, $h) } else { $l.AutoSize = $true }
        $l.Font = $fuente
        $l.ForeColor = $color
        return $l
    }
    $penBorde = New-Object System.Drawing.Pen($C_BORDE, 1)

    function Nueva-Tarjeta($hall) {
        $col = $C_AZUL; $ico = 'i'
        switch (('' + $hall.Estado)) {
            'OK'       { $col = $C_OK;   $ico = [char]0x2714 }
            'ATENCION' { $col = $C_WARN; $ico = '!' }
            'REVISAR'  { $col = $C_BAD;  $ico = [char]0x2716 }
        }
        if ($null -eq $col)    { $col = [System.Drawing.Color]::FromArgb(37, 99, 235) }
        if ($null -eq $C_CARD) { $C_CARD = [System.Drawing.Color]::White }
        $p = New-Object System.Windows.Forms.Panel
        $p.Size = New-Object System.Drawing.Size(516, 116)
        $p.BackColor = $C_CARD
        $p.Margin = New-Object System.Windows.Forms.Padding(6)
        $p.Add_Paint({ param($s, $e) $e.Graphics.DrawRectangle($penBorde, 0, 0, $s.Width - 2, $s.Height - 2) })
        $bar = New-Object System.Windows.Forms.Panel
        $bar.Location = New-Object System.Drawing.Point(0, 0)
        $bar.Size = New-Object System.Drawing.Size(7, 116)
        $bar.BackColor = $col
        $ic = Nuevo-Label $ico 16 12 28 26 $F_ICO $col
        $t1 = Nuevo-Label $hall.Titulo 48 10 340 20 $F_H2 $C_TXT
        $v1 = Nuevo-Label $hall.Valor 48 34 456 26 $F_VAL $col
        $v1.AutoEllipsis = $true
        $d1 = Nuevo-Label $hall.Detalle 48 62 456 46 $F_SUB $C_SUB
        $d1.AutoEllipsis = $true
        [void]$p.Controls.Add($bar)
        [void]$p.Controls.Add($ic)
        [void]$p.Controls.Add($t1)
        [void]$p.Controls.Add($v1)
        [void]$p.Controls.Add($d1)
        return $p
    }

    $vista1 = New-Object System.Windows.Forms.Panel
    $vista1.Dock = 'Fill'; $vista1.BackColor = $C_FONDO
    $vista2 = New-Object System.Windows.Forms.Panel
    $vista2.Dock = 'Fill'; $vista2.BackColor = $C_FONDO; $vista2.Visible = $false
    $vista3 = New-Object System.Windows.Forms.Panel
    $vista3.Dock = 'Fill'; $vista3.BackColor = $C_FONDO; $vista3.Visible = $false
    [void]$form.Controls.Add($vista3)
    [void]$form.Controls.Add($vista2)
    [void]$form.Controls.Add($vista1)

    # ---------------- Vista 1: bienvenida ----------------
    [void]$vista1.Controls.Add((Nuevo-Label 'Diagnostico del equipo' 40 30 700 40 $F_TIT $C_TXT))
    [void]$vista1.Controls.Add((Nuevo-Label 'Analiza el hardware y el software, muestra el estado de cada componente y genera un informe completo para revisar a detalle. No modifica nada del sistema.' 42 78 1000 40 $F_TXT $C_SUB))

    $tarjeta = New-Object System.Windows.Forms.Panel
    $tarjeta.Location = New-Object System.Drawing.Point(40, 130)
    $tarjeta.Size = New-Object System.Drawing.Size(1030, 330)
    $tarjeta.BackColor = $C_CARD
    $tarjeta.Add_Paint({ param($s, $e) $e.Graphics.DrawRectangle($penBorde, 0, 0, $s.Width - 2, $s.Height - 2) })
    [void]$vista1.Controls.Add($tarjeta)

    [void]$tarjeta.Controls.Add((Nuevo-Label 'Un clic, 11 pruebas, 2 a 4 minutos' 30 22 600 30 $F_H2 $C_TXT))
    $lista = Nuevo-Label ("-  Hardware: procesador, memoria RAM, grafica, discos y salud del SSD" + [Environment]::NewLine +
                          "-  Rendimiento: que proceso consume CPU y memoria, y los picos raros" + [Environment]::NewLine +
                          "-  Estabilidad: errores del sistema, apagados inesperados, errores de disco" + [Environment]::NewLine +
                          "-  Arranque: cuanto tarda, que lo frena y todo lo que arranca solo" + [Environment]::NewLine +
                          "-  Seguridad: Defender, exclusiones, amenazas y procesos sin firma valida" + [Environment]::NewLine +
                          "-  Controladores: dispositivos sin driver, drivers viejos y los que pueden actualizarse" + [Environment]::NewLine +
                          "-  Programas instalados y actualizaciones de Windows (incluye drivers pendientes)" + [Environment]::NewLine +
                          "-  Velocidad real del disco (escritura y lectura)" + [Environment]::NewLine +
                          "-  Informe completo en HTML y TXT con todo el detalle") 30 62 620 210 $F_TXT $C_SUB
    [void]$tarjeta.Controls.Add($lista)

    $btnCorrer = New-Object System.Windows.Forms.Button
    $btnCorrer.Text = 'ANALIZAR ESTE EQUIPO'
    $btnCorrer.Size = New-Object System.Drawing.Size(330, 78)
    $btnCorrer.Location = New-Object System.Drawing.Point(668, 96)
    $btnCorrer.BackColor = $C_AZUL
    $btnCorrer.ForeColor = [System.Drawing.Color]::White
    $btnCorrer.FlatStyle = 'Flat'
    $btnCorrer.FlatAppearance.BorderSize = 0
    $btnCorrer.Font = $F_BTN
    $btnCorrer.Cursor = 'Hand'
    [void]$tarjeta.Controls.Add($btnCorrer)
    [void]$tarjeta.Controls.Add((Nuevo-Label 'Se pediran permisos de Administrador: varias pruebas los necesitan.' 668 186 340 40 $F_SUB $C_GRIS))

    $btnOpc = New-Object System.Windows.Forms.Button
    $btnOpc.Text = 'Opciones avanzadas'
    $btnOpc.Size = New-Object System.Drawing.Size(170, 28)
    $btnOpc.Location = New-Object System.Drawing.Point(40, 482)
    $btnOpc.FlatStyle = 'Flat'
    $btnOpc.BackColor = $C_FONDO
    $btnOpc.Cursor = 'Hand'
    [void]$vista1.Controls.Add($btnOpc)

    $opc = New-Object System.Windows.Forms.Panel
    $opc.Location = New-Object System.Drawing.Point(40, 516)
    $opc.Size = New-Object System.Drawing.Size(1030, 90)
    $opc.BackColor = $C_CARD
    $opc.Visible = $false
    $opc.Add_Paint({ param($s, $e) $e.Graphics.DrawRectangle($penBorde, 0, 0, $s.Width - 2, $s.Height - 2) })
    [void]$vista1.Controls.Add($opc)

    $chkDisco = New-Object System.Windows.Forms.CheckBox
    $chkDisco.Text = 'Incluir prueba de velocidad del disco (recomendado)'
    $chkDisco.Location = New-Object System.Drawing.Point(20, 18); $chkDisco.AutoSize = $true; $chkDisco.Checked = $true
    [void]$opc.Controls.Add($chkDisco)
    $chkUpd = New-Object System.Windows.Forms.CheckBox
    $chkUpd.Text = 'Consultar Windows Update: actualizaciones y drivers (tarda 1 a 3 minutos mas)'
    $chkUpd.Location = New-Object System.Drawing.Point(20, 48); $chkUpd.AutoSize = $true
    [void]$opc.Controls.Add($chkUpd)
    [void]$opc.Controls.Add((Nuevo-Label 'Segundos de muestreo de CPU:' 590 20 200 20 $F_TXT $C_TXT))
    $numM = New-Object System.Windows.Forms.NumericUpDown
    $numM.Location = New-Object System.Drawing.Point(790, 18); $numM.Size = New-Object System.Drawing.Size(70, 24)
    $numM.Minimum = 0; $numM.Maximum = 600; $numM.Value = $MuestreoSegundos
    [void]$opc.Controls.Add($numM)

    # ---------------- Vista 2: progreso ----------------
    $cont2 = New-Object System.Windows.Forms.Panel
    $cont2.Dock = 'Fill'
    $cont2.Padding = New-Object System.Windows.Forms.Padding(40, 0, 40, 40)
    [void]$vista2.Controls.Add($cont2)
    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Multiline = $true; $txtLog.ScrollBars = 'Vertical'; $txtLog.WordWrap = $false
    $txtLog.ReadOnly = $true; $txtLog.Font = $F_MONO
    $txtLog.BackColor = [System.Drawing.Color]::White
    $txtLog.BorderStyle = 'FixedSingle'
    $txtLog.Dock = 'Fill'
    [void]$cont2.Controls.Add($txtLog)

    $cab2 = New-Object System.Windows.Forms.Panel
    $cab2.Dock = 'Top'; $cab2.Height = 150; $cab2.BackColor = $C_FONDO
    [void]$vista2.Controls.Add($cab2)
    [void]$cab2.Controls.Add((Nuevo-Label 'Analizando el equipo...' 40 30 700 40 $F_TIT $C_TXT))
    $lblPaso = Nuevo-Label 'Iniciando...' 42 80 1000 24 $F_H2 $C_AZUL
    [void]$cab2.Controls.Add($lblPaso)
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(42, 112); $bar.Size = New-Object System.Drawing.Size(1020, 18)
    [void]$cab2.Controls.Add($bar)

    # ---------------- Vista 3: tablero de resultados ----------------
    $flow = New-Object System.Windows.Forms.FlowLayoutPanel
    $flow.Dock = 'Fill'
    $flow.AutoScroll = $true
    $flow.BackColor = $C_FONDO
    $flow.Padding = New-Object System.Windows.Forms.Padding(32, 8, 32, 8)
    [void]$vista3.Controls.Add($flow)

    $btnBar = New-Object System.Windows.Forms.Panel
    $btnBar.Dock = 'Bottom'; $btnBar.Height = 78; $btnBar.BackColor = $C_FONDO
    [void]$vista3.Controls.Add($btnBar)

    $btnGuardar = New-Object System.Windows.Forms.Button
    $btnGuardar.Text = 'Guardar informe completo...'
    $btnGuardar.Size = New-Object System.Drawing.Size(250, 44)
    $btnGuardar.Location = New-Object System.Drawing.Point(40, 16)
    $btnGuardar.BackColor = $C_OK; $btnGuardar.ForeColor = [System.Drawing.Color]::White
    $btnGuardar.FlatStyle = 'Flat'; $btnGuardar.FlatAppearance.BorderSize = 0
    $btnGuardar.Font = $F_BTN; $btnGuardar.Cursor = 'Hand'
    [void]$btnBar.Controls.Add($btnGuardar)

    $btnAbrir = New-Object System.Windows.Forms.Button
    $btnAbrir.Text = 'Abrir informe (HTML)'
    $btnAbrir.Size = New-Object System.Drawing.Size(230, 44)
    $btnAbrir.Location = New-Object System.Drawing.Point(302, 16)
    $btnAbrir.FlatStyle = 'Flat'; $btnAbrir.Font = $F_TXT; $btnAbrir.Cursor = 'Hand'
    [void]$btnBar.Controls.Add($btnAbrir)

    $btnCopiar = New-Object System.Windows.Forms.Button
    $btnCopiar.Text = 'Copiar resumen'
    $btnCopiar.Size = New-Object System.Drawing.Size(200, 44)
    $btnCopiar.Location = New-Object System.Drawing.Point(544, 16)
    $btnCopiar.FlatStyle = 'Flat'; $btnCopiar.Font = $F_TXT; $btnCopiar.Cursor = 'Hand'
    [void]$btnBar.Controls.Add($btnCopiar)

    $btnRepetir = New-Object System.Windows.Forms.Button
    $btnRepetir.Text = 'Analizar de nuevo'
    $btnRepetir.Size = New-Object System.Drawing.Size(200, 44)
    $btnRepetir.Location = New-Object System.Drawing.Point(756, 16)
    $btnRepetir.FlatStyle = 'Flat'; $btnRepetir.Font = $F_TXT; $btnRepetir.Cursor = 'Hand'
    [void]$btnBar.Controls.Add($btnRepetir)

    $cab3 = New-Object System.Windows.Forms.Panel
    $cab3.Dock = 'Top'; $cab3.Height = 212; $cab3.BackColor = $C_FONDO
    [void]$vista3.Controls.Add($cab3)

    $lblResumenTitulo = Nuevo-Label 'Resultado del analisis' 40 26 900 40 $F_TIT $C_TXT
    [void]$cab3.Controls.Add($lblResumenTitulo)
    $lblResumenSub = Nuevo-Label '' 42 72 1000 20 $F_TXT $C_SUB
    [void]$cab3.Controls.Add($lblResumenSub)

    $banner = New-Object System.Windows.Forms.Panel
    $banner.Location = New-Object System.Drawing.Point(40, 104)
    $banner.Size = New-Object System.Drawing.Size(1030, 92)
    $banner.BackColor = $C_CARD
    $banner.Add_Paint({ param($s, $e) $e.Graphics.DrawRectangle($penBorde, 0, 0, $s.Width - 2, $s.Height - 2) })
    [void]$cab3.Controls.Add($banner)
    $lblBannerIco = Nuevo-Label 'i' 20 22 50 40 $F_BIG $C_AZUL
    [void]$banner.Controls.Add($lblBannerIco)
    $lblBannerTitulo = Nuevo-Label '' 78 16 900 30 $F_H2 $C_TXT
    [void]$banner.Controls.Add($lblBannerTitulo)
    $lblBannerDetalle = Nuevo-Label '' 78 48 900 36 $F_TXT $C_SUB
    [void]$banner.Controls.Add($lblBannerDetalle)

    # ---------------- Estado y funcionalidad ----------------
    $script:log = $txtLog; $script:lblPaso = $lblPaso; $script:bar = $bar
    $script:btnCorrer = $btnCorrer; $script:numM = $numM
    $script:chkDisco = $chkDisco; $script:chkUpd = $chkUpd; $script:opc = $opc
    $script:ocupado = $false; $script:form = $form

    $global:UI = {
        param($t)
        if ($script:lblPaso) { $script:lblPaso.Text = $t }
        if ($script:log) {
            $script:log.AppendText($t + [Environment]::NewLine)
            $script:log.SelectionStart = $script:log.TextLength
            $script:log.ScrollToCaret()
        }
        if ($script:bar) {
            $v = $script:bar.Value + 2
            if ($v -gt 96) { $v = 96 }
            $script:bar.Value = $v
        }
    }

    function Texto-Resumen {
        $l = New-Object System.Collections.ArrayList
        [void]$l.Add('RESUMEN DE LA PC ' + $env:COMPUTERNAME + '   ' + (Get-Date -Format 'yyyy-MM-dd HH:mm'))
        [void]$l.Add('RESULTADO GLOBAL: ' + $global:EstadoGlobal + '   (' + $global:NRevisar + ' para revisar, ' + $global:NAtender + ' para atender)')
        [void]$l.Add('')
        foreach ($x in $global:Hallazgos) {
            [void]$l.Add(('[' + $x.Estado + '] ' + $x.Titulo + ': ' + $x.Valor))
            if ($x.Detalle) { [void]$l.Add('        ' + $x.Detalle) }
        }
        return ($l -join [Environment]::NewLine)
    }

    function Cargar-Tablero {
        $nOk   = $C_OK;   if ($null -eq $nOk)   { $nOk   = [System.Drawing.Color]::FromArgb(22, 163, 74) }
        $nWarn = $C_WARN; if ($null -eq $nWarn) { $nWarn = [System.Drawing.Color]::FromArgb(217, 119, 6) }
        $nBad  = $C_BAD;  if ($null -eq $nBad)  { $nBad  = [System.Drawing.Color]::FromArgb(220, 38, 38) }
        $flow.Controls.Clear()
        foreach ($x in $global:Hallazgos) { [void]$flow.Controls.Add((Nueva-Tarjeta $x)) }
        $lblResumenSub.Text = ((Get-Date -Format 'yyyy-MM-dd HH:mm') + '   |   ' + $global:Hallazgos.Count + ' componentes analizados   |   informe guardado en: ' + $OutDir)
        if ($global:EstadoGlobal -eq 'OK') {
            $banner.BackColor = [System.Drawing.Color]::FromArgb(233, 248, 238)
            $lblBannerIco.Text = [char]0x2714; $lblBannerIco.ForeColor = $nOk
            $lblBannerTitulo.ForeColor = $nOk
            $lblBannerTitulo.Text = 'Todo en orden'
            $lblBannerDetalle.Text = 'No se detecto nada que requiera revision. Podes guardar el informe completo para ver el detalle.'
        } elseif ($global:EstadoGlobal -eq 'ATENCION') {
            $banner.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 233)
            $lblBannerIco.Text = '!'; $lblBannerIco.ForeColor = $nWarn
            $lblBannerTitulo.ForeColor = $nWarn
            $lblBannerTitulo.Text = ('' + $global:NAtender + ' punto(s) para tener en cuenta')
            $lblBannerDetalle.Text = 'No son fallas criticas, pero conviene revisarlos. Mira las tarjetas en amarillo y guarda el informe completo.'
        } else {
            $banner.BackColor = [System.Drawing.Color]::FromArgb(254, 238, 238)
            $lblBannerIco.Text = [char]0x2716; $lblBannerIco.ForeColor = $nBad
            $lblBannerTitulo.ForeColor = $nBad
            $lblBannerTitulo.Text = ('' + $global:NRevisar + ' punto(s) que conviene revisar')
            $lblBannerDetalle.Text = 'Revisa las tarjetas en rojo. Guarda el informe completo y envialo para el analisis detallado.'
        }
    }

    $btnOpc.Add_Click({ $script:opc.Visible = -not $script:opc.Visible })

    $btnAbrir.Add_Click({
        if (Test-Path $global:Htm) { Start-Process $global:Htm }
        else { [System.Windows.Forms.MessageBox]::Show('Todavia no hay informe. Ejecuta el analisis primero.') | Out-Null }
    })
    $btnCopiar.Add_Click({
        if ($global:Hallazgos) { Set-Clipboard -Value (Texto-Resumen); [System.Windows.Forms.MessageBox]::Show('Resumen copiado. Pegalo en el chat.') | Out-Null }
        else { [System.Windows.Forms.MessageBox]::Show('Todavia no hay resultados.') | Out-Null }
    })
    $btnRepetir.Add_Click({
        $vista3.Visible = $false; $vista2.Visible = $false; $vista1.Visible = $true; $vista1.BringToFront()
    })
    $btnGuardar.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = 'Elegi la carpeta donde guardar el informe completo (por ejemplo un pendrive)'
        $dlg.ShowNewFolderButton = $true
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $n = 0
            foreach ($f in @($global:Htm, $global:Txt, (Join-Path $OutDir 'muestreo_cpu.csv'), $global:CsvMon)) {
                if ($f -and (Test-Path $f)) { Copy-Item $f -Destination $dlg.SelectedPath -Force; $n++ }
            }
            [System.Windows.Forms.MessageBox]::Show(('Se copiaron ' + $n + ' archivos a:' + [Environment]::NewLine + $dlg.SelectedPath), 'Informe guardado', 'OK', 'Information') | Out-Null
        }
    })

    $btnCorrer.Add_Click({
        if ($script:ocupado) { return }
        $script:ocupado = $true
        $script:btnCorrer.Enabled = $false
        $script:bar.Value = 0
        $script:log.Clear()
        $global:Bombear = $true
        $script:MuestreoSegundos = [int]$script:numM.Value
        $script:TestDisco = [bool]$script:chkDisco.Checked
        $script:ActualizacionesPendientes = [bool]$script:chkUpd.Checked
        $script:DriversPendientes = [bool]$script:chkUpd.Checked
        $vista1.Visible = $false; $vista3.Visible = $false; $vista2.Visible = $true; $vista2.BringToFront()
        try {
            [void](Ejecutar-Diagnostico)
            $script:bar.Value = 100
            $errTablero = ''
            try { Cargar-Tablero } catch { $errTablero = $_.Exception.Message }
            $global:Bombear = $false
            $vista2.Visible = $false; $vista3.Visible = $true; $vista3.BringToFront()
            try { Set-Clipboard -Value (Texto-Resumen) } catch { }
            if ($errTablero) {
                [System.Windows.Forms.MessageBox]::Show('El informe completo se genero y se guardo bien, pero fallo el dibujado del tablero:' + [Environment]::NewLine + $errTablero + [Environment]::NewLine + [Environment]::NewLine + 'Usa "Abrir informe (HTML)" o "Guardar informe completo..." para ver los resultados.', 'Aviso', 'OK', 'Warning') | Out-Null
            } elseif ($global:EstadoGlobal -ne 'OK') {
                $msj = 'Analisis terminado.'
                if ($global:EstadoGlobal -eq 'ATENCION') { $msj += [Environment]::NewLine + [Environment]::NewLine + 'Hay ' + $global:NAtender + ' punto(s) para tener en cuenta.' }
                else { $msj += [Environment]::NewLine + [Environment]::NewLine + 'Hay ' + $global:NRevisar + ' punto(s) que conviene revisar.' }
                $msj += [Environment]::NewLine + [Environment]::NewLine + 'El resumen ya esta copiado al portapapeles.' + [Environment]::NewLine + 'Usa "Guardar informe completo..." para llevar el detalle.'
                [System.Windows.Forms.MessageBox]::Show($msj, 'Analisis terminado', 'OK', 'Information') | Out-Null
            }
        } catch {
            $det = 'Error inesperado: ' + $_.Exception.Message
            try { if ($_.InvocationInfo) { $det += [Environment]::NewLine + [Environment]::NewLine + 'Linea ' + $_.InvocationInfo.ScriptLineNumber + ': ' + ('' + $_.InvocationInfo.Line).Trim() } } catch { }
            [System.Windows.Forms.MessageBox]::Show($det, 'Error', 'OK', 'Error') | Out-Null
            $vista2.Visible = $false; $vista1.Visible = $true; $vista1.BringToFront()
        } finally {
            $script:ocupado = $false
            $script:btnCorrer.Enabled = $true
            $global:Bombear = $false
        }
    })

    if ($Autotest) {
        $Error.Clear()
        $global:Hallazgos = Rec-Hallazgos
        Cargar-Tablero
        $n = $flow.Controls.Count
        $ctrl = 0
        foreach ($c in $flow.Controls) { $ctrl += $c.Controls.Count }
        $prom = 0
        if ($n -gt 0) { $prom = [int]($ctrl / $n) }
        $form.Dispose()
        Write-Host ('AUTOTEST GUI: tarjetas=' + $n + ' | controles por tarjeta=' + $prom + ' (esperado 5) | errores capturados=' + $Error.Count + ' | estado global=' + $global:EstadoGlobal)
        $i = 0
        foreach ($e in $Error) { $i++; Write-Host ('   ERROR ' + $i + ' L' + $e.InvocationInfo.ScriptLineNumber + ': ' + $e.Exception.Message) }
        return $true
    }
    [void]$form.ShowDialog()
    $form.Dispose()
    return $true
}

# ============================================================================
#  ARRANQUE DEL PROGRAMA
# ============================================================================
if ($TestGUI) {
    [void](Mostrar-GUI -Autotest)
    exit 0
}

if ($Auto) {
    $global:Bombear = $false
    $global:UI = { param($t) Write-Host ('   ' + $t) -ForegroundColor DarkGray }
    Write-Host ''
    Write-Host ('INFORME DE PC   ' + $env:COMPUTERNAME + '   ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -ForegroundColor Cyan
    if (-not $esAdmin) { Write-Host 'AVISO: sin permisos de Administrador, varias secciones quedaran vacias.' -ForegroundColor Yellow }
    Write-Host ('Carpeta de salida: ' + $OutDir) -ForegroundColor Cyan
    [void](Ejecutar-Diagnostico)
    if ($MonitorSegundos -gt 0) {
        Write-Host ''
        Write-Host ('Monitoreando picos durante ' + $MonitorSegundos + ' segundos (umbral ' + $MonitorUmbral + '%)...') -ForegroundColor Cyan
        [void](Ejecutar-Monitoreo -segundos $MonitorSegundos -umbral $MonitorUmbral)
    }
    Write-Host ''
    Write-Host '============= RESUMEN (copia esto y envialo) =============' -ForegroundColor Green
    if ($global:Sec.Contains('0. RESUMEN DE HALLAZGOS')) { foreach ($l in $global:Sec['0. RESUMEN DE HALLAZGOS']) { Write-Host $l } }
    if ($global:Sec.Contains('1. IDENTIFICACION DEL EQUIPO Y DEL SISTEMA')) { foreach ($l in $global:Sec['1. IDENTIFICACION DEL EQUIPO Y DEL SISTEMA']) { Write-Host $l } }
    Write-Host '===========================================================' -ForegroundColor Green
    Write-Host ('Informe HTML : ' + $global:Htm)
    Write-Host ('Informe TXT  : ' + $global:Txt)
    Write-Host ('CSV de CPU   : ' + (Join-Path $OutDir 'muestreo_cpu.csv'))
    try {
        $txt = 'RESUMEN DE LA PC ' + $env:COMPUTERNAME + [Environment]::NewLine
        if ($global:Sec.Contains('0. RESUMEN DE HALLAZGOS')) { $txt += ($global:Sec['0. RESUMEN DE HALLAZGOS'] -join [Environment]::NewLine) }
        if ($global:Sec.Contains('1. IDENTIFICACION DEL EQUIPO Y DEL SISTEMA')) { $txt += [Environment]::NewLine + ($global:Sec['1. IDENTIFICACION DEL EQUIPO Y DEL SISTEMA'] -join [Environment]::NewLine) }
        Set-Clipboard -Value $txt
        Write-Host 'Resumen copiado al portapapeles.' -ForegroundColor Green
    } catch { }
    exit 0
}

[void](Mostrar-GUI)
