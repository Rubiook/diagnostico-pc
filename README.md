# 🖥️ Diagnóstico de PC — InformePC

**Aplicación portable para Windows** (hecha en PowerShell) que revisa el hardware y el software del
equipo, muestra el estado de cada componente en un **tablero de colores** (verde / naranja / rojo) y
genera un **informe completo en HTML y TXT**.

> ⚠️ **No modifica nada del sistema: solo lee.** Lo único que escribe en el disco es la carpeta donde
> guarda el informe. No instala nada y no necesita .NET extra.

## ✨ Qué hace

- Corre **11 pruebas** sobre el equipo: hardware, rendimiento, procesos sospechosos, arranque
  automático, seguridad, controladores, eventos, energía y disco.
- Arma un **tablero de estado** con una tarjeta por componente y un resumen global.
- **Revisa si hay que actualizar los controladores (drivers)**: dispositivos sin driver o con error,
  drivers muy antiguos y drivers nuevos disponibles en Windows Update.
- Genera un **informe completo en HTML** (para leer o imprimir) y **TXT** (para copiar y pegar),
  más los CSV del muestreo de CPU.
- Incluye un **monitor de picos de CPU** con un botón para marcar el momento exacto en que sube el
  ventilador (para cazar el proceso culpable).

## 📦 Qué hay en la carpeta `app`

| Archivo | Para qué sirve |
|---|---|
| `Ejecutar_Informe_PC.bat` | **Se ejecuta este**: abre la interfaz gráfica (se auto-eleva a Administrador) |
| `Informe_Consola.bat` | Modo consola, sin ventana |
| `INFORME_PC.ps1` | El programa completo (es el único archivo de código) |
| `LEEME.txt` | Manual de uso en texto plano |

## ⬇️ Cómo se usa

1. **Descargá el programa**
   - Botón verde **`Code` → `Download ZIP`** (baja todo el repositorio), o
   - la última versión con el ZIP del programa en **[Releases](../../releases)**.
2. **Descomprimilo** donde quieras (por ejemplo en el Escritorio).
3. Si Windows bloquea los archivos por venir de internet, desbloqueálos una vez (PowerShell):
   `Get-ChildItem "ruta\app" -Recurse | Unblock-File`
4. Doble clic en **`app\Ejecutar_Informe_PC.bat`** y aceptá los permisos de Administrador
   (varias pruebas los necesitan para leer el registro y los eventos).
5. Apretá **ANALIZAR ESTE EQUIPO** y esperá de **2 a 4 minutos**.
6. Al terminar:
   - el **resumen queda copiado al portapapeles** (pegalo donde quieras),
   - **Guardar informe completo...** copia el HTML + TXT + CSV a la carpeta que elijas
     (ideal para llevarlos en un pendrive),
   - **Abrir informe (HTML)** lo abre en el navegador.

### Modo consola (sin ventana)

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\app\INFORME_PC.ps1 -Auto
```

Opciones útiles:

| Opción | Qué hace |
|---|---|
| `-MuestreoSegundos 90` | Muestreo de CPU más largo (por defecto 45) |
| `-TestDisco` | Agrega la prueba de velocidad del disco |
| `-DriversPendientes` | Consulta Windows Update para ver si hay drivers nuevos (usa internet) |
| `-ActualizacionesPendientes` | Consulta Windows Update: actualizaciones de Windows + drivers |
| `-MonitorSegundos 600` | Monitorea 10 minutos y graba `monitor_picos.csv` |
| `-MonitorUmbral 15` | % de CPU para considerar "pico" (por defecto 20) |
| `-OutDir "C:\Temp\rep"` | Cambia la carpeta de salida |
| `-NoElevar` | No pedir elevación (para pruebas) |

## 🔍 Qué revisa en detalle

| # | Prueba | Qué mira |
|---|---|---|
| 1 | Identificación | Equipo, modelo, SKU, BIOS/UEFI, versión de Windows, uptime, firmware |
| 2 | Hardware | CPU, módulos de RAM (detecta *single channel*), GPU y driver, discos, volúmenes, salud SMART, batería |
| 3 | Rendimiento | Top 20 procesos por CPU, top 10 por RAM, carga por núcleo |
| 4 | Muestreo de CPU | 45 s (configurable) con los picos más altos y el proceso culpable |
| 5 | Procesos sospechosos | Sin firma digital válida, corriendo desde AppData/Temp, conexiones de red con su proceso |
| 6 | Arranque automático | Administrador de tareas, claves Run/RunOnce, carpetas de Inicio, tareas programadas no-Microsoft, servicios |
| 7 | Seguridad | Defender (estado, exclusiones, amenazas), firewall, RDP/WinRM, cuentas y administradores, Winlogon/AppInit, hosts, DNS |
| 8 | Controladores | Dispositivos sin driver o con error, drivers más antiguos, drivers nuevos en Windows Update, rastros de Driver Booster/IObit, programas instalados y hotfixes |
| 9 | Eventos | Errores agrupados, tiempo de arranque (Id 100) y qué lo frena (Id 101), apagados inesperados, errores de disco, WHEA, limitación térmica del CPU |
| 10 | Energía | Plan de energía, límites del procesador y boost, inicio rápido, temperaturas ACPI |
| 11 | Velocidad del disco | Escritura y lectura (escribe y borra 256 MB) |

## 🚦 El tablero de estado

| Color | Significado |
|---|---|
| 🟢 **OK** | Todo normal |
| 🟠 **ATENCIÓN** | No es una falla, pero conviene mirarlo |
| 🔴 **REVISAR** | Hay algo que deberías revisar |

Incluye tarjetas de: equipo y Windows, procesador, memoria RAM, gráficos y driver, almacenamiento y
espacio libre, velocidad del disco, Windows Update, seguridad (Defender), programas de arranque,
procesos sin firma válida, picos de CPU, energía y temperatura, estabilidad del sistema, arranque de
Windows y **controladores (drivers)**.

## 🔌 Revisión de controladores (drivers)

Esta parte funciona **sin internet** y sin modificar nada:

- Lee el estado de **cada dispositivo** y traduce el código de error del Administrador de
  dispositivos (por ejemplo `28` = *no tiene los controladores instalados*, `10` = *no puede
  iniciarse*, `43` = *Windows lo detuvo por un error*).
- Lista los **controladores más antiguos** de las clases importantes (video, red, audio, bluetooth,
  cámara, USB, almacenamiento, impresora, batería...), descartando las fechas falsas que muchos
  fabricantes escriben en sus archivos `.inf` (por ejemplo `1968` o `21/06/2006`).
- Detecta si el adaptador de video está usando el **driver genérico de Microsoft** en lugar del del
  fabricante.

Y si tildás la casilla **Consultar Windows Update** (o usás `-DriversPendientes`), además consulta
Windows Update y te dice **qué controladores hay para actualizar**, con modelo, fabricante y fecha.
El informe incluye una guía de **cómo actualizarlos en orden** (y por qué **no** conviene usar
"actualizadores" tipo Driver Booster / IObit).

> 🛈 Windows Update solo ofrece los drivers que Microsoft publica ahí. Para el driver más nuevo de tu
> placa de video o del chipset, siempre conviene bajarlo del sitio oficial del fabricante.

## 🔒 Privacidad

El informe se guarda **solo en tu equipo**. Nada se envía a internet por sí solo: la única parte que
usa la red es la consulta opcional a Windows Update. El contenido puede incluir nombres de usuario,
rutas y nombres de equipos de tu red: revisalo antes de compartirlo públicamente.

## 📋 Requisitos

- Windows 10 u 11 (cualquier edición).
- PowerShell 5.1 (ya viene con Windows). No hace falta instalar nada más.
- Permisos de Administrador para algunas pruebas (registro de arranque, eventos, exclusiones de
  Defender). Si no los das, esas secciones quedan vacías pero el resto funciona igual.

## 🗂️ Estructura del repositorio

```
diagnostico-pc/
├─ app/                      <- el programa (esto es lo que descarga la gente)
│  ├─ Ejecutar_Informe_PC.bat
│  ├─ Informe_Consola.bat
│  ├─ INFORME_PC.ps1
│  └─ LEEME.txt
├─ README.md
├─ LICENSE
└─ .gitignore
```

## 🛠️ Problemas frecuentes

**"No se puede cargar el archivo porque la ejecución de scripts está deshabilitada"**
→ Ejecutá los `.bat` (ya usan `-ExecutionPolicy Bypass`), o a mano:
`powershell -NoProfile -ExecutionPolicy Bypass -File .\app\INFORME_PC.ps1`

**Windows no me deja abrir el `.bat` o dice que el archivo está bloqueado**
→ Los archivos bajados de internet quedan bloqueados:
`Get-ChildItem "ruta\app" -Recurse | Unblock-File`

**Algunas secciones dicen "sin datos" o "requiere Administrador"**
→ Ejecutá `Ejecutar_Informe_PC.bat` y aceptá el cartel de UAC.

## 📄 Licencia

MIT — mirá [LICENSE](LICENSE). Podés usarlo, modificarlo y compartirlo libremente.

---

Hecho para diagnosticar equipos Windows sin instalar nada. Si te sirvió, dejale una ⭐ al repo.

