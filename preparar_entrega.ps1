<#
    Prepara el paquete que se le envia a otra persona.

        .\preparar_entrega.ps1              usa el .exe existente si esta al dia
        .\preparar_entrega.ps1 -Recompilar  fuerza .\build.ps1 antes de empaquetar
        .\preparar_entrega.ps1 -Carpeta     version en carpeta (util si el antivirus
                                            bloquea el ejecutable de un solo archivo)

    Genera entrega\App Alertas <version>.zip con el programa, una base de datos
    VACIA (sin rutas, correos ni contrasenas de este equipo) y la documentacion.
#>
param(
    [switch]$Recompilar,
    [switch]$Carpeta
)

# Los ejecutables escriben su progreso en stderr; con "Stop" PowerShell lo
# trataria como error fatal. Se controla por codigo de salida.
$ErrorActionPreference = "Continue"
Set-Location -Path $PSScriptRoot

function Invocar($descripcion, $bloque) {
    & $bloque
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        throw "$descripcion (codigo $LASTEXITCODE)."
    }
}

function CopiarVerificado($origen, $destino) {
    # Copy-Item falla en silencio si el destino esta bloqueado (abierto en Excel).
    Copy-Item -Force $origen $destino -ErrorAction SilentlyContinue
    $o = Get-Item $origen
    $d = Get-Item $destino -ErrorAction SilentlyContinue
    if (-not $d -or $d.Length -ne $o.Length) {
        throw ("No se pudo copiar '$origen' a '$destino'. Si ese archivo esta abierto " +
               "en Excel o en otro programa, cierralo y vuelve a intentarlo.")
    }
}

Write-Host "== App Alertas: preparar entrega ==" -ForegroundColor Cyan

# --- version --------------------------------------------------------------- #
$coincidencia = Select-String -Path "app\__init__.py" -Pattern '__version__\s*=\s*"([^"]+)"'
if (-not $coincidencia) { throw "No se pudo leer __version__ de app\__init__.py." }
$version = $coincidencia.Matches[0].Groups[1].Value
Write-Host "Version: $version"

# --- compilar si hace falta ------------------------------------------------- #
$exe = "dist\App Alertas\App Alertas.exe"
$hayQueCompilar = $Recompilar -or -not (Test-Path $exe)

if (-not $hayQueCompilar) {
    $fechaExe = (Get-Item $exe).LastWriteTime
    $fuenteNueva = Get-ChildItem -Recurse -Filter *.py -Path "app" |
        Where-Object { $_.LastWriteTime -gt $fechaExe } |
        Select-Object -First 1
    if ($fuenteNueva) {
        Write-Host "El codigo cambio despues de compilar ($($fuenteNueva.Name)); se recompila."
        $hayQueCompilar = $true
    }
}

if ($hayQueCompilar) {
    if ($Carpeta) {
        Invocar "Fallo la compilacion" { & "$PSScriptRoot\build.ps1" -Carpeta }
    } else {
        Invocar "Fallo la compilacion" { & "$PSScriptRoot\build.ps1" }
    }
} else {
    Write-Host "Se reutiliza el ejecutable ya compilado."
}

# --- armar la carpeta de entrega -------------------------------------------- #
$nombre  = "App Alertas $version"
$destino = "entrega\$nombre"

Remove-Item -Recurse -Force $destino -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $destino | Out-Null

try {

Write-Host "Copiando el programa ..."
if ($Carpeta) {
    Copy-Item -Recurse -Force "dist\App Alertas\*" $destino
    # Se regeneran mas abajo, limpios.
    Remove-Item -Force "$destino\BD_Reportes.xlsx", "$destino\LEEME.md" -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force "$destino\logs" -ErrorAction SilentlyContinue
} else {
    Copy-Item -Force $exe "$destino\App Alertas.exe"
}
New-Item -ItemType Directory -Force -Path "$destino\logs" | Out-Null

Write-Host "Generando una base de datos vacia ..."
$bd = Join-Path (Resolve-Path $destino) "BD_Reportes.xlsx"
Invocar "No se pudo generar la plantilla de base de datos" {
    python crear_bd.py "$bd" --plantilla --forzar
}

Write-Host "Copiando la documentacion ..."
CopiarVerificado "README.md" "$destino\LEEME.md"
# Se reescribe con BOM para que el Bloc de notas muestre bien los acentos.
Get-Content "recursos\PRIMEROS PASOS.txt" -Encoding UTF8 |
    Out-File "$destino\PRIMEROS PASOS.txt" -Encoding utf8

# --- revision antes de enviar ----------------------------------------------- #
Write-Host "Revisando que no se filtren datos de este equipo ..."
Invocar "La revision de la entrega encontro problemas" {
    python revisar_entrega.py "$destino"
}

}
catch {
    # Una carpeta incompleta es peor que ninguna: se puede confundir con una
    # entrega valida y enviarse sin base de datos ni manual.
    Remove-Item -Recurse -Force $destino -ErrorAction SilentlyContinue
    Write-Host ""
    Write-Host "Se canceló la entrega y se borró la carpeta incompleta." -ForegroundColor Yellow
    throw
}

# --- comprimir --------------------------------------------------------------- #
$zip = "entrega\$nombre.zip"
Remove-Item -Force $zip -ErrorAction SilentlyContinue
Write-Host "Comprimiendo ..."
Compress-Archive -Path "$destino\*" -DestinationPath $zip -CompressionLevel Optimal

$archivo = Get-Item $zip
Write-Host ""
Write-Host "Listo: $($archivo.FullName)" -ForegroundColor Green
Write-Host ("Tamano: {0:N1} MB" -f ($archivo.Length / 1MB))
Write-Host ""
Write-Host "Envia ese unico archivo .zip." -ForegroundColor Yellow
Write-Host "NO envies tu propia BD_Reportes.xlsx: su hoja SMTP guarda tu contrasena en texto plano." -ForegroundColor Yellow
