<#
    Genera la version portable de App Alertas.

        .\build.ps1              -> un unico .exe (portable)
        .\build.ps1 -Carpeta     -> carpeta con el .exe y sus dependencias (arranca antes)
        .\build.ps1 -Limpiar     -> borra build/ y dist/ antes de compilar

    Resultado: dist\App Alertas\  con el .exe, BD_Reportes.xlsx y LEEME.txt
#>
param(
    [switch]$Carpeta,
    [switch]$Limpiar
)

# Los ejecutables (pip, PyInstaller) escriben su progreso en stderr; con "Stop"
# PowerShell lo trataria como error fatal. Se controla por codigo de salida.
$ErrorActionPreference = "Continue"
Set-Location -Path $PSScriptRoot

function Invocar($descripcion, $bloque) {
    & $bloque
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        throw "$descripcion (codigo $LASTEXITCODE)."
    }
}

Write-Host "== App Alertas: empaquetado ==" -ForegroundColor Cyan

if ($Limpiar) {
    Write-Host "Limpiando build/ y dist/ ..."
    Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
}

# --- dependencias --------------------------------------------------------- #
Write-Host "Comprobando dependencias ..."
Invocar "No se pudieron instalar las dependencias" { python -m pip install --quiet --upgrade -r requirements.txt }

# --- icono ---------------------------------------------------------------- #
if (-not (Test-Path "recursos\app.ico")) {
    Write-Host "Generando el icono ..."
    python recursos\crear_icono.py
}

# --- pruebas -------------------------------------------------------------- #
Write-Host "Ejecutando las pruebas ..."
Invocar "Las pruebas no pasaron; se cancela el empaquetado" { python pruebas.py }

# --- compilacion ---------------------------------------------------------- #
$modo = if ($Carpeta) { "--onedir" } else { "--onefile" }
Write-Host "Compilando con PyInstaller ($modo) ..."

$argumentos = @(
    $modo,
    "--windowed",
    "--name", "App Alertas",
    "--icon", "recursos\app.ico",
    "--noconfirm",
    "--clean",
    "--collect-submodules", "openpyxl",
    "--add-data", "recursos\app.ico;recursos",
    "--exclude-module", "tkinter",
    "--exclude-module", "PySide6.QtWebEngineCore",
    "--exclude-module", "PySide6.QtQuick",
    "--exclude-module", "PySide6.Qt3DCore",
    "--exclude-module", "PySide6.QtMultimedia",
    "--exclude-module", "PySide6.QtCharts",
    "--exclude-module", "PySide6.QtDataVisualization",
    # Pillow NO se puede excluir: matplotlib.colors lo importa.
    "--exclude-module", "scipy",
    "--exclude-module", "pandas",
    "App_Alertas.py"
)
Invocar "PyInstaller termino con errores" { python -m PyInstaller @argumentos }

# --- entrega -------------------------------------------------------------- #
$salida = "dist\App Alertas"
if (-not $Carpeta) {
    New-Item -ItemType Directory -Force -Path $salida | Out-Null
    Move-Item -Force "dist\App Alertas.exe" "$salida\App Alertas.exe"
}

if (Test-Path "BD_Reportes.xlsx") {
    Copy-Item -Force "BD_Reportes.xlsx" "$salida\BD_Reportes.xlsx"
} else {
    python crear_bd.py "$salida\BD_Reportes.xlsx"
}
Copy-Item -Force "README.md" "$salida\LEEME.md"
New-Item -ItemType Directory -Force -Path "$salida\logs" | Out-Null

$exe = Get-Item "$salida\App Alertas.exe"
Write-Host ""
Write-Host "Listo: $($exe.FullName)" -ForegroundColor Green
Write-Host ("Tamano: {0:N1} MB" -f ($exe.Length / 1MB))
Write-Host "Copia la carpeta '$salida' completa a donde quieras usarla."
