<#
.SYNOPSIS
    Crea un acceso directo en el Escritorio para abrir la interfaz de
    "Lista de Precios Giro Vivo" con doble clic, usando el ícono propio.

.DESCRIPTION
    Corré este script UNA sola vez (o cada vez que quieras volver a crear
    el acceso directo). Apunta a abrir_lista_de_precios.bat y usa
    scripts\icono.ico como ícono del acceso directo.

    Uso:
        powershell -ExecutionPolicy Bypass -File crear_acceso_directo.ps1
#>

$raizProyecto = $PSScriptRoot
$batDestino = Join-Path $raizProyecto "abrir_lista_de_precios.bat"
$icono = Join-Path $raizProyecto "scripts\icono.ico"

if (-not (Test-Path $batDestino)) {
    Write-Error "No se encontró '$batDestino'. Corré este script desde la raíz del proyecto."
    exit 1
}
if (-not (Test-Path $icono)) {
    Write-Warning "No se encontró '$icono'. El acceso directo se crea igual, pero con el ícono por defecto de .bat."
}

$escritorio = [Environment]::GetFolderPath("Desktop")
$rutaAccesoDirecto = Join-Path $escritorio "Lista de Precios Giro Vivo.lnk"

$shell = New-Object -ComObject WScript.Shell
$accesoDirecto = $shell.CreateShortcut($rutaAccesoDirecto)
$accesoDirecto.TargetPath = $batDestino
$accesoDirecto.WorkingDirectory = $raizProyecto
if (Test-Path $icono) {
    $accesoDirecto.IconLocation = $icono
}
$accesoDirecto.Description = "Abrir la interfaz de Lista de Precios Giro Vivo"
$accesoDirecto.Save()

Write-Host "Acceso directo creado en: $rutaAccesoDirecto"
