@echo off
setlocal enabledelayedexpansion
title Instalador de Microsoft Office - Digitalcode Spa
chcp 65001 >nul
mode con cols=100 lines=40
cd /d "%~dp0"
cls

echo ===================================================
echo  Instalador de Microsoft Office   
echo  Creado por Digitalcode Spa   
echo ===================================================
echo.

:: Verificar permisos de Administrador
net session >nul 2>&1
if %errorLevel% NEQ 0 (
    echo ERROR: Se requieren permisos de Administrador.
    echo Reiniciando con privilegios elevados...
    powershell -Command "Start-Process cmd -ArgumentList '/c %~s0' -Verb RunAs"
    exit
)

:: Verificar si setup.exe existe
if not exist "%~dp0\setup.exe" (
    echo ERROR: No se encontró setup.exe en la carpeta actual.
    pause
    exit /b
)

:: Verificar si Office ya está instalado
wmic product where "name like 'Microsoft Office%%'" get name | findstr /I "Microsoft Office" >nul 2>&1
if %errorlevel% EQU 0 (
    echo ERROR: Office ya está instalado. Desinstálelo antes de continuar.
    pause
    exit /b
)

:: Menú principal
echo ==============================================
echo           SELECCIONE EL TIPO DE OFFICE  
echo ==============================================
echo [1] Versiones RETAIL
echo [2] Versiones LTSC
echo ==============================================
set /p tipo="Ingrese la opción: "
if "%tipo%"=="1" goto RETAIL
if "%tipo%"=="2" goto LTSC
echo Opción inválida. Intente de nuevo.
pause
exit /b

:RETAIL
echo ==============================================
echo            VERSIONES OFFICE RETAIL  
echo ==============================================
set "opciones[1]=Office 2019 Pro Plus (Retail)|ProPlus2019Retail|Current"
set "opciones[2]=Office 2021 Pro Plus (Retail)|ProPlus2021Retail|Current"
set "opciones[3]=Office 2024 Pro Plus (Retail)|ProPlus2024Retail|Current"
set "opciones[4]=Office 2019 Standard (Retail)|Standard2019Retail|Current"
set "opciones[5]=Office 2021 Standard (Retail)|Standard2021Retail|Current"
set "opciones[6]=Office 2024 Standard (Retail)|Standard2024Retail|Current"
goto SELECCION

:LTSC
echo ==============================================
echo            VERSIONES OFFICE LTSC  
echo ==============================================
set "opciones[1]=Office 2019 Pro Plus LTSC|ProPlus2019Volume|PerpetualVL2019"
set "opciones[2]=Office 2021 Pro Plus LTSC|ProPlus2021Volume|PerpetualVL2021"
set "opciones[3]=Office 2024 Pro Plus LTSC|ProPlus2024Volume|PerpetualVL2024"
set "opciones[4]=Office 2019 Standard LTSC|Standard2019Volume|PerpetualVL2019"
set "opciones[5]=Office 2021 Standard LTSC|Standard2021Volume|PerpetualVL2021"
set "opciones[6]=Office 2024 Standard LTSC|Standard2024Volume|PerpetualVL2024"

:SELECCION
echo ==============================================
for /L %%i in (1,1,6) do (
    for /F "tokens=1 delims=|" %%a in ("!opciones[%%i]! ") do echo [%%i]  %%a
)
echo ==============================================
set /p version="Ingrese la opción: "

:: Obtener ProductID y Channel correctamente sin espacios extras
for /L %%i in (1,1,6) do (
    if "!version!"=="%%i" (
        for /F "tokens=2,3 delims=|" %%a in ("!opciones[%%i]! ") do (
            set "ProductID=%%a"
            set "Channel=%%b"
        )
    )
)

:: Eliminar espacios en blanco al final de Channel y ProductID
for /F "tokens=* delims= " %%a in ("%ProductID%") do set "ProductID=%%a"
for /F "tokens=* delims= " %%a in ("%Channel%") do set "Channel=%%a"

if not defined ProductID (
    echo Opción inválida. Intente de nuevo.
    pause
    exit /b
)

:: Selección de arquitectura
echo ==============================================
echo           SELECCIONE LA ARQUITECTURA  
echo ==============================================
echo [1] 32-bit
echo [2] 64-bit
echo ==============================================
set /p arch="Ingrese la opción: "
if "%arch%"=="1" set "OfficeArch=32"
if "%arch%"=="2" set "OfficeArch=64"
if not defined OfficeArch (
    echo Opción inválida. Intente de nuevo.
    pause
    exit /b
)

:: Selección de idioma
echo ==============================================
echo             SELECCIONE EL IDIOMA  
echo ==============================================
echo [1] Español (es-ES)
echo [2] Inglés (en-US)
echo [3] Portugués (pt-BR)
echo ==============================================
set /p lang="Ingrese la opción: "
if "%lang%"=="1" set "LangID=es-es"
if "%lang%"=="2" set "LangID=en-us"
if "%lang%"=="3" set "LangID=pt-br"
if not defined LangID (
    echo Opción inválida. Intente de nuevo.
    pause
    exit /b
)

:: Crear configuration.xml con valores correctos sin mostrarlo al usuario
if exist "%~dp0\configuration.xml" del /f /q "%~dp0\configuration.xml"
(
echo ^<Configuration^>
echo     ^<Add OfficeClientEdition="%OfficeArch%" Channel="%Channel%"^>
echo         ^<Product ID="%ProductID%"^>
echo             ^<Language ID="%LangID%" /^>
echo         ^</Product^>
echo     ^</Add^>
echo     ^<Display Level="Full" AcceptEULA="TRUE" /^>
echo ^</Configuration^>
) > configuration.xml

:: Mostrar solo mensaje de éxito sin contenido del archivo
echo Archivo de configuración creado.

:: Instalar Office
echo Instalando Office... Espere
echo ==============================================
"%~dp0\setup.exe" /configure "%~dp0\configuration.xml"
set errorCode=%errorlevel%

:: Manejo de errores
if %errorCode% NEQ 0 (
    echo ERROR: La instalación falló con código %errorCode%.
    del /f /q "%~dp0\configuration.xml"
    pause
    exit /b
)

:: Limpieza final
del /f /q "%~dp0\configuration.xml"
echo Instalación completada correctamente.
pause
exit