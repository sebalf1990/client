@echo off
setlocal enabledelayedexpansion

REM ===========================================================================
REM  AO20 - build del cliente VB6 por CLI (plan 05.001)
REM
REM  Gemelo de dev\server\build_tmp.bat (nacido del plan 30.001). Deja
REM  EVIDENCIA auditable en build-stamp.txt: flags reales, resultado y hash
REM  del binario antes/despues.
REM
REM  REGLA OPERATIVA (2026-08-05): esta es la UNICA via valida para compilar
REM  el cliente por CLI. El uso directo de vb6.exe sigue bloqueado por hook.
REM  La regla vieja "cliente solo desde IDE" existia porque un build CLI puede
REM  pisar archivos que el IDE tiene abiertos: el guard de vb6.exe de abajo
REM  neutraliza exactamente ese riesgo.
REM
REM  OJO: el /d de la linea de comandos es DECORATIVO (VB6 lo ignora si el
REM  .vbp tiene CondComp). Los flags reales son los del CondComp de
REM  Argentum20.vbp y este script los copia al stamp tal cual.
REM ===========================================================================

set "VB6=C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe"
set "HERE=%~dp0"
set "PROJ=!HERE!Argentum20.vbp"
set "EXE=!HERE!Argentum.exe"
set "LOG=!HERE!vb6build.log"
set "STAMP=!HERE!build-stamp.txt"

echo ============================================================
echo  AO20 - build del cliente
echo ============================================================

if not exist "!VB6!" goto :sin_compilador
if not exist "!PROJ!" goto :sin_proyecto

REM --- GUARD 1: el IDE VB6 no puede estar abierto (riesgo de pisar archivos) -
tasklist /FI "IMAGENAME eq vb6.exe" 2>nul | find /I "vb6.exe" >nul
if not errorlevel 1 goto :ide_abierto

REM --- GUARD 2: Argentum.exe no puede estar corriendo -----------------------
tasklist /FI "IMAGENAME eq Argentum.exe" 2>nul | find /I "Argentum.exe" >nul
if not errorlevel 1 goto :cliente_corriendo

REM --- flags reales: el CondComp del .vbp -----------------------------------
set "CONDCOMP=(sin CondComp en el .vbp)"
for /f "tokens=* delims=" %%C in ('findstr /B /C:"CondComp" "!PROJ!"') do set "CONDCOMP=%%C"
echo  flags reales: !CONDCOMP!
echo.

REM --- hash previo ----------------------------------------------------------
set "HASH_ANTES=(no habia binario)"
if exist "!EXE!" call :hash "!EXE!" HASH_ANTES

REM --- log limpio (el /out de VB6 APPENDEA) ---------------------------------
if exist "!LOG!" del /q "!LOG!" >nul 2>&1

REM --- compilar -------------------------------------------------------------
echo [1/2] Compilando...
"!VB6!" /make "!PROJ!" /out "!LOG!"
set "VB6_RC=!ERRORLEVEL!"

set "OK=0"
if not exist "!LOG!" goto :sin_log
findstr /C:"succeeded" "!LOG!" >nul 2>&1 && set "OK=1"
findstr /C:"failed" "!LOG!" >nul 2>&1 && set "OK=0"
:sin_log

REM --- hash posterior -------------------------------------------------------
set "HASH_DESPUES=(no se genero binario)"
if exist "!EXE!" call :hash "!EXE!" HASH_DESPUES

REM --- stamp auditable ------------------------------------------------------
echo [2/2] Escribiendo build-stamp.txt...
> "!STAMP!" echo # AO20 build stamp del cliente - lo genera build_client.bat, no editar a mano
>> "!STAMP!" echo fecha          = %DATE% %TIME%
>> "!STAMP!" echo flags          = !CONDCOMP!
>> "!STAMP!" echo proyecto       = !PROJ!
>> "!STAMP!" echo vb6_exit_code  = !VB6_RC!
>> "!STAMP!" echo build_ok       = !OK!
>> "!STAMP!" echo hash_antes     = !HASH_ANTES!
>> "!STAMP!" echo hash_despues   = !HASH_DESPUES!
if exist "!LOG!" >> "!STAMP!" echo # ---- vb6build.log ----
if exist "!LOG!" type "!LOG!" >> "!STAMP!"

echo.
echo ------------------------------------------------------------
if "!OK!"=="0" goto :fallo
if "!HASH_ANTES!"=="!HASH_DESPUES!" goto :sin_cambio
echo  BUILD OK - binario nuevo
echo  hash: !HASH_DESPUES!
echo  evidencia: !STAMP!
echo ------------------------------------------------------------
exit /b 0

:sin_cambio
echo  BUILD OK - pero el binario NO CAMBIO
echo  El hash es igual al de antes: o no habia nada que recompilar,
echo  o VB6 no pudo escribir el archivo.
echo  evidencia: !STAMP!
echo ------------------------------------------------------------
exit /b 0

:fallo
echo  BUILD FALLIDO
echo  detalle en: !LOG!
if exist "!LOG!" type "!LOG!"
echo ------------------------------------------------------------
exit /b 1

:sin_compilador
echo [ERROR] No se encontro el compilador VB6 en:
echo         !VB6!
exit /b 1

:sin_proyecto
echo [ERROR] No se encontro el proyecto: !PROJ!
exit /b 1

:ide_abierto
echo [ERROR] Hay un vb6.exe ABIERTO. Un build CLI puede pisar archivos que el
echo         IDE tiene abiertos (la razon historica de la regla "solo IDE").
echo         Cerra el IDE VB6 y volve a intentar.
exit /b 1

:cliente_corriendo
echo [ERROR] Argentum.exe esta CORRIENDO. VB6 no va a poder sobreescribir el
echo         binario. Cerra el cliente y volve a intentar.
exit /b 1

REM --- subrutina: hash sha256 de %1 en la variable %2 -----------------------
REM  No usar certutil: en Windows en espanol imprime texto extra que se cuela
REM  en el parseo y produce un falso "el binario NO CAMBIO".
:hash
set "_h="
for /f "usebackq delims=" %%H in (`powershell -NoProfile -Command "(Get-FileHash -LiteralPath '%~1' -Algorithm SHA256).Hash"`) do if not defined _h set "_h=%%H"
set "%~2=!_h!"
goto :eof
