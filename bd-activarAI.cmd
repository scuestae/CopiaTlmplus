: -------------------------------------------------------------
: Eurotronic 2018 - bd-arcivarAI.cmd - ver. 1.2
: -------------------------------------------------------------
: Activar el After Imaging en un base de datos arrancada (online)
: Requisitos:
: - copiar este archivo en la carpeta TLMP del servidor
: - en esta carpeta debe existir la carpeta BASES y la carpeta TMP
: - En modo comando escribir:
:   bd-activarAI.cmd NombreBaseDatos 
: 
: -------------------------------------------------------------
: Procesos que se ejecutan en este comando:
: ACTIVAR AFTER IMAGE a una bdatos existente
: 1) crear un fichero tlmplusAI.st con las lineas
:    a c:\tlmp\bases\tlmplus.a1
:    a c:\tlmp\bases\tlmplus.a2
:    a c:\tlmp\bases\tlmplus.a3
: 2) ejecutar el comando que escribe para escribir en la bdatos
:    call %dlc%\bin\prostrct add tlmplus tlmplusAI.st
: 3) realiar una copia de seguridad completa
:    %DLC%\BIN\probkup bases\tlmplus bak\tlmplus.bck
: 4) activar el AI
:    %dlc%\bin\rfutil bases\tlmplus -C aimage begin
: -------------------------------------------------------------
: Si fuera necesario tambien puede activar el AI en una bdatos
: que esta arrancada. Por ejemplo para salvar una transaccion.
: 1) la bdatos debe tener habilitada el area AI segun los pasos
:    descritos en el apartado anterior.
: 2) ejecutar el comando:
:    call %dlc%\bin\probkup online bases\tlmplus bak\tlmplus.bck enableai
: -------------------------------------------------------------
 	
@echo off
SETLOCAL ENABLEEXTENSIONS
SETLOCAL ENABLEDELAYEDEXPANSION

:: --- historial de versiones
SET _VERSION=1.0 (c)_Eurotronic &:: inicial
SET _VERSION=1.1 (c)_Eurotronic &:: probado en cliente fnt
SET _VERSION=1.2 (c)_Eurotronic &:: se completa los comentarios iniciales

:: --- definir variables
SET _PathTlmp=%~d0\tlmp
SET _PathTlmpAI=%~d0\tlmp

IF "%1" NEQ "" GOTO 010
 ECHO. Comando para activar el After Imaging en una base de datos OpenEdge
 ECHO. 1) se incorporan los ficheros a1, a2, a3 a la estructura de db-name
 ECHO. 2) se hace una copia online a la carpeta destino path-destino
 ECHO. Sintaxis:
 ECHO.   bd-activarAI ^<db-name^> ^<path-destino^>
 ECHO.   Ejemplo: bd_activarAI tlmplus f:\CopiaTLmplus\Bases1     
 ECHO.   %_VERSION%
GOTO FIN

:010
IF "%2" NEQ "" GOTO 015
 ECHO ERROR: falta el segundo parametro
GOTO FIN

:015
IF EXIST bases GOTO 020
 ECHO ERROR: no se encuentra la carpeta .\bases
GOTO FIN

:020
IF EXIST %_PathTlmp%\bases\%1.db GOTO 030
 ECHO ERROR: No existe el archivo %1.db
GOTO FIN

:030
IF EXIST %2 GOTO 040
 ECHO ERROR: No existe la carpeta destino 
GOTO FIN

:040
:: --- comprobar si la base de datos ya tienen el AI activado
%dlc%\bin\_dbutil prostrct list bases\%1    tmp\temp.st > NUL
FIND /I "%1.a1" < tmp\temp.st
IF %ERRORLEVEL% EQU 0 GOTO 060

    :: --- crear fichero .st temporal para definir los ficheros AI
    ::     deben existir 3 extent como minimo.
    ECHO a  %_PathTlmpAI%\bases\%1.a1 >  tmp\%1.st
    ECHO a  %_PathTlmpAI%\bases\%1.a2 >> tmp\%1.st
    ECHO a  %_PathTlmpAI%\bases\%1.a3 >> tmp\%1.st
    ::

    :050
    :: --- sumar la configuracion a la bdatos
    %dlc%\bin\_dbutil prostrct addonline bases\%1 tmp\%1.st
    IF %ERRORLEVEL% EQU 0 GOTO 060
      ECHO ERROR anadiendo los ficheros AI a la estructura
      ECHO Se intenta activar en la primera copia.
    GOTO 060

:060
:: --- actualizar el fichero de .st
%dlc%\bin\_dbutil prostrct list bases\%1 bases\%1.st > NUL

:: --- hacer un backup activando el AI
call %dlc%\bin\probkup online bases\%1 %2\%1.bck enableai
IF %ERRORLEVEL% EQU 0 GOTO 070
  ECHO ERROR en el backup de la base de datos
GOTO FIN

:070
  ECHO Proceso ejecutado correctamente.
  %dlc%\bin\_dbutil prostrct list bases\%1 %2\%1.st > NUL
:FIN