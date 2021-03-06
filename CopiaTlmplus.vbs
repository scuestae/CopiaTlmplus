'--------------------------------------------------------------------------
' Eurotronic 2017 - CopiaTlmplus.vbs - ver. 4.8
' -------------------------------------------------------------------------
' vbscript para copiar las bases Tlmplus a un directorio
' sintaxis:
'
' cscript CopiaTlmplus.vbs DirectorioDestino
'
' Ejemplo:
' cscript CopiaTlmplus.vbs F:\COPIA_BASES
'
' Fichero de log:       CopiaTlmplus.log
' Fichero de control:   Copia_x
' Fichero de registro:  <DirectorioDestino>\Copia_LOG.txt
'--------------------------------------------------------------------------
' Actualizaciones
'----------------
' Se hace un backup de Progress de las bases en lugar de copiar sus ficheros
' Se cambia la copia de archivos por el comando ROBOCOPY,  por tanto
' no se borra el directorio destino, sino que se actualiza el existente
' Se controla que la base de datos no tenga activado el AI
' Se revisan los mensajes del log de la copia
' Se copian los ficheros <bdatos>.lg por si hubiera que revisarlo
' Se permite enviar el cuerpo del email en formato html
' Se habilita la codificación Unicode en la script y en el email
' Se controla que todas las bases de datos estén paradas (4.8) 
' Se añade tarea de cálculo de costes
' Se añada copia adicional en carpeta "\N"
'
' --------------
' Funcionamiento
' --------------
' Dentro del DirectorioDestino de copia se crean tantas carpetas como indique la 
' variable NumeroCopias. El nombre de estas carpetas es un número:
' 1, 2, 3, etc.
' cada carpeta contiene una copia, por ejemplo de cada día, cuando se completan 
' todas las copias se sobrescribe la carpeta 1.
' También se crea un fichero de control con el nombre de la copia  
' seguido del número de la copia realizada. Cuando el fichero de control no existe, 
' por ejemplo la primera vez, se copia sobre la carpeta 1.
' CarpetaDestino\
'               \1
'               \1\Bases
'               \1\GesDoc
'               \1\Formularios
'               \2
'               \2\Bases
'               \2\GesDoc
'               \2\Formularios
'               \Copia_2
'               \N (opcional, copia para cloudbackup)
' La carpeta Bases contiene un fichero los ficheros necesarios para una restauración
' *.bck  : fichero backup de la base de datos
' *.st   : fichero con la estructura de volumnes de la bdatos
' *.a?   : fichero AI con las transacciones despues del backup
' *.lg   : fichero lg antes de truncarse
' ----------------
' Puesta en marcha
' ----------------
' - Copiar este archivo en una carpeta, por ejemplo la de instalación de Tlmplus
' - Editarlo y modificar los valores de las variables requeridas para la copia.
'   Estas van a depender de cada instalación. Las mas importantes son:
'   DirectoriosOrigen : contiene todas las carpetas que se van a incluir en la copia
'   BasesDeDatos      : contiene las bases de datos que se van copiar por PROBACKUP
'   DirectorioTlmp    : contiene la unidad y la carpeta donde estA instalado Tlmplus
' -----------
' Ejecución
' -----------
' Lo mas práctico es crear una tarea programada que se ejecute todos los días a 
' una hora en la que no haya ningun usuario trabajando, ya que la script realiza
' la parada, truncado y borrado de los ficheros temporales. 
' El siguiente ejemplo ejecuta la script situada en la carpeta C:\Tlmp y copia 
' las carpetas especificadas en la variable DirectoriosOrigen y en la 
' carpeta G:\Copias
' 
' cscript C:\Tlmp\CopiaTlmplus.vbs G:\CopiaTlmplus
' -------------
' RECUPERACIÓN
' -------------
' si no existe la base de datos copiar la estructura
' COPY bak\tlmplus.st bases
'
' ejecutar la restauracion de fichero .bck
' call %dlc%\bin\prorest bases\tlmplus bak\tlmplus.bck
'
' recuperar las transacciones posteriores a la copia anterior 
' call %dlc%\bin\rfutil bases\tlmplus -C roll forward verbose -a bak\tlmplus.a2
'
' arrancar la base de datos 
' call %dlc%\bin\dbman -start -db tlmplus

' hacer un backup activando el AI
' call %dlc%\bin\probkup online bases\%1 %2\%1.bck enableai
'
' en este ejemplo se recupera la base tlmplus desde el directorio bak 
' siempre se deben restaurar las 3 bases de datos  
'--------------------------------------------------------------------------
'
OPTION EXPLICIT

DIM DirectoriosOrigen, BasesDeDatos, DirectorioTlmp, NumeroDeCopias, cFicheroControlOld, fs, nt
DIM lEnViarEmail, ret, lHtml, lAutentificacion, lSSL, cFicheroControlNew, lCopiaEnCarpetaN 
DIM cServidor, cParaEmail, cDeEmail, cAsunto, cMensaje, cAdjunto, FicheroLog, lRecalculosTLMPLUS
DIM cPuerto, cPassword, cUsuario, FicheroControl, lNuevoLOG, nErrores, FicheroLogRC, DirectorioCarpetaN
DIM BasesManenimiento, AppServerAgentes

CONST COPYRIGTH = "(C) Eurotronic ver. 4.8"
CONST SI = TRUE
CONST NO = FALSE
CONST TIEMPOESPERA = 10000
CONST OPCIONESR = " /COPY:DT /MIR /IT /NP /NFL /R:10 /W:1"

' bases de datos y agentes AppServer a los que se les hace mantenimiento
BasesManenimiento  = ARRAY( "tlmplus", "tlmplus1", "tlmplus2", "tlmbd-c",  "tlmbd-d", "tlmp-web", "tlmp-gmao" )
AppServerAgentes   = ARRAY( "ZSSINBBDD", "ZSTlmp1", "ZSTlmp12", "ZSTlmp2", "tlmp-web", "App-SatE" ) 

'--------------------------------------------------------------------------
' INDICAR LOS VALORES REQUERIDOS PARA LA COPIA
'--------------------------------------------------------------------------
' Posibles directorios origen:
' 1) \\servidor01\tlmp            : copia todo el recurso de red
' 2) d:\tlmp\bases                : copia toda la carpeta
' 3) d:\tlmp\bases\tlmplus*.*     : copia solo archivos especificados
' Nota: NO PONER LOS CARACTERES \ NI * AL FINAL, SALVO EN EN CASO 3)
'--------------------------------------------------------------------------

DirectoriosOrigen  = ARRAY( "C:\TLMP\GESDOC" ) ' ARRAY( "E:\TLMP\GESDOC", "E:\TLMP\FORMULARIOS" )
BasesDeDatos       = ARRAY( "tlmplus", "tlmplus1", "tlmplus2" ) ', "tlmp-web" )


DirectorioTlmp     = "C:\TLMP" '"E:\TLMP"
NumeroDeCopias     = 4
lCopiaEnCarpetaN   = SI 
lEnviarEmail       = SI
lRecalculosTLMPLUS = SI 

cParaEmail         = "usuario@eurotronic.es" 
cAsunto            = "EMPRESA: Copia Tlmplus finalizada" 

cDeEmail           = "Copia bases Tlmplus <usuario@eurotronic.es>"
cServidor          = "smtp.empresa.es"
lHtml              = SI
cPuerto            = 587
cUsuario           = "usuario@eurotronic.es"
cPassword          = "********"
lAutentificacion   = SI
lSSL               = NO

lNuevoLOG          = SI
FicheroLog         = ".\" & LEFT(WScript.ScriptName, INSTRREV(WScript.ScriptName,".")-1) & ".log" 
FicheroControl     = "Copia_"
FicheroLogRC       = "Copia_LOG.txt"

'--------------------------------------------------------------------------
Main()
'--------------------------------------------------------------------------

'--------------------------------------------------------------------------
' Inicio de la ejecución
'--------------------------------------------------------------------------
SUB Main()

    DIM nCopia, nCopiaAnterior, DirectorioNCopia, nReturn, I, DirectorioNbases, DirectorioNbasesAnterior
    DIM nSalida, objArgs, DirectorioNCopiaAnterior, DestinoDirectorio, OrigenDirectorio
    
    SET fs = Wscript.CREATEOBJECT("Scripting.FileSystemObject")
    SET nt = WScript.CREATEOBJECT("WScript.Network")

    ' código de salida de la script
    nSalida = 1
    ' contar los errores no criticos
    nErrores = 0 
    
    '---OBTENER LOS ARGUMENTOS
    SET objArgs = WScript.Arguments
    IF WScript.Arguments.Count = 1 THEN
        '--- DIRECTORIO DONDE EXPORTAR
        DestinoDirectorio = objArgs.Unnamed.Item(0)
    ELSE
        WScript.Echo COPYRIGTH & vbCrLF &  "Sintaxis: cscript " & WScript.ScriptName & " <DirectorioDestino> "
        wscript.Quit(nSalida)
    END IF
    
    ' -----------------------------
    ' PREPARAR DESTINO DE COPIA
    ' -----------------------------
    ' --- Reiniciar el LOG en cada ejecución
    IF lNuevoLOG THEN
    	CrearLog
    END IF

    WriteLog "--- Inicio--- " & DATE & " " & TIME  
    WriteLog "Equipo\usuario   : " & nt.ComputerName & "\" & nt.UserName 
    WriteLog "Programa de copia: " & WScript.ScriptFullName & " " & COPYRIGTH
    WriteLog "Carpeta destino  : " & DestinoDirectorio  
    WriteLog "Fichero registro : " & DestinoDirectorio & "\" & FicheroLogRC

    ' CarpetaDestino\
    '               \1
    '               \1\bases

    ' --- OBTENER EL ORDEN DE LA COPIA
    nCopia = NumeroCopia(DestinoDirectorio)
    nCopiaAnterior = NumeroCopiaAnterior(nCopia)
    
    DirectorioNCopia = DestinoDirectorio & "\" & nCopia
    DirectorioNCopiaAnterior = DestinoDirectorio & "\" & nCopiaAnterior
    DirectorioCarpetaN = DestinoDirectorio & "\N" 

    DirectorioNbases = DirectorioNCopia & "\bases"
    DirectorioNbasesAnterior = DirectorioNCopiaAnterior & "\bases"

    ' --- CREAR CARPETA DESTINO
    IF NOT CrearCarpetaDestino(DestinoDirectorio) THEN
        Salir nSalida, DestinoDirectorio & "\" & FicheroLogRC
    END IF
    ' --- CREAR LA SUBCARPETA DE COPIA 1, 2, 3, ETC.
    IF NOT CrearCarpetaDestino(DirectorioNCopia) THEN
        Salir nSalida, DestinoDirectorio & "\" & FicheroLogRC
    END IF
    ' --- CREAR LA SUBCARPETA DE COPIA ANTERIOR 1, 2, 3, ETC.
    IF NOT CrearCarpetaDestino(DirectorioNCopiaAnterior) THEN
        Salir nSalida, DestinoDirectorio & "\" & FicheroLogRC
    END IF
    ' --- CREAR LA SUBCARPETA DE COPIA 1\bases, 2\bases, ETC.
    IF NOT CrearCarpetaDestino(DirectorioNbases) THEN
        Salir nSalida, DestinoDirectorio & "\" & FicheroLogRC
    END IF
    ' --- CREAR LA SUBCARPETA DE COPIA ANTERIOR 1\bases, 2\bases, ETC.
    IF NOT CrearCarpetaDestino(DirectorioNbasesAnterior) THEN
        Salir nSalida, DestinoDirectorio & "\" & FicheroLogRC
    END IF

    ' -----------------------------
    ' BACKUP BASES DE DATOS
    ' -----------------------------
    ' --- Parar Bases de datos
    AccionBases "stop"
    BorrarTMP

    ' copiar bases si todas están paradas
    IF ComprobarBasesParadas = 0 THEN
        
        ' ---- borrar el contenido del directorio destino\ncopia\bases
        fs.DeleteFile DirectorioNbases & "\*.*",TRUE
        IF ERR.number <>0 THEN
            CALL NERROR("ERROR: borrando el contendido de {0}", DestinoDirectorio)	
            Salir nSalida, DestinoDirectorio & "\" & FicheroLogRC
        END IF

        ' --- copiar bases de datos
        FOR I=0 TO UBOUND(BasesDeDatos)
            ' --- backup
            WriteLog "------------------------------"
            WriteLog F2( "Backup base de datos: {0} a la carpeta destino: {1}", BasesDeDatos(I), DirectorioNbases  )
            nReturn = BackupBaseDatos( BasesDeDatos(I), DirectorioNbases, DirectorioNbasesAnterior )
            IF nReturn <> 0 THEN 
                nErrores = nErrores + 1
                WriteLog F2( "ERROR en BackupBaseDatos(): {0}, código de salida: {1}", BasesDeDatos(I), nReturn )
            END IF
            WriteLog "------------------------------"
        NEXT 

        TruncarBi
        TruncarLg
    ELSE 
        nErrores = nErrores + 1
    END IF

    ' --- Iniciar Bases de datos
    AccionBases "start"

    ' ----------------------------------------
    ' COPIAR CARPETAS: GESDOC, FORMULARIOS, N 
    ' ----------------------------------------
    IF CopiaCarpetas(DestinoDirectorio, DirectorioNCopia) THEN
        WriteLog "Copia realizada correctamente. "
        nSalida = 0
    ELSE
        WriteLog "Copia NO REALIZADA !!!. "
        nSalida = 1
    END IF
    WriteLog "Consulte el fichero adjunto para ver los detalles de la copia. "

    ' -----------------------------
    ' ACTUALIZAR FICHERO CONTROL 
    ' -----------------------------
    IF nErrores = 0 THEN
        ActualizarFicheroControl
    END IF

    ' -----------------------------
    ' RECÁLCULOS DIARIOS TLMPLUS
    ' -----------------------------
    If lRecalculosTLMPLUS Then 
        RecalculoTlmplus
    End if

    Salir nSalida, DestinoDirectorio & "\" & FicheroLogRC

END SUB

' -------------------------------------------------------------------------
' Backup base-datos a-carpeta-destino
'--------------------------------------------------------------------------
FUNCTION BackupBaseDatos( BaseDatos, DirectorioDestino, DirectorioDestinoAnterior )
    DIM oShell, nReturn, cFicheroLleno
    SET oShell = WScript.CREATEOBJECT("Wscript.Shell")
    
    BackupBaseDatos = 0

    ' --- comprobar que la bdatos esta parada
    nReturn = oShell.Run( F2("{0}\dlc\bin\_proutil {0}\bases\{1} -C busy", DirectorioTlmp, BaseDatos), 0, TRUE )
    IF nReturn <> 0 THEN
        WriteLog F1( "ERROR: La base de datos {0} no esta parada.", BaseDatos )   
        BackupBaseDatos = 9
        EXIT FUNCTION
    END IF 

    ' --- saber el AI esta activado y copiar si uno de los 3 AI esta lleno
    nReturn = CopiarVaciarAILleno( BaseDatos, DirectorioDestinoAnterior )
    IF nReturn = 0 THEN ' la bdatos SI tiene activado el AI
        ' --- comprobar que hay un AI vacío
        IF NOT ResultadoEjecucion( F2("CMD.EXE /C {0}\dlc\bin\_rfutil {0}\bases\{1} -C aimage extent list", DirectorioTlmp, BaseDatos), "Vacia") THEN
            WriteLog F1( "ERROR: La base de datos {0} no tiene un AI vacío.", BaseDatos )
            BackupBaseDatos = 7   
            EXIT FUNCTION
        END IF 
    END IF

    ' --- realizar el backup
    WriteLog F3("Realizar backup : {0}\dlc\bin\probkup {0}\bases\{1} {2}\{1}.bck -verbose", DirectorioTlmp, BaseDatos, DirectorioDestino)
    nReturn = oShell.Run( F3("{0}\dlc\bin\probkup {0}\bases\{1} {2}\{1}.bck -verbose", DirectorioTlmp, BaseDatos, DirectorioDestino), 0, TRUE )
    IF nReturn <> 0 THEN
        WriteLog F3("{0}\dlc\bin\probkup {0}\bases\{1} {2}\{1}.bck", DirectorioTlmp, BaseDatos, DirectorioDestino)
        WriteLog "ERROR: Realizando el backup, código de salida: " & nReturn
        BackupBaseDatos = 5
        EXIT FUNCTION
    END IF
    
    ' --- verificar la copia
    WriteLog F3("Verificar backup: {0}\dlc\bin\prorest {0}\bases\{1} {2}\{1}.bck -vf", DirectorioTlmp, BaseDatos, DirectorioDestino)
    nReturn = oShell.Run( F3("{0}\dlc\bin\prorest {0}\bases\{1} {2}\{1}.bck -vf", DirectorioTlmp, BaseDatos, DirectorioDestino), 0, TRUE )
    IF nReturn <> 0 THEN
        WriteLog F3("{0}\dlc\bin\prorest {0}\bases\{1} {2}\{1}.bck -vf", DirectorioTlmp, BaseDatos, DirectorioDestino)
        WriteLog "ERROR: Verificando el backup, código de salida: " & nReturn
        BackupBaseDatos = 3
    END IF

    ' --- copiar estructura por si hubiera que recuperar
    nReturn = oShell.Run( F3("{0}\dlc\bin\_dbutil prostrct list {0}\bases\{1} {2}\{1}.st", DirectorioTlmp, BaseDatos, DirectorioDestino), 0, TRUE )
    IF nReturn <> 0 THEN
        WriteLog F3("{0}\dlc\bin\_dbutil prostrct list {0}\bases\{1} {2}\{1}.st", DirectorioTlmp, BaseDatos, DirectorioDestino)
        WriteLog "ERROR: Copiando estructura, código de salida: " & nReturn
        BackupBaseDatos = 2
    END IF 

    ' --- copiar fichero log por si hubiera que revisarlo
    WriteLog F3( "Copiar LG actual: {0}\bases\{1}.lg {2}",  DirectorioTlmp, BaseDatos , DirectorioDestino )
    nReturn = oShell.Run( F3("CMD.EXE /C COPY /Y {0}\bases\{1}.lg {2}", DirectorioTlmp, BaseDatos, DirectorioDestino), 0, TRUE )
    IF nReturn <> 0 THEN 
        WriteLog "--- ERROR al copiar LG actual: " & BaseDatos
        BackupBaseDatos = 1
    END IF
    
    ' --- saber que fichero AI esta lleno
    CopiarVaciarAILleno BaseDatos, DirectorioDestinoAnterior 

    SET oShell = Nothing
END FUNCTION

'--------------------------------------------------------------------------
' Crear la carpeta destino
'--------------------------------------------------------------------------
FUNCTION CrearCarpetaDestino(carpeta)
    DIM o, oShell, nReturn
    
    CrearCarpetaDestino = FALSE

    SET oShell = WScript.CREATEOBJECT("Wscript.Shell")

    ON ERROR RESUME NEXT 
    ' comprobar el acceso al destino y esperar para recuperar la conexion
    SET o = fs.GetFolder(carpeta)
    IF Err.Number <> 0 Then 
	    Err.Clear
	    WScript.Sleep(TIEMPOESPERA)
    END IF
    
    ' comprobar si existe la carpeta destino
    IF Not fs.FolderExists(carpeta) THEN
        fs.CreateFolder(carpeta)
        WriteLog F1("Directorio creado: {0}", carpeta)
        ' si hubo error en el paso anterior
        IF ERR.number <>0 THEN
            CALL NERROR("Error al crear el directorio destino {0}", carpeta)	
            SET oShell = Nothing
            EXIT FUNCTION
		END IF
    END IF
    
    ' comprobar si existe la carpeta destino N
    IF lCopiaEnCarpetaN AND Not fs.FolderExists(DirectorioCarpetaN) THEN
        fs.CreateFolder(DirectorioCarpetaN)
        WriteLog F1("Directorio creado: {0}", DirectorioCarpetaN)
        ' si hubo error en el paso anterior
        IF ERR.number <>0 THEN
            CALL NERROR("Error al crear el directorio destino {0}", DirectorioCarpetaN)	
            SET oShell = Nothing
            EXIT FUNCTION
		END IF
    END IF
    CrearCarpetaDestino = TRUE
    SET oShell = Nothing

END FUNCTION

'--------------------------------------------------------------------------
'  lee las lineas de salida de un comando y busca una cadena
'--------------------------------------------------------------------------
FUNCTION ResultadoEjecucion(cComando, cBuscar)
    ' cComando : comando a ejecutar, debe empezar por "CMD.EXE /C "
    ' cBuscar  : ejecucion correcta si se encuentra
    Dim ObjExec, objShell
    Dim strFromProc, cTipo
    
    IF cBuscar = "" THEN
        cTipo = "_DEVOLVER_"
    ELSE
        cTipo = "_BUSCAR_"
    END IF

    ResultadoEjecucion = FALSE

    WriteLog cComando

    Set objShell = WScript.CreateObject("WScript.Shell")
    Set ObjExec = objShell.Exec(cComando)

    Do
        strFromProc = ObjExec.StdOut.ReadLine()
        WriteLog strFromProc
        ' si cBuscar es vacío devolver en ella el resultado
        IF cTipo = "_DEVOLVER_" AND cBuscar = "" THEN 
            cBuscar = strFromProc
        END IF

        IF cTipo = "_BUSCAR_" AND InStr( UCASE(strFromProc), UCASE(cBuscar) ) THEN
            'WScript.Echo "ENCONTRADO"
            ResultadoEjecucion = TRUE    
        END IF
    Loop While Not ObjExec.Stdout.atEndOfStream
    
    SET objShell= Nothing
    SET ObjExec = Nothing
END FUNCTION

'--------------------------------------------------------------------------
'  Comprueba AI lleno, lo copia Y lo vacia
'--------------------------------------------------------------------------
FUNCTION CopiarVaciarAILleno( BaseDatos, DirectorioDestinoAnterior )
    ' BaseDatos         : base de datoa a  comprobar
    ' DirectorioDestino : carpeta donde copiarlo
    ' retorna 0 si el AI esta activado y 1 si no esa activado
    DIM nReturn, cFicheroLleno, oShell
    
    CopiarVaciarAILleno = 0
    SET oShell = WScript.CREATEOBJECT("Wscript.Shell")
    
    ' --- saber si AI esta activado
    nReturn = oShell.Run( F2("{0}\dlc\bin\_rfutil {0}\bases\{1} -C aimage extent list", DirectorioTlmp, BaseDatos), 0, TRUE )
    ' 0 - si esta activado
    ' 2 - si no esta activado   
    IF nReturn = 0 THEN 
        ' --- saber que fichero AI esta lleno
        ResultadoEjecucion F2("CMD.EXE /C {0}\dlc\bin\_rfutil {0}\bases\{1} -C aimage extent full", DirectorioTlmp, BaseDatos), cFicheroLleno
        IF fs.FileExists( cFicheroLleno ) THEN
            ' --- saber que fichero esta lleno
            WScript.Sleep(TIEMPOESPERA)
            WriteLog F2( "Copiar AI lleno : {0} a {1}.",  cFicheroLleno, DirectorioDestinoAnterior )
            nReturn = oShell.Run( F2("CMD.EXE /C COPY /Y {0} {1}", cFicheroLleno, DirectorioDestinoAnterior), 0, TRUE )
            IF nReturn <> 0 THEN 
                WriteLog "--- ERROR al copiar AI lleno: " & cFicheroLleno
            END IF

            ' --- vaciar el fichero lleno
            WriteLog "Vaciar AI lleno : " & cFicheroLleno  
            nReturn = oShell.Run( F2("{0}\dlc\bin\_rfutil {0}\bases\{1} -C aimage extent empty", DirectorioTlmp, BaseDatos), 0, TRUE )   
            IF nReturn <> 0 THEN 
                WriteLog "--- ERROR al vaciar AI lleno: " & cFicheroLleno
            END IF
        ELSE 
            ' saber si esta activado el AI en la base de datos
            ' WriteLog "Fichero no encontrado: " & cFicheroLleno
        END IF
    ELSE ' 2
        CopiarVaciarAILleno = 1
    END IF
    SET oShell = Nothing

END FUNCTION

'--------------------------------------------------------------------------
' Comprobar si las bases de datos están paradas
'--------------------------------------------------------------------------
FUNCTION ComprobarBasesParadas()
    Dim I, oShell, nReturn

    SET oShell = WScript.CREATEOBJECT("Wscript.Shell")
    ComprobarBasesParadas = 0 
    FOR I=0 TO UBOUND(BasesDeDatos)
        nReturn = oShell.Run( F2("{0}\dlc\bin\_proutil {0}\bases\{1} -C busy", DirectorioTlmp,  BasesDeDatos(I) ), 0, TRUE ) 
        IF nReturn <> 0 THEN
            WriteLog F1( "ERROR: La base de datos {0} no esta parada.", BasesDeDatos(I) )   
            ComprobarBasesParadas = 1
            EXIT FOR
        END IF    
    NEXT
    SET oShell = Nothing
END FUNCTION

'--------------------------------------------------------------------------
' Copiar las carpeta al destino
'--------------------------------------------------------------------------
FUNCTION CopiaCarpetas(DestinoDirectorio, DirectorioNCopia)
    DIM Carpeta, I, CopiarOrigen, lCarpeta
    DIM nFicherosAcopiar, nCarpetasAcopiar, CarpetaOrigen, lFichero
    DIM FicherosOrigen, A, cRoboCopy
    DIM oShell, nReturn
    
    SET oShell = WScript.CREATEOBJECT("Wscript.Shell")

    CopiaCarpetas = FALSE
    
    ON ERROR RESUME NEXT

    ' --- COPIA DE LAS CARPETAS INDICADAS EN EL ARRAY DirectoriosOrigen
    FOR I=0 TO UBOUND(DirectoriosOrigen)
    	
	    WriteLog F2("Copiar: {0} al directorio destino: {1}", DirectoriosOrigen(i), DirectorioNCopia)
		lCarpeta = TRUE
		lFichero = TRUE
		nCarpetasAcopiar = 0
		nFicherosAcopiar = 0
	    CopiarOrigen = DirectoriosOrigen(i)
		CarpetaOrigen = DirectoriosOrigen(i)
        FicherosOrigen= "*.*"
        IF I=0 THEN 
            A = ":" 
        ELSE 
            A = "+:"
        END IF

		' se puden dar tres casos. 
		' --- 1) el origen es un recurso de red. Ej. \\SERVIDOR01\TLMP
		IF LEFT( DirectoriosOrigen(i), 2) ="\\" AND  INSTRREV( LEFT( DirectoriosOrigen(i), INSTRREV( DirectoriosOrigen(i), "\")-1), "\" )= 2  THEN
		        ' determinar la carpeta destino segun el origen
				Carpeta=MID(DirectoriosOrigen(i), INSTRREV(DirectoriosOrigen(i),"\")+1)
		        ' en este caso aNadir todo, si no da error
		        CopiarOrigen = DirectoriosOrigen(i) '& "\*" 		        

		' --- 2) el origen es una carpeta. Ej. C:\TLMPLUS\BASES 
		ELSEIF fs.FolderExists( DirectoriosOrigen(i) ) THEN
			' carpeta destino = carpeta origen 
			Carpeta=MID(DirectoriosOrigen(i), INSTRREV(DirectoriosOrigen(i),"\")+1)
			
			lFichero = FALSE
			
		' --- 3) el origen es un conjunto de fichero. Ej. C:\TLMPLUS\BASES\TLMPLUS*.*
		ELSE
			Carpeta=MID( _
					   left( DirectoriosOrigen(i), _
                       INSTRREV( DirectoriosOrigen(i), "\")-1), _
						   INSTRREV(left( DirectoriosOrigen(i), _
                           INSTRREV( DirectoriosOrigen(i), "\")-1),"\")+1)
			CarpetaOrigen = LEFT( DirectoriosOrigen(i), INSTRREV( DirectoriosOrigen(i), "\")-1)
            FicherosOrigen = MID(DirectoriosOrigen(i), INSTRREV( DirectoriosOrigen(i), "\")+1)
            lCarpeta = FALSE
		
		END IF
	  
		' si hay carpetas, copiar
		' Comprobar el acceso a la carpeta origen
		nCarpetasAcopiar = fs.GetFolder( CarpetaOrigen).SubFolders.Count 
		nFicherosAcopiar = fs.GetFolder( CarpetaOrigen).Files.Count
		
		IF ERR.number = -2147024863 OR ERR.number = -2147024832 OR ERR.number = -2147024873 THEN
			ERR.clear
		END IF
        ' ejecutar ROBOCOPY

        ' -- comando ROBOCOPY --
        cRoboCopy = "ROBOCOPY " & CarpetaOrigen & " " & _ 
                        DirectorioNCopia & "\" & Carpeta  & " " & _ 
                        FicherosOrigen & _ 
                        OPCIONESR & _ 
                        " /UNILOG" & A & DestinoDirectorio & "\" & FicheroLogRC

        WriteLog F1("{0}", cRoboCopy  )
  
        nReturn = oShell.Run( cRoboCopy , 0, true )

        WriteLog F2("código del resultado: {0} = {1}", nReturn, ResultadoRC(nReturn))
        
    NEXT
    ' realizar copia en carpeta N para cloud backup.
    ' se copia la última copia \CopiaTlmplus\x sobre la \CopiaTlmplus\N
    IF lCopiaEnCarpetaN THEN 
        WriteLog "Copia adicional para sincronizar en la nube"
        cRoboCopy = "ROBOCOPY " & DirectorioNCopia & " " & _ 
            DirectorioCarpetaN & " " & _ 
            FicherosOrigen & _ 
            OPCIONESR & _ 
            " /UNILOG" & A & DestinoDirectorio & "\" & FicheroLogRC
        WriteLog F1( "{0}", cRoboCopy  ) 
        nReturn = oShell.Run( cRoboCopy , 0, true ) 
        WriteLog F2("código del resultado: {0} = {1}", nReturn, ResultadoRC(nReturn))
    END IF

    ' si hay errores salir de la funcion 
    IF nErrores = 0 THEN 
        CopiaCarpetas = TRUE
    END IF

    SET oShell = Nothing

END FUNCTION

'--------------------------------------------------------------------------
'  actualizar el fichero de control
'--------------------------------------------------------------------------
SUB ActualizarFicheroControl()
 	' Si todo correcto actualizar el fichero de control
 	IF cFicheroControlOld <> "" THEN
 		fs.DeleteFile cFicheroControlOld,TRUE    
		IF ERR.number <>0 THEN
			CALL NERROR("Error al borrar el fichero de control {0}", cFicheroControlOld)
			nErrores = nErrores  + 1
		END IF
	END IF
 	fs.CreateTextFile(cFicheroControlNew) ' se crea el nuevo
	IF ERR.number <>0 THEN
		CALL NERROR("Error al crear el fichero de control {0}", cFicheroControlNew)
		nErrores = nErrores  + 1
    END IF
END SUB

'--------------------------------------------------------------------------
'  crear un directorio o carpeta
'--------------------------------------------------------------------------
FUNCTION CrearDirectorio(d) 
	' d : directorio 
	CrearDirectorio = TRUE
	fs.CreateFolder(d)
	
	IF ERR.number <>0 THEN
		CALL NERROR(" Error al crear el directorio destino {0}", d) 
		CrearDirectorio = FALSE
	ELSE
		WriteLog F1(" Directorio creado: {0}", d)
	END IF

END FUNCTION

'--------------------------------------------------------------------------
' Mensaje de error estandard
'--------------------------------------------------------------------------
FUNCTION NERROR( cMensaje, cC0)
	WriteLog F3(cMensaje & vbCRLF & " Error:({1}) {2} ", cC0, ERR.Number, ERR.description)
	ERR.clear	
END FUNCTION

'--------------------------------------------------------------------------
'  Borrar ficheros temporales
'--------------------------------------------------------------------------
SUB BorrarTMP()
   ' 
   DIM oShell

   SET oShell = WScript.CREATEOBJECT("Wscript.Shell")
    
   WriteLog "Borrar ficheros temporales: " & "\TMP\*.*;" &  "\PROTRACE.*;" & "\DBI*;" &  "\SRT*" 
      
   on error resume Next
   fs.DeleteFile DirectorioTlmp & "\TMP\*.*", TRUE
   fs.DeleteFile DirectorioTlmp & "\PROTRACE.*", TRUE
   fs.DeleteFile DirectorioTlmp & "\DBI*", TRUE
   fs.DeleteFile DirectorioTlmp & "\LBI*", TRUE
   fs.DeleteFile DirectorioTlmp & "\SRT*", TRUE
   on error goto 0

   SET oShell = Nothing
	 
   WScript.Sleep(TIEMPOESPERA)
   
END SUB

'--------------------------------------------------------------------------
'  parar/inciar bases tlmplus
'--------------------------------------------------------------------------
SUB AccionBases(cAccion)
   ' cAccion : "stop" - Parar, "start" - Iniciar
   DIM I, oShell, nReturn
   ' Respetar el orden indicado en https://knowledgebase.progress.com/articles/Article/P19023

   SET oShell = WScript.CREATEOBJECT("Wscript.Shell")

   WriteLog UCASE(cAccion) & " bases de datos"
   
    If cAccion="stop" THEN
        ' 1 parar apps server
        WriteLog "... parar AppServer"
        FOR I=0 TO UBOUND(AppServerAgentes)
            ' intentar la parada normal según https://knowledgebase.progress.com/articles/Article/P110644
            nReturn = oShell.Run( F2( "{0}\dlc\bin\asbman -name {1} -stop", DirectorioTlmp,  AppServerAgentes(I) ), 0, TRUE ) 
            IF nReturn <> 0 THEN
                ' intentar la parada de emergencia (kill)
                nReturn = oShell.Run( F2( "{0}\dlc\bin\asbman -name {1} -kill", DirectorioTlmp,  AppServerAgentes(I) ), 0, TRUE ) 
                IF nReturn <> 0 THEN
                    WriteLog F1( "ERROR: parando el agente AppServer {0}.", AppServerAgentes(I) ) 
                END IF
            END IF    
        NEXT

        WScript.Sleep(TIEMPOESPERA*3) ' el tiempo de respuesta puede se excesivo
       
        ' 2 parar base datos
        FOR I=0 TO UBOUND(BasesManenimiento)
            nReturn = oShell.Run( F2( "{0}\dlc\bin\dbman -stop -db {1}", DirectorioTlmp,  BasesManenimiento(I) ), 0, TRUE ) 
            IF nReturn <> 0 THEN
                WriteLog F1( "ERROR: parando la base de datos {0}.", BasesManenimiento(I) )   
            END IF    
        NEXT

        ' eliminar procesos java.exe
        'WriteLog "... eliminar procesos java.exe"
        'oShell.run "taskkill.exe /F /IM java.exe", 0, True
    End If

    If cAccion="start" THEN
        WriteLog "... arrancar AppServer"
        ' 1 arrancar bases de datos
        FOR I=0 TO UBOUND(BasesManenimiento)
            nReturn = oShell.Run( F2( "{0}\dlc\bin\dbman -start -db {1}", DirectorioTlmp,  BasesManenimiento(I) ), 0, TRUE ) 
            IF nReturn <> 0 THEN
                WriteLog F1( "ERROR: arrancando la base de datos {0}.", BasesManenimiento(I) )   
            END IF    
        NEXT
    
        WScript.Sleep(TIEMPOESPERA*3)

        ' 2 arrancar appserver
        FOR I=0 TO UBOUND(AppServerAgentes)
            nReturn = oShell.Run( F2( "{0}\dlc\bin\asbman -name {1} -start", DirectorioTlmp,  AppServerAgentes(I) ), 0, TRUE ) 
            IF nReturn <> 0 THEN
                WriteLog F1( "ERROR: parando el agente AppServer {0}.", AppServerAgentes(I) )   
            END IF    
        NEXT

    End If

   SET oShell = Nothing
   
   WScript.Sleep(TIEMPOESPERA)
   
END SUB

'--------------------------------------------------------------------------
' Recálculo en Tlmplus
'--------------------------------------------------------------------------
SUB RecalculoTlmplus()
	DIM oShell, nReturn, c, i, aProceso, aTexto
	SET oShell = WScript.CREATEOBJECT("Wscript.Shell")

    aProceso = Array ( "p10tar90000-640.r",    "p10tar90000-650.r",               "p10tar90000-540.r", "ptarfictmp010.r",   "p10tar01903-020.r" )
	aTexto   = Array ( "Pendiente de recibir", "Pendiente de servir y reservado", "Stock en tránsito", "Tablas temporales", "Reconstrucción de costes" ) 

	WriteLog "Inicio recálculos en la base de datos: " & DATE & " " & TIME
    For i=0 To UBound(aProceso)

        nReturn = oShell.Run( F2( "{0}\dlc\bin\prowin32.exe -p {0}\prog\{1} -ininame {0}\tlmp.ini -Wa -wpp", DirectorioTlmp,  aProceso(i) ), 0, true )
        WriteLog F3(" {0} {1} código del resultado: {2}.", aTexto(i) & SPACE(35-LEN( aTexto(i))), TIME, nReturn)
    
    Next 
	WriteLog "Fin recálculos en la base de datos: " & DATE & " " & TIME
End SUB

'--------------------------------------------------------------------------
'  Truncar bi bases tlmplus
'--------------------------------------------------------------------------
SUB TruncarBi()
    DIM I, oShell, nReturn

    SET oShell = WScript.CREATEOBJECT("Wscript.Shell")

    WriteLog "Truncar los ficheros BI."
    FOR I=0 TO UBOUND(BasesManenimiento)
        nReturn = oShell.Run( F2("{0}\dlc\bin\_proutil {0}\bases\{1} -C truncate bi", DirectorioTlmp,  BasesManenimiento(I) ), 0, TRUE ) 
        IF nReturn <> 0 THEN
            WriteLog F1( "ERROR: Truncando el fichero BI de {0}.", BasesManenimiento(I) )   
        END IF    
    NEXT

    WScript.Sleep(TIEMPOESPERA)
    SET oShell = Nothing

END SUB

'--------------------------------------------------------------------------
'  Truncar lg bases tlmplus
'--------------------------------------------------------------------------
SUB TruncarLg()
   DIM I, oShell, nReturn

   SET oShell = WScript.CREATEOBJECT("Wscript.Shell")

    WriteLog "Truncar los ficheros LG." 
    FOR I=0 TO UBOUND(BasesManenimiento)
        nReturn = oShell.Run( F2("{0}\dlc\bin\_dbutil prolog {0}\bases\{1}", DirectorioTlmp,  BasesManenimiento(I) ), 0, TRUE ) 
        'IF nReturn <> 0 THEN ' !! SE DESACTIVA YA QUE SIEMPRE DEVUELVE ERRORLEVEL=2
        '    WriteLog F1( "ERROR: Truncando el fichero LG de {0}.", BasesManenimiento(I) )   
        'END IF    
    NEXT
    
    WScript.Sleep(TIEMPOESPERA)
    SET oShell = Nothing

END SUB

'--------------------------------------------------------------------------
'  códigos de retorno de ROBOCOPY
'--------------------------------------------------------------------------
FUNCTION ResultadoRC(nReturn)
    ' if errorlevel 16 echo ***FATAL ERROR*** & goto end
    ' if errorlevel 15 echo FAIL MISM XTRA COPY & goto end
    ' if errorlevel 14 echo FAIL MISM XTRA & goto end
    ' if errorlevel 13 echo FAIL MISM COPY & goto end
    ' if errorlevel 12 echo FAIL MISM & goto end
    ' if errorlevel 11 echo FAIL XTRA COPY & goto end
    ' if errorlevel 10 echo FAIL XTRA & goto end
    ' if errorlevel 9 echo FAIL COPY & goto end
    ' if errorlevel 8 echo FAIL & goto end
    ' if errorlevel 7 echo MISM XTRA COPY & goto end
    ' if errorlevel 6 echo MISM XTRA & goto end
    ' if errorlevel 5 echo MISM COPY & goto end
    ' if errorlevel 4 echo MISM & goto end
    ' if errorlevel 3 echo XTRA COPY & goto end
    ' if errorlevel 2 echo XTRA & goto end
    ' if errorlevel 1 echo COPY & goto end
    ' if errorlevel 0 echo ?no change? & goto end
    
    IF nReturn = 0 THEN
        ResultadoRC = "sin cambios"
    ELSEIF nReturn >= 1 AND nReturn <= 7 THEN
        ResultadoRC = "ficheros copiados"
    ELSEIF nReturn >= 8 THEN
        ResultadoRC = "error en la copia !!!"
        nErrores = nErrores  + 1 
    END IF
   
END FUNCTION

'--------------------------------------------------------------------------
' Salir de la script
'--------------------------------------------------------------------------
SUB Salir(nSalida, cFicheroAdjunto)
    DIM c
    IF nSalida = 0 THEN
	    c = " sin errores "
    ELSE
	    c = " CON ERRORES !!! "
    END IF
    WriteLog "--- Fin   ---" & c & DATE & " " & TIME  
    
    IF lEnviarEmail = TRUE THEN
	    
        cAdjunto = cFicheroAdjunto
        cMensaje = cMensaje
        ' ini - convertir en html
        cMensaje = Replace("<meta http-equiv='Content-Type' content='text/html; charset=utf-8'><pre style='white-space:pre-wrap'>", "'", CHR(34)) & cMensaje & "</pre>"
        ' fin - convertir en hmtl
        cAsunto  = cAsunto & c
        
        ret = FALSE
        ret = Enviar_Mail_CDO(cServidor, _
            cParaEmail, _
            cDeEmail, _
            cAsunto, _
            cMensaje, _
            lHtml, _
            cAdjunto, _
            cPuerto, _
            cUsuario, _
            cPassword, _
            lAutentificacion, _
            lSSL)
    END If
        
    wscript.Quit(nSalida)

END SUB

'--------------------------------------------------------------------------
' Cambiar los atributos de ficheros antes de borrarlos
'--------------------------------------------------------------------------
FUNCTION CambiarAtributos(Directorio)
	Dim fl, fc, fd
	
	Set fc = fs.GetFolder(Directorio).Files
	Set fd = fs.GetFolder(Directorio).SubFolders
	' recorre ficheros
	For Each fl in fc  
		If fl.attributes <> 0 Then
			fl.attributes = 0 ' fl.attributes - 32
		End If	  
	Next
	' recorre las subcarpetas
    For Each fl in fd	
		CambiarAtributos(fl.path )
	Next
END FUNCTION

'--------------------------------------------------------------------------
' Handle wmi Job object
'--------------------------------------------------------------------------
FUNCTION WMIJobCompleted(outParam)   
    DIM WMIJob, jobState, nPorcentaje

    SET WMIJob = objWMIService.GET(outParam.Job)

    WMIJobCompleted = TRUE

    jobState = WMIJob.JobState

    WHILE jobState = JobRunning or jobState = JobStarting

	    IF WMIJob.PercentComplete MOD 25 = 0 and nPorcentaje <> WMIJob.PercentComplete THEN
		    nPorcentaje = WMIJob.PercentComplete
            WriteLog F1("En proceso... {0}% completado.", WMIJob.PercentComplete)
        END IF

        WScript.Sleep(2000)
        SET WMIJob = objWMIService.GET(outParam.Job)
        jobState = WMIJob.JobState
    WEND

    IF (jobState <> JobCompleted) THEN
        WriteLog F1("Código de error:{0}", WMIJob.ErrorCode)
        WriteLog F1("Descripción del error:{0}", WMIJob.ErrorDescription)
        WMIJobCompleted = FALSE
    END IF
    SET WMIJob = Nothing
END FUNCTION

'--------------------------------------------------------------------------
' Create the console log files.
'--------------------------------------------------------------------------
FUNCTION WriteLog(line)
    DIM fileStream
    
    SET fileStream = fs.OpenTextFile(FicheroLog, 8, True, -1) 
    '     8=ForAppending, True=Crear si no existe, -1 = Unicode,
    WScript.Echo line
    fileStream.WriteLine line
    fileStream.Close
    
    ' -- guardar para enviar por correo
    cMensaje = cMensaje &  line & vbCrLf
    SET fileStream = Nothing
END FUNCTION
'--------------------------------------------------------------------------
' Crear log
'--------------------------------------------------------------------------
SUB CrearLog()
 	DIM fileStream
    'Creamos el fichero de Log
	Set fileStream = fs.CreateTextFile(FicheroLog, True, True)
	fileStream.Close
	Set fileStream = Nothing
END SUB

'--------------------------------------------------------------------------
' The string formatting functions to avoid string concatenation.
'--------------------------------------------------------------------------
FUNCTION F3(myString, arg0, arg1, arg2) 
	F3 = F2(myString, arg0, arg1)
    F3 = REPLACE(F3, "{2}", arg2)
END FUNCTION

FUNCTION F2(myString, arg0, arg1)
    F2 = F1(myString, arg0)
    F2 = REPLACE(F2, "{1}", arg1)
END FUNCTION

FUNCTION F1(myString, arg0)
    F1 = REPLACE(myString, "{0}", arg0)
END FUNCTION

'--------------------------------------------------------------------------
' calcular número de copia
'--------------------------------------------------------------------------
FUNCTION NumeroCopia(cDirectorio)
	DIM cFichero, i
	cFicheroControlOld = ""
	cFichero = cDirectorio & "\" & FicheroControl
	NumeroCopia = 1
	FOR i=1 TO NumeroDeCopias
		IF fs.FileExists(cFichero & i ) THEN
			cFicheroControlOld = cFichero & i 
			NumeroCopia = i + 1		
			IF NumeroCopia = NumeroDeCopias + 1  THEN
				NumeroCopia = 1
			END IF
			EXIT FOR
		END IF
	NEXT 
	cFicheroControlNew = cFichero & NumeroCopia
END FUNCTION

'--------------------------------------------------------------------------
' calcular número de anterior
'--------------------------------------------------------------------------
FUNCTION NumeroCopiaAnterior(nCopia)
    NumeroCopiaAnterior = nCopia - 1
    IF NumeroCopiaAnterior = 0 THEN NumeroCopiaAnterior = NumeroDeCopias
END FUNCTION

'--------------------------------------------------------------------------
'  Enviar correo electronico
'--------------------------------------------------------------------------
FUNCTION Enviar_Mail_CDO(Servidor_SMTP , _
			Para , _
			De , _
			Asunto , _
			Mensaje , _
			Html , _
			Path_Adjunto , _
			Puerto , _
			Usuario , _
			Password , _
			Usar_Autentificacion, _
			Usar_SSL) 
    ' Variable de objeto Cdo.Message
    DIM Obj_Email 
          
    ' Crea un Nuevo objeto CDO.Message
    SET Obj_Email = CreateObject ("cdo.Message")
    
    Obj_Email.BodyPart.Charset = "utf-8" 

    ' Indica el servidor Smtp para poder enviar el Mail ( puede ser el nombre _
    '  del servidor o su direcci?n IP )
    Obj_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Servidor_SMTP
    
    Obj_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    
    ' Puerto. Por defecto se usa el puerto 25, en el caso de Gmail se usan los puertos _
    '  465 o  el puerto 587 ( este ?ltimo me dio error )
    
    Obj_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CLNG(Puerto)
    
    ' Indica el tipo de autentificaci?n con el servidor de correo _
    ' El valor 0 no requiere autentificarse, el valor 1 es con autentificaci?n
    Obj_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = ABS(Usar_Autentificacion)
    
        ' Tiempo m?ximo de espera en segundos para la conexi?n
    Obj_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30

    ' Configura las opciones para el login en el SMTP
    IF Usar_Autentificacion THEN

        ' Id de usuario del servidor Smtp ( en el caso de gmail, debe ser la direcci?n de correro _
        ' mas el @gmail.com )
        Obj_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = Usuario

        ' Password de la cuenta
        Obj_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Password

        ' Indica si se usa SSL para el env?o. En el caso de Gmail requiere que est? en True
        Obj_Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = Usar_SSL
    
    END IF
    
    ' DirecciOn del Destinatario
    Obj_Email.TO = Para
    
    ' DirecciOn del remitente
    Obj_Email.From = De
    
    ' Asunto del mensaje
    Obj_Email.Subject = Asunto
    
    ' Cuerpo del mensaje
    If Html Then 
        Obj_Email.HTMLBody = Mensaje
        Obj_Email.HTMLBodyPart.Charset = "utf-8"
    Else
        Obj_Email.TextBody = Mensaje
    End If

    'Ruta del archivo adjunto   
    IF Path_Adjunto <> vbNullString THEN
        Obj_Email.AddAttachment (Path_Adjunto)
    END IF
 
    ' Actualiza los datos antes de enviar
    Obj_Email.Configuration.Fields.Update
      
    ' EnvIa el email
    Obj_Email.Send

    IF ERR.Number = 0 THEN
       Enviar_Mail_CDO = TRUE
       WriteLog F2("Notificado por correo a {0} mediante el servidor {1}", Para, Servidor_SMTP)
    ELSE
       CALL NERROR("Error en la notificaci?n por correo a {0}", Para)
       Enviar_Mail_CDO = FALSE	
    END IF
    
    ' Descarga la referencia
    IF Not Obj_Email Is Nothing THEN
        SET Obj_Email = Nothing
    END IF
    
    ON ERROR GOTO 0

END FUNCTION

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  APAGAR EL EQUIPO
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function ApagarWindows(tnAction)
    Dim loWmiService, laWmiInstance, loInstance, loInstanceWin

    Set loWmiService  = GetObject("winmgmts:{(Shutdown)}")
    Set laWmiInstance = loWmiService.InstancesOf("win32_operatingsystem")
    For Each loInstance In laWmiInstance
        loInstance.win32shutdown(tnAction)
    Next
End Function
