
'-----------------------------------------------------------------------------------------------------------------------
'											FUNCIONES DE ESPERA DIN�MICA T...
'-----------------------------------------------------------------------------------------------------------------------
function BucleDeEsperaValorVision(iglobal,Variable,Espera,Reintentos) 
        Encontrado = false 
        Intentos = 0 
        Coordenadas = "0" 
        
        do While (Coordenadas = "0" OR Coordenadas = "0|0") AND Intentos < Reintentos 
                iGlobal.accion "acAIA", "EXEC|AIA.Vision|Enumerar" 
                Coordenadas = iGlobal.accion("acAIA","EXEC|AIA.Vision|BuscarPatron||"& Variable &"|90|CENTRO") 
                wscript.sleep Espera 
                Intentos = Intentos + 1
				fWriteLog RutaLogs, "Script: INFO-[" & now() & "]-intentos " & Intentos
				'test Coordenadas
        loop 

        If Coordenadas <> "0" AND Coordenadas <> "0|0" Then 
                BucleDeEsperaValorVision = true 
        else 
                BucleDeEsperaValorVision = false 
        End If 
end function  


Function BucleEsperarVentanaACA(Venentana, iteraciones, tiempo)	

	BucleEsperarVentanaACA = false
	flagexiste = false
		i = 0	
	While not flagexiste AND i < iteraciones		
		
		If iGlobal.accion("acAIA","EXEC|AIA.ACA|Existe|$"& Venentana &"|^"& Venentana) Then 
			flagexiste = true
		end if
		wscript.sleep tiempo
		i = i + 1
	Wend	
	If i = iteraciones Then
		BucleEsperarVentanaACA = false
	Else
		BucleEsperarVentanaACA = true
	End If	
End Function



Function BucleEsperaValorWeb(variable, iteraciones, tiempo)
	check = ""
	i = 0
	While check = "" AND i < iteraciones 
		check = iglobal.accion("acAIA","get|"& variable)
		wscript.sleep tiempo
		i = i + 1
	Wend
	
	If i = iteraciones Then
		BucleEsperaValorWeb = false
	Else
		BucleEsperaValorWeb = true
	End If
End Function

Function BucleEsperaValorConcreto(variable, valor, iteraciones, tiempo)
	check = ""
	i = 0
	While check <> valor AND i < iteraciones 
		check = iglobal.parse(Variable)
		wscript.sleep tiempo
		i = i + 1
	Wend
	
	If i = iteraciones Then
		BucleEsperaValorConcreto = false
	Else
		BucleEsperaValorConcreto = true
	End If

End Function
'-----------------------------------------------------------------------------------------------------------------------
'											CLEAR IE CACHE
'-----------------------------------------------------------------------------------------------------------------------
Function ClearIECache()
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351", 1, True
    Set objShell = Nothing
End Function






'-----------------------------------------------------------------------------------------------------------------------
'											PASAR VALORES ENTRE SCRIPTS
'-----------------------------------------------------------------------------------------------------------------------
function enviarValorScript (valor)
	Dim objShell, objEnvVar 
	Set objShell = CreateObject("WScript.Shell") 
	Set objEnvVar = objShell.Environment("Volatile")

	objEnvVar(variable) = valor
	Set objEnvVar = Nothing 
	Set objShell = Nothing 

end function

function recogerValorScript
	
	Dim objEnvVar 
	Dim objShell
	Set objShell = CreateObject("WScript.Shell") 
	Set objEnvVar = objShell.Environment("Volatile") 
	recogerValorScript=objEnvVar(variable) 
	Set objEnvVar = Nothing 
		
end function


REM |--------------------------------|
REM  Rellena con '0'
REM |--------------------------------|
Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function

REM |--------------------------------|
REM  Get fecha y hora actual
REM |--------------------------------|
Function getFecha(fecha)
	
	if not IsNull(fecha) or fecha="" then
		getFecha = Year(Date) & "-" & LPad(Month(Date),"0",2) & "-" & LPad(Day(Date), "0", 2) & " " & LPad(Hour(Time), "0", 2) & ":" & LPad(Minute(Time), "0", 2) & ":" & LPad(Second(Time), "0", 2)
	else
		aux=Split(fecha," ")
		
		fecha_split=aux(0)
		hora_split=aux(1)
		
		fecha_aux=Split(fecha_split,"/")
		hora_aux=Split(hora_split,":")
		
		dia = fecha_aux(0)
		mes = fecha_aux(1)
		ano = fecha_aux(2)
		
		hora = hora_aux(0)
		min = hora_aux(1)
		seg = hora_aux(2)
		
		getFecha = ano & "-" & LPad(mes,"0",2) & "-" & LPad(dia, "0", 2) & " " & LPad(hora, "0", 2) & ":" & LPad(min, "0", 2) & ":" & LPad(seg, "0", 2)
	end if
End Function


'-----------------------------------------------------------------------------------------------------------------------
'											FUNCION TEST, LOG Y CERRAR PROCESO
'-----------------------------------------------------------------------------------------------------------------------
function test (texto)
	iGlobal.accion "acEnviarValores", texto & "=#test"
	iGlobal.accion "acTestVariables", "#test"
	wscript.sleep 600
end function


function fWriteLog(RutaLogs,texto)
	Set objFsoLog = CreateObject("Scripting.FileSystemObject")  
	Dim objTextStream
	fecha = iglobal.parse("@FechaHoyDDMMYYYY")
	nFichero = RutaLogs & iglobal.parse("@puesto") & fecha & "Log.txt" 
	If (objFsoLog.FileExists(nFichero)) Then
		Set logOutput = objFsoLog.OpenTextFile(nFichero, 8, True)
		logOutput.WriteLine(texto)	'Escritura en fichero
		logOutput.Close	'Cerrar fichero
		'Validar numeros de lineas de fichero log
		Const ForReading = 1
		Set objTextFile = objFsoLog.OpenTextFile(nFichero, ForReading)
		objTextFile.ReadAll
		numero_lineas= objTextFile.Line	'Obtiene cantidad de lineas de fichero log
		Set objTextFile = Nothing
		IF numero_lineas >= 9999 Then
			nFichero2 = RutaLogs & iglobal.parse("@puesto") & fecha & "_" & LPad(Hour(Time), "0", 2) & "-" & LPad(Minute(Time), "0", 2) & "-" & LPad(Second(Time), "0", 2) & "_" & "Log.txt"
			wscript.sleep 2000
			objFsoLog.CopyFile nFichero, nFichero2	'Reemplazar nombre de fichero log
			wscript.sleep 2000
			Set MiArchivo = objFsoLog.GetFile(nFichero)
			MiArchivo.Delete
			wscript.sleep 2000
			Set logOutput = objFsoLog.OpenTextFile(nFichero, 8, True)
			logOutput.WriteLine(texto)	'Escritura en fichero			
			logOutput.Close	'Cerrar fichero
		End If
	Else
		Set logOutput = objFsoLog.OpenTextFile(nFichero, 8, True)
		logOutput.WriteLine(texto)	'Escritura en fichero
		logOutput.Close	'Cerrar fichero
	End If
	
	Set logOutput = Nothing 
	Set objFsoLog = Nothing
end function

function cierroProceso(proceso)
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colProcesses = objWMIService.ExecQuery _
		("Select * from Win32_Process Where Name like '%"& proceso &"%'")
	For Each objProcess in colProcesses
		If IsObject(objProcess) Then                    
		   on Error Resume next
			   objProcess.Terminate() 
		   On Error Goto 0
		end if
	Next
end function
function FinalizarProcesos(iGlobal,rutaLogs)	
	'Cerrar procesos
	test "CIERRA TODOS LOS PROCESOS"
	'CerrarConexion Conexion
	CerrarConexion ConexionMakro
	'fWriteLog rutaLogs,comentario
	ScriptFilename99 = "99_funciones.vbs"
	fWriteLog rutaLogs, "Script: " & ScriptFilename99 & " Funcion EsperarEtiquetaBrowser INFO-[" & now() & "]-CIERRA TODOS LOS PROCESOS " 
	fWriteLog rutaLogs,"----------------------------------FIN PROCESO CERRAR CREAR USUARIO--------------------------------"
	iGlobal.accion "acFinProcedimiento",""
	CierroProceso "EXCEL.EXE" 
	CierroProceso "iexplore.exe"
	CierroProceso "chromedriver.exe" 
	CierroProceso "chrome.exe" 
	CierroProceso "iGlobal.exe" 
	CierroProceso "wscript.exe" 
	wscript.quit(0)
end function
'-----------------------------------------------------------------------------------------------------------------------
'											BUSCAR ELEMENTO POR VISION
'-----------------------------------------------------------------------------------------------------------------------
function buscarElemento_vision (ventana, rutaImagen, repeticiones)
	'ventana sin $
	'msgbox "buscarElemento_vision"
	oWS.SendKeys "^0"
	timevision1=timer()
	REM iGlobal.accion "acAIA", "STOP|AIA.Vision"
	REM iGlobal.accion "acAIA", "INIT|AIA.Vision"
	
	wscript.sleep 500
	iglobal.accion "acAIA", "EXEC|AIA.ACA|FOCO|$"& ventana &"|^"& ventana
	wscript.sleep 500
	iglobal.accion "acInteraccion", "SCRIPT:$"& ventana &"=MAXIMIZA"
	oWS.SendKeys "^0"
	wscript.sleep 500
	i = 0
	encontrado = false
	do while not encontrado and i<repeticiones
		iGlobal.accion "acAIA", "EXEC|AIA.Vision|Enumerar"
		iGlobal.accion "acAIA", "EXEC|AIA.Vision|MoverCursor||100|0"
		wscript.sleep 200
		
		Coordenadas = iGlobal.accion("acAIA", "EXEC|AIA.Vision|BuscarPatron||" & rutaImagen &"|90|CENTRO")
		buscarElemento_vision=Coordenadas
		wscript.sleep 100
		If Coordenadas <> "" AND Coordenadas <> "0|0" Then
		   encontrado = true
		   'arrayCoordenadas=Split(Coordenadas, "|")
			'test "Vision: Coordenadas encontradas:  " & Coordenadas
		   iGlobal.accion "acAIA", "EXEC|AIA.Vision|MoverCursor||" & Coordenadas
		   wscript.sleep 1000
		   iGlobal.accion "acAIA", "EXEC|AIA.Vision|Click||IZQUIERDO"
		   wscript.sleep 1000 
		   iGlobal.accion "acAIA", "EXEC|AIA.Vision|MoverCursor||100|0" 
		else
			iGlobal.accion "acAIA", "EXEC|AIA.Vision|MoverCursor||100|0" 
			'test "Vision no encontrado!!!!"
			buscarElemento_vision=false
			'wscript.quit(11)
		end if
		i = i + 1
	

	loop
		timevision2=timer()
		'test "Coordenadas: " & Coordenadas
		'test "Tiempo Vision:" & timevision2-timevision1
end function


function buscarElemento_vision2 (ventana, rutaImagen2, repeticiones, modoBuscarPatron, accion)
	
	Dim RutaLogs
	if accion = "" then
		accion = "IZQUIERDO"
	end if
	
	RutaLogs = iglobal.parse("#RutaLogs")
	iglobal.accion "acAIA", "EXEC|AIA.ACA|FOCO|$"& ventana &"|^"& ventana
	wscript.sleep 300
	'iglobal.accion "acInteraccion", "SCRIPT:$"& ventana &"=MAXIMIZA"
	'oWS.SendKeys "^0"
	wscript.sleep 200

	i = 0
	encontrado = false
	do while not encontrado and i<repeticiones
		iGlobal.accion "acAIA", "EXEC|AIA.Vision|Enumerar"
		iGlobal.accion "acAIA", "EXEC|AIA.Vision|MoverCursor||100|0"
		wscript.sleep 200		
		Coordenadas = iGlobal.accion("acAIA", "EXEC|AIA.Vision|BuscarPatron||" & rutaImagen2 &"|90|" & modoBuscarPatron)	
		wscript.sleep 100
		If Coordenadas <> "" AND Coordenadas <> "0|0" Then
		   encontrado = true
		   iGlobal.accion "acAIA", "EXEC|AIA.Vision|MoverCursor||" & Coordenadas
		   wscript.sleep 300
		   iGlobal.accion "acAIA", "EXEC|AIA.Vision|Click||" & accion
		end if
		i = i + 1
		wscript.sleep 1000
	loop

	if not encontrado then
		fWriteLog RutaLogs,"ERROR-["&now()&"] Funcion buscarElemento_vision2 - El elemento no ha sido encontrado por vision: " & rutaImagen2
		REM fWriteLog RutaLogs,"ERROR-["&now()&"] Funcion buscarElemento_vision2 - El elemento no ha sido encontrado por vision: " & rutaImagen2
		iGlobal.accion "acAIA", "EXEC|AIA.Vision|MoverCursor||100|0" 
		buscarElemento_vision2 = false
	else
		fWriteLog RutaLogs,"OK-["&now()&"] Funcion buscarElemento_vision2 - El elemento se ha encontrado por vision: " & rutaImagen2
		REM fWriteLog RutaLogs,"OK-["&now()&"] Funcion buscarElemento_vision2 - El elemento se ha encontrado por vision: " & rutaImagen2
		wscript.sleep 100 
		iGlobal.accion "acAIA", "EXEC|AIA.Vision|MoverCursor||100|0" 
		buscarElemento_vision2 = true
	end if
end function

function buscarElementovision (ventana, rutaImagen2, repeticiones, modoBuscarPatron)
	Dim RutaLogs
	RutaLogs = iglobal.parse("#RutaLogs")
	fWriteLog RutaLogs,"INFO-["&now()&"] Funcion buscarElementovision - ventana: " & ventana
	fWriteLog RutaLogs,"INFO-["&now()&"] Funcion buscarElementovision - rutaImagen2: " & rutaImagen2
	iglobal.accion "acAIA", "EXEC|AIA.ACA|FOCO|$"& ventana &"|^"& ventana
	wscript.sleep 300
	'iglobal.accion "acInteraccion", "SCRIPT:$"& ventana &"=MAXIMIZA"
	'oWS.SendKeys "^0"
	'wscript.sleep 200
	
	i = 0
	encontrado = false
	do while not encontrado and i<repeticiones
		iGlobal.accion "acAIA", "EXEC|AIA.Vision|Enumerar"
		'iGlobal.accion "acAIA", "EXEC|AIA.Vision|MoverCursor||100|0"
		wscript.sleep 200
		
		Coordenadas = iGlobal.accion("acAIA", "EXEC|AIA.Vision|BuscarPatron||" & rutaImagen2 &"|90|" & modoBuscarPatron)
		fWriteLog RutaLogs,"INFO-["&now()&"] Funcion buscarElementovision - Coordenadas: " & Coordenadas
		wscript.sleep 100
		If Coordenadas <> "" AND Coordenadas <> "0|0" Then
		   encontrado = true
		end if
		i = i + 1
		wscript.sleep 1000
	loop

	if not encontrado then
		fWriteLog RutaLogs,"ERROR-["&now()&"] Funcion buscarElementovision - El elemento no ha sido encontrado por vision: " & rutaImagen2
		buscarElementovision = false
	else
		fWriteLog RutaLogs,"OK-["&now()&"] Funcion buscarElementovision - El elemento se ha encontrado por vision: " & rutaImagen2
		wscript.sleep 100 
		buscarElementovision = true
	end if
end function



'-----------------------------------------------------------------------------------------------------------------------
'									EXISTE BT BLOQUEAR
'-----------------------------------------------------------------------------------------------------------------------

function existeBTBloquear (ventana, tituloVentana, variableBoton)
'ejemplo invocaci�n: existeBTBloquear "$V_Siebel","Siebel Energy","$Siebel_BT_bloquear"
'Esta funcion comprueba si existe o no el bot�n "Bloquear"
'Si no existe el boton, significa que la solicitud ya esta bloqueada
'Si existe el boton, la solicitun no se encuentra bloqueada
'Una solicitud debe ser bloqueada para poder tramitarla 
'Si existeBTBloquear=true---- el bot�n existe asi que no esta bloqueada
'Si existeBTBloquear=false---- el bot�n NO existe asi que ya esta bloqueada
	valor = ""
	i=0
	do while valor = "" and i<2
		'iGlobal.accion "acAIA", "EXEC|AIA.Siebel|Enumerar|$V_Siebel|Siebel Energy"
		iGlobal.accion "acAIA", "EXEC|AIA.Siebel|Enumerar|"& ventana &"|" & tituloVentana
		valor = iGlobal.accion("acAIA","EXEC|AIA.Siebel|Clase|" & variableBoton)
		i = i + 1
		'test "buscando bt bloqueado" & i
	loop
	
	if valor <> "" then
		existeBTBloquear=true
	else
		existeBTBloquear=false
	end if
	
end function'-----------------------------------------------------------------------------------------------------------------------
'											OTROS
'-----------------------------------------------------------------------------------------------------------------------

' C�DIGO PARA MEDIR TIEMPOS
		REM timevision1=timer()
		REM timevision2=timer()
		REM msgbox "tiempo Vision:" & timevision2-timevision1
		
'NORMALIZA

function revisarFormato(aux)
	'Esta funcion se usa para pasar datos de BBDD a iGlobal, no pueden dejarse retornos de carro, saltos de l�nea ni espacios al principio y final
	aux=replace (aux,"*","")
	aux=replace(replace(aux, chr(10), ""), chr(13), "")
	revisarFormato= trim (aux)
		
end function

function enviaKeys(oWS,cadena)
	i=1
	do while i<Len(cadena)+1
		aux = Mid(cadena,i,1)
		oWS.SendKeys aux
		wscript.sleep 200
		i=i+1
	loop
end function

function Normaliza(palabra)
    
    Dim arrWrapper(1)
    Dim arrReplace(18)
    Dim arrReplaceWith(18)
    
    arrWrapper(0) = arrReplace
    arrWrapper(1) = arrReplace
    
    arrWrapper(0)(0) = ""&chr(225)&""
    arrWrapper(0)(1) = ""&chr(233)&""
    arrWrapper(0)(2) = ""&chr(237)&""
    arrWrapper(0)(3) = ""&chr(243)&""
    arrWrapper(0)(4) = ""&chr(250)&""
	arrWrapper(0)(5) = "�"
    arrWrapper(0)(6) = "�"
    arrWrapper(0)(7) = "�"
    arrWrapper(0)(8) = "�"
    arrWrapper(0)(9) = "�"
    arrWrapper(0)(10) = "�"
    arrWrapper(0)(11) = "�"
    arrWrapper(0)(12) = "�"
    arrWrapper(0)(13) = "�"
    arrWrapper(0)(14) = "�"
	arrWrapper(0)(15) = "�"
	arrWrapper(0)(16) = "�"
	arrWrapper(0)(17) = "*"
	arrWrapper(0)(18) = "+"
	
    arrWrapper(1)(0) = ""& ChrW(&H301)&"a"
    arrWrapper(1)(1) = ""& ChrW(&H301)&"e"
    arrWrapper(1)(2) = ""& ChrW(&H301)&"i"
    arrWrapper(1)(3) = ""& Chr(162) &"�"
    arrWrapper(1)(4) = ""& ChrW(&H301)&"u"
	arrWrapper(1)(5) = ""& ChrW(&H301)&"a"
    arrWrapper(1)(6) = ""& ChrW(&H301)&"e"
    arrWrapper(1)(7) = ""& ChrW(&H301)&"i"
    arrWrapper(1)(8) = ""& ChrW(&H301)&"o"
    arrWrapper(1)(9) = ""& ChrW(&H301)&"u"
	arrWrapper(1)(10) = ""& ChrW(&H301)&"A"
    arrWrapper(1)(11) = ""& ChrW(&H301)&"E"
    arrWrapper(1)(12) = ""& ChrW(&H301)&"I"
    arrWrapper(1)(13) = ""& ChrW(&H301)&"O"
    arrWrapper(1)(14) = ""& ChrW(&H301)&"U"
	arrWrapper(1)(15) = ""& chr(241) & "" '"�"
	arrWrapper(1)(16) = ""& chr(209) &"" '"�"
	arrWrapper(1)(17) = ""& chr(42) &"" '"*"
	arrWrapper(1)(18) = ""& chr(42) &"" '"+ por *"

    
    For N = 0 To 17     
        palabra = Replace(palabra, arrWrapper(0)(N), arrWrapper(1)(N), 1, -1, 0)		
    Next
    
    Normaliza = palabra
end function

'-----------------------------------------------------------------------------------------------------------------------
'											BUSCAR ELEMENTO POR VISION
'-----------------------------------------------------------------------------------------------------------------------
'**
'* Esta funcion busca el elemento por AIA.Vision y hace click
'*
'* @param ventana: Ventana principal donde se encuentra el elemento
'* @param rutaImagen: Ruta al patr�n del elemento
'* @param repeticiones: N�mero de repeticiones que debemos ejecutar como intentos de b�squeda
'* @param modoBuscarPatron: Modo de b�squeda del patr�n (CENTRO, ESQUINASUPERIORIZQUIERDA...)
'*
'* @return buscarElemento_vision2: Devuelve true si el elemento ha sido encontrado y false de lo contrario
'*
'**
function buscarElemento_vision2 (ventana, rutaImagen2, repeticiones, modoBuscarPatron, accion)
	
	Dim RutaLogs
	if accion = "" then
		accion = "IZQUIERDO"
	end if
	
	RutaLogs = iglobal.parse("#RutaLogs")
	iglobal.accion "acAIA", "EXEC|AIA.ACA|FOCO|$"& ventana &"|^"& ventana
	wscript.sleep 300
	'iglobal.accion "acInteraccion", "SCRIPT:$"& ventana &"=MAXIMIZA"
	'oWS.SendKeys "^0"
	wscript.sleep 200

	i = 0
	encontrado = false
	do while not encontrado and i<repeticiones
		iGlobal.accion "acAIA", "EXEC|AIA.Vision|Enumerar"
		iGlobal.accion "acAIA", "EXEC|AIA.Vision|MoverCursor||100|0"
		wscript.sleep 200		
		Coordenadas = iGlobal.accion("acAIA", "EXEC|AIA.Vision|BuscarPatron||" & rutaImagen2 &"|90|" & modoBuscarPatron)	
		wscript.sleep 100
		If Coordenadas <> "" AND Coordenadas <> "0|0" Then
		   encontrado = true
		   iGlobal.accion "acAIA", "EXEC|AIA.Vision|MoverCursor||" & Coordenadas
		   wscript.sleep 300
		   iGlobal.accion "acAIA", "EXEC|AIA.Vision|Click||" & accion
		end if
		i = i + 1
		wscript.sleep 1000
	loop

	if not encontrado then
		fWriteLog RutaLogs,"ERROR-["&now()&"] Funcion buscarElemento_vision2 - El elemento no ha sido encontrado por vision: " & rutaImagen2
		REM fWriteLog RutaLogs,"ERROR-["&now()&"] Funcion buscarElemento_vision2 - El elemento no ha sido encontrado por vision: " & rutaImagen2
		iGlobal.accion "acAIA", "EXEC|AIA.Vision|MoverCursor||100|0" 
		buscarElemento_vision2 = false
	else
		fWriteLog RutaLogs,"OK-["&now()&"] Funcion buscarElemento_vision2 - El elemento se ha encontrado por vision: " & rutaImagen2
		REM fWriteLog RutaLogs,"OK-["&now()&"] Funcion buscarElemento_vision2 - El elemento se ha encontrado por vision: " & rutaImagen2
		wscript.sleep 100 
		iGlobal.accion "acAIA", "EXEC|AIA.Vision|MoverCursor||100|0" 
		buscarElemento_vision2 = true
	end if
end function

function buscarElementovision (ventana, rutaImagen2, repeticiones, modoBuscarPatron)
	Dim RutaLogs
	RutaLogs = iglobal.parse("#RutaLogs")
	fWriteLog RutaLogs,"INFO-["&now()&"] Funcion buscarElementovision - ventana: " & ventana
	fWriteLog RutaLogs,"INFO-["&now()&"] Funcion buscarElementovision - rutaImagen2: " & rutaImagen2
	iglobal.accion "acAIA", "EXEC|AIA.ACA|FOCO|$"& ventana &"|^"& ventana
	wscript.sleep 300
	'iglobal.accion "acInteraccion", "SCRIPT:$"& ventana &"=MAXIMIZA"
	'oWS.SendKeys "^0"
	'wscript.sleep 200
	
	i = 0
	encontrado = false
	do while not encontrado and i<repeticiones
		iGlobal.accion "acAIA", "EXEC|AIA.Vision|Enumerar"
		'iGlobal.accion "acAIA", "EXEC|AIA.Vision|MoverCursor||100|0"
		wscript.sleep 200
		
		Coordenadas = iGlobal.accion("acAIA", "EXEC|AIA.Vision|BuscarPatron||" & rutaImagen2 &"|90|" & modoBuscarPatron)
		fWriteLog RutaLogs,"INFO-["&now()&"] Funcion buscarElementovision - Coordenadas: " & Coordenadas
		wscript.sleep 100
		If Coordenadas <> "" AND Coordenadas <> "0|0" Then
		   encontrado = true
		end if
		i = i + 1
		wscript.sleep 1000
	loop

	if not encontrado then
		fWriteLog RutaLogs,"ERROR-["&now()&"] Funcion buscarElementovision - El elemento no ha sido encontrado por vision: " & rutaImagen2
		buscarElementovision = false
	else
		fWriteLog RutaLogs,"OK-["&now()&"] Funcion buscarElementovision - El elemento se ha encontrado por vision: " & rutaImagen2
		wscript.sleep 100 
		buscarElementovision = true
	end if
end function

'* Esta funcion encuentra dato en vision tomando coordenadas, tipo de accion, tiempo y hace click
Function buscarElto_vision (ventana, rutaImagen, repeticiones,plusx,plusy,TipoAccion,tiempo,lugarclick)
	wscript.sleep 200
	i = 0
	encontrado = false
	do while not encontrado and i<repeticiones
		iGlobal.accion "acAIA", "EXEC|AIA.Vision|MoverCursor||100|0"
		wscript.sleep 500
		iGlobal.accion "acAIA", "EXEC|AIA.Vision|Enumerar"
		wscript.sleep 200
		Coordenadas = iGlobal.accion("acAIA", "EXEC|AIA.Vision|BuscarPatron||" & rutaImagen & "|90|" & lugarclick)
		'buscarElemento_vision=Coordenadas
		wscript.sleep 100
		If Coordenadas <> "" AND Coordenadas <> "0|0" Then
		   encontrado = true
		   arrayAux = Split(coordenadas, "|")
		   coordAux = arrayAux(0) + plusx & "|" & arrayAux(1) + plusy
		   iGlobal.accion "acAIA", "EXEC|AIA.Vision|MoverCursor||" & coordAux
		   wscript.sleep 100
		   iGlobal.accion "acAIA", "EXEC|AIA.Vision|Click||" & TipoAccion
		end if
		i = i + 1
		wscript.sleep tiempo
	loop

	if i = repeticiones then
		iGlobal.accion "acAIA", "EXEC|AIA.Vision|MoverCursor||100|0" 
		buscarElto_vision = false
	else
		buscarElto_vision = true
	end if
End Function

'* Comprobar carga de Siebel
function f_comprobarCargaSiebel(ventana,texto)
	if iGlobal.accion ("acAIA","EXEC|AIA.Siebel|esperar_descarga") = -1 then
		wscript.quit(-1)
	else
		iGlobal.accion "acAIA","EXEC|AIA.Siebel|Enumerar|$" & ventana & "|" & texto
	end if
end function

'*******************************************************************SCREENSHOTS KO*********************************************************************************************
Function ScreenshotsKO (name_script)
	On Error Resume Next
	Err.Clear
	Maquina = iglobal.parse("@puesto")
	iGlobal.accion "acAIA","EXEC|AIA.Siebel|esperar_descarga"
	wscript.sleep 3000
	timeStamp = replace (replace (replace (replace(iglobal.parse("@Timestamp")," ",""), "/", ""), ":",""), ",", "")
	RutaScreenshots1 = iglobal.parse("#RutaRpaInputErrores")	
	'iGlobal.accion "acAIA", "EXEC|AIA.Vision|ScreenShot||" & RutaScreenshots1 & Maquina & "_" & CUPS & "_" & timeStamp & "_KO.jpg"
	nombrepathfile = RutaScreenshots1 & Maquina & "_" & timeStamp & "_KO.jpg"
	nombrefile = Maquina & "_" & timeStamp & "_KO.jpg"
	iglobal.accion "acEnviarValores",  nombrefile & "=#RutaScreenshots"
	
	iGlobal.accion "acAIA", "EXEC|AIA.Vision|ScreenShot||" & nombrepathfile
	'fWriteLog RutaLogs,"["&now()&"]" & name_script & ":  KO::: RUTA SCREENSHOTS:  " & nombrefile
	fWriteLog RutaLogs,"["&now()&"]" & name_script & ":  KO::: NOMBRE ARCHIVO SCREENSHOTS:  " & nombrefile
	If Err.Number <> 0 Then 
		ScreenshotsKO = false
	 else 
		ScreenshotsKO = true
	End If
End Function

' Funci�n que valida los mensajes de una ventana emergente
's�lo se pasa el mensajes
function VentanasEmergentes (mensaje)
	VentanasEmergentes = 0
	if InStr(UCase(mensaje),"OTRO USUARIO MODIFIC")>0 OR InStr(UCase(mensaje),"LTIMO COMENTARIO")>0 then
		VentanasEmergentes = 1
	elseif InStr(UCase(mensaje),"NO SE PUEDE GUARDAR EL REGISTRO")>0 then
		VentanasEmergentes = 1
	elseif InStr(UCase(mensaje),"SESSION WARNING")>0 then
		VentanasEmergentes = 1
	elseif InStr(UCase(mensaje),"NO SE ENCUENTRA DISPONIBLE")>0 then
		VentanasEmergentes = 1
	elseif InStr(UCase(mensaje),"MENSAJE")>0 and not InStr(UCase(mensaje),"ERRORWF")>0 then
		VentanasEmergentes = 1
	elseif InStr(UCase(mensaje),"COMANDO DE CERRAR")>0 then
		VentanasEmergentes = 1
	elseif InStr(UCase(mensaje),"YA HAY UN EXPLORADOR DE WEB")>0 then
		VentanasEmergentes = 1
	elseif InStr(UCase(mensaje),"EL USUARIO")>0 then
		VentanasEmergentes = 1
	elseif InStr(UCase(mensaje),"SE HA EXCEDIDO EL TIEMPO")>0 then
		VentanasEmergentes = 1
	end if
	if VentanasEmergentes = 1 then
		fWriteLog RutaLogs, "[" & now() & "] [Fun-VentanaEmergente] - WARNING-VENTANA EMERGENTE ENCONTRADA: " & mensaje
		fWriteLog RutaLogs,"REINICIO POR VERTANA EMERGENTE..."
		wscript.sleep 1000
		iglobal.accion "acInteraccion","SCRIPT:$MensajeSiebel_BT_aceptar=CLICK"
		iglobal.accion "acInteraccion","SCRIPT:$MensajeSiebel_BT_aceptar=CLICK"
		wscript.sleep 1000
		iglobal.accion "acInteraccion","SCRIPT:$MensajeSiebel_BT_aceptar=CLICK"
		iglobal.accion "acInteraccion","SCRIPT:$MensajeSiebel_BT_aceptar=CLICK"
	end if
end function

'-----------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------
'											SENTENCIAS SQL
'-----------------------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------------
'											BUSCAR ELEMENTO POR WEB
'Enumera y busca un elemento catalogado en web
'parametros:
'elemento: varible aplicacion  catalogada en web o la  variable con la ventana principal
'bandera:  0 - buscar ventana 1 Buscar variable
'-----------------------------------------------------------------------------------------------------------------------
function buscarElemento_web(elemento,bandera)	
	buscarElemento_web = true
	valor = ""
	i = 0	
	do while valor = "" and i<10
		iglobal.accion "acAIA", "EXEC|AIA.Web|ENUMERAR|$" & elemento
		wscript.sleep 1000
		if bandera then
			valor = iglobal.parse("$" & elemento)
		else
			valor = iglobal.accion("acAIA", "EXEC|AIA.Web|GET|$" & elemento)
		end if
		REM msgbox valor
		i = i + 1
		
	loop
	REM msgbox valor
	if i = 10 then
		buscarElemento_web = false
	end if	
end function

function buscar_elementoIglobal(elemento,iteraciones,tiempo)
	buscar_elementoIglobal = true
	i = 0	
	flaglogado = ""			
	While  flaglogado = "" and i < iteraciones 			
		flaglogado = iglobal.parse("$" & elemento) 
		wscript.sleep 1000	
		i = i + 1
	Wend		
	if i = iteraciones then
		buscar_elementoIglobal = false		
	end if
end function 		
	
function buscarElemento_web2(elemento,bandera,iteraciones, tiempo)	
	buscarElemento_web2 = true
	valor = ""
	i = 0
	'msgbox bandera	
	do while valor = "" and i<iteraciones
		iglobal.accion "acAIA", "EXEC|AIA.Web|ENUMERAR|$" & elemento
		wscript.sleep tiempo
		if bandera then
			valor = iglobal.parse("$" & elemento)
			'MSGBOX "1"
		else
			valor = iglobal.accion("acAIA", "EXEC|AIA.Web|GET|$" & elemento)
			'MSGBOX "2"
		end if
		'msgbox valor
		i = i + 1		
	loop
	REM msgbox valor
	if i = iteraciones then
		buscarElemento_web2 = false
	end if	
end function
Function BucleEsperaValor(variable, iteraciones, tiempo)
	check = ""
	i = 0
	REM test "BucleEsperaValor: " &i &variable
	While check = "" AND i < iteraciones 
		check = iglobal.parse(Variable)
		wscript.sleep tiempo
		i = i + 1
	Wend
	
	If i = iteraciones Then
		BucleEsperaValor = false
	Else
		BucleEsperaValor = true
	End If
End Function
function agregarfocoventana(ventana)
	V1_VENTANA = ventana
	wscript.sleep 500
	iGlobal.accion "acAIA","EXEC|AIA.ACA|FOCO|$" & V1_VENTANA & "|^" & V1_VENTANA
	wscript.sleep 500
	iGlobal.accion "acInteraccion", "SCRIPT:$" & V1_VENTANA & "=MAXIMIZA"
end function 
'ESPERA UN ELEMENTO HASTA SER CARGADO
Public Function EsperarEtiquetaBrowser(Pattern,Retry,TimeMs)
	ScriptFilename99 = "99_funciones.vbs"
	Dim Found
	Found = false
	Dim Count
	Count = 0
	DO WHILE NOT Found AND Count < Retry
		Found = iGlobal.accion ("acAIA","EXEC|AIA.Browser|BuscarEtiqueta|"& Pattern)
		Count = Count + 1
		IF NOT Found THEN
			wscript.sleep TimeMs	
		END IF
	loop
	fWriteLog RutaLogs, "Script: " & ScriptFilename99 & " Funcion EsperarEtiquetaBrowser INFO-[" & now() & "]-EsperarEtiqueta " & Found &"-Patron:"& Pattern &" Reintentos:"& Count
	EsperarEtiquetaBrowser = Found
End Function

 function identificadorunico()
	Dim thisday , thistime
	thisday = Date
	thistime = Time
	identificadorunico =iglobal.parse("@puesto") &"_"& "" &Year(thisday)& "" &Month(thisday) & "" & Day(thisday) &""&  Hour(thistime) & "" & Minute(thistime) & "" & Second(thistime) 
 end function

function borrarArchivo(Ruta)
	dim filesys
     borrarArchivo = false     
    if Ruta <> "" then
		Set filesys = CreateObject("Scripting.FileSystemObject") 
		 filesys.CreateTextFile Ruta, True 
		 If filesys.FileExists(Ruta) Then 
			filesys.DeleteFile (Ruta) 
			borrarArchivo = true 
		 End If 	
	end if	
end function

'Hace el login y obtiene el Token 
'Retorna True (OK ) o False  (KO)
'ORDEN PARAMENTROS
'1 Opcion (tipo de accion)      "login"
'2 host (ip)					 hostAranda
'3 path (resto dominio login)    pathAranada
'4 usuario						 userAranda 	
'5 pass							 passAranda
'6 Ruta de log                   RutaLogs
'7 Ruta respuesta api rest       rutaRespuestaws 
Function InicioDeSesion(hostAranda, pathAranada, userAranda, passAranda, RutaLogs, rutaRespuestaws)
	'if (iglobal.parse("#FlagDebug")) then msgbox "Funcion InicioDeSesion" end if 
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-INICIO FUNCION InicioDeSesion"
	InicioDeSesion = FALSE
	borrarArchivo(rutaRespuestaws & "response.json")
	ows.Run "node " & rutaJavascript & "rest.js login "& hostAranda &" "& pathAranada &" "& userAranda  &" "& passAranda &" "& RutaLogs &" "& rutaRespuestaws	
	Set json = New VbsJson	
	iglobal.accion "acEnviarValores", "=#tokenSesionAranda" 'almacena el token de la sesion	
	'Esperar qeu el archivo se cree
	ruta = rutaRespuestaws & "response.json"
	iteraciones = 5
	tiempo = 3000
	flagfileexist = BucleEsperaArchivoJson(ruta, iteraciones, tiempo)
	if not flagfileexist then
		fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FALLO INICIO DE SESI�N NO SE GENERO JSON"		
	else 
		'Se puede agregar aca una validacion si archivo existe
		Set fso = WScript.CreateObject("Scripting.Filesystemobject")
		'str = fso.OpenTextFile(rutaFiles & "jsonblanco.txt").ReadAll
		str = fso.OpenTextFile(rutaRespuestaws & "response.json").ReadAll
		'msgbox InStr(str,"FailureOnLicense")
		if InStr(str,"FailureOnLicense") <> 0 Then		
			'if (iglobal.parse("#FlagDebug")) then msgbox str end if 
			fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FALLO INICIO DE SESI�N: " & str
			fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FIN FUNCION InicioDeSesion " & InicioDeSesion
			exit function
			'wscript.quit(-1)
		end if		
		str = "{""token"":" & str & "}"
		Set jsonToken = json.Decode(str)
		For Each token In jsonToken
			For Each jsoncuerpo In jsonToken(token)
				'if modelo <> "" then
				For Each x In jsoncuerpo
					if  x = "Field" then
						if jsoncuerpo("Field") = "sessionId" then
							'msgbox " --> "&jsoncuerpo("Field") &" : "& jsoncuerpo("Value")
							iglobal.accion "acEnviarValores", jsoncuerpo("Value") & "=#tokenSesionAranda"
						end if
						if jsoncuerpo("Field") = "userId" then
							'msgbox " --> "&jsoncuerpo("Field") &" : "& jsoncuerpo("Value")
							iglobal.accion "acEnviarValores", jsoncuerpo("Value") & "=#UserId"
						end if
					end if
				'end if
				next
			next		
		next
		'if (iglobal.parse("#FlagDebug")) then msgbox iglobal.parse("#tokenSesionAranda") &" "& iglobal.parse("#UserId") end if 
		fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-TOKEN OBTENIDO: " & iglobal.parse("#tokenSesionAranda")	
		fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-IDUSUARIO OBTENIDO: " & iglobal.parse("#UserId")	
		InicioDeSesion = TRUE				
	end if 	
		fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FIN FUNCION InicioDeSesion " & InicioDeSesion
End Function
'cierra la sesion
'Retorna True (OK ) o False  (KO)
'ORDEN PARAMENTROS
'1 Opcion (tipo de accion)      "logout"
'2 host (ip)					 hostAranda
'3 path (resto dominio logout)   logoutAranda
'4 token de sesion	     		 tokenSesion 	
'5 Ruta de log                   RutaLogs
'6 Ruta respuesta api rest       rutaRespuestaws 
Function CierreDeSesion(hostAranda, logoutAranda,tokenSesion, RutaLogs, rutaRespuestaws)	
	'if (iglobal.parse("#FlagDebug")) then msgbox "Funcion CierreDeSesion " & iglobal.parse("#tokenSesionAranda") end if 
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-INICIO FUNCION CierreDeSesion"
	borrarArchivo(rutaRespuestaws & "response.json")
	'msgbox hostAranda &" -> "& logoutAranda &" -> "& tokenSesion &" -> "& RutaLogs
	ows.Run "node " & rutaJavascript & "rest.js logouts "& hostAranda &" "& logoutAranda &" "& tokenSesion &" "& RutaLogs &" "& rutaRespuestaws	
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FIN FUNCION CierreDeSesion"
End Function

'LISTA LOS CASOS DE USUARIO
'Retorna True (OK ) o False  (KO)
'ORDEN PARAMENTROS
'1 Opcion (tipo de accion)      "listcasos"
'2 host (ip)					 hostAranda
'3 path (resto dominio logout)   pathAranda
'4 token de sesion	     		 tokenSesion 	
'5 Ruta de log                   RutaLogs
'6 Ruta respuesta api rest       rutaRespuestaws 
'7 Id del proyeto			     ProjectIdAranda 
'8 Id del tipo de caso           ItemTypeAranda 
' Function ListarCasosUsuario(hostAranda, pathAranda,tokenSesion, RutaLogs, rutaRespuestaws ,ProjectIdAranda, ItemTypeAranda)	
Function ListarCasosUsuario(hostAranda, pathAranda,tokenSesion, RutaLogs, rutaRespuestaws ,ProjectIdAranda, ItemTypeAranda)	
	'if (iglobal.parse("#FlagDebug")) then msgbox "Funcion ListarCasosUsuario " & iglobal.parse("#tokenSesionAranda") end if 
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-INICIO FUNCION ListarCasosUsuario"
	ListarCasosUsuario = false
	borrarArchivo(rutaRespuestaws & "response.json")
	'ows.Run "node " & rutaJavascript & "rest.js listcasos "& hostAranda &" "& pathAranada &" "& userAranda  &" "& passAranda &" "& RutaLogs &" "& rutaRespuestaws &" "& ProjectIdAranda &" "& ItemTypeAranda	
	'msgbox rutaRespuestaws
	ows.Run "node " & rutaJavascript & "rest.js listcasos "& hostAranda &" "& pathAranda &" "& tokenSesion &" "& RutaLogs &" "& rutaRespuestaws &" "& ProjectIdAranda &" "& ItemTypeAranda
	ruta = rutaRespuestaws & "response.json"
	iteraciones = 5
	tiempo = 3000
	flagfileexist = BucleEsperaArchivoJson(ruta, iteraciones, tiempo)
	if not flagfileexist then
		fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FALLO EL PROCESO DE LISTAR CASOS NO SE GENERO EL ARCHIVO JSON"		
	else 
		Set fso = WScript.CreateObject("Scripting.Filesystemobject")
		'str = fso.OpenTextFile(rutaFiles & "jsonblanco.txt").ReadAll
		str = fso.OpenTextFile(rutaRespuestaws & "response.json").ReadAll		
		if InStr(str,"InvalidToken") <> 0 Then		
			'if (iglobal.parse("#FlagDebug")) then msgbox str end if 
			fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FALLO EL PROCESO DE LISTAR CASOS: " & str			
		else
			ListarCasosUsuario = true
		end if	
	end if
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FIN FUNCION ListarCasosUsuario RTA: " & ListarCasosUsuario
End Function

'Espera que se cree el archivo recibido por APIREST
'ORDEN PARAMENTROS
'ruta 			Ruta y nombre del archivo
'itereacciones 	Cantidad de reintentos
'tiempo         Tiempo de espera que se cree el archivo
Function BucleEsperaArchivoJson(ruta, iteraciones, tiempo)	
	Set fileSys = CreateObject("Scripting.FileSystemObject") 
	flagexiste = true
		i = 0	
	While flagexiste AND i < iteraciones 		
		Set fileSys = CreateObject("Scripting.FileSystemObject")
		If fileSys.FileExists(ruta) Then 
			flagexiste = false
		end if
		wscript.sleep tiempo
		i = i + 1
	Wend	
	If i = iteraciones Then
		BucleEsperaArchivoJson = false
	Else
		BucleEsperaArchivoJson = true
	End If	
End Function

function renovarSesion()
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-INICIO FUNCION renovarSesion"
	renovarSesion = false
	ruta = iGlobal.parse("#RutaRespuestaws") & "response.json"
	borrarArchivo(ruta)
	ScriptFilename = "99_funciones.vbs"
	pathGenerico = iGlobal.parse("#PathGenericoAranda") &""& iGlobal.parse("#RenoSesionAranda")
	rta = ows.Run ("node " & iGlobal.parse("#RutaJavascript") & "rest.js renovarsesion "& iGlobal.parse("#hostAranda") &" "& pathGenerico &" "& iglobal.parse("#tokenSesionAranda")  &" "& iGlobal.parse("#RutaLogs") &" "& iGlobal.parse("#RutaRespuestaws"),1,true)
	iteraciones = 5
	tiempo = 3000
	flagfileexist = BucleEsperaArchivoJson(ruta, iteraciones, tiempo)
	if not flagfileexist then
		fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FALLO EL PROCESO RENOVAR SESION. NO SE GENERO EL ARCHIVO JSON"		
		fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FIN FUNCION renovarSesion RTA: " & renovarSesion
		exit function
	else 
		Set fso = WScript.CreateObject("Scripting.Filesystemobject")
		'str = fso.OpenTextFile(rutaFiles & "jsonblanco.txt").ReadAll
		str = fso.OpenTextFile(ruta).ReadAll		
		if not InStr(str,"true") <> 0  Then				
			fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FALLO EL PROCESO RENOVAR SESION." & str	
			fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FIN FUNCION renovarSesion RTA: " & renovarSesion	
			exit function	
		end if	
		renovarSesion = true
	end if	
	renovarSesion = true
	
end function

 'AGREGA DATOS AL FORMTAO DE SOLICITUD EXCEL
'Retorna True (OK ) o False  (KO)
'ORDEN PARAMENTROS
'1 parametro con opcion de debe realizar el CASE.
Function AgregarDatosFormatoNovedades(parametro)
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-INICIO FUNCION AgregarDatosFormatoNovedades"	
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-PARAMETRO " & parametro
	AgregarDatosFormatoNovedades = false
	'Definici�n de variables
	 Dim objExcel
	 Dim objWorkbook
	 Dim Hoja : Hoja = 1
	 Dim max_col : max_col = 20 
	 Dim max_row : max_row = 371
	 Dim col : col = 2
	 Dim row : row = 5
	'Abrir Excel	
	Dim rutaExcel : rutaExcel = iGlobal.parse("#RutaNombreFormatoSolicitud")
	if rutaExcel <> "" then
		Set filesys = CreateObject("Scripting.FileSystemObject") 		
		 If NOT filesys.FileExists(rutaExcel) Then 
			fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FIN FUNCION AgregarDatosFormatoNovedades RTA: " & AgregarDatosFormatoNovedades	
			exit function
		 End If 	
	end if
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-ABRE EL EXCEL  " & rutaExcel
	AbrirExcel objExcel, objWorkbook, rutaExcel
	select case parametro	
		case "impresion"
			fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-INGRESA LOS DATOS CON EL CODIGO DE IMPRESION" 			
			DarValorCelda objExcel,1,25,8,codigoImpresion			
			especialistaImpresion = quitarAcentos(especialistaImpresion)
			DarValorCelda objExcel,1,27,8,especialistaImpresion
			DarValorCelda objExcel,1,28,8,dateModImpresion
		case "mbs"
			fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-INGRESA LOS DATOS CON EL USUARIO MBS" 			
			DarValorCelda objExcel,1,25,4,usuarioMBS
			especialistaMBS = quitarAcentos(especialistaMBS)	
			DarValorCelda objExcel,1,27,4,especialistaMBS			
			DarValorCelda objExcel,1,28,4,dateModMBS	
		case "galeon"
			fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-INGRESA LOS DATOS CON EL USUARIO GALEON" 			
			DarValorCelda objExcel,1,25,10,usuarioGALEON
			especialistaGALEON = quitarAcentos(especialistaGALEON)	
			DarValorCelda objExcel,1,27,10,especialistaGALEON			
			DarValorCelda objExcel,1,28,10,dateModGALEON		
		case "mafis"
			fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-INGRESA LOS DATOS CON EL USUARIO MAFIS" 			
			DarValorCelda objExcel,1,25,9,usuarioMAFIS
			especialistaMAFIS = quitarAcentos(especialistaMAFIS)	
			DarValorCelda objExcel,1,27,9,especialistaMAFIS			
			DarValorCelda objExcel,1,28,9,dateModMAFIS	
	end select
	AgregarDatosFormatoNovedades = true
	GuardarLibro objExcel, ""
	CerrarExcel objExcel, objWorkbook
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FIN FUNCION AgregarDatosFormatoNovedades RTA: " & AgregarDatosFormatoNovedades	
end function

'RETORNA EL VALOR DE UNA CELDA DEL FORMATO SOLICITUD NOVEDADES
'Retorna EL VALOR DE LA CELDA O VACIO
'ORDEN PARAMENTROS
'FILA     
'COLUMNA
Function ObtenerDatosFormatoNovedades(fila,columna)
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-INICIO FUNCION ObtenerDatosFormatoNovedades"	
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-PARAMETRO " & parametro
	ObtenerDatosFormatoNovedades = ""
	'Definici�n de variables
	 Dim objExcel
	 Dim objWorkbook	
	'Abrir Excel	
	Dim rutaExcel : rutaExcel = iGlobal.parse("#RutaNombreFormatoSolicitud")
	if rutaExcel <> "" then
		Set filesys = CreateObject("Scripting.FileSystemObject") 		
		 If NOT filesys.FileExists(rutaExcel) Then 
			fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FIN FUNCION ObtenerDatosFormatoNovedades RTA: " & ObtenerDatosFormatoNovedades	
			exit function
		 End If 	
	end if
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-ABRE EL EXCEL  " & rutaExcel
	AbrirExcel objExcel, objWorkbook, rutaExcel
    ObtenerDatosFormatoNovedades = ObtenerValorCelda(objExcel,1,fila,columna)
	CerrarExcel objExcel, objWorkbook
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FIN FUNCION ObtenerDatosFormatoNovedades RTA: " & ObtenerDatosFormatoNovedades	
end function
function quitarAcentos(testx)
	quitarAcentos = ""
	testx = Replace(testx,chr(225),"a")
	testx = Replace(testx,chr(233),"e")
	testx = Replace(testx,chr(237),"i")
	testx = Replace(testx,chr(243),"o")
	testx = Replace(testx,chr(250),"u")	
	testx = Replace(testx,chr(241),"n")	
	testx = Replace(testx,chr(193),"A")
	testx = Replace(testx,chr(201),"E")
	testx = Replace(testx,chr(205),"I")
	testx = Replace(testx,chr(211),"O")
	testx = Replace(testx,chr(218),"U")	
	testx = Replace(testx,chr(209),"N")	
	quitarAcentos = testx
end function
'LEE Y EXTRAE LOS VALORES DEL EXCEL DE CONTACTOS  DE LA HOJA DE CONFIGURACION Y LOS ASIGNA A LAS VARIABLES IGLOBAL
function LeerConfiguracionInicial()
	LeerConfiguracionInicial = false	
	Dim objExcel
	Dim objWorkbook
	Dim rutaExcel :  rutaExcel = iGlobal.parse("#RutaConfig") & "configuracion.xlsx"
	HojaMC = 1
	filGN = 2 'fila
	colValor= 2 'valor variable
	colVariable= 1 'Nombre de la variable
	AbrirExcel objExcel, objWorkbook, rutaExcel
	registros = true
	fWriteLog RutaLogs,"Script: " & ScriptFilename & "-INFO-[" & now() & "](RPA) - VARIABLES DE CONFIGURACION"
	do while registros		
		nombre = ObtenerValorCelda (objExcel,HojaMC,filGN,colVariable)	
		nombre = trim(nombre)
		if nombre <> "" then
			valor  = ObtenerValorCelda (objExcel,HojaMC,filGN,colValor)
			valor = trim(valor)
			iglobal.accion "acEnviarValores", valor & "=" & nombre	
			fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]--VARIABLE: " & 	nombre & " --VALOR: " & iGlobal.parse(nombre)
			LeerConfiguracionInicial = true
			filGN = filGN + 1
		else
			registros = false	
		end if
	loop
	CerrarExcel objExcel, objWorkbook	
end function
function esNuloTexto(parametro)
	esNuloTexto = ""
	if IsNull(parametro)  then 
	 parametro = "" 
	end if
	esNuloTexto= parametro
end function 

function ErrCatch(numError,bloque,script) 
    If numError <> 0 Then
	    fWriteLog RutaLogs, "Script: " & script & " INFO-[" & now() & "]-OCURRIO UNA EXCEPCION EN EL BLOQUE : " & Bloque & " --NUMERO ERROR: " & numError
		iglobal.accion "acEnviarValores","-EXCEPCI�N DEL SISTEMA -SCRIPT: " & script & " BLOQUE: " & bloque & " ERROR: " & numError & "=#ErrorProceso"
		wscript.quit(5)			
	end if
End function 
'BORARA TODOS LOS ARCHIVOS DE UNA RUTA 
'Retorna  N/A
'ORDEN PARAMETROS
'rutaborrar Indica la ruta donde se van a borrar los archivos
Function BorrarArchivos(rutaborrar)
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-INICIO FUNCION BorrarArchivos"
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-RUTA: "&rutaborrar
	Set objFSO = CreateObject("Scripting.FileSystemObject")	
	If  objFSO.FolderExists(rutaborrar) Then		
		borrar = ""
		borrar = rutaborrar &"\*.*"		
		objFSO.DeleteFile(borrar), DeleteReadOnly
	end if
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FIN FUNCION BorrarArchivos"
end function
'el nombre debe llegar AIA.Browser|BuscarEtiqueta
function esperar_Browser(nombre,navegador,etiqueta,atributo,valor,posicion,iteraciones)
	esperar_Browser = false
	i = 0	
	encontrado = false			
	'msgbox encontrado
	do while encontrado <> "true" and i < iteraciones 
	   encontrado = iglobal.accion ("acAIA","EXEC|AIA.Browser|BuscarEtiqueta|{nombre:" & nombre & ";navegador:" & navegador & ";etiqueta:" & etiqueta & ";atributo:" & atributo & ";valor:" & valor & ";posicion:" & posicion & ";}")
	   'msgbox encontrado
	   wscript.sleep 1000	
	   i = i + 1
	loop
	'msgbox iteraciones & " i " & i
	if i < iteraciones then
	 esperar_Browser = true 
	end if
end function 

Function ValidarDatos(parametro)		
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-INICIO FUNCION ValidarDatos"
	REM id_casoAux				0
	REM num_caso				1 
	REM tipo_envio				2 
	REM cedula				3 
	REM ss_ticket_solicitud		4 
	REM cun					5 
	REM nombre_completo			6
	REM telefono				7
	REM direccion				8	
	REM departamento 			9
	REM ciudad				10
	REM negocio_empresa 		11
	REM nombre_empresa			12 
	REM area 					13
	REM ruta_file 				14
	REM buzon					15
	REM correo 				16
	REM envios 				17
	REM solicitud 				18
	REM radicado 				19
	REM asesor 				20
	REM sede 					21
	REM adjunto				22	
	ValidarDatos = true
	Dim arrMatrizCargos
		arrEtiquetas = Array("ID CASO","NUMERO DE CASO","TIPO DE ENVIO","CEDULA","SS/TICKET/SOLICITUD","CUN","NOMBRE COMPLETO","TELEFONO","DIRECCION","DEPARTAMENTO","CIUDAD","NEGOCIO EMPRESA","NOMBRE EMPRESA","AREA","RUTA ARCHIVOS","BUZON","CORREO","ENVIOS","SOLICITUD","RADICADO","ASESOR","SEDE","ADJUNTO")
		arrMatrizEnvios =Split(iGlobal.parse("#MatrizEnvios"),"||")
		for x=0 to UBound(arrMatrizEnvios)   
			
			if parametro = "EF" then
				if x = 0 or x = 1 or x = 2 or x = 4 or x = 5 or x = 6 or x = 7 or x = 8 or x = 9 or x = 10 or x = 11 or x = 12 or x = 13 or x = 14 then
					if arrMatrizEnvios(x) = "" then			
						fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-DATOS CAPTURADOS: "& arrEtiquetas(x) &":"& arrMatrizEnvios(x) 
						ValidarDatos = false
					end if
				end if	
			elseif parametro = "EC" then
				if x = 0 or x = 1 or x = 2 or x = 15 or x = 16 or x = 6 or x = 4 or x = 14 then
					if arrMatrizEnvios(x) = "" then			
						fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-DATOS CAPTURADOS: "& arrEtiquetas(x) &":"& arrMatrizEnvios(x) 
						ValidarDatos = false
					end if
				end if	
			elseif parametro = "NC" then
				if x = 0 or x = 1 or x = 2 or x = 6 or x = 12 or x = 3 or x = 8 or x = 10 or x = 9 or x = 7 or x = 17 or x = 18 or x = 19 or x = 20 or x = 21 or x = 13 or x = 22 then
					if arrMatrizEnvios(x) = "" then			
						fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-DATOS CAPTURADOS: "& arrEtiquetas(x) &":"& arrMatrizEnvios(x) 
						ValidarDatos = false
					end if
				end if
			end if
		next			
	fWriteLog RutaLogs, "Script: " & ScriptFilename & " INFO-[" & now() & "]-FIN FUNCION ValidarDatos " & ValidarDatos
end function
function f_damePorcentaje (Conexion) ' Nueva version	
	fWriteLog RutaLogs, "[" & now() & "] f_damePorcentaje: 001"
	SQL = "SELECT COUNT(*) AS TOTALES FROM psr_input  where DATE(fecha_inicio) = current_date()"
	'SQL = "SELECT COUNT(*) AS TOTALES FROM crear_usuario  where DATE(Fecha_fin) = '2019-07-15'"

	'wscript.echo SQL
	EjecutarQuery Conexion, SQL, ResultSetBD, RA
	fWriteLog RutaLogs, "[" & now() & "] f_damePorcentaje: 002"
	totales = ResultSetBD(0)
	fWriteLog RutaLogs, "[" & now() & "] totales: " & totales
	'msgbox "totales:" & totales

	SQL = "SELECT COUNT(*) AS TOTALES FROM psr_input  where DATE(fecha_inicio) = current_date() AND estadoAutomatizacion <> 'PROCESANDO' AND estadoAutomatizacion <> 'PENDIENTE'"
	'SQL = "SELECT COUNT(*) AS TOTALES FROM crear_usuario  where DATE(Fecha_fin) = '2019-07-15' AND estado <> 'PENDIENTE'"
	EjecutarQuery Conexion, SQL, ResultSetBD, RA

	completados = ResultSetBD(0)

	fWriteLog RutaLogs, "[" & now() & "] completados: " & completados
	'MsgBox "completados:" & completados


	'porcentaje = (completados/totales) * 100
	porcentaje = (CInt(completados) * 100) 
	porcentaje = porcentaje/CInt(totales)

	porcentaje = Round(porcentaje,2)

	f_damePorcentaje = porcentaje
	'f_damePorcentaje = 0
end function

Function AbrirConexion()
	AbrirConexion = true
	if AbrirConexion Then
		fWriteLog RutaLogs,"Script: " & ScriptFilename & "-INFO-[" & now() & "](RPA) - Abrir conexion con BD de proceso de Gestion PSR"
		contBD=0
		do while not ConectarBD_DSN(BD_CONEXION,DSN_CONEXION,USER_CONEXION,PASS_CONEXION) and contBD < 3	
			fWriteLog RutaLogs,"Script: " & ScriptFilename & "-INFO-[" & now() & "](RPA) - ERROR EN LA CONEXION A LA BASE DE DATOS(gestion_psr) --- REINTENTO: " & contBD
			wscript.sleep 120000 ' 2 min
			contBD = contBD + 1	
		loop
		if contBD = 3 then
			fWriteLog RutaLogs,"Script: " & ScriptFilename & "-ERROR-[" & now() & "](RPA) - Error en la conexion a la base de datos"
			AbrirConexion = false	
		end if
	end if
	fWriteLog RutaLogs, "Script: " & ScriptFilename & "-INFO-[" & now() & "](RPA) - Datos de conexi?n:" & vbCrLf & BD_CONEXION
End Function