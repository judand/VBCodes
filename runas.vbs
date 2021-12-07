' -------------------------------------------------------------------------------------------------'
'	Nombre: .vbs
'	Creacion: 15.01.20 -  - <codigo Jira>
'	Descripcion: 
'	Entrada: No tiene parametros de entrada
'	Salida:escribe en archivo salida true o false
'	Modificaciones: 
'	- <codigo Jira> 
' -------------------------------------------------------------------------------------------------'

'-------------------------------< DECLARACION DE OBJETOS e iniciacion >----------------------------
'--------------------------------------------------------------------------------------------------
DIM oWS
SET oWS = WScript.CreateObject("WScript.Shell")
Set iglobal = CreateObject("iGlobal.iViewObj")
'set rutaLibs = iGlobal.parse("#RutaLib"
'set strRutaLogs = iGlobal.parse("#RutaLogs")
rutaLibs="C:\Users\Administrador\Documents\lib\"
strRutaLogs= "C:\Users\Administrador\Documents\SmurfitKappa documentos\desarrollo\LOGS\"
'********** Importar Librerias **********
Import rutaLibs & "99_funciones.vbs"

'************************************************SCRIPT********************************************


'vari="javierf.duran@gmail.com"
'varnueva=Split(vari,"@")
'msgbox varnueva(0)
fWriteLog  strRutaLogs, "INFO["& now() &"]-runas.vbs-Inicia ejecucion"
oWS.Run "cmd.exe", 1, False
Wscript.Sleep 2000 'need to give time for window to open.
oWS.AppActivate "cmd" 'make sure we grab the right window to send password to
wscript.sleep 3000
SendKey "Colombia..4272", 1000
oWS.SendKeys "{ENTER}"
wscript.sleep 3000
'oWS.AppActivate "cmd"
'oWS.SendKeys "C:\Users\duranja\Documents\Smurfit_crearusuarios\DESARROLLO\src\bd_conection.vbs" 'send the password to the waiting window.
'wscript.sleep 200
'oWS.SendKeys "{ENTER}"
wscript.quit
'************************************************FUNCIONES*****************************************
Private Sub Import(ByVal filename)	
	Dim fso, sh, code, dir
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set sh  = CreateObject("WScript.Shell")
	filename = Trim(sh.ExpandEnvironmentStrings(filename))
	If Not (Left(filename, 2) = "\\" Or Mid(filename, 2, 2) = ":\") Then
		If Not fso.FileExists(fso.GetAbsolutePathName(filename)) Then
			For Each dir In Split(sh.ExpandEnvironmentStrings("%PATH%"), ";")
				If fso.FileExists(fso.BuildPath(dir, filename)) Then
					filename = fso.BuildPath(dir, filename)
					Exit For
				End If
			Next
		End If
		filename = fso.GetAbsolutePathName(filename)
	End If
	code = fso.OpenTextFile(filename).ReadAll
	ExecuteGlobal code
	Set fso = Nothing
	Set sh  = Nothing
End Sub

Public Sub sendKey(key , waitMS ) 'funcion que usa sendkeys con espera de tiempo entre cada caracter
	  Dim endTime
	 if Len(key)>0 then
		for i=1 to Len(key)
			'msgbox  Mid(key,i,1)
			oWS.SendKeys Mid(key,i,1)
			wscript.sleep waitMS
		next
	 end if
End Sub