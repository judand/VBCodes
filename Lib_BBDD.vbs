
'-----------------------------------------------------------------------------------------------------------------------
'											FUNCIONES DE BBDD
'-----------------------------------------------------------------------------------------------------------------------

function ConectarBD(Conexion,CadenaConexion)
	Set Conexion = CreateObject("ADODB.Connection")
	On Error Resume Next
		Conexion.Open CadenaConexion
		ConectarBD = true
	If Err.Number <> 0 Then
			ConectarBD = false
			Err.Clear
		end if
	On Error Goto 0
end function

' function EjecutarQuery(Conexion,SQL,Resultado,RA)
	' RutaLogs = iglobal.parse("#RutaLogs")
	' On Error Resume Next
		' set Resultado = Conexion.execute(SQL,RA)
		' EjecutarQuery = true
	' If Err.Number <> 0 Then
			' fWriteLog RutaLogs,"Err.Description: " & Err.Description
			' EjecutarQuery = false
			' Err.Clear
		' end if
	' On Error Goto 0
' end function

function EjecutarQuery_2(Conexion,SQL,Resultado,RA)
	RutaLogs = iglobal.parse("#RutaLogs")
	On Error Resume Next
		fWriteLog RutaLogs,"99_funciones.vbs - EjecutarQuery_2 - SQL[" & SQL & "]"
		set Resultado = Conexion.execute(SQL,RA)
		fWriteLog RutaLogs,"99_funciones.vbs - EjecutarQuery_2 - Conexion.State[" & Conexion.State & "]"
		EjecutarQuery_2 = true
	If Err.Number <> 0 Then
			fWriteLog RutaLogs,"99_funciones.vbs - EjecutarQuery_2 - Err.Description: " & Err.Description & " -- Err.Number: " & Err.Number
			EjecutarQuery_2 = false
			Err.Clear
	end if
	On Error Goto 0
end function

function CerrarConexion(Conexion)
	On Error Resume Next
		Conexion.Close
		CerrarConexion = true
		If Err.Number <> 0 Then
			CerrarConexion = false
			Err.Clear
		end if
	On Error Goto 0
end function

Function ConectarBD_DSN(Conexion,DSN,Usuario,Contrasena)
	fWriteLog RutaLogs, "99_funciones.vbs - ConectarBD_DSN: " & vbCrLf & "Usuario: " & Usuario & vbCrLf & "Contrase�a: " & Contrasena
	Set Conexion = CreateObject("ADODB.Connection")
	On Error Resume Next
		Conexion.Open DSN, Usuario, Contrasena
		ConectarBD_DSN = true
		If Err.Number <> 0 Then
			fWriteLog RutaLogs, "onectarBD_DSN: " & vbCrLf & "Err.Number: " & Err.Number & vbCrLf & "Err.Description: " & Err.Description
			ConectarBD_DSN = false
			Err.Clear
		end if
	On Error Goto 0
end Function
'----------------------------------------------CONEXION SQL SERVER -------------------------------

DIM Connection

function ConectarBDSQL(Connection,user,pass,esquema)'PENDIENTE PASAR LOS DATOS DINAMICAMENTE
	'ConnString="DRIVER={SQL Server};SERVER=srvcaysqlvm;UID=iglobal;" & _ 
	'"PWD=iglobal;DATABASE=iGlobal"
	ConnString="DRIVER={SQL Server};SERVER=srvcaysqlvm;UID="& user &";" & _ 
	"PWD="& pass &";DATABASE=" & esquema
	
	Set Connection = CreateObject("ADODB.Connection")
	On Error Resume Next
		Connection.Open ConnString
		ConectarBDSQL = true
	If Err.Number <> 0 Then
			ConectarBDSQL = false
			Err.Clear
		end if
	On Error Goto 0
end function


'--------------------------------------------- FIN SQL SEVER------------------------------------------'
Dim Conexion 
'Conectar a una Base de datos 
function ConectarBD(Conexion,CadenaConexion)
	Set Conexion = CreateObject("ADODB.Connection")
	f_iReintentos = 0
	f_conectada = false
	do while f_iReintentos < 3 and not f_Conectada
		On Error Resume Next
			Conexion.Open CadenaConexion
			ConectarBD = true
			f_Conectada = true
		If Err.Number <> 0 Then
				ConectarBD = false
				Err.Clear
				f_Conectada = false
			end if
		On Error Goto 0
		f_iReintentos = f_iReintentos + 1
	loop
end function
'Ejecutar una Query y devolver un resultado
function EjecutarQuery(Conexion,SQL,Resultado,RA)
	f_iReintentos = 0
	f_conectada = false
	do while f_iReintentos < 3 and not f_Conectada
		On Error Resume Next
			set Resultado = Conexion.execute(SQL,RA)
			EjecutarQuery = true
			f_Conectada = true
		If Err.Number <> 0 Then
				EjecutarQuery = false
				f_Conectada = false
				Err.Clear
			end if
		On Error Goto 0
		f_iReintentos = f_iReintentos + 1
	loop
end function
'Cerrar la conexión a la Base de datos
function CerrarConexion(Conexion)
	On Error Resume Next
		Conexion.Close
		CerrarConexion = true
		If Err.Number <> 0 Then
			CerrarConexion = false
			Err.Clear
		end if
	On Error Goto 0
end function