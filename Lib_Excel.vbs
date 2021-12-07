'******************************************************************************************
'*************************************EXCEL************************************************
'******************************************************************************************
'Abrir Excel
'objExcel: Devuelve objeto de Excel abierto
'objWorkbook: Devuelve el Workbooks del Excel abierto
'RutaFichero: Ruta del fichero Excel a abrir
function AbrirExcel(objExcel,objWorkbook,RutaFichero)
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open(RutaFichero)
end function

'Cerrar Excel
'objExcel: Objeto Excel abierto
'objWorkbook: Objeto Workbooks del Excel abierto
function CerrarExcel(objExcel,objWorkbook)
	objExcel.DisplayAlerts = False
	objExcel.Quit
	Set objWorkbook = Nothing
	Set objExcel = Nothing
end function

'Obtener Valor de celda Excel
'objExcel: Objeto Excel abierto
'Hoja: Hoja de donde se van a tomar los valores
'Fila: Fila de donde se va a tomar el valor (1...N)
'Columna: Columna de donde se va a tomar el valor (1...N)
function ObtenerValorCelda(objExcel,Hoja,Fila,Columna)
	ObtenerValorCelda = objExcel.Sheets(Hoja).Cells(Fila,Columna).Value
end function
'******************************************************************************************
'******************************************************************************************
function DarValorCelda(objExcel,Hoja,Fila,Columna,Valor)
	objExcel.Sheets(Hoja).Cells(Fila,Columna).Value = Valor
end function

function GuardarLibro(objExcel,Ruta)
	if(Ruta = "") then
		objExcel.ActiveWorkbook.Save
	else
		objExcel.ActiveWorkBook.SaveAs Ruta
	end if
end function

function ObtenerNumeroHojas(objExcel)
	ObtenerNumeroHojas = objExcel.Sheets.Count
end function

'Obtiene el nombre de la hoja actual
function ObtenerNombreHoja(objExcel,n_hoja)
	ObtenerNombreHoja = objExcel.Sheets(n_hoja).Name
end function

'obtiene numero de filas de la hoja actual
function ObtenerFilas(objExcel,n_hoja)
	ObtenerFilas = objExcel.Sheets(n_hoja).UsedRange.Rows.Count 
end function
function ObtenerColumnas(objExcel,n_hoja)
	ObtenerColumnas = objExcel.Sheets(n_hoja).UsedRange.Columns.Count 
end function