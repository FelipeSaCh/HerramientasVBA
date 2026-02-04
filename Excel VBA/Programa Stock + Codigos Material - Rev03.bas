Attribute VB_Name = "Módulo1"
Sub CopiarAConAsteriscos()
    Dim ws As Worksheet
    Dim wsSend As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Codigo Partes")
    Set wsSend = ThisWorkbook.Sheets("SEND")
    
    ' Última fila con datos en columna A
    ultimaFila = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row
    
    ' Recorrer desde la fila 2
    For i = 8 To ultimaFila
        If ws.Cells(i, "M").Value <> "" Then
            wsSend.Cells(i, "A").Value = "*" & ws.Cells(i, "M").Value & "*"
        End If
    Next i
    
    ' Copiar columna A de SEND al portapapeles
    ultimaFila = wsSend.Cells(wsSend.Rows.Count, "A").End(xlUp).Row
    If ultimaFila >= 8 Then
        wsSend.Range("A1:A" & ultimaFila).Copy
    Else
        MsgBox "Proceso completado pero no hay datos para copiar al portapapeles.", vbExclamation
    End If
End Sub


Sub CopiarA()
    Dim ws As Worksheet
    Dim wsSend As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Codigos SAP")
    Set wsSend = ThisWorkbook.Sheets("SEND")
    
    ' Última fila con datos en columna A
    ultimaFila = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row
    
    ' Recorrer desde la fila 2
    For i = 8 To ultimaFila
        If ws.Cells(i, "M").Value <> "" Then
            wsSend.Cells(i, "A").Value = "" & ws.Cells(i, "M").Value & ""
        End If
    Next i
    
    ' Copiar columna A de SEND al portapapeles
    ultimaFila = wsSend.Cells(wsSend.Rows.Count, "A").End(xlUp).Row
    If ultimaFila >= 8 Then
        wsSend.Range("A1:A" & ultimaFila).Copy
    Else
        MsgBox "Proceso completado pero no hay datos para copiar al portapapeles.", vbExclamation
    End If
End Sub

Sub EjecutarMMBEenSAP()
    Dim sapGui As Object
    Dim applicationSAP As Object
    Dim connection As Object
    Dim session As Object
    Dim conexionExitosa As Boolean
    
    ' Conectar a SAP
    conexionExitosa = ConectarSAP(sapGui, applicationSAP, connection, session)
    
    If Not conexionExitosa Then
        MsgBox "No se pudo conectar a SAP. El proceso se cancelará.", vbCritical
        Exit Sub
    End If
    
    ' Ejecutar el script de MMBE
    EjecutarScriptMMBE session
    
End Sub

Sub EjecutarMMBEenSAPV2()
    Dim sapGui As Object
    Dim applicationSAP As Object
    Dim connection As Object
    Dim session As Object
    Dim conexionExitosa As Boolean
    
    ' Conectar a SAP
    conexionExitosa = ConectarSAP(sapGui, applicationSAP, connection, session)
    
    If Not conexionExitosa Then
        MsgBox "No se pudo conectar a SAP. El proceso se cancelará.", vbCritical
        Exit Sub
    End If
    
    ' Ejecutar el script de MMBE
    EjecutarScriptMMBE_V2 session
    
End Sub


Sub EjecutarScriptMMBE_V2(session As Object)
    On Error GoTo ErrorHandler
    
    ' Maximizar ventana
    session.findById("wnd[0]").maximize
    
    ' Ingresar a transacción MMBE
    session.findById("wnd[0]/tbar[0]/okcd").Text = "MMBE"
    session.findById("wnd[0]").sendVKey 0
    
    ' Ejecutar acciones del script grabado
    session.findById("wnd[0]/usr/ctxtMS_MATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtMS_WERKS-LOW").Text = ""
    session.findById("wnd[0]").sendVKey 4
    session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB019/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[3,24]").Text = ""
    session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB019").Select
    session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB019/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/btnG_SELFLD_TAB-MORE[2,56]").press
    
    session.findById("wnd[2]/tbar[0]/btn[24]").press
    session.findById("wnd[2]/tbar[0]/btn[8]").press
    session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB019/ssubSUBSCR_PRESEL:SAPLSDH4:0220/chkG_SELPOP_STATE-BUTTON").SetFocus
    session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB019/ssubSUBSCR_PRESEL:SAPLSDH4:0220/chkG_SELPOP_STATE-BUTTON").Selected = True
    Esperar 3
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 13
    session.findById("wnd[1]").sendVKey 14
    session.findById("wnd[2]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[2]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
    session.findById("wnd[2]/tbar[0]/btn[0]").press
    session.findById("wnd[1]").Close
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    
    Exit Sub
    
ErrorHandler:
End Sub

Sub EjecutarScriptMMBE(session As Object)
    On Error GoTo ErrorHandler
    
    ' Maximizar ventana
    session.findById("wnd[0]").maximize
    
    ' Ingresar a transacción MMBE
    session.findById("wnd[0]/tbar[0]/okcd").Text = "MMBE"
    session.findById("wnd[0]").sendVKey 0
    
    ' Ejecutar acciones del script grabado
    session.findById("wnd[0]/usr/ctxtMS_MATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtMS_WERKS-LOW").Text = ""
    session.findById("wnd[0]").sendVKey 4
    session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB019/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[3,24]").Text = ""
    session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB019").Select
    session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB019/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/btnG_SELFLD_TAB-MORE[0,56]").press
    
    session.findById("wnd[2]/tbar[0]/btn[24]").press
    session.findById("wnd[2]/tbar[0]/btn[8]").press
    session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB019/ssubSUBSCR_PRESEL:SAPLSDH4:0220/chkG_SELPOP_STATE-BUTTON").SetFocus
    session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB019/ssubSUBSCR_PRESEL:SAPLSDH4:0220/chkG_SELPOP_STATE-BUTTON").Selected = True
    Esperar 3
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 13
    session.findById("wnd[1]").sendVKey 14
    session.findById("wnd[2]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[2]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
    session.findById("wnd[2]/tbar[0]/btn[0]").press
    session.findById("wnd[1]").Close
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    
    Exit Sub
    
ErrorHandler:
End Sub


Sub Esperar(segundos As Integer)
    Application.Wait (Now + TimeValue("00:00:" & Format(segundos, "00")))
End Sub

Sub PegarPortapapelesEnDATA()
    Dim wsData As Worksheet
    Dim ultimaFila As Long
    Dim rango As Range
    
    ' Referencia a la hoja DATA
    Set wsData = ThisWorkbook.Sheets("DATA")
    
    ' Limpiar datos anteriores
    wsData.Cells.clear
    
    ' Pegar directamente sin seleccionar
    On Error Resume Next
    wsData.Range("A1").PasteSpecial Paste:=xlPasteAll
    If Err.Number <> 0 Then
        MsgBox "No hay datos en el portapapeles o no se puede pegar.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' El resto del código permanece igual...
    ' Encontrar la última fila con datos en columna A
    ultimaFila = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    If ultimaFila < 1 Then
        MsgBox "No se encontraron datos en el portapapeles.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Seleccionar el rango con datos
    Set rango = wsData.Range("A1:A" & ultimaFila)
    
    ' Usar Text-to-Columns nativo de Excel
    rango.TextToColumns _
        Destination:=rango.Cells(1, 1), _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=False, _
        Semicolon:=False, _
        Comma:=False, _
        Space:=False, _
        Other:=True, _
        OtherChar:="|", _
        FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), _
                        Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1))
    
    ' Ocultar columna A
    wsData.Columns("A:E").Hidden = True

    ' Ajustar el ancho de las columnas
    wsData.Columns("B:K").AutoFit
    
    Application.ScreenUpdating = True
    
    ' MsgBox "Proceso completado: " & (ultimaFila) & " filas procesadas con Text-to-Columns.", vbInformation
End Sub

Sub consolidarC()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, newRow As Long
    Dim dict As Object
    Dim key As String
    Dim keyVar As Variant
    Dim tempDictB As Object, tempDictE As Object
    
    Set ws = ThisWorkbook.Sheets("DATA")
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Procesar datos desde fila 4
    For i = 4 To lastRow
        key = ws.Cells(i, "D").Value ' identificador en columna D
        
        ' Validar que la clave no esté vacía
        If key <> "" Then
            If Not dict.Exists(key) Then
                ' Para cada identificador guardamos dos diccionarios: B y E
                Set tempDictB = CreateObject("Scripting.Dictionary")
                Set tempDictE = CreateObject("Scripting.Dictionary")
                dict.Add key, Array(tempDictB, tempDictE)
            End If
            
            ' Agregar valores de columna B (sin repetir)
            If ws.Cells(i, "B").Value <> "" Then
                dict(key)(0)(ws.Cells(i, "B").Value) = 1
            End If
            
            ' Agregar valores de columna E (sin repetir)
            If ws.Cells(i, "E").Value <> "" Then
                dict(key)(1)(ws.Cells(i, "E").Value) = 1
            End If
        End If
    Next i
    
    ' Limpiar columnas F, G, H
    If lastRow >= 4 Then
        ws.Range("F1:H" & lastRow + 100).ClearContents
    End If
    
    ' Escribir resultados
    newRow = 2
    For Each keyVar In dict.Keys
        ' Columna F: Identificador único (columna D)
        ws.Cells(newRow, "F").Value = keyVar
        
        ' Columna G: Valores consolidados de B
        If dict(keyVar)(0).Count > 0 Then
            ws.Cells(newRow, "G").Value = Join(dict(keyVar)(0).Keys, ", ")
        Else
            ws.Cells(newRow, "G").Value = ""
        End If
        
        ' Columna H: Valores consolidados de E (agrupados)
        If dict(keyVar)(1).Count > 0 Then
            ws.Cells(newRow, "H").Value = Join(dict(keyVar)(1).Keys, ", ")
        Else
            ws.Cells(newRow, "H").Value = ""
        End If
        
        newRow = newRow + 1
    Next keyVar
    
    ' Encabezados
    ws.Range("F1").Value = "IDENTIFICADOR (Col D)"
    ws.Range("G1").Value = "VALORES Col B"
    ws.Range("H1").Value = "VALORES Col E"
    
    ' Formato
    ws.Range("F1:H1").Font.Bold = True
    ws.Columns("F:H").AutoFit
    
    ' MsgBox "Consolidación completada en columnas F, G y H" & vbCrLf & _
           "Identificadores únicos: " & dict.Count, vbInformation
End Sub

Sub CompararLibros()
    Dim wb As Workbook
    Dim wsOrigen As Worksheet, wsDestino As Worksheet, wsCopy As Worksheet
    Dim ultimaFilaCopy As Long, ultimaFilaOrigen As Long
    Dim i As Long, j As Long, filaDestino As Long
    Dim encontrado As Boolean
    Dim valorBuscar As String
    
    '--- Asignar libros y hojas
    Set wb = ThisWorkbook
    Set wsOrigen = wb.Sheets("DATA")
    Set wsDestino = wb.Sheets("FILTRO")
    Set wsCopy = wb.Sheets("Codigo Partes")
    
    '--- Últimas filas con datos
    ultimaFilaCopy = wsCopy.Cells(wsCopy.Rows.Count, "M").End(xlUp).Row
    ultimaFilaOrigen = wsOrigen.Cells(wsOrigen.Rows.Count, "G").End(xlUp).Row
    
    '--- Inicializar fila de destino (comienza en 2)
    filaDestino = 2
    
    '--- Recorrer la columna A de origen (Codigos ingresados) empezando en fila 4
    For i = 8 To ultimaFilaCopy
        valorBuscar = Trim(wsCopy.Cells(i, "M").Value)
        encontrado = False
        
        ' Buscar en columna G de origen (DATA)
        For j = 2 To ultimaFilaOrigen
            If InStr(1, wsOrigen.Cells(j, "G").Value, valorBuscar, vbTextCompare) > 0 Then
                ' Coincidencia encontrada -> Copiar datos a FILTRO en filaDestino
                ' Forzar formato de texto en columna B
                wsDestino.Cells(filaDestino, "B").NumberFormat = "@"
                wsDestino.Cells(filaDestino, "B").Value = CStr(wsOrigen.Cells(j, "G").Value)
                
                wsDestino.Cells(filaDestino, "A").Value = wsOrigen.Cells(j, "F").Value
                wsDestino.Cells(filaDestino, "C").Value = wsOrigen.Cells(j, "H").Value
                encontrado = True
                Exit For
            End If
        Next j
        
        ' Si no encontró coincidencia
        If Not encontrado Then
            ' Forzar formato de texto en columna B también para valores no encontrados
            wsDestino.Cells(filaDestino, "B").NumberFormat = "@"
            wsDestino.Cells(filaDestino, "B").Value = CStr(valorBuscar)
            wsDestino.Cells(filaDestino, "A").Value = "No existen valores dentro de SAP"
            wsDestino.Cells(filaDestino, "C").Value = "No existen valores dentro de SAP"
        End If
        
        ' Incrementar la fila de destino para el próximo registro
        filaDestino = filaDestino + 1
    Next i
    
    ' MsgBox "Comparación finalizada.", vbInformation
End Sub

Sub CompararLibros_V2()
    Dim wb As Workbook
    Dim wsOrigen As Worksheet, wsDestino As Worksheet, wsCopy As Worksheet
    Dim ultimaFilaCopy As Long, ultimaFilaOrigen As Long
    Dim i As Long, j As Long, filaDestino As Long
    Dim encontrado As Boolean
    Dim valorBuscar As String
    
    '--- Asignar libros y hojas
    Set wb = ThisWorkbook
    Set wsOrigen = wb.Sheets("DATA")
    Set wsDestino = wb.Sheets("FILTRO")
    Set wsCopy = wb.Sheets("Codigos SAP")
    
    '--- Últimas filas con datos
    ultimaFilaCopy = wsCopy.Cells(wsCopy.Rows.Count, "M").End(xlUp).Row
    ultimaFilaOrigen = wsOrigen.Cells(wsOrigen.Rows.Count, "G").End(xlUp).Row
    
    '--- Inicializar fila de destino (comienza en 2)
    filaDestino = 2
    
    '--- Recorrer la columna A de origen (Codigos ingresados) empezando en fila 4
    For i = 8 To ultimaFilaCopy
        valorBuscar = Trim(wsCopy.Cells(i, "M").Value)
        encontrado = False
        
        ' Buscar en columna G de origen (DATA)
        For j = 2 To ultimaFilaOrigen
            If InStr(1, wsOrigen.Cells(j, "F").Value, valorBuscar, vbTextCompare) > 0 Then
                ' Coincidencia encontrada -> Copiar datos a FILTRO en filaDestino
                ' Forzar formato de texto en columna B
                wsDestino.Cells(filaDestino, "B").NumberFormat = "@"
                wsDestino.Cells(filaDestino, "B").Value = CStr(wsOrigen.Cells(j, "G").Value)
                
                wsDestino.Cells(filaDestino, "A").Value = wsOrigen.Cells(j, "F").Value
                wsDestino.Cells(filaDestino, "C").Value = wsOrigen.Cells(j, "H").Value
                encontrado = True
                Exit For
            End If
        Next j
        
        ' Si no encontró coincidencia
        If Not encontrado Then
            ' Forzar formato de texto en columna B también para valores no encontrados
            wsDestino.Cells(filaDestino, "B").NumberFormat = "@"
            wsDestino.Cells(filaDestino, "B").Value = CStr(valorBuscar)
            wsDestino.Cells(filaDestino, "A").Value = "No existen valores dentro de SAP"
            wsDestino.Cells(filaDestino, "C").Value = "No existen valores dentro de SAP"
        End If
        
        ' Incrementar la fila de destino para el próximo registro
        filaDestino = filaDestino + 1
    Next i
    
    ' MsgBox "Comparación finalizada.", vbInformation
End Sub
Sub Rectificacion()
    Dim wb As Workbook
    Dim wsFiltro As Worksheet, wsRectificacion As Worksheet
    Dim ultimaFilaFiltro As Long
    Dim i As Long, filaDestino As Long
    Dim valorFormateado As String
    Dim datosParaClipboard As String
    Dim contador As Integer
    
    '--- Asignar libros y hojas
    Set wb = ThisWorkbook
    Set wsFiltro = wb.Sheets("FILTRO")
    Set wsRectificacion = wb.Sheets("SEND")
    
    '--- Limpiar hoja RECTIFICACION antes de empezar
    wsRectificacion.Cells.ClearContents
    
    '--- Última fila con datos en FILTRO
    ultimaFilaFiltro = wsFiltro.Cells(wsFiltro.Rows.Count, "A").End(xlUp).Row
    
    '--- Empezar en fila 1 de RECTIFICACION
    filaDestino = 1
    datosParaClipboard = ""
    contador = 0
    
    '--- Recorrer FILTRO y copiar valores donde columna A tenga "No existen valores dentro de SAP"
    For i = 2 To ultimaFilaFiltro
        If wsFiltro.Cells(i, "A").Value = "No existen valores dentro de SAP" Then
            ' Formatear el valor de columna B
            valorFormateado = FormatearRectificacion(wsFiltro.Cells(i, "B").Value)
            
            ' Copiar valor formateado a RECTIFICACION
            wsRectificacion.Cells(filaDestino, "A").Value = valorFormateado
            
            ' Agregar al string para clipboard (separado por saltos de línea)
            If datosParaClipboard <> "" Then
                datosParaClipboard = datosParaClipboard & vbCrLf
            End If
            datosParaClipboard = datosParaClipboard & valorFormateado
            
            filaDestino = filaDestino + 1
            contador = contador + 1
        End If
    Next i
    
    '--- Copiar al portapapeles si hay datos
    If contador > 0 Then
        ' Crear objeto DataObject para acceder al portapapeles
        Dim objData As Object
        Set objData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        
        ' Poner datos en el portapapeles
        objData.SetText datosParaClipboard
        objData.PutInClipboard
        
        ' Liberar objeto
        Set objData = Nothing
        
        ' MsgBox "Rectificación completada. Se copiaron " & contador & " valores a la hoja SEND." & vbCrLf & _
               "Los valores también se han copiado al portapapeles.", vbInformation
               
        Call EjecutarMMBEenSAP
        Call PegarPortapapelesEnDATA
        Call consolidarC
        Call CompararFILTROconDATA

    Else
        MsgBox "No se encontraron valores para rectificar.", vbInformation
    End If
    
    
End Sub

Function FormatearRectificacion(valor As String) As String
    Dim resultado As String
    
    ' Verificar si el valor no está vacío
    If Trim(valor) = "" Then
        FormatearRectificacion = ""
        Exit Function
    End If
    
    ' Eliminar todos los guiones "-"
    resultado = Replace(valor, "-", "")
    
    ' Agregar asteriscos al inicio y al final
    resultado = "*" & resultado & "*"
    
    ' Devolver el resultado formateado
    FormatearRectificacion = resultado
End Function

Sub CompararFILTROconDATA()
    Dim wb As Workbook
    Dim wsOrigen As Worksheet, wsDestino As Worksheet
    Dim ultimaFilaDestino As Long, ultimaFilaOrigen As Long
    Dim i As Long, j As Long
    Dim encontrado As Boolean
    Dim valorBuscar As String, valorOrigen As String
    
    '--- Asignar libros y hojas
    Set wb = ThisWorkbook
    Set wsOrigen = wb.Sheets("DATA")
    Set wsDestino = wb.Sheets("FILTRO")
    
    
    '--- Últimas filas con datos
    ultimaFilaDestino = wsDestino.Cells(wsDestino.Rows.Count, "B").End(xlUp).Row
    ultimaFilaOrigen = wsOrigen.Cells(wsOrigen.Rows.Count, "G").End(xlUp).Row
    
    '--- Recorrer la columna B de FILTRO (destino)
    For i = 2 To ultimaFilaDestino ' Empieza en fila 2, asumiendo encabezado
        valorBuscar = LimpiarTexto(Trim(wsDestino.Cells(i, "B").Value))
        encontrado = False
        
        ' Buscar en columna G de DATA (origen)
        For j = 2 To ultimaFilaOrigen
            valorOrigen = LimpiarTexto(Trim(wsOrigen.Cells(j, "G").Value))
            
            If InStr(1, valorOrigen, valorBuscar, vbTextCompare) > 0 Then
                ' Coincidencia encontrada -> Reemplazar datos en FILTRO
                wsDestino.Cells(i, "A").Value = wsOrigen.Cells(j, "F").Value
                wsDestino.Cells(i, "B").Value = wsOrigen.Cells(j, "G").Value
                wsDestino.Cells(i, "C").Value = wsOrigen.Cells(j, "H").Value
                encontrado = True
                Exit For
            End If
        Next j
        
        ' Si no encontró coincidencia, mantener los valores actuales
        ' (No se hace nada, los valores se quedan como están)
    Next i
    
    ' MsgBox "Comparación FILTRO-DATA finalizada.", vbInformation
End Sub

Function LimpiarTexto(texto As String) As String
    Dim resultado As String
    
    ' Verificar si el texto está vacío
    If Trim(texto) = "" Then
        LimpiarTexto = ""
        Exit Function
    End If
    
    ' Eliminar guiones, espacios y otros caracteres no deseados
    resultado = Replace(texto, "-", "")
    resultado = Replace(resultado, " ", "")
    resultado = Replace(resultado, "_", "")
    resultado = Replace(resultado, ".", "")
    
    ' Devolver el texto limpio
    LimpiarTexto = resultado
End Function

Sub CopiarFiltroACC()
    Dim wb As Workbook
    Dim wsFiltro As Worksheet, wsCC As Worksheet
    Dim ultimaFilaFiltro As Long, ultimaFilaCC As Long
    Dim i As Long, filaDestino As Long
    Dim valorCelda As String
    
    '--- Asignar libros y hojas
    Set wb = ThisWorkbook
    Set wsFiltro = wb.Sheets("FILTRO")
    Set wsCC = wb.Sheets("CC")
    
    '--- Limpiar hoja CC desde la fila 2 hacia abajo
    ultimaFilaCC = wsCC.Cells(wsCC.Rows.Count, "A").End(xlUp).Row
    If ultimaFilaCC >= 2 Then
        wsCC.Range("A2:A" & ultimaFilaCC).ClearContents
    End If
    
    '--- Última fila con datos en FILTRO
    ultimaFilaFiltro = wsFiltro.Cells(wsFiltro.Rows.Count, "A").End(xlUp).Row
    
    '--- Empezar en fila 2 de CC (conservando encabezados si los hay)
    filaDestino = 2
    
    '--- Recorrer FILTRO columna A
    For i = 2 To ultimaFilaFiltro ' Empieza en fila 2, asumiendo encabezado
        valorCelda = Trim(wsFiltro.Cells(i, "A").Value)
        
        ' Copiar solo si NO contiene "No existen valores dentro de SAP"
        If valorCelda <> "No existen valores dentro de SAP" Then
            wsCC.Cells(filaDestino, "A").Value = valorCelda
            filaDestino = filaDestino + 1
        End If
    Next i
    
    '--- Mensaje de confirmación
    Dim contador As Integer
    contador = filaDestino - 2
    If contador > 0 Then
        
        ' MsgBox "Se copiaron " & contador & " valores a la hoja CC.", vbInformation
    Else
        MsgBox "No se encontraron valores para copiar.", vbInformation
    End If
End Sub


'--------------------------------------------
'--------------PROCESO DE STOCK--------------
'--------------------------------------------

Sub ProcesoSAP_MMBE_FIND()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Variables para SAP
    Dim sapGui As Object, applicationSAP As Object, connection As Object, session As Object
    
    ' Variables para Excel
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ultimaFila As Long, fila As Long
    Dim item As String, valorFind As String
    
    ' Especificar libro y hoja exactos
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("CC")
    
    ' Encontrar última fila en columna H (desde H4)
    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Verificar que hay datos (empezando desde H4)
    If ultimaFila < 2 Then
        MsgBox "No hay datos para procesar en la columna A(desde A4)", vbExclamation
        Exit Sub
    End If
    
    ' Conectar a SAP
    If Not ConectarSAP(sapGui, applicationSAP, connection, session) Then
        MsgBox "No se pudo conectar a SAP", vbCritical
        Exit Sub
    End If
    
    ' Recorrer cada item desde H4 hasta el final
    For fila = 2 To ultimaFila
        item = Trim(ws.Cells(fila, "A").Value)
        
        ' Saltar celdas vacías
        If item = "" Then
            ws.Cells(fila, "B").Value = "0"
            GoTo SiguienteItem
        End If
        
        ' Procesar el item en SAP y obtener valor del FIND
        valorFind = ProcesarItemMMBE(session, item)
        
        ' Escribir resultado en columna I
        ws.Cells(fila, "B").Value = valorFind
        
        ' Pequeña pausa entre items
        Esperar 2
        
SiguienteItem:
        DoEvents
    Next fila
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Function ProcesarItemMMBE(session As Object, item As String) As String
    On Error GoTo ErrorHandler
    
    Dim valorFind As String
    valorFind = "0" ' Valor por defecto
    
    ' Ingresar transacción MMBE
    session.findById("wnd[0]/tbar[0]/okcd").Text = "MMBE"
    session.findById("wnd[0]").sendVKey 0
    Esperar 1
    
    ' Limpiar campo e ingresar item
    session.findById("wnd[0]/usr/ctxtMS_MATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtMS_MATNR-LOW").caretPosition = 0
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").Select
    session.findById("wnd[1]/usr/txtV-LOW").Text = "PRACMT"
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 6
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtMS_MATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtMS_MATNR-LOW").caretPosition = 0
    session.findById("wnd[0]/usr/ctxtMS_MATNR-LOW").Text = item
    session.findById("wnd[0]").sendVKey 0
    Esperar 1
    
    ' Ejecutar (F8)
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    Esperar 3
    
    ' Obtener valor del FIND
    valorFind = ObtenerValorFind(session)
    
    ' Volver atrás (F3) para siguiente item
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    Esperar 1
    
    ProcesarItemMMBE = valorFind
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    On Error GoTo 0
    ProcesarItemMMBE = "0"
End Function

Function ObtenerValorFind(session As Object) As String
    On Error GoTo ErrorHandler
    
    Dim findText As String
    findText = "0" ' Valor por defecto
    
    ' Maximizar ventana principal
    session.findById("wnd[0]").maximize
    
    ' Navegar hasta la ventana donde está el campo de búsqueda (wnd[3])
    session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell/shellcont[1]/shell[0]").pressButton "DETAILS"
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem "&DETAIL"
    session.findById("wnd[2]/usr/cntlGRID/shellcont/shell").setCurrentCell 1, "VALUE"
    session.findById("wnd[2]/usr/cntlGRID/shellcont/shell").selectedRows = "1"
    session.findById("wnd[2]/usr/cntlGRID/shellcont/shell").contextMenu
    session.findById("wnd[2]/usr/cntlGRID/shellcont/shell").selectContextMenuItem "&FIND"
    
    ' Leer el valor directamente del campo txtGS_SEARCH-VALUE en wnd[3]
    On Error Resume Next
    If session.findById("wnd[3]/usr/txtGS_SEARCH-VALUE").Exists Then
        session.findById("wnd[3]/usr/txtGS_SEARCH-VALUE").SetFocus
        session.findById("wnd[3]/usr/txtGS_SEARCH-VALUE").caretPosition = 0
        findText = session.findById("wnd[3]/usr/txtGS_SEARCH-VALUE").Text
    End If
    On Error GoTo 0
    
    ' Si está vacío, mantener "0"
    If findText = "" Then findText = "0"
    
    ' Cerrar ventanas en orden inverso
    session.findById("wnd[3]").Close
    session.findById("wnd[2]").Close
    session.findById("wnd[1]").Close
    
    ObtenerValorFind = findText
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    ' Intentar cerrar cualquier ventana emergente en caso de error
    session.findById("wnd[3]").Close
    session.findById("wnd[2]").Close
    session.findById("wnd[1]").Close
    On Error GoTo 0
    ObtenerValorFind = "0"
End Function

Function ConectarSAP(sapGui As Object, applicationSAP As Object, connection As Object, session As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim WshShell As Object
    Dim usuario As String, contrasena As String, extSAP As String, pathSAP As String
    Dim i As Integer
    Dim sesionEncontrada As Boolean
    Dim transactionActual As String
    Dim conn As Object, sess As Object
    Dim conexionExistente As Boolean
    
    ' Obtener credenciales desde la hoja (solo por si son necesarias)
    usuario = ThisWorkbook.Sheets("RUTA SAP").Range("A2").Value
    contrasena = ThisWorkbook.Sheets("RUTA SAP").Range("A1").Value
    extSAP = ThisWorkbook.Sheets("RUTA SAP").Range("A3").Value
    pathSAP = ThisWorkbook.Sheets("RUTA SAP").Range("A4").Value
    
    ' Intentar obtener SAPGUI si ya está abierto
    On Error Resume Next
    Set sapGui = GetObject("SAPGUI")
    On Error GoTo 0
    
    If sapGui Is Nothing Then
        ' Si no está abierto SAPGUI, intentar abrir SAP Logon solo si hay ruta
        If pathSAP <> "" And Dir(pathSAP) <> "" Then
            Set WshShell = CreateObject("WScript.Shell")
            WshShell.Run Chr(34) & pathSAP & Chr(34), 1, False
            
            ' Esperar hasta que SAPGUI esté disponible (máx 30 segundos)
            For i = 1 To 30
                DoEvents
                On Error Resume Next
                Set sapGui = GetObject("SAPGUI")
                On Error GoTo 0
                If Not sapGui Is Nothing Then Exit For
                Application.Wait (Now + TimeValue("00:00:01"))
            Next i
        End If
        
        If sapGui Is Nothing Then
            MsgBox "SAP no está en ejecución. Abre SAP manualmente y vuelve a intentar.", vbInformation
            ConectarSAP = False
            Exit Function
        End If
    End If
    
    ' Obtener motor de scripting
    On Error Resume Next
    Set applicationSAP = sapGui.GetScriptingEngine
    If applicationSAP Is Nothing Then
        MsgBox "No se pudo obtener el motor de scripting de SAP", vbCritical
        ConectarSAP = False
        Exit Function
    End If
    On Error GoTo 0
    
    sesionEncontrada = False
    conexionExistente = False
    
    ' PRIMERA ESTRATEGIA: Buscar sesión ya en pantalla principal (SAP Easy Access)
    For Each conn In applicationSAP.Connections
        For Each sess In conn.Sessions
            On Error Resume Next
            transactionActual = sess.Info.Transaction
            If Err.Number = 0 Then
                ' Verificar si está en pantalla principal (SAP Easy Access)
                If EstaEnPantallaPrincipal(sess) Then
                    Set connection = conn
                    Set session = sess
                    sesionEncontrada = True
                    conexionExistente = True
                    Exit For
                End If
            End If
            On Error GoTo 0
        Next sess
        If sesionEncontrada Then Exit For
    Next conn
    
    ' SEGUNDA ESTRATEGIA: Buscar sesión que esté en login o disponible
    If Not sesionEncontrada Then
        For Each conn In applicationSAP.Connections
            For Each sess In conn.Sessions
                On Error Resume Next
                transactionActual = sess.Info.Transaction
                If Err.Number = 0 Then
                    ' Verificar si está en pantalla de login
                    If EstaEnPantallaLogin(sess) Then
                        Set connection = conn
                        Set session = sess
                        sesionEncontrada = True
                        conexionExistente = True
                        Exit For
                    ' Verificar si es una sesión disponible (SESSION_MANAGER o similar)
                    ElseIf transactionActual = "SESSION_MANAGER" Or transactionActual = "" Or transactionActual = "S000" Then
                        Set connection = conn
                        Set session = sess
                        sesionEncontrada = True
                        conexionExistente = True
                        Exit For
                    End If
                End If
                On Error GoTo 0
            Next sess
            If sesionEncontrada Then Exit For
        Next conn
    End If
    
    ' TERCERA ESTRATEGIA: Si hay conexión pero no sesión disponible, usar credenciales
    If Not sesionEncontrada And extSAP <> "" Then
        ' Buscar si ya existe la conexión a extSAP
        For Each conn In applicationSAP.Connections
            If UCase(conn.Description) = UCase(extSAP) Then
                Set connection = conn
                conexionExistente = True
                Exit For
            End If
        Next conn
        
        ' Si no existe la conexión, crear una nueva
        If Not conexionExistente Then
            On Error Resume Next
            Set connection = applicationSAP.OpenConnection(extSAP, True)
            If connection Is Nothing Then
                MsgBox "No se pudo establecer conexión con: " & extSAP, vbCritical
                ConectarSAP = False
                Exit Function
            End If
            On Error GoTo 0
        End If
        
        ' Usar la primera sesión disponible
        Set session = connection.Sessions(0)
        sesionEncontrada = True
        Esperar 2
        
        ' Solo hacer login si es necesario
        On Error Resume Next
        transactionActual = session.Info.Transaction
        If Err.Number = 0 Then
            If transactionActual = "S000" Or transactionActual = "" Or EstaEnPantallaLogin(session) Then
                ' Verificar credenciales antes de intentar login
                If usuario <> "" And contrasena <> "" Then
                    RealizarLogin session, usuario, contrasena
                Else
                    MsgBox "Se requiere login pero no hay credenciales. Por favor, inicia sesión manualmente.", vbInformation
                    ConectarSAP = False
                    Exit Function
                End If
            End If
        End If
        On Error GoTo 0
    End If
    
    If Not sesionEncontrada Then
        MsgBox "No se encontró ninguna sesión de SAP disponible. " & _
               "Por favor, abre SAP manualmente y asegúrate de estar en la pantalla principal.", vbInformation
        ConectarSAP = False
        Exit Function
    End If
    
    ' Manejar múltiples logins si es necesario
    ManejarMultiplesLogins session
    
    ' Verificar que la sesión es válida
    If Not VerificarSesionValida(session) Then
        MsgBox "No se pudo establecer una sesión válida con SAP", vbCritical
        ConectarSAP = False
        Exit Function
    End If
    
    ConectarSAP = True
    Exit Function
    
ErrorHandler:
    ConectarSAP = False
    MsgBox "Error al conectar con SAP: " & Err.Description & " (Error #" & Err.Number & ")", vbCritical
End Function

' Función auxiliar para verificar si estamos en pantalla principal
Function EstaEnPantallaPrincipal(session As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim transactionActual As String
    Dim screenTitle As String
    
    transactionActual = session.Info.Transaction
    
    ' Verificar si estamos en transacción de Easy Access
    If transactionActual = "SESSION_MANAGER" Or transactionActual = "" Then
        ' Verificar elementos típicos de la pantalla principal
        On Error Resume Next
        screenTitle = session.findById("wnd[0]/titl").Text
        If InStr(1, screenTitle, "SAP Easy Access", vbTextCompare) > 0 Then
            EstaEnPantallaPrincipal = True
            Exit Function
        End If
        
        ' Verificar por la barra de menú de SAP
        If Not session.findById("wnd[0]/mbar/menu[0]/menu[3]", False) Is Nothing Then
            EstaEnPantallaPrincipal = True
            Exit Function
        End If
        
        ' Verificar por campo de transacción
        If Not session.findById("wnd[0]/tbar[0]/okcd", False) Is Nothing Then
            EstaEnPantallaPrincipal = True
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    EstaEnPantallaPrincipal = False
    Exit Function
    
ErrorHandler:
    EstaEnPantallaPrincipal = False
End Function

' Función auxiliar para verificar si estamos en pantalla de login
Function EstaEnPantallaLogin(session As Object) As Boolean
    On Error Resume Next
    
    ' Buscar campos típicos de login
    If Not session.findById("wnd[0]/usr/txtRSYST-BNAME", False) Is Nothing Then
        EstaEnPantallaLogin = True
        Exit Function
    End If
    
    If Not session.findById("wnd[0]/usr/txtP_USER", False) Is Nothing Then
        EstaEnPantallaLogin = True
        Exit Function
    End If
    
    If Not session.findById("wnd[0]/usr/txtMANDT", False) Is Nothing Then
        EstaEnPantallaLogin = True
        Exit Function
    End If
    
    ' Verificar por texto de login
    Dim titleText As String
    titleText = session.findById("wnd[0]/titl").Text
    If InStr(1, titleText, "SAP", vbTextCompare) > 0 And _
       (InStr(1, titleText, "Logon", vbTextCompare) > 0 Or _
        InStr(1, titleText, "Login", vbTextCompare) > 0) Then
        EstaEnPantallaLogin = True
        Exit Function
    End If
    
    EstaEnPantallaLogin = False
End Function

' Función para realizar login solo cuando sea necesario
Sub RealizarLogin(session As Object, usuario As String, contrasena As String)
    On Error Resume Next
    
    ' Método tradicional de login
    If Not session.findById("wnd[0]/usr/txtRSYST-BNAME", False) Is Nothing Then
        session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = usuario
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = contrasena
        session.findById("wnd[0]").sendVKey 0
        Esperar 3
    ' Método para versiones más recientes
    ElseIf Not session.findById("wnd[0]/usr/txtP_USER", False) Is Nothing Then
        session.findById("wnd[0]/usr/txtP_USER").Text = usuario
        session.findById("wnd[0]/usr/txtP_PASSWORD").Text = contrasena
        session.findById("wnd[0]").sendVKey 0
        Esperar 3
    Else
        ' Método genérico
        LoginGenerico session, usuario, contrasena
    End If
End Sub

' Función para manejar múltiples logins de forma robusta
Private Sub ManejarMultiplesLogins(session As Object)
    On Error Resume Next
    
    ' Verificar si existe ventana de múltiple login
    If session.findById("wnd[1]", False) Is Nothing Then Exit Sub
    
    ' Diferentes métodos según la versión de SAP
    If Not session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2", False) Is Nothing Then
        ' Método estándar
        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.createSession
        Esperar 2
    ElseIf Not session.findById("wnd[1]/usr/radSPOP-OPTION1", False) Is Nothing Then
        ' Método alternativo 1
        session.findById("wnd[1]/usr/radSPOP-OPTION1").Select
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    ElseIf Not session.findById("wnd[1]/tbar[0]/btn[0]", False) Is Nothing Then
        ' Simplemente presionar OK
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    ElseIf Not session.findById("wnd[1]/tbar[0]/btn[12]", False) Is Nothing Then
        ' Presionar Enter (btn[12] es typically Enter)
        session.findById("wnd[1]/tbar[0]/btn[12]").press
    End If
    
    On Error GoTo 0
End Sub

' Función para login genérico (busca campos de usuario/contraseña)
Private Sub LoginGenerico(session As Object, usuario As String, contrasena As String)
    On Error Resume Next
    
    ' Buscar diferentes patrones de campos de login
    Dim posiblesCampos(1 To 6) As String
    posiblesCampos(1) = "wnd[0]/usr/txtRSYST-BNAME"
    posiblesCampos(2) = "wnd[0]/usr/txtP_USER"
    posiblesCampos(3) = "wnd[0]/usr/txtUSERNAME"
    posiblesCampos(4) = "wnd[0]/usr/txtMY_USER"
    posiblesCampos(5) = "wnd[0]/usr/txtNAME"
    posiblesCampos(6) = "wnd[0]/usr/txtUSER"
    
    Dim campoUsuario As String, campoPassword As String
    Dim i As Integer
    
    For i = 1 To 6
        If Not session.findById(posiblesCampos(i), False) Is Nothing Then
            campoUsuario = posiblesCampos(i)
            ' Intentar encontrar campo de password correspondiente
            If InStr(campoUsuario, "RSYST-BNAME") > 0 Then
                campoPassword = Replace(campoUsuario, "RSYST-BNAME", "RSYST-BCODE")
            ElseIf InStr(campoUsuario, "P_USER") > 0 Then
                campoPassword = Replace(campoUsuario, "P_USER", "P_PASSWORD")
            ElseIf InStr(campoUsuario, "USERNAME") > 0 Then
                campoPassword = Replace(campoUsuario, "USERNAME", "PASSWORD")
            Else
                campoPassword = Replace(campoUsuario, "USER", "PASSWORD")
            End If
            
            If Not session.findById(campoPassword, False) Is Nothing Then
                session.findById(campoUsuario).Text = usuario
                session.findById(campoPassword).Text = contrasena
                Exit For
            End If
        End If
    Next i
    
    On Error GoTo 0
End Sub

' Función para verificar que la sesión es válida
Private Function VerificarSesionValida(session As Object) As Boolean
    On Error Resume Next
    
    Dim testTransaction As String
    testTransaction = session.Info.Transaction
    
    If Err.Number = 0 Then
        ' Verificar que no estamos en pantalla de login
        If testTransaction = "S000" Or testTransaction = "" Then
            VerificarSesionValida = False
        Else
            VerificarSesionValida = True
        End If
    Else
        VerificarSesionValida = False
    End If
    
    On Error GoTo 0
End Function



Sub clear()

    Dim wsData As Worksheet, wsFiltro As Worksheet
    Dim wsCC As Worksheet, wsRectificacion As Worksheet
    Dim wsSend As Worksheet
    
    Set wsData = ThisWorkbook.Sheets("DATA")
    Set wsFiltro = ThisWorkbook.Sheets("FILTRO")
    Set wsCC = ThisWorkbook.Sheets("CC")
    Set wsRectificacion = ThisWorkbook.Sheets("RECTIFICACION")
    Set wsSend = ThisWorkbook.Sheets("SEND")
    
    wsData.Cells.Range("A1:E10000").clear
    wsData.Cells.Range("F2:H10000").clear
    wsFiltro.Cells.Range("A2:C10000").clear
    wsFiltro.Cells.Range("E2:E10000").clear
    wsCC.Cells.ClearContents
    wsRectificacion.Cells.clear
    wsSend.Cells.clear
    
End Sub

Sub sendStock()
    Dim wsFiltro As Worksheet, wsCC As Worksheet
    Dim ultimaFilaCC As Long, ultimaFilaFiltro As Long
    Dim i As Long, filaDestino As Long
    Dim valorOriginal As String
    Dim valorFormateado As String
    
    Set wsFiltro = ThisWorkbook.Sheets("FILTRO")
    Set wsCC = ThisWorkbook.Sheets("CC")
    
    '--- Últimas filas con datos
    ultimaFilaCC = wsCC.Cells(wsCC.Rows.Count, "B").End(xlUp).Row
    ultimaFilaFiltro = wsFiltro.Cells(wsFiltro.Rows.Count, "C").End(xlUp).Row
    
    '--- Iniciar en fila 2 de ambas hojas
    filaDestino = 2
    
    '--- Recorrer CC columna B y copiar a FILTRO columna F
    For i = 2 To ultimaFilaCC
        ' Verificar si en FILTRO columna C existe "No existen valores dentro de SAP"
        If filaDestino <= ultimaFilaFiltro Then
            If wsFiltro.Cells(filaDestino, "C").Value <> "No existen valores dentro de SAP" Then
                ' Obtener el valor original con formato
                valorOriginal = wsCC.Cells(i, "B").Value
                
                ' Preservar el formato original
                If IsNumeric(valorOriginal) Then
                    ' Si es numérico, formatear con separador de miles
                    valorFormateado = Format(valorOriginal, "#,##0")
                Else
                    ' Si no es numérico, mantener el valor original
                    valorFormateado = valorOriginal
                End If
                
                ' Copiar de CC columna B a FILTRO columna E con formato
                wsFiltro.Cells(filaDestino, "E").Value = valorFormateado
                
                ' Opcional: Aplicar formato de número a la celda
                wsFiltro.Cells(filaDestino, "E").NumberFormat = "#,##0"
                
                filaDestino = filaDestino + 1
            Else
                ' Saltar esta fila en FILTRO y mantener el índice de CC
                filaDestino = filaDestino + 1
                i = i - 1 ' Mantener el mismo valor de CC para la siguiente iteración
            End If
        Else
            ' Si nos quedamos sin filas en FILTRO, salir del loop
            Exit For
        End If
    Next i

End Sub

Sub generarInforme()
    
    'DEFINIR ESTAS VARIABLES SEGÚN TUS NECESIDADES
    Dim nombreHojaOrigen As String
    Dim rutaDestino As String
    Dim nombreLibroDestino As String
    
    'ASIGNAR VALORES A LAS VARIABLES (MODIFICA ESTOS VALORES)
    nombreHojaOrigen = "FILTRO" 'Ejemplo: "Hoja1"
    nombreLibroDestino = "informe de stock" 'Ejemplo: "LibroDestino.xlsx"
    rutaDestino = ThisWorkbook.Sheets("RUTA SAP").Range("A5").Value
    
    
    'Declaración de variables
    Dim wbOrigen As Workbook
    Dim wbDestino As Workbook
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim nombreArchivoCompleto As String
    Dim contador As Integer
    Dim nombreBase As String
    Dim extension As String
    
    'Verificar que las variables estén definidas
    If nombreHojaOrigen = "" Or rutaDestino = "" Or nombreLibroDestino = "" Then
        MsgBox "Por favor, define todas las variables necesarias antes de ejecutar el macro.", vbExclamation
        Exit Sub
    End If
    
    'Asegurar que la ruta termine con \
    If Right(rutaDestino, 1) <> "\" Then
        rutaDestino = rutaDestino & "\"
    End If
    
    'Generar nombre de archivo único
    nombreArchivoCompleto = GenerarNombreUnico(rutaDestino, nombreLibroDestino, "xlsx")
    
    On Error GoTo ErrorHandler
    
    'Usar el libro actual donde se ejecuta la macro
    Set wbOrigen = ThisWorkbook
    Set wsOrigen = wbOrigen.Sheets(nombreHojaOrigen)
    
    'Crear nuevo libro
    Set wbDestino = Workbooks.Add
    Set wsDestino = wbDestino.Sheets(1)
    
    'Encontrar la última fila con datos (usando la columna A como referencia)
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, "A").End(xlUp).Row
    
    'Copiar columnas con formato
    'Columna A -> G
    If ultimaFila >= 1 Then
        wsOrigen.Range("A1:A" & ultimaFila).Copy
        wsDestino.Range("G1").PasteSpecial Paste:=xlPasteAll
    End If
    
    'Columna B -> H
    If ultimaFila >= 1 Then
        wsOrigen.Range("B1:B" & ultimaFila).Copy
        wsDestino.Range("H1").PasteSpecial Paste:=xlPasteAll
    End If
    
    'Columna C -> K
    If ultimaFila >= 1 Then
        wsOrigen.Range("C1:C" & ultimaFila).Copy
        wsDestino.Range("K1").PasteSpecial Paste:=xlPasteAll
    End If
    
    'Columna D -> J
    If ultimaFila >= 1 Then
        wsOrigen.Range("D1:D" & ultimaFila).Copy
        wsDestino.Range("J1").PasteSpecial Paste:=xlPasteAll
    End If
    
    'Columna E -> I
    If ultimaFila >= 1 Then
        wsOrigen.Range("E1:E" & ultimaFila).Copy
        wsDestino.Range("I1").PasteSpecial Paste:=xlPasteAll
    End If
    
    'Limpiar el portapapeles
    Application.CutCopyMode = False
    
    'Ajustar el ancho de las columnas para que se vea el contenido
    wsDestino.Columns("G:K").AutoFit
    
    'Guardar el libro destino con el nombre único
    wbDestino.SaveAs Filename:=nombreArchivoCompleto, FileFormat:=xlOpenXMLWorkbook
    
    'Cerrar el libro destino (opcional - puedes comentar esta línea si quieres dejarlo abierto)
    wbDestino.Close SaveChanges:=False
    
   
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    'Limpiar el portapapeles en caso de error
    Application.CutCopyMode = False
    
    'Cerrar el libro destino si está abierto
    If Not wbDestino Is Nothing Then
        wbDestino.Close SaveChanges:=False
    End If
End Sub

'Función para generar nombres de archivo únicos
Function GenerarNombreUnico(ruta As String, nombreBase As String, extension As String) As String
    Dim nombreCompleto As String
    Dim contador As Integer
    Dim nombreArchivo As String
    
    'Asegurar que la extensión no tenga el punto
    If Left(extension, 1) = "." Then
        extension = Mid(extension, 2)
    End If
    
    'Primer intento con el nombre original
    nombreCompleto = ruta & nombreBase & "." & extension
    
    'Si el archivo no existe, usar este nombre
    If Dir(nombreCompleto) = "" Then
        GenerarNombreUnico = nombreCompleto
        Exit Function
    End If
    
    'Si existe, buscar un nombre único
    contador = 1
    Do
        nombreArchivo = ruta & nombreBase & " (" & contador & ")." & extension
        If Dir(nombreArchivo) = "" Then
            GenerarNombreUnico = nombreArchivo
            Exit Function
        End If
        contador = contador + 1
    Loop While contador < 1000 'Límite de seguridad
    
    'Si llegamos aquí, usar timestamp como último recurso
    nombreArchivo = ruta & nombreBase & "_" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & "." & extension
    GenerarNombreUnico = nombreArchivo
End Function


'--------------------------------------------
'--------------MAIN--------------
'--------------------------------------------

Sub EjecutarTetrapak()
    
    
    Dim respuesta As VbMsgBoxResult
    
    Dim existen As Boolean
    Dim filaEncontrada As Long
    
    Call codigo_enviado(existen, filaEncontrada)
        
    If existen = True Then
        MsgBox "EXISTEN valores en ambas columnas en la fila: " & filaEncontrada, vbInformation
        Exit Sub
    Else
        Call clear
    End If
        
    ' === Paso 1: Copiar con asteriscos ===
    Call CopiarAConAsteriscos

    
    ' === Paso 2: Ejecutar MMBE en SAP ===
    respuesta = MsgBox("¿Deseas ejecutar la transacción MMBE en SAP ahora?", vbYesNo + vbQuestion, "Confirmar ejecución")
    If respuesta = vbYes Then
        Call EjecutarMMBEenSAP
    
    Else
        MsgBox "Ejecución de SAP omitida por el usuario.", vbExclamation
        Exit Sub
    End If
    
    ' === Paso 3: Pegar en DATA ===
    Call PegarPortapapelesEnDATA

  
    Call consolidarC
    Call CompararLibros
    Call Rectificacion
    Call CopiarFiltroACC
    Call ProcesoSAP_MMBE_FIND
    Call sendStock
    Call CopiarDatosFiltroACodigoPartes

End Sub




Sub EjecutarKHS()

    Dim respuesta As VbMsgBoxResult

    Dim existen As Boolean
    Dim filaEncontrada As Long
    
    Call codigo_enviado(existen, filaEncontrada)
        
    If existen = True Then
        MsgBox "EXISTEN valores en ambas columnas en la fila: " & filaEncontrada, vbInformation
        Exit Sub
    Else
        Call clear
    End If

    ' === Paso 1: Copiar con asteriscos ===
    Call CopiarAConAsteriscos
   
    ' === Paso 2: Ejecutar MMBE en SAP ===
    respuesta = MsgBox("¿Deseas ejecutar la transacción MMBE en SAP ahora?", vbYesNo + vbQuestion, "Confirmar ejecución")
    If respuesta = vbYes Then
        Call EjecutarMMBEenSAP
    Else
        MsgBox "Ejecución de SAP omitida por el usuario.", vbExclamation
        Exit Sub
    End If
    
    ' === Paso 3: Pegar en DATA ===
    Call PegarPortapapelesEnDATA

    
    Call consolidarC
    Call CompararLibros
    Call CopiarFiltroACC
    Call ProcesoSAP_MMBE_FIND
    Call sendStock
    Call CopiarDatosFiltroACodigoPartes

End Sub

Sub EjecutarSap()

    Dim respuesta As VbMsgBoxResult

    Dim existen As Boolean
    Dim filaEncontrada As Long
    
    Call codigo_enviado(existen, filaEncontrada)
        
    If existen = True Then
        MsgBox "EXISTEN valores en ambas columnas en la fila: " & filaEncontrada, vbInformation
        Exit Sub
    Else
        Call clear
    End If

    ' === Paso 1: Copiar con asteriscos ===
    Call CopiarA
   
    ' === Paso 2: Ejecutar MMBE en SAP ===
    respuesta = MsgBox("¿Deseas ejecutar la transacción MMBE en SAP ahora?", vbYesNo + vbQuestion, "Confirmar ejecución")
    If respuesta = vbYes Then
        Call EjecutarMMBEenSAPV2
    Else
        MsgBox "Ejecución de SAP omitida por el usuario.", vbExclamation
        Exit Sub
    End If
    
    ' === Paso 3: Pegar en DATA ===
    Call PegarPortapapelesEnDATA

    
    Call consolidarC
    Call CompararLibros_V2
    Call CopiarFiltroACC
    Call ProcesoSAP_MMBE_FIND
    Call sendStock
    Call CopiarDatosFiltroACodigoSap


End Sub




Public Sub codigo_enviado(ByRef existen As Boolean, ByRef filaEncontrada As Long)
    Dim ws As Worksheet
    Dim i As Long
    Dim ultimaFila As Long
    
    Set ws = ThisWorkbook.Sheets("Codigo Partes")
    existen = False
    filaEncontrada = 0
    
    ' Encontrar la última fila con datos
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultimaFila < 3 Then ultimaFila = 3 ' Mínimo hasta fila 3
    
    ' Buscar desde fila 4
    For i = 8 To ultimaFila
        If ws.Cells(i, 1).Value <> "" And ws.Cells(i, 2).Value <> "" Then
            existen = True
            filaEncontrada = i
            Exit For
        End If
    Next i
    
End Sub


Sub CopiarDatosFiltroACodigoPartes()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFilaOrigen As Long
    Dim ultimaFilaDestino As Long
    
    ' Establecer referencia a las hojas de trabajo
    Set wsOrigen = ThisWorkbook.Worksheets("FILTRO")
    Set wsDestino = ThisWorkbook.Worksheets("Codigo Partes")
    
    ' Encontrar la última fila con datos en la hoja FILTRO (columna A)
    ultimaFilaOrigen = wsOrigen.Cells(wsOrigen.Rows.Count, "A").End(xlUp).Row
    
    ' Verificar si hay datos para copiar (excluyendo encabezados)
    If ultimaFilaOrigen <= 1 Then
        MsgBox "No hay datos para copiar en la hoja FILTRO.", vbInformation
        Exit Sub
    End If
    
    ' Encontrar la última fila con datos en la hoja Codigo Partes (columna Q)
    ultimaFilaDestino = wsDestino.Cells(wsDestino.Rows.Count, "Q").End(xlUp).Row
    
    ' Si la última fila es la primera, empezar en la fila 2 (para respetar encabezados)
    ' Si ya hay datos, empezar después del último dato
    If ultimaFilaDestino = 1 Then
        ultimaFilaDestino = 2
    Else
        ultimaFilaDestino = ultimaFilaDestino + 1
    End If
    
    ' Copiar datos desde FILTRO (A:E, fila 2 hasta última fila con datos)
    ' hacia Codigo Partes (Q:U, a partir de la fila determinada)
    wsOrigen.Range("A2:E" & ultimaFilaOrigen).Copy _
        Destination:=wsDestino.Range("Q" & ultimaFilaDestino)
    
    ' Mensaje de confirmación
    Dim cantidadFilas As Long
    cantidadFilas = ultimaFilaOrigen - 1 ' Excluyendo encabezado
    
    ' Limpiar referencias
    Set wsOrigen = Nothing
    Set wsDestino = Nothing
End Sub

Sub CopiarDatosFiltroACodigoSap()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFilaOrigen As Long
    Dim ultimaFilaDestino As Long
    
    ' Establecer referencia a las hojas de trabajo
    Set wsOrigen = ThisWorkbook.Worksheets("FILTRO")
    Set wsDestino = ThisWorkbook.Worksheets("Codigos SAP")
    
    ' Encontrar la última fila con datos en la hoja FILTRO (columna A)
    ultimaFilaOrigen = wsOrigen.Cells(wsOrigen.Rows.Count, "A").End(xlUp).Row
    
    ' Verificar si hay datos para copiar (excluyendo encabezados)
    If ultimaFilaOrigen <= 1 Then
        MsgBox "No hay datos para copiar en la hoja FILTRO.", vbInformation
        Exit Sub
    End If
    
    ' Encontrar la última fila con datos en la hoja Codigo Partes (columna Q)
    ultimaFilaDestino = wsDestino.Cells(wsDestino.Rows.Count, "Q").End(xlUp).Row
    
    ' Si la última fila es la primera, empezar en la fila 2 (para respetar encabezados)
    ' Si ya hay datos, empezar después del último dato
    If ultimaFilaDestino = 1 Then
        ultimaFilaDestino = 2
    Else
        ultimaFilaDestino = ultimaFilaDestino + 1
    End If
    
    ' Copiar datos desde FILTRO (A:E, fila 2 hasta última fila con datos)
    ' hacia Codigo Partes (Q:U, a partir de la fila determinada)
    wsOrigen.Range("A2:E" & ultimaFilaOrigen).Copy _
        Destination:=wsDestino.Range("Q" & ultimaFilaDestino)
    
    ' Mensaje de confirmación
    Dim cantidadFilas As Long
    cantidadFilas = ultimaFilaOrigen - 1 ' Excluyendo encabezado
    
    ' Limpiar referencias
    Set wsOrigen = Nothing
    Set wsDestino = Nothing
End Sub


