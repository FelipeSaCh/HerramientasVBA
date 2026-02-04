Attribute VB_Name = "Módulo1"
Public sapGui As Object
Public applicationSAP As Object
Public connection As Object
Public session As Object
Public WshShell As Object
Public usuario As String
Public contrasena As String
Public pathSAP As String
Public extSAP As String
Public solicitud As String
Public textoEncontrado As Boolean

Sub Extraer_Lineas()

    ' === CONFIGURACIÓN INICIAL ===
    application.ScreenUpdating = False
    application.EnableEvents = False
    application.Calculation = xlCalculationManual

    Dim ws As Worksheet, wsFinal As Worksheet, wsData As Worksheet
    Dim lastRow As Long, monthIndex As Long, colOutput As Long
    Dim months As Variant
    Dim datos As Variant
    Dim registros() As Variant, colores() As Long
    Dim i As Long, cantReg As Long, filaDestino As Long
    Dim temp1, temp2, temp3, temp4, temp5
    Dim tempColor As Long
    Dim outputRow As Long
    Dim monthCol As Long, colCheck As Long
    Dim j As Long, k As Long
    Dim arr() As Long
    
    ' Diccionarios para acumular conteos
    Dim dictTotal As Object
    Set dictTotal = CreateObject("Scripting.Dictionary")
    
    Dim dictMeses As Object
    Set dictMeses = CreateObject("Scripting.Dictionary")
    
    ' Definir hojas de trabajo
    Set ws = ThisWorkbook.Sheets("Planeación de ParadasRev01")
    Set wsFinal = ThisWorkbook.Sheets("Revisiones automaticas general")
    Set wsData = ThisWorkbook.Sheets("Conteo de dias")
    
    ' Array con los nombres de los meses
    months = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", _
                   "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    
    ' Última fila con datos en columna B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    
    ' Limpiar hojas de salida
    wsFinal.Cells.Clear
    wsData.Cells.Clear
    
    ' Columnas iniciales donde están los INI y FIN
    monthCol = 4   ' Columna D ? INI
    colCheck = 5   ' Columna E ? FIN
    colOutput = 1  ' Columna inicial en hoja final
    
    ' Cargar toda la tabla en memoria (más rápido que leer celda a celda)
    datos = ws.Range(ws.Cells(3, 2), ws.Cells(lastRow, 5 + 2 * (UBound(months)))).Value
    
    ' === RECORRER LOS MESES ===
    For monthIndex = LBound(months) To UBound(months)
        
        ' Escribir título del mes en hoja final
        With wsFinal.Range(wsFinal.Cells(1, colOutput), wsFinal.Cells(1, colOutput + 4))
            .Merge
            .Value = months(monthIndex)
            .Font.Bold = True
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(200, 220, 250)
        End With
        
        ' Encabezados de columnas en hoja final
        wsFinal.Cells(2, colOutput).Resize(1, 5).Value = Array("Planta", "Línea", "INI", "FIN", "Conteo")
        
        ' === CONTAR REGISTROS DEL MES ===
        cantReg = 0
        For i = 1 To UBound(datos)
            If (datos(i, monthCol - 1) <> "") Or (datos(i, colCheck - 1) <> "") Then
                cantReg = cantReg + 1
            End If
        Next i
        
        ' === SI HAY REGISTROS, PROCESARLOS ===
        If cantReg > 0 Then
            ReDim registros(1 To cantReg, 1 To 5)
            ReDim colores(1 To cantReg)
            filaDestino = 1
            
            ' Llenar matriz con los registros del mes
            For i = 1 To UBound(datos)
                If (datos(i, monthCol - 1) <> "") Or (datos(i, colCheck - 1) <> "") Then
                    
                    ' Guardar planta y línea
                    registros(filaDestino, 1) = datos(i, 1)
                    registros(filaDestino, 2) = datos(i, 2)
                    
                    ' Determinar el último día real del mes (maneja bisiestos)
                    Dim fechaBase As Date, ultDiaMes As Long
                    fechaBase = DateSerial(Year(Date), monthIndex + 1, 1)
                    ultDiaMes = Day(DateSerial(Year(fechaBase), Month(fechaBase) + 1, 0))
                    
                    ' INI vacío ? asignar día 1
                    If datos(i, monthCol - 1) = "" Then
                        registros(filaDestino, 3) = 1
                    Else
                        registros(filaDestino, 3) = datos(i, monthCol - 1)
                    End If
                    
                    ' FIN vacío ? asignar último día del mes
                    If datos(i, colCheck - 1) = "" Then
                        registros(filaDestino, 4) = ultDiaMes
                    Else
                        registros(filaDestino, 4) = datos(i, colCheck - 1)
                    End If
                    
                    ' Calcular número de días (conteo)
                    If registros(filaDestino, 4) - registros(filaDestino, 3) = 0 Then
                        registros(filaDestino, 5) = 1
                    Else
                        registros(filaDestino, 5) = registros(filaDestino, 4) - registros(filaDestino, 3) + 1
                    End If
                    
                    ' Acumular días en total general
                    If Not dictTotal.Exists(registros(filaDestino, 2)) Then
                        dictTotal.Add registros(filaDestino, 2), registros(filaDestino, 5)
                    Else
                        dictTotal(registros(filaDestino, 2)) = dictTotal(registros(filaDestino, 2)) + registros(filaDestino, 5)
                    End If
                    
                    ' Acumular días en diccionario por mes
                    If Not dictMeses.Exists(registros(filaDestino, 2)) Then
                        ReDim arr(0 To 11)
                        arr(monthIndex) = registros(filaDestino, 5)
                        dictMeses.Add registros(filaDestino, 2), arr
                    Else
                        arr = dictMeses(registros(filaDestino, 2))
                        arr(monthIndex) = arr(monthIndex) + registros(filaDestino, 5)
                        dictMeses(registros(filaDestino, 2)) = arr
                    End If
                    
                    ' Guardar color de la celda original
                    colores(filaDestino) = ws.Cells(i + 2, monthCol).Interior.Color
                    
                    filaDestino = filaDestino + 1
                End If
            Next i
            
            ' === ORDENAR REGISTROS POR INI (menor a mayor) ===
            For j = 1 To cantReg - 1
                For k = j + 1 To cantReg
                    If registros(j, 3) > registros(k, 3) Then
                        temp1 = registros(j, 1): temp2 = registros(j, 2)
                        temp3 = registros(j, 3): temp4 = registros(j, 4): temp5 = registros(j, 5)
                        tempColor = colores(j)
                        
                        registros(j, 1) = registros(k, 1): registros(j, 2) = registros(k, 2)
                        registros(j, 3) = registros(k, 3): registros(j, 4) = registros(k, 4): registros(j, 5) = registros(k, 5)
                        colores(j) = colores(k)
                        
                        registros(k, 1) = temp1: registros(k, 2) = temp2
                        registros(k, 3) = temp3: registros(k, 4) = temp4: registros(k, 5) = temp5
                        colores(k) = tempColor
                    End If
                Next k
            Next j
            
            ' === VOLCAR MATRIZ ORDENADA EN LA HOJA FINAL ===
            outputRow = 3
            wsFinal.Cells(outputRow, colOutput).Resize(cantReg, 5).Value = registros
            
            ' Aplicar colores a cada fila
            For i = 1 To cantReg
                wsFinal.Range(wsFinal.Cells(outputRow + i - 1, colOutput), _
                              wsFinal.Cells(outputRow + i - 1, colOutput + 4)).Interior.Color = colores(i)
            Next i
        End If
        
        ' Avanzar columnas para el siguiente mes
        monthCol = monthCol + 2
        colCheck = colCheck + 2
        colOutput = colOutput + 5
    Next monthIndex
    
    ' === CREAR TABLA RESUMEN EN HOJA wsData ===
    Dim key As Variant
    Dim filaResumen As Long
    
    ' Encabezados
    wsData.Cells(1, 1).Value = "Línea"
    wsData.Cells(1, 2).Value = "Total Días"
    For monthIndex = 0 To 11
        wsData.Cells(1, 3 + monthIndex).Value = months(monthIndex)
    Next monthIndex
    
    wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, 14)).Font.Bold = True
    wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, 14)).Interior.Color = RGB(200, 220, 250)
    
    ' Llenar resumen con los acumulados
    filaResumen = 2
    For Each key In dictTotal.Keys
        wsData.Cells(filaResumen, 1).Value = key
        wsData.Cells(filaResumen, 2).Value = dictTotal(key)
        
        arr = dictMeses(key)
        For monthIndex = 0 To 11
            wsData.Cells(filaResumen, 3 + monthIndex).Value = arr(monthIndex)
        Next monthIndex
        
        filaResumen = filaResumen + 1
    Next key
    
    ' === FORMATO CONDICIONAL AL TOTAL ===
    Dim rngTotal As Range
    Set rngTotal = wsData.Range("B2:B" & filaResumen - 1)
    rngTotal.FormatConditions.Delete
    
    ' Rojo si > 10
    With rngTotal.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="10")
        .Interior.Color = RGB(255, 150, 150)
        .Font.Color = RGB(156, 0, 6)
        .Font.Bold = True
    End With
    
    ' Verde si <= 10
    With rngTotal.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLessEqual, Formula1:="10")
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
        .Font.Bold = True
    End With
    
    ' === RESTAURAR CONFIGURACIÓN Y FINALIZAR ===
    application.ScreenUpdating = True
    application.EnableEvents = True
    application.Calculation = xlCalculationAutomatic
    Call CopiarContenido
    

End Sub


Function CopiarContenido() As Boolean
    Dim wsOrigen As Worksheet, wsDestino As Worksheet
    
    Set wsOrigen = ThisWorkbook.Sheets("Planeación de ParadasRev01")
    Set wsDestino = ThisWorkbook.Sheets("Revisiones automaticas general")
    
    wsOrigen.Range("B47:C50").Copy
    wsDestino.Range("A23").PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    
    application.CutCopyMode = False
    CopiarContenido = True
End Function

Sub GenerarTablas()

    Dim ws As Worksheet, wsOut As Worksheet
    Dim dict As Object, dict2 As Object, dict3 As Object
    Dim header As Variant, val As Variant, val2 As Variant, val3 As Variant
    Dim arr As Variant, arr2 As Variant, arr3 As Variant
    Dim i As Long, j As Long, k As Long, col As Long
    Dim lastRow As Long, startRow As Long
    Dim nombreRango As String
    Dim parts As Variant
    Dim newArr() As String, oldSize As Long, idx As Long
    Dim headers1 As Variant, headers2 As Variant, headers3 As Variant
    Dim headerCount As Long
    
    ' Asignar hojas
    Set ws = ThisWorkbook.Sheets("Planeación de ParadasRev01")
    Set wsOut = ThisWorkbook.Sheets("Revision de estado")
    Set dict = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")
    Set dict3 = CreateObject("Scripting.Dictionary")
    
    ' Limpiar solo el rango especificado
    wsOut.Range("A1:BS100").Clear
    
    ' --- PRIMERA TABLA (col B = encabezados, col C = valores) ---
    ' Primero recolectar TODOS los encabezados únicos
    For i = 3 To 32
        header = Trim(ws.Cells(i, 2).Value) ' Columna B
        If Len(header) > 0 Then
            dict(header) = "" ' Solo para asegurar que existe la clave
        End If
    Next i
    
    ' Ahora procesar los valores para cada encabezado
    For i = 3 To 32
        header = Trim(ws.Cells(i, 2).Value) ' Columna B
        val = Trim(ws.Cells(i, 3).Value)    ' Columna C
        
        ' Normalizar valores
        If Len(val) > 0 Then
            val = Replace(val, " ", "_")
            val = Replace(val, "-", "_")
            val = Replace(val, ".", "_")
        Else
            val = "" ' Mantener valores vacíos
        End If
        
        If Len(header) > 0 Then
            If Not dict.Exists(header) Then
                dict(header) = Array(val)
            Else
                If IsArray(dict(header)) Then
                    arr = dict(header)
                    ReDim Preserve arr(UBound(arr) + 1)
                    arr(UBound(arr)) = val
                    dict(header) = arr
                Else
                    dict(header) = Array(val)
                End If
            End If
        End If
    Next i
    
    ' Escribir PRIMERA TABLA - asegurar que todas las columnas se creen
    col = 1
    headers1 = dict.Keys
    For Each header In headers1
        nombreRango = Replace(header, " ", "_")
        wsOut.Cells(1, col).Value = nombreRango
        
        ' Verificar si hay valores para este encabezado
        If IsArray(dict(header)) Then
            arr = dict(header)
            For j = LBound(arr) To UBound(arr)
                wsOut.Cells(j + 2, col).Value = arr(j)
            Next j
        Else
            ' Si no hay valores, dejar columna vacía pero crear el nombre de rango
            wsOut.Cells(2, col).Value = "" ' Al menos una celda vacía
        End If
        
        lastRow = wsOut.Cells(wsOut.Rows.Count, col).End(xlUp).row
        If lastRow < 2 Then lastRow = 2 ' Asegurar al menos una fila de datos
        
        On Error Resume Next
        ThisWorkbook.Names(nombreRango).Delete
        On Error GoTo 0
        ThisWorkbook.Names.Add Name:=nombreRango, _
            RefersTo:=wsOut.Range(wsOut.Cells(2, col), wsOut.Cells(lastRow, col))
        
        col = col + 1
    Next header
    
 ' --- SEGUNDA TABLA (col C = encabezados, col AE = valores) ---
    ' Primero recolectar TODOS los encabezados únicos
    For i = 3 To 32
        header = Trim(ws.Cells(i, 3).Value) ' Columna C
        If Len(header) > 0 Then
            dict2(header) = "" ' Solo para asegurar que existe la clave
        End If
    Next i
    
    ' Ahora procesar los valores
    For i = 3 To 32
        header = Trim(ws.Cells(i, 3).Value) ' Columna C
        val2 = Trim(ws.Cells(i, 31).Value)  ' Columna AE
        
        If Len(header) > 0 Then
            If Len(val2) > 0 Then
                val2 = Replace(val2, ",", "|")
                val2 = Replace(val2, "-", "|")
                val2 = Replace(val2, "(", "|")
                val2 = Replace(val2, ")", "|")
                val2 = Replace(val2, " ", "|")
                
                parts = Split(val2, "|")
                For j = LBound(parts) To UBound(parts)
                    parts(j) = Trim(parts(j))
                Next j
            Else
                ' Si no hay valor, crear array vacío
                parts = Array("")
            End If
            
            If Not dict2.Exists(header) Then
                dict2(header) = parts
            Else
                If IsArray(dict2(header)) Then
                    arr2 = dict2(header)
                    ' Si el array actual está vacío, reemplazarlo
                    If UBound(arr2) = 0 And arr2(0) = "" Then
                        dict2(header) = parts
                    Else
                        oldSize = UBound(arr2)
                        ReDim Preserve newArr(oldSize + UBound(parts) + 1)
                        For k = 0 To oldSize
                            newArr(k) = arr2(k)
                        Next k
                        For idx = 0 To UBound(parts)
                            newArr(oldSize + 1 + idx) = parts(idx)
                        Next idx
                        dict2(header) = newArr
                    End If
                Else
                    dict2(header) = parts
                End If
            End If
        End If
    Next i
    
    ' Escribir SEGUNDA TABLA desde columna 12 - todas las columnas
    startRow = 1
    col = 12
    headers2 = dict2.Keys
    For Each header In headers2
        nombreRango = Replace(header, " ", "_")
        wsOut.Cells(startRow, col).Value = nombreRango
        
        ' Verificar si hay valores para este encabezado
        If IsArray(dict2(header)) Then
            arr2 = dict2(header)
            For j = LBound(arr2) To UBound(arr2)
                If Len(arr2(j)) > 0 Then ' Solo escribir valores no vacíos
                    wsOut.Cells(startRow + j + 1, col).Value = arr2(j)
                End If
            Next j
        End If
        
        ' Asegurar que haya al menos una celda para el nombre de rango
        If wsOut.Cells(startRow + 1, col).Value = "" Then
            wsOut.Cells(startRow + 1, col).Value = ""
        End If
        
        lastRow = wsOut.Cells(wsOut.Rows.Count, col).End(xlUp).row
        If lastRow <= startRow Then lastRow = startRow + 1
        
        On Error Resume Next
        ThisWorkbook.Names(nombreRango).Delete
        On Error GoTo 0
        ThisWorkbook.Names.Add Name:=nombreRango, _
            RefersTo:=wsOut.Range(wsOut.Cells(startRow + 1, col), wsOut.Cells(lastRow, col))
        
        col = col + 1
    Next header
    

    ' --- TERCERA TABLA (col C = encabezados, col AH = valores) ---
    ' Primero recolectar TODOS los encabezados únicos
    For i = 3 To 32
        header = Trim(ws.Cells(i, 3).Value) ' Columna C
        If Len(header) > 0 Then
            dict3(header) = "" ' Solo para asegurar que existe la clave
        End If
    Next i
    
    ' Ahora procesar los valores
    For i = 3 To 32
        header = Trim(ws.Cells(i, 3).Value) ' Columna C
        val3 = Trim(ws.Cells(i, 32).Value)  ' Columna AH
        
        If Len(header) > 0 Then
            If Len(val3) > 0 Then
                val3 = Replace(val3, ",", "|")
                val3 = Replace(val3, "-", "|")
                val3 = Replace(val3, "(", "|")
                val3 = Replace(val3, ")", "|")
                val3 = Replace(val3, " ", "|")
         
                parts = Split(val3, "|")
                For j = LBound(parts) To UBound(parts)
                    parts(j) = Trim(parts(j))
                Next j
            Else
                ' Si no hay valor, crear array vacío
                parts = Array("")
            End If
            
            If Not dict3.Exists(header) Then
                dict3(header) = parts
            Else
                If IsArray(dict3(header)) Then
                    arr3 = dict3(header)
                    ' Si el array actual está vacío, reemplazarlo
                    If UBound(arr3) = 0 And arr3(0) = "" Then
                        dict3(header) = parts
                    Else
                        oldSize = UBound(arr3)
                        ReDim Preserve newArr(oldSize + UBound(parts) + 1)
                        For k = 0 To oldSize
                            newArr(k) = arr3(k)
                        Next k
                        For idx = 0 To UBound(parts)
                            newArr(oldSize + 1 + idx) = parts(idx)
                        Next idx
                        dict3(header) = newArr
                    End If
                Else
                    dict3(header) = parts
                End If
            End If
        End If
    Next i
    
    ' Escribir TERCERA TABLA desde columna 42 - todas las columnas
    startRow = 1
    col = 42
    headers3 = dict3.Keys
    For Each header In headers3
        nombreRango = Replace(header, " ", "_") & "_T3"
        wsOut.Cells(startRow, col).Value = nombreRango
        
        ' Verificar si hay valores para este encabezado
        If IsArray(dict3(header)) Then
            arr3 = dict3(header)
            For j = LBound(arr3) To UBound(arr3)
                If Len(arr3(j)) > 0 Then ' Solo escribir valores no vacíos
                    wsOut.Cells(startRow + j + 1, col).Value = arr3(j)
                End If
            Next j
        End If
        
        ' Asegurar que haya al menos una celda para el nombre de rango
        If wsOut.Cells(startRow + 1, col).Value = "" Then
            wsOut.Cells(startRow + 1, col).Value = ""
        End If
        
        lastRow = wsOut.Cells(wsOut.Rows.Count, col).End(xlUp).row
        If lastRow <= startRow Then lastRow = startRow + 1
        
        On Error Resume Next
        ThisWorkbook.Names(nombreRango).Delete
        On Error GoTo 0
        ThisWorkbook.Names.Add Name:=nombreRango, _
            RefersTo:=wsOut.Range(wsOut.Cells(startRow + 1, col), wsOut.Cells(lastRow, col))
        
        col = col + 1
    Next header
    
    wsOut.Columns.AutoFit
    
    MsgBox "Datos actualizados. Columnas generadas: " & _
           (dict.Count + dict2.Count + dict3.Count) & " en total.", vbInformation

End Sub

Sub LimpiarCeldas()
    ' Cambia el rango según lo necesites
    With ThisWorkbook.Sheets("Consultas")
        .Range("D3:E100").ClearContents
    End With
End Sub
Sub abrirsap()
    On Error Resume Next ' Preventivo inicial
    application.DisplayAlerts = False
    
    Dim posiciones As Integer
    posiciones = 1
    
    ' Lectura de rutas y credenciales
    usuario = ThisWorkbook.Sheets("RUTA SAP").Range("A2").Value
    contrasena = ThisWorkbook.Sheets("RUTA SAP").Range("A1").Value
    pathSAP = ThisWorkbook.Sheets("RUTA SAP").Range("A4").Value
    solicitud = ThisWorkbook.Sheets("Consultas").Range("A1").Value ' Asegúrate que la variable solicitud exista
    
    Dim rutaBase As String
    rutaBase = ThisWorkbook.Sheets("Consultas").Range("A2").Value
    If Right(rutaBase, 1) <> "\" Then rutaBase = rutaBase & "\"
  
    Dim SapGuiAuto As Object, sapApp As Object, connection As Object, session As Object
    
    ' Abrir SAP Logon
    Call Shell(pathSAP, vbMinimizedFocus)
    application.Wait Now + TimeValue("00:00:08")

    Set sapGui = GetObject("SAPGUI")
    Set Appl = sapGui.GetScriptingEngine
    
    ' Abrir Conexión
    Set connection = Appl.OpenConnection("ALB [Productivo-Nuevo]", True)
    If connection Is Nothing Then
        MsgBox "No se pudo abrir la conexión ALB [Productivo-Nuevo]", vbCritical
        Exit Sub
    End If
    
    Set session = connection.Children(0)
    
    ' Ingresar usuario y contraseña
    session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = usuario
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = contrasena
    session.findById("wnd[0]").sendVKey 0 ' Enter

    ' Manejar Multi-Logon (Punto crítico en otros equipos)
    application.Wait Now + TimeValue("00:00:04")
    
    If Not session.findById("wnd[1]", False) Is Nothing Then
        If Not session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2", False) Is Nothing Then
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select
            session.findById("wnd[1]/tbar[0]/btn[0]").press
        End If
    End If

    ' --- CORRECCIÓN CLAVE: Re-vincular y manejo de nueva sesión ---
    DoEvents
    Set session = connection.Children(0) ' Re-enfocar sesión principal
    
    On Error Resume Next
    session.createSession ' Intentar crear sesión
    application.Wait Now + TimeValue("00:00:03")
    
    ' Si createSession funcionó, usamos la última. Si falló, seguimos en la 0.
    If connection.Sessions.Count > 1 Then
        Set session = connection.Sessions(connection.Sessions.Count - 1)
    Else
        Set session = connection.Children(0)
    End If
    On Error GoTo 0
    
    Set session = connection.Sessions(connection.Sessions.Count - 1)
        ' Abrir transacción ME53N y buscar la solicitud
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "ME53N"
    session.findById("wnd[0]").sendVKey 0
    
    ' Esperar carga de la transacción
    Esperar 3
    
    ' Buscar la solicitud
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-BANFN").Text = solicitud
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressToolbarContextButton "&MB_VARIANT"
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressToolbarButton "&MB_VARIANT"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = -1
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectColumn "VARIANT"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").pressColumnHeader "VARIANT"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").contextMenu
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectContextMenuItem "&FILTER"
    session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "//PRACMT"
    session.findById("wnd[2]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
    
    ' Esperar resultados de búsqueda
    Esperar 3
    
    ' --- NUEVA FUNCIONALIDAD: OBTENER DATOS DIRECTAMENTE DEL GRID ---
    Dim grid As Object
    Dim totalFilas As Integer
    Dim fila As Integer
    Dim textoColumnaG As String
    Dim textoColumnaH As String
    Dim filaDestinoExcel As Integer
    ThisWorkbook.Sheets("Consultas").Range("G3:J500").ClearContents
    ThisWorkbook.Sheets("Consultas").Range("L3:Q500").ClearContents
    Dim moneda As String
    Dim valorTotal As String
    Dim cantidad As String
    Dim centro As String
    Dim fecha As String

    ' Obtener referencia al grid
    
        On Error Resume Next
    Set grid = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")
    On Error GoTo 0
    
    ' Si falló, presiona el botón y reintenta
    If grid Is Nothing Then
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB1:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4000/btnDYN_4000-BUTTON").press
        
        ' Reintento
        On Error Resume Next
        Set grid = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")
        On Error GoTo 0
    End If
    
    ' Validación final (opcional pero recomendado)
    If grid Is Nothing Then
        MsgBox "No fue posible cargar el GRID después del reintento.", vbCritical
        Exit Sub
    End If
    
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT12").Select
    ' Configurar las columnas que nos interesan
    grid.currentCellColumn = "TXZ01"
    grid.selectColumn "TXZ01"
    grid.deselectColumn "EBELN"
    
    ' Obtener el número total de filas en el grid
    totalFilas = grid.rowCount
    filaDestinoExcel = 3 ' Comenzar en la fila 3 (columna G)
    
    ' Limpiar columnas G y H antes de escribir nuevos datos
    ThisWorkbook.Sheets("Consultas").Range("G3:H500").ClearContents
    
    ' Recorrer todas las filas del grid y extraer los datos
    For fila = 0 To totalFilas - 1
        ' Obtener texto de la columna TXZ01 (irá a columna G)
        grid.currentCellRow = fila
        grid.currentCellColumn = "TXZ01"
        textoColumnaG = grid.GetCellValue(fila, "TXZ01")
        
        ' Obtener texto de la columna EBELN (irá a columna H)
        textoColumnaH = grid.GetCellValue(fila, "EBELN")
        textoColumnaO = grid.GetCellValue(fila, "STATUSICON")
        moneda = grid.GetCellValue(fila, "WAERS")
        valorTotal = grid.GetCellValue(fila, "PREIS")
        cantidad = grid.GetCellValue(fila, "MENGE")
        centro = grid.GetCellValue(fila, "WERKS")
        fecha = grid.GetCellValue(fila, "FRGDT")
        ' Escribir en Excel - columna G (TXZ01) y columna H (EBELN)
        ThisWorkbook.Sheets("Consultas").Cells(filaDestinoExcel, 8).Value = textoColumnaG  ' Columna G = 7
        ThisWorkbook.Sheets("Consultas").Cells(filaDestinoExcel, 7).Value = textoColumnaH  ' Columna H = 8
        ThisWorkbook.Sheets("Consultas").Cells(filaDestinoExcel, 20).Value = textoColumnaO  ' Columna L =14
        ThisWorkbook.Sheets("Consultas").Cells(filaDestinoExcel, 13).Value = moneda      ' Columna L
        ThisWorkbook.Sheets("Consultas").Cells(filaDestinoExcel, 14).Value = valorTotal ' Columna M
        ThisWorkbook.Sheets("Consultas").Cells(filaDestinoExcel, 15).Value = cantidad
        ThisWorkbook.Sheets("Consultas").Cells(filaDestinoExcel, 16).Value = centro
        ThisWorkbook.Sheets("Consultas").Cells(filaDestinoExcel, 17).Value = fecha
        

        ' Incrementar fila destino en Excel
        filaDestinoExcel = filaDestinoExcel + 1
    Next fila
    
   
    ' --- FIN DE EXTRACCIÓN DIRECTA DEL GRID ---
    ' --- NUEVA FUNCIONALIDAD: LECTURA DIRECTA DE TEXTO EN COLUMNA ICON ---
    Dim textoBuscar As String
    Dim textoActual As String
    Dim textoAnterior As String
    Dim contadorRegistros As Integer
    Dim j As Integer
    Dim procesoTerminado As Boolean
    Dim textoEncontrado As Boolean
    Dim textoIcon As String
    Dim textoVacio As Boolean
    Dim ultimoTextoDescripcion As String
    Dim penultimoTextoDescripcion As String
    
    textoBuscar = "@5D\QEs posible liberar@"  ' El texto específico que buscas
    contadorRegistros = 0
    textoAnterior = ""
    filaDestinoExcel = 3 ' Comenzar en la fila 3 (celda K3)
    procesoTerminado = False
    
    On Error Resume Next
    ThisWorkbook.Sheets("Consultas").Range("I3:J5000").ClearContents
    ThisWorkbook.Sheets("Consultas").Range("L3:L5000").ClearContents
    ' Loop Do While que se ejecuta hasta que el texto se repita
    i = 1
    Do While Not procesoTerminado
        ' Leemos el texto del combo ANTES de presionar el botón (para la ventana actual)
        textoActual = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST").Text
    
        ' Comparamos con la iteración anterior para detectar fin del ciclo
        If textoActual = textoAnterior Then
            procesoTerminado = True
        Else
            ' Actualizamos la variable de iteración anterior
            textoAnterior = textoActual
            
            ' Obtener el grid de liberaciones de la ventana ACTUAL
            Set grid = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT12/ssubTABSTRIPCONTROL1SUB:SAPLMERELVI:1101/cntlRELEASE_INFO_ITEM/shellcont/shell")
            
            ' --- PROCESO DE LECTURA DIRECTA - LEER TEXTO DE LA COLUMNA ICON ---
            textoEncontrado = False
            textoVacio = True ' Inicialmente asumimos que está vacío
            
            ' --- NUEVA FUNCIONALIDAD: LECTURA DE LA COLUMNA DESCRIPTION ---
            ultimoTextoDescripcion = ""
            penultimoTextoDescripcion = ""
            
            ' Obtener el número total de filas en el grid
            totalFilas = grid.rowCount
            
            ' Recorrer todas las filas de las columnas ICON y DESCRIPTION
            For j = 0 To totalFilas - 1
                ' Leer directamente el texto de la celda en columna ICON, fila j
                On Error Resume Next ' Por si hay errores al leer celdas vacías
                textoIcon = grid.GetCellValue(j, "ICON")
                On Error GoTo 0
                
                ' Verificar si la celda no está vacía
                If textoIcon <> "" Then
                    textoVacio = False ' Hay al menos una celda con texto
                    
                    ' Verificar si el texto buscado está contenido en el texto de la celda
                    If InStr(1, textoIcon, textoBuscar, vbTextCompare) > 0 Then
                        textoEncontrado = True
                        ' Exit For ' Descomenta si quieres optimizar
                    End If
                End If
                
                ' --- LECTURA DE LA COLUMNA DESCRIPTION ---
                ' Configurar la columna DESCRIPTION como la columna actual para lectura
                grid.currentCellColumn = "DESCRIPTION"
                
                ' Leer el texto de la celda en columna DESCRIPTION, fila j
                On Error Resume Next
                Dim textoDescripcion As String
                textoDescripcion = grid.GetCellValue(j, "DESCRIPTION")
                On Error GoTo 0
                
                ' Almacenar el penúltimo y último texto no vacío
                If textoDescripcion <> "" Then
                    penultimoTextoDescripcion = ultimoTextoDescripcion
                    ultimoTextoDescripcion = textoDescripcion
                End If
            Next j
            
            ' *** ESCRIBIR EN EXCEL CON LOS DATOS CORRECTOS ***
            ' Escribir el texto actual en la columna K
            ThisWorkbook.Sheets("Consultas").Cells(filaDestinoExcel, 10).Value = textoActual ' Columna K = 11
            
            ' --- DECISIÓN SOBRE QUÉ ESCRIBIR EN COLUMNA J ---
            If textoVacio Then
                ' Si todas las celdas de ICON están vacías
                With ThisWorkbook.Sheets("Consultas").Cells(filaDestinoExcel, 9)
                    .Value = "Sin liberar"
                    .Interior.Color = RGB(255, 200, 200) ' Rojo claro
                End With
            
            ElseIf textoEncontrado Then
                ' Si se encontró el texto buscado
                With ThisWorkbook.Sheets("Consultas").Cells(filaDestinoExcel, 9)
                    .Value = "Sin liberar"
                    .Interior.Color = RGB(255, 200, 200) ' Rojo claro
                End With
                contadorRegistros = contadorRegistros + 1
            
            Else
                ' Si hay texto pero no coincide con el buscado
                With ThisWorkbook.Sheets("Consultas").Cells(filaDestinoExcel, 9)
                    .Value = "Liberada"
                    .Interior.Color = RGB(200, 255, 200) ' Verde claro
                End With
            End If

            
            ' --- ESCRIBIR EL TEXTO DE DESCRIPTION EN COLUMNA L ---
            ' Preferir el penúltimo texto si el último está vacío
            Dim textoFinalDescripcion As String
            If ultimoTextoDescripcion = "" And penultimoTextoDescripcion <> "" Then
                textoFinalDescripcion = penultimoTextoDescripcion
            Else
                textoFinalDescripcion = ultimoTextoDescripcion
            End If
            
            ' Escribir en columna L (columna 13)
            ThisWorkbook.Sheets("Consultas").Cells(filaDestinoExcel, 12).Value = textoFinalDescripcion
            
            ' Presionar el botón para cambiar a la SIGUIENTE vista/pantalla
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT002").press
            
            ' Incrementar la fila para el próximo registro
            filaDestinoExcel = filaDestinoExcel + 1
            i = i + 1 ' Incrementar contador (opcional, para control)
            
            ' Agregar un pequeño delay para estabilidad
            Esperar 0.3
        End If
        
    Loop ' Continuar hasta que procesoTerminado = True
    
    On Error GoTo 0
        ' Mensaje final resumiendo los resultados
    If contadorRegistros > 0 Then
        
    Else
        ' Verificar si hubor casos de "Sin posibilidad de liberación"
        Dim ultimaFila As Integer
        Dim k As Integer
        Dim sinPosibilidadCount As Integer
        sinPosibilidadCount = 0
        
        ultimaFila = ThisWorkbook.Sheets("Consultas").Cells(ThisWorkbook.Sheets("Consultas").Rows.Count, "J").End(xlUp).row
        For k = 3 To ultimaFila
            If ThisWorkbook.Sheets("Consultas").Cells(k, 10).Value = "Sin posibilidad de liberación" Then
                sinPosibilidadCount = sinPosibilidadCount + 1
            End If
        Next k
        
    End If
        ' --- NUEVA FUNCIONALIDAD: REVISAR FILAS G VACÍAS CON J = "LIBERADA" ---

    
    ' --- VERIFICAR SI HAY ÓRDENES SIN LIBERAR ---
    Dim hayOrdenesSinLiberar As Boolean
    hayOrdenesSinLiberar = False
    
    ' Buscar "Con posibilidad de liberación" en la columna J de la hoja actual
    ultimaFilaRevisar = ThisWorkbook.Sheets("Consultas").Cells(ThisWorkbook.Sheets("Consultas").Rows.Count, "I").End(xlUp).row
    For k = 3 To ultimaFilaRevisar
        If ThisWorkbook.Sheets("Consultas").Cells(k, "I").Value = "Con posibilidad de liberación" Then
            hayOrdenesSinLiberar = True
            Exit For
        End If
    Next k
    
    ' Mensaje final según el resultado
    ' --- FIN DE NUEVAS FUNCIONALIDADES ---
    
    ' Cerrar sesión
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    Call ProcesarPedidosSAP
    ' Limpiar objetos

End Sub


Sub ProcesarPedidosSAP()
    Dim wsConsultas As Worksheet
    Dim wsPedidos As Worksheet
    Dim pedidosUnicos As Collection
    Dim tieneDatos As Boolean
    Dim i As Long
    On Error GoTo ErrorHandler
    
    ' Verificar hojas

    Set wsConsultas = ThisWorkbook.Worksheets("Consultas")
    Set wsPedidos = ThisWorkbook.Worksheets("Pedidos")

 
 
    ' Verificar datos en primeras 10 filas
    tieneDatos = False
    For i = 3 To 12
        If Trim(wsConsultas.Cells(i, "G").Value) <> "" Then
            tieneDatos = True
            Exit For
        End If
    Next i
    
    If Not tieneDatos Then
        MsgBox "No existen pedidos", vbInformation
        Exit Sub
    End If
    
    ' Obtener pedidos únicos
    Set pedidosUnicos = ObtenerPedidosUnicos(wsConsultas)
    
    If pedidosUnicos.Count = 0 Then
        Exit Sub
    End If
    
    ' Guardar pedidos en hoja Pedidos
    GuardarPedidosEnHoja wsPedidos, pedidosUnicos
    
    ' Ejecutar SAP con todos los pedidos
    EjecutarSAPConMultiplesPedidos pedidosUnicos
    
    ' Limpiar
    Set pedidosUnicos = Nothing
    Exit Sub
    
ErrorHandler:
 
    ' Continuar sin mostrar error
    Set pedidosUnicos = Nothing
End Sub


' ---------------------------
' FUNCIÓN PARA OBTENER PEDIDOS ÚNICOS
' ---------------------------
Function ObtenerPedidosUnicos(wsConsultas As Worksheet) As Collection
    Dim dict As Object
    Dim pedidosUnicos As New Collection
    Dim lastRow As Long
    Dim i As Long
    Dim valor As String
    Dim key As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Obtener última fila
    lastRow = wsConsultas.Cells(wsConsultas.Rows.Count, "G").End(xlUp).row
    
    If lastRow < 3 Then
        Set ObtenerPedidosUnicos = pedidosUnicos
        Exit Function
    End If
    
    ' Recorrer columna G para obtener valores únicos
    For i = 3 To lastRow
        valor = Trim(wsConsultas.Cells(i, "G").Value)
        
        If valor <> "" Then
            key = CStr(valor)
            
            If Not dict.Exists(key) Then
                dict.Add key, valor
            End If
        End If
    Next i
    
    ' Pasar valores del diccionario a la colección
    Dim dictItem As Variant
    For Each dictItem In dict.Items
        On Error Resume Next ' Ignorar errores de duplicados
        pedidosUnicos.Add dictItem, CStr(dictItem)
        On Error GoTo 0
    Next dictItem
    
    Set dict = Nothing
    Set ObtenerPedidosUnicos = pedidosUnicos
End Function

' ---------------------------
' PROCEDIMIENTO PARA GUARDAR PEDIDOS EN HOJA
' ---------------------------
Sub GuardarPedidosEnHoja(wsPedidos As Worksheet, pedidosUnicos As Collection)
    Dim i As Long
    
    ' Limpiar columna A
    wsPedidos.Columns("A").ClearContents
    
    ' Escribir encabezado
    wsPedidos.Cells(1, "A").Value = "Pedidos Únicos"
    wsPedidos.Cells(1, "A").Font.Bold = True
    
    ' Escribir pedidos
    For i = 1 To pedidosUnicos.Count
        wsPedidos.Cells(i + 1, "A").Value = pedidosUnicos(i)
    Next i
    
    ' Ordenar y formatear
    If pedidosUnicos.Count > 0 Then
        wsPedidos.Range("A2:A" & pedidosUnicos.Count + 1).Sort _
            key1:=wsPedidos.Range("A2"), _
            Order1:=xlAscending, _
            header:=xlNo
        wsPedidos.Columns("A").AutoFit
    End If
End Sub

' ---------------------------
' PROCEDIMIENTO PARA EJECUTAR SAP CON MÚLTIPLES PEDIDOS
' ---------------------------
Sub EjecutarSAPConMultiplesPedidos(pedidosUnicos As Collection)
    Dim sapGui As Object
    Dim appSAP As Object
    Dim conn As Object
    Dim session As Object
    Dim i As Integer
    Dim j As Integer
    Dim sesionEncontrada As Boolean
    Dim esRelease770 As Boolean
    Dim WshShell As Object
    
    On Error GoTo ErrorHandlerSAP
    
    ' Intentar obtener SAPGUI
    On Error Resume Next
    Set sapGui = GetObject("SAPGUI")
    On Error GoTo 0
    
    If sapGui Is Nothing Then
        ' Abrir SAP Logon
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.Run Chr(34) & pathSAP & Chr(34), 1, False
        
        ' Esperar hasta que SAPGUI esté disponible
        For i = 1 To 30
            DoEvents
            On Error Resume Next
            Set sapGui = GetObject("SAPGUI")
            On Error GoTo 0
            If Not sapGui Is Nothing Then Exit For
            Esperar2 1
        Next i
        
        If sapGui Is Nothing Then
            Exit Sub
        End If
    End If
    
    ' Obtener motor de scripting
    Set appSAP = sapGui.GetScriptingEngine
    sesionEncontrada = False
    
    ' Buscar sesión activa
    For Each conn In appSAP.Children
        If InStr(1, conn.Description, extSAP, vbTextCompare) > 0 Then
            For Each session In conn.Children
                If session.Info.Transaction = "SESSION_MANAGER" Or _
                   session.Info.Transaction = "" Then
                    sesionEncontrada = True
                    Exit For
                End If
            Next
            If sesionEncontrada Then Exit For
        End If
    Next
    
    ' Si no hay sesión activa, crear una
    If Not sesionEncontrada Then
        Set conn = appSAP.OpenConnection(extSAP, True)
        Set session = conn.Sessions(0)
        
        ' Loguearse
        On Error Resume Next
        If session.findById("wnd[0]/usr/txtRSYST-BNAME", False).Exists Then
            esRelease770 = False
            session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = usuario
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = contrasena
            session.findById("wnd[0]").sendVKey 0
        ElseIf session.findById("wnd[0]/usr/txtPASSWORD_FIELD", False).Exists Then
            esRelease770 = True
            session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "100"
            session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = usuario
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = contrasena
            session.findById("wnd[0]").sendVKey 0
        Else
            session.findById("wnd[0]/usr/txtRSYST-BNAME", False).Text = usuario
            session.findById("wnd[0]/usr/pwdRSYST-BCODE", False).Text = contrasena
            session.findById("wnd[0]").sendVKey 0
        End If
        On Error GoTo 0
        
        Esperar2 3
        
        ' Manejar multi login
        On Error Resume Next
        If session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2", False).Exists Then
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.createSession
        ElseIf session.findById("wnd[1]/usr/radSPOPLI-OPTION1", False).Exists Then
            session.findById("wnd[1]/usr/radSPOPLI-OPTION1").Select
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.createSession
        End If
        On Error GoTo 0
        
        Esperar2 2
        Set session = conn.Sessions(conn.Sessions.Count - 1)
    End If
    
    ' Verificar pantalla principal
    Do While session.Info.Transaction = "S000" Or session.Info.Transaction = ""
        Esperar 1
    Loop
    
    ' -------------------------
    ' EJECUTAR ME2N CON MÚLTIPLES PEDIDOS
    ' -------------------------
    session.findById("wnd[0]").maximize
    
    ' Ingresar transacción ME2N
    session.findById("wnd[0]/tbar[0]/okcd").Text = "ME2N"
    session.findById("wnd[0]").sendVKey 0
    
    ' Seleccionar opción de menú
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").Select
    
    ' Ingresar PRACMT
    session.findById("wnd[1]/usr/txtV-LOW").Text = "PRACMT"
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 6
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' Hacer clic en botón de selección múltiple
    session.findById("wnd[0]/usr/btn%_EN_EBELN_%_APP_%-VALU_PUSH").press
    Esperar2 1
    
    ' Ingresar pedidos en la tabla de selección múltiple
    For i = 1 To pedidosUnicos.Count
        ' Calcular fila en la tabla (base 0)
        Dim fila As Integer
        fila = i - 1
        
        ' Construir el ID del campo según la fila
        Dim fieldId As String
        fieldId = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & fila & "]"
        
        ' Ingresar el pedido
        On Error Resume Next
        session.findById(fieldId).Text = pedidosUnicos(i)
        On Error GoTo 0
        
        ' Pequeña pausa entre cada entrada
        Esperar2 0.5
    Next i
    
    ' Si hay más de un pedido, establecer foco en el último
    If pedidosUnicos.Count > 0 Then
        Dim lastFieldId As String
        lastFieldId = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGle/ctxtRSCSEL_255-SLOW_I[1," & (pedidosUnicos.Count - 1) & "]"
    End If
    
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[23]").press
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, "BNFPO"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "WERKS"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "SUPERFIELD"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "EMATN"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "TXZ01"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "EINDT"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "MENGE"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "WEMNG"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "MGLIEF"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "MEINS"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "NETPR"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "WAERS"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "EBELN"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "EBELP"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "NETWR"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "FRGZU"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "BANFN"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "BNFPO"
    session.findById("wnd[0]/tbar[1]/btn[45]").press
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
     
    ' ESPERAR MÁS TIEMPO PARA QUE SAP GENERE EL REPORTE
    Esperar2 5
    
    ' Salir de SAP
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/btn[15]").press ' F3
    On Error GoTo 0
    
    ' ESPERAR ANTES DE COPIAR
    Esperar2 3
    
    ' Limpiar objetos SAP
    Set session = Nothing
    Set conn = Nothing
    Set appSAP = Nothing
    Set sapGui = Nothing
    
    ' LLAMAR AL PROCEDIMIENTO PARA PEGAR DESPUÉS DE CERRAR SAP
    ' Esperar un momento adicional para asegurar que los datos estén en portapapeles
    Esperar2 2
    
    ' Intentar pegar los datos
    Call limpiarinforme
    
    Exit Sub
    
ErrorHandlerSAP:
    ' Manejar error silenciosamente
    On Error Resume Next
    If Not session Is Nothing Then
        session.findById("wnd[0]/tbar[0]/btn[15]").press
    End If
        
    Set session = Nothing
    Set conn = Nothing
    Set appSAP = Nothing
    Set sapGui = Nothing
    
    ' Intentar pegar de todas formas (puede haber datos en portapapeles)
  
    
End Sub

' ---------------------------
' PROCEDIMIENTO PARA LIMPIAR INFORME
' ---------------------------
Sub limpiarinforme()
    Dim wsInforme As Worksheet
    
    On Error Resume Next
    Set wsInforme = ThisWorkbook.Worksheets("Informe de pedidos")
    If wsInforme Is Nothing Then
        Set wsInforme = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsInforme.Name = "Informe de pedidos"
    End If
    On Error GoTo 0
    
    ' Limpiar hoja sin usar Select
    wsInforme.Cells.Clear
    wsInforme.Cells.ClearFormats
    
    Call PegarPortapapelesEnInforme
End Sub

' ---------------------------
' VERSIÓN CORREGIDA - SIN USAR SELECT
' ---------------------------
Sub PegarPortapapelesEnInforme()
    Dim wsInforme As Worksheet
    Dim ws As Worksheet
    
    ' Crear/obtener hoja
    On Error Resume Next
    Set wsInforme = ThisWorkbook.Worksheets("Informe de pedidos")
    If wsInforme Is Nothing Then
        Set wsInforme = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsInforme.Name = "Informe de pedidos"
    End If
    On Error GoTo 0
    
    ' Limpiar hoja
    wsInforme.Cells.ClearContents
    wsInforme.Cells.ClearFormats
    
    ' Activar hoja y seleccionar A1
    wsInforme.Activate
    wsInforme.Range("A1").Select
    
    ' Pausa para asegurar activación
    DoEvents
    Esperar 1
    
    ' Intentar pegar directamente - Excel generalmente maneja esto bien
    On Error Resume Next
    wsInforme.Paste
    If Err.Number <> 0 Then
        ' Si falla, intentar pegado especial
        Err.Clear
        wsInforme.Range("A1").PasteSpecial Paste:=xlPasteValues
    End If
    On Error GoTo 0
    
    ' Autoajustar columnas
    wsInforme.Columns.AutoFit

    
    ' Llamar al siguiente procedimiento
    Call CompararYActualizarPedidos
End Sub

' ---------------------------
' PROCEDIMIENTO DE ESPERA MEJORADO
' ---------------------------
Sub Esperar2(segundos As Single)
    Dim inicio As Single
    inicio = Timer
    Do While Timer < inicio + segundos
        DoEvents
    Loop
End Sub

' Función para esperar segundos
Sub Esperar(segundos As Integer)
    application.Wait (Now + TimeValue("00:00:" & Format(segundos, "00")))
End Sub

' Función para esperar que un archivo exista
Function EsperarArchivo(rutaArchivo As String, maxEsperaSegundos As Integer) As Boolean
    Dim i As Integer
    EsperarArchivo = False
    
    For i = 1 To maxEsperaSegundos
        If Dir(rutaArchivo) <> "" Then
            ' Archivo encontrado, esperar 1 segundo adicional para asegurar que esté completo
            Esperar 1
            EsperarArchivo = True
            Exit Function
        End If
        Esperar 1
    Next i
    
    MsgBox "No se pudo encontrar el archivo: " & vbCrLf & rutaArchivo, vbExclamation
End Function

' ==========================
' OBTENER SOLICITUD (modificado para columnas D, E y F3)
' ==========================
Sub obtenerSolicitud()
    Dim ws As Worksheet
    Dim ultimaFilaD As Long, ultimaFilaE As Long
    Dim i As Long
    Dim contadorD As Long, contadorE As Long, contadorF As Long, contadorTotal As Long
    Dim regex As Object, matches As Object
    Dim valorCelda As String
    
    Set ws = ThisWorkbook.Sheets("Consultas")
    ultimaFilaD = ws.Cells(ws.Rows.Count, "D").End(xlUp).row
    ultimaFilaE = ws.Cells(ws.Rows.Count, "E").End(xlUp).row
    solicitud = ""
    contadorD = 0
    contadorE = 0
    contadorF = 0
    contadorTotal = 0
    
    ' Crear objeto RegExp
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    ' Busca bloques de 6 a 20 dígitos (ajusta según tu caso)
    regex.Pattern = "\d{6,20}"
    
    ' Buscar en columna D
    For i = 3 To ultimaFilaD
        valorCelda = Trim(ws.Cells(i, "D").Value)
        If valorCelda <> "" Then
            ' Extraer solo los números de la celda
            Set matches = regex.Execute(valorCelda)
            If matches.Count > 0 Then
                contadorD = contadorD + 1
                ' Solo guardar la primera solicitud encontrada en D
                If contadorD = 1 Then
                    solicitud = CStr(matches(0).Value) ' Primer número encontrado, como texto
                End If
            End If
        End If
    Next i
    
    ' Buscar en columna E
    For i = 3 To ultimaFilaE
        valorCelda = Trim(ws.Cells(i, "E").Value)
        If valorCelda <> "" Then
            ' Extraer solo los números de la celda
            Set matches = regex.Execute(valorCelda)
            If matches.Count > 0 Then
                contadorE = contadorE + 1
                ' Solo guardar si no se encontró ya uno en columna D
                If contadorD = 0 And contadorE = 1 Then
                    solicitud = CStr(matches(0).Value) ' Primer número encontrado, como texto
                End If
            End If
        End If
    Next i
    
    ' Buscar SOLO en celda F3 (no en toda la columna F)
    valorCelda = Trim(ws.Range("F3").Value)
    If valorCelda <> "" Then
        ' Extraer solo los números de la celda
        Set matches = regex.Execute(valorCelda)
        If matches.Count > 0 Then
            contadorF = contadorF + 1
            ' Solo guardar si no se encontró ya uno en columnas D o E
            If contadorD = 0 And contadorE = 0 And contadorF = 1 Then
                solicitud = CStr(matches(0).Value) ' Primer número encontrado, como texto
            End If
        End If
    End If
    
    contadorTotal = contadorD + contadorE + contadorF
    
    ' Lógica de validación mejorada
    If contadorTotal = 0 Then
        MsgBox "No se encontró ninguna solicitud en las columnas D, E ni en F3.", vbExclamation
        solicitud = ""
        Exit Sub
    ElseIf contadorTotal > 1 Then
        ' Verificar combinaciones específicas
        If contadorD >= 1 And contadorE >= 1 Then
            MsgBox "Se encontraron solicitudes en las columnas D y E. Solo debe haber una solicitud en D, E o F3.", vbCritical
        ElseIf contadorD >= 1 And contadorF >= 1 Then
            MsgBox "Se encontraron solicitudes en la columna D y en F3. Solo debe haber una solicitud en D, E o F3.", vbCritical
        ElseIf contadorE >= 1 And contadorF >= 1 Then
            MsgBox "Se encontraron solicitudes en la columna E y en F3. Solo debe haber una solicitud en D, E o F3.", vbCritical
        ElseIf contadorD > 1 Then
            MsgBox "Se encontró más de una solicitud en la columna D. Solo debe haber una solicitud.", vbCritical
        ElseIf contadorE > 1 Then
            MsgBox "Se encontró más de una solicitud en la columna E. Solo debe haber una solicitud.", vbCritical
        Else
            MsgBox "Se encontró más de una solicitud. Solo debe haber una solicitud en D, E o F3.", vbCritical
        End If
        solicitud = ""
        Exit Sub
    End If
    
    ' Si llegamos aquí, hay exactamente una solicitud válida
End Sub

' ==========================
' MAIN (sin cambios)
' ==========================
Sub main()
    Call obtenerSolicitud
    If solicitud <> "" Then
        Call abrirsap
    End If
End Sub

Sub Actualizar()
    application.ScreenUpdating = False
    application.Calculation = xlCalculationManual
    application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Paso 1: Aplicar fórmulas y colores
    Call AplicarFormulaYColores
    
    ' Paso 2: Iniciar actualización automática
    ' Call AutoUpdateOnColorChange
    
    application.ScreenUpdating = True
    application.Calculation = xlCalculationAutomatic
    application.EnableEvents = True
      
    Exit Sub
    
ErrorHandler:
    application.ScreenUpdating = True
    application.Calculation = xlCalculationAutomatic
    application.EnableEvents = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Sub AplicarFormulaYColores()
    Dim wsDestino As Worksheet
    Dim i As Long
    
    Set wsDestino = ThisWorkbook.Sheets("Cronograma")
    
    application.ScreenUpdating = False
    application.Calculation = xlCalculationManual
    
    ' Aplicar fórmulas en AC (suma diferencias)
    For i = 4 To 34
        wsDestino.Range("AC" & i).Formula = "=SumaDiferenciasConColor(A" & i & ":AB" & i & ")"
    Next i
    
    ' Aplicar fórmulas en AD (contar sin color) - ¡NUEVA LÍNEA!
    For i = 4 To 34
        wsDestino.Range("AD" & i).Formula = "=ContarSinColor(E" & i & ":AB" & i & ")"
    Next i
    
    ' Aplicar colores
    Call AplicarColoresProporcionales
    
    application.ScreenUpdating = True
    application.Calculation = xlCalculationAutomatic
    ThisWorkbook.RefreshAll
    
    
End Sub

Function SumaDiferenciasConColor(rangoFila As Range) As Double
    Dim suma As Double
    Dim j As Long
    Dim colorReferencia As Long
    Dim wsReferencia As Worksheet
    Dim celdaReferencia As Range
    
    ' Referencia a la hoja "Planeación de ParadasRev01"
    Set wsReferencia = ThisWorkbook.Sheets("Planeación de ParadasRev01")
    Set celdaReferencia = wsReferencia.Range("B48")
    colorReferencia = celdaReferencia.Interior.Color
    
    On Error GoTo ErrorHandler
    
    suma = 0
    
    ' Calcular las diferencias: (F-E) + (H-G) + (J-I) + ... + (AB-AA)
    For j = 6 To 28 Step 2 ' Columnas F, H, J, L, ..., AB (6,8,10,...,28)
        ' Verificar colores y calcular diferencia
        If rangoFila.Cells(1, j).Interior.Color = colorReferencia Then
            ' Celda par tiene el color, cuenta como 0
            If rangoFila.Cells(1, j - 1).Interior.Color <> colorReferencia Then
                suma = suma - rangoFila.Cells(1, j - 1).Value
            End If
        ElseIf rangoFila.Cells(1, j - 1).Interior.Color = colorReferencia Then
            ' Celda impar tiene el color, cuenta solo la par
            suma = suma + rangoFila.Cells(1, j).Value
        Else
            ' Ninguna tiene el color, resta normal
            suma = suma + (rangoFila.Cells(1, j).Value - rangoFila.Cells(1, j - 1).Value)
        End If
    Next j
    
    SumaDiferenciasConColor = suma
    Exit Function
    
ErrorHandler:
    SumaDiferenciasConColor = 0
    
End Function

Function CV(celda As Range) As Variant
    Dim colorReferencia As Long
    Dim wsReferencia As Worksheet
    
    On Error GoTo ErrorHandler
    
    ' Referencia a la hoja "Planeación de ParadasRev01"
    Set wsReferencia = ThisWorkbook.Sheets("Planeación de ParadasRev01")
    colorReferencia = wsReferencia.Range("B48").Interior.Color
    
    ' Verificar si la celda tiene el color de referencia
    If celda.Interior.Color = colorReferencia Then
        CV = 0
    Else
        CV = celda.Value
    End If
    
    Exit Function
    
ErrorHandler:
    CV = celda.Value
End Function

' Sub AutoUpdateOnColorChange()
    ' Programa la actualización automática cada 5 segundos
    ' Application.OnTime Now + TimeValue("00:00:05"), "ForzarActualizacion"
' End Sub

Sub ForzarActualizacion()
    ' Actualizar cálculo y volver a programar
    ThisWorkbook.RefreshAll
    application.Calculate
    Call AplicarColoresProporcionales ' Actualizar colores también
    Call AutoUpdateOnColorChange ' Volver a programar
End Sub

Sub AplicarColoresProporcionales()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim celdaOrigen As Range
    Dim celdaDestino As Range
    Dim valorOrigen As String
    Dim valorDestino As String
    Dim filaOrigen As Long
    Dim filaDestino As Long
    Dim columnaOrigen As Long
    
    ' Configurar las hojas (ORIGEN: Planeación, DESTINO: Consultas)
    Set wsOrigen = ThisWorkbook.Sheets("Planeación de ParadasRev01")
    Set wsDestino = ThisWorkbook.Sheets("Cronograma")
    
    application.ScreenUpdating = False
    application.Calculation = xlCalculationManual
    
    ' Recorrer cada fila en el destino (Consultas)
    For Each celdaDestino In wsDestino.Range("D4:D" & wsDestino.Cells(wsDestino.Rows.Count, "D").End(xlUp).row)
        If Not IsEmpty(celdaDestino.Value) Then
            valorDestino = CStr(celdaDestino.Value)
            
            ' Buscar coincidencia en el origen (Planeación)
            For Each celdaOrigen In wsOrigen.Range("C3:C" & wsOrigen.Cells(wsOrigen.Rows.Count, "C").End(xlUp).row)
                If Not IsEmpty(celdaOrigen.Value) Then
                    valorOrigen = CStr(celdaOrigen.Value)
                    
                    If valorOrigen = valorDestino Then
                        filaOrigen = celdaOrigen.row
                        filaDestino = celdaDestino.row
                        
                        ' Aplicar colores proporcionales desde columna D/AA a E/AB
                        For columnaOrigen = 4 To 27 ' Columnas D a AA (4 a 27)
                            ' Obtener el color de la celda de origen
                            Dim colorOrigen As Long
                            colorOrigen = wsOrigen.Cells(filaOrigen, columnaOrigen).Interior.Color
                            
                            ' Aplicar el color a la celda de destino
                            wsDestino.Cells(filaDestino, columnaOrigen + 1).Interior.Color = colorOrigen
                        Next columnaOrigen
                        
                        Exit For
                    End If
                End If
            Next celdaOrigen
        End If
    Next celdaDestino
    
    application.ScreenUpdating = True
    application.Calculation = xlCalculationAutomatic
End Sub

Function ContarSinColor(rango As Range) As Long
    Dim celda As Range
    Dim contador As Long
    Dim colorReferencia As Long
    Dim wsRef As Worksheet
    Dim valorCelda As Variant
    
    On Error GoTo ErrorHandler
    
    Set wsRef = ThisWorkbook.Sheets("Planeación de ParadasRev01")
    colorReferencia = wsRef.Range("B48").Interior.Color
    
    contador = 0
    For Each celda In rango
        ' Manejar errores específicos de BUSCARV
        If IsError(celda.Value) Then
            ' Saltar celdas con errores (#N/A, #VALOR!, etc.)
        Else
            ' Verificar que no esté vacío y sea numérico positivo
            If celda.Value <> "" And IsNumeric(celda.Value) Then
                If CDbl(celda.Value) > 0 Then
                    If celda.Interior.Color <> colorReferencia Then
                        contador = contador + 1
                    End If
                End If
            End If
        End If
    Next celda
    
    ContarSinColor = contador \ 2
    Exit Function
    
ErrorHandler:
    ContarSinColor = contador \ 2  ' Retornar lo contado hasta el momento
End Function


Sub rectificar_tabla()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Consultas")
    Dim rutas As Worksheet
    Set rutas = ThisWorkbook.Sheets("RUTA SAP")
    
    'Verificar si ambos rangos están completamente vacíos
    If WorksheetFunction.CountA(ws.Range("L3:L2000")) = 0 And _
       WorksheetFunction.CountA(ws.Range("G3:J2000")) = 0 Then
        
        MsgBox "No se generará informe, tabla vacía"
        Exit Sub
    ElseIf rutas.Range("A5").Value = "" Then
    
        MsgBox "No existe ruta configurada"
        Exit Sub
        
    Else
        Call generarInforme
    End If
End Sub

Sub generarInforme()
    
    'DEFINIR VARIABLES
    Dim nombreHojaOrigen As String
    Dim rutaDestino As String
    Dim nombreLibroDestino As String
    
    'ASIGNAR VALORES
    nombreHojaOrigen = "Consultas"
    nombreLibroDestino = "informe de stock"
    rutaDestino = ThisWorkbook.Sheets("RUTA SAP").Range("A5").Value
    
    'Declaración de variables
    Dim wbOrigen As Workbook
    Dim wbDestino As Workbook
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFila As Long
    Dim nombreArchivoCompleto As String
    
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
    
    'Encontrar la última fila con datos
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, "I").End(xlUp).row
    
    'Verificar si hay datos para copiar
    If ultimaFila < 2 Then
        MsgBox "No hay datos para generar el informe.", vbInformation
        wbDestino.Close SaveChanges:=False
        Exit Sub
    End If
    
    'Desactivar actualización de pantalla para mayor velocidad
    application.ScreenUpdating = False
    
    'Copiar columnas CON VALORES (no fórmulas)
    'Columna G -> A
    wsOrigen.Range("G1:G" & ultimaFila).Copy
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteValues
    
    'Columna H -> B
    wsOrigen.Range("H1:H" & ultimaFila).Copy
    wsDestino.Range("B1").PasteSpecial Paste:=xlPasteValues
    
    'Columna I -> C
    wsOrigen.Range("I1:I" & ultimaFila).Copy
    wsDestino.Range("C1").PasteSpecial Paste:=xlPasteValues
    
    'Columna J -> D
    wsOrigen.Range("J1:J" & ultimaFila).Copy
    wsDestino.Range("D1").PasteSpecial Paste:=xlPasteValues
    
    'Columna K -> E (IMPORTANTE: Esta columna tiene la fórmula de "Borrado")
    wsOrigen.Range("K1:K" & ultimaFila).Copy
    wsDestino.Range("E1").PasteSpecial Paste:=xlPasteValues
    
    'Columna L -> F
    wsOrigen.Range("L1:L" & ultimaFila).Copy
    wsDestino.Range("F1").PasteSpecial Paste:=xlPasteValues
    
    'Columna M -> G
    wsOrigen.Range("M1:M" & ultimaFila).Copy
    wsDestino.Range("G1").PasteSpecial Paste:=xlPasteValues
    
    'Columna N -> H
    wsOrigen.Range("N1:N" & ultimaFila).Copy
    wsDestino.Range("H1").PasteSpecial Paste:=xlPasteValues
    
    'Columna O -> I
    wsOrigen.Range("O1:O" & ultimaFila).Copy
    wsDestino.Range("I1").PasteSpecial Paste:=xlPasteValues
    
    'Columna P -> J
    wsOrigen.Range("P1:P" & ultimaFila).Copy
    wsDestino.Range("J1").PasteSpecial Paste:=xlPasteValues
    
    'Columna Q -> K
    wsOrigen.Range("Q1:Q" & ultimaFila).Copy
    wsDestino.Range("K1").PasteSpecial Paste:=xlPasteValues
    
    'NUEVAS COLUMNAS AÑADIDAS:
    'Columna R -> L
    wsOrigen.Range("R1:R" & ultimaFila).Copy
    wsDestino.Range("L1").PasteSpecial Paste:=xlPasteValues
    
    'Columna S -> M
    wsOrigen.Range("S1:S" & ultimaFila).Copy
    wsDestino.Range("M1").PasteSpecial Paste:=xlPasteValues
    
    'Limpiar el portapapeles
    application.CutCopyMode = False
    
    'Agregar encabezados si no existen
    If wsDestino.Range("A1").Value = "" Then
        wsDestino.Range("A1").Value = "Columna G"
        wsDestino.Range("B1").Value = "Columna H"
        wsDestino.Range("C1").Value = "Columna I"
        wsDestino.Range("D1").Value = "Columna J"
        wsDestino.Range("E1").Value = "Columna K"
        wsDestino.Range("F1").Value = "Columna L"
        wsDestino.Range("G1").Value = "Columna M"
        wsDestino.Range("H1").Value = "Columna N"
        wsDestino.Range("I1").Value = "Columna O"
        wsDestino.Range("J1").Value = "Columna P"
        wsDestino.Range("K1").Value = "Columna Q"
        wsDestino.Range("L1").Value = "Columna R"
        wsDestino.Range("M1").Value = "Columna S"
    End If
    
    'Ajustar el ancho de las columnas
    wsDestino.Columns("A:M").AutoFit
    
    'Formato de encabezados
    With wsDestino.Range("A1:M1")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242) 'Azul claro
        .HorizontalAlignment = xlCenter
    End With
    
    'Aplicar bordes a los datos
    If ultimaFila > 1 Then
        With wsDestino.Range("A1:M" & ultimaFila)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    End If
    
    'COLOREAR FILAS CON "BORRADO" EN ROJO
    Dim fila As Long
    For fila = 2 To ultimaFila
        If UCase(Trim(wsDestino.Cells(fila, "E").Value)) = "BORRADO" Then
            'Colorear toda la fila en rojo
            wsDestino.Rows(fila).Interior.Color = RGB(255, 199, 206) 'Rojo claro
            wsDestino.Rows(fila).Font.Color = RGB(156, 0, 6) 'Rojo oscuro
        End If
    Next fila
    
    'Reactivar actualización de pantalla
    application.ScreenUpdating = True
    
    'Guardar el libro destino
    wbDestino.SaveAs Filename:=nombreArchivoCompleto, FileFormat:=xlOpenXMLWorkbook
    
    'Preguntar si abrir el archivo
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Deseas abrir el archivo generado?" & vbCrLf & _
                       "Ruta: " & nombreArchivoCompleto, vbYesNo + vbQuestion, "Informe generado")
    
    If respuesta = vbYes Then
        Workbooks.Open nombreArchivoCompleto
    Else
        wbDestino.Close SaveChanges:=False
        MsgBox "Archivo generado exitosamente en:" & vbCrLf & nombreArchivoCompleto, vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    'Reactivar actualización de pantalla en caso de error
    application.ScreenUpdating = True
    application.CutCopyMode = False
    
    MsgBox "Error: " & Err.Description, vbCritical
    
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

Sub limpiar_tabla()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Consultas")
    
    ws.Range("G3:J2000").Clear
    ws.Range("L3:S2000").Clear

End Sub

' ==========================
' OBTENER SOLICITUDES (versión silenciosa)
' ==========================
Sub obtenerSolicitudesMultiples()
    Dim ws As Worksheet
    Dim ultimaFilaF As Long
    Dim i As Long
    Dim regex As Object, matches As Object
    Dim valorCelda As String
    Dim solicitudes() As String
    Dim contadorSolicitudes As Long
    Dim j As Long
    Dim solicitudActual As String
    
    application.ScreenUpdating = False
    application.EnableEvents = False
    application.DisplayAlerts = False
    
    Set ws = ThisWorkbook.Sheets("Consultas")
    ultimaFilaF = ws.Cells(ws.Rows.Count, "F").End(xlUp).row
    
    If ultimaFilaF < 3 Then
        application.ScreenUpdating = True
        application.EnableEvents = True
        application.DisplayAlerts = True
        Exit Sub
    End If
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "\d{6,20}"
    
    ReDim solicitudes(1 To ultimaFilaF - 2)
    contadorSolicitudes = 0
    
    For i = 3 To ultimaFilaF
        valorCelda = Trim(ws.Cells(i, "F").Value)
        If valorCelda <> "" Then
            Set matches = regex.Execute(valorCelda)
            If matches.Count > 0 Then
                contadorSolicitudes = contadorSolicitudes + 1
                solicitudes(contadorSolicitudes) = CStr(matches(0).Value)
            End If
        End If
    Next i
    
    If contadorSolicitudes > 0 Then
        ReDim Preserve solicitudes(1 To contadorSolicitudes)
        
        For j = 1 To contadorSolicitudes
            solicitudActual = solicitudes(j)
            ws.Range("F3").Value = solicitudActual
            Call ProcesarSolicitudIndividual(solicitudActual, j, contadorSolicitudes)
            
            If j < contadorSolicitudes Then
                Esperar 1
            End If
        Next j
        
        ws.Range("F3").ClearContents
    End If
    
    application.ScreenUpdating = True
    application.EnableEvents = True
    application.DisplayAlerts = True
End Sub

' ==========================
' FUNCIÓN PARA PROCESAR UNA SOLICITUD INDIVIDUAL (silenciosa)
' ==========================
Sub ProcesarSolicitudIndividual(numeroSolicitud As String, indice As Long, total As Long)
    Dim ws As Worksheet
    Dim wsRutas As Worksheet
    Dim rutaDestino As String
    Dim nombreLibroDestino As String
    Dim wbDestino As Workbook
    Dim wsDestino As Worksheet
    Dim wsOrigen As Worksheet
    Dim ultimaFila As Long
    Dim nombreArchivoCompleto As String
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("Consultas")
    Set wsRutas = ThisWorkbook.Sheets("RUTA SAP")
    
    solicitud = numeroSolicitud
    
    Call abrirsap
    
    If WorksheetFunction.CountA(ws.Range("L3:L2000")) = 0 And _
       WorksheetFunction.CountA(ws.Range("G3:J2000")) = 0 Then
        Exit Sub
    ElseIf wsRutas.Range("A5").Value = "" Then
        Exit Sub
    End If
    
    nombreLibroDestino = "informe de solicitud " & numeroSolicitud
    rutaDestino = wsRutas.Range("A5").Value
    
    If Right(rutaDestino, 1) <> "\" Then
        rutaDestino = rutaDestino & "\"
    End If
    
    nombreArchivoCompleto = GenerarNombreUnico(rutaDestino, nombreLibroDestino, "xlsx")
    
    Set wsOrigen = ThisWorkbook.Sheets("Consultas")
    
    Set wbDestino = Workbooks.Add
    Set wsDestino = wbDestino.Sheets(1)
    
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, "I").End(xlUp).row
    
    If ultimaFila >= 1 Then
        'Copiar todas las columnas G a Q de una vez
        wsOrigen.Range("G2:Q" & ultimaFila).Copy
        wsDestino.Range("A1").PasteSpecial Paste:=xlPasteAll
    End If
        
    application.CutCopyMode = False
    
    wsDestino.Columns("A:J").AutoFit
    
    wbDestino.SaveAs Filename:=nombreArchivoCompleto, FileFormat:=xlOpenXMLWorkbook
    wbDestino.Close SaveChanges:=False
    
    Call limpiar_tabla
    
    Esperar 0
    
    Exit Sub
    
ErrorHandler:
    application.CutCopyMode = False
    
    If Not wbDestino Is Nothing Then
        On Error Resume Next
        wbDestino.Close SaveChanges:=False
        On Error GoTo 0
    End If
    
    On Error Resume Next
    Call limpiar_tabla
    On Error GoTo 0
End Sub

' ==========================
' MAIN PARA PROCESAMIENTO MÚLTIPLE (silencioso)
' ==========================
Sub mainMultiples()
    Call obtenerSolicitudesMultiples
End Sub



Sub CompararYActualizarPedidos()
    Dim wsInforme As Worksheet
    Dim wsConsultas As Worksheet
    Dim dictInforme As Object
    Dim lastRowInforme As Long
    Dim lastRowConsultas As Long
    Dim i As Long, j As Long
    Dim numeroInforme As String
    Dim textoConsulta As String
    Dim numeroConsulta As String
    Dim posInicio As Long, posFin As Long
    Dim encontrado As Boolean
    Dim valorF As String
    
    On Error GoTo ErrorHandler
    
    
    
    ' Verificar que existen las hojas
    Set wsInforme = ThisWorkbook.Worksheets("Informe de pedidos")
    Set wsConsultas = ThisWorkbook.Worksheets("Consultas")
    wsConsultas.Range("R3:S10000").Clear
    ' Crear diccionario para búsqueda rápida
    Set dictInforme = CreateObject("Scripting.Dictionary")
    
    ' Obtener última fila en Informe de pedidos (columna R)
    lastRowInforme = wsInforme.Cells(wsInforme.Rows.Count, "R").End(xlUp).row
    
    ' Llenar diccionario con datos del Informe
    For i = 2 To lastRowInforme  ' Asumiendo fila 1 es encabezado
        numeroInforme = Trim(wsInforme.Cells(i, "R").Value)
        If numeroInforme <> "" Then
            ' Usar el número como clave, guardar valores de I y F
            If Not dictInforme.Exists(numeroInforme) Then
                ' Obtener y limpiar valor de columna F (reemplazar . por /)
                valorF = CStr(wsInforme.Cells(i, "F").Value)
                If valorF <> "" Then
                    valorF = Replace(valorF, ".", "/")
                End If
                
                dictInforme.Add numeroInforme, Array(wsInforme.Cells(i, "I").Value, valorF)
            End If
        End If
    Next i
    
    ' Obtener última fila en Consultas (columna J)
    lastRowConsultas = wsConsultas.Cells(wsConsultas.Rows.Count, "J").End(xlUp).row
    
    ' Recorrer Consultas y buscar coincidencias
    For i = 2 To lastRowConsultas  ' Asumiendo fila 1 es encabezado
        textoConsulta = Trim(wsConsultas.Cells(i, "J").Value)
        
        If textoConsulta <> "" Then
            ' Buscar número entre corchetes
            posInicio = InStr(1, textoConsulta, "[")
            posFin = InStr(1, textoConsulta, "]")
            
            If posInicio > 0 And posFin > 0 And posFin > posInicio Then
                ' Extraer número entre corchetes
                numeroConsulta = Mid(textoConsulta, posInicio + 1, posFin - posInicio - 1)
                numeroConsulta = Trim(numeroConsulta)
                
                ' Verificar si existe en el diccionario
                If dictInforme.Exists(numeroConsulta) Then
                    ' Obtener valores del diccionario
                    Dim valores As Variant
                    valores = dictInforme(numeroConsulta)
                    
                    ' Copiar a columnas S y R de Consultas
                    wsConsultas.Cells(i, "S").Value = valores(0)  ' Columna I del Informe
                    wsConsultas.Cells(i, "R").Value = valores(1)  ' Columna F del Informe (con / en lugar de .)
                    
                    encontrado = True
                End If
            End If
        End If
    Next i
    
    ' Aplicar reemplazo también a los datos existentes en columna R de Consultas
    For i = 2 To lastRowConsultas
        If Not IsEmpty(wsConsultas.Cells(i, "R").Value) Then
            Dim valorExistente As String
            valorExistente = CStr(wsConsultas.Cells(i, "R").Value)
            If valorExistente <> "" Then
                wsConsultas.Cells(i, "R").Value = Replace(valorExistente, ".", "/")
            End If
        End If
    Next i
    wsConsultas.Activate

    If encontrado Then
        MsgBox "Proceso completado. Datos actualizados en la hoja Consultas.", vbInformation
    Else
        MsgBox "No se encontraron coincidencias.", vbInformation
    End If
    
    ' Limpiar
    Set dictInforme = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Set dictInforme = Nothing
End Sub
