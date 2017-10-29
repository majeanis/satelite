Attribute VB_Name = "ExpExcel"
Option Explicit

Public Type rRegParametros
    Nombre       As String
    Valor        As String
End Type
Public gaRegParametros() As rRegParametros

Public Type rRegColumnas
    Nombre          As String
    Titulo          As String
    Formato         As String
    Ancho           As Long
    AlineamientoHor As Long
    Tipo            As Long
End Type
Public gaRegColumnas() As rRegColumnas

Global Const wc_tipo_dato_integer = 1
Global Const wc_tipo_dato_float = 2
Global Const wc_tipo_dato_fecha = 3
Global Const wc_tipo_dato_hora = 4

Global Const gsSignoMenor = ""
Global Const gsSignoComillas = ""

Dim gsNodo            As String
Dim gsNodoHijo        As String
Function ExportToExcelFromFile(ByVal sFile As String) As Boolean
    '<V1.3.0>
    Dim Row                     As Long
    Dim Col                     As Long
    Dim nLastRow                As Long
    Dim nLastCol                As Long
    Dim ObjExcel                As Excel.Application
    Dim nHoja                   As Integer
    Dim bHojaFound              As Boolean
    Dim lsControl               As Control
    
    Dim nIndice                 As Integer
    Dim nItem                   As Integer
    Dim nCampos                 As Integer
    Dim sTextoOld               As String
    Dim nIniRow                 As Long
    Dim nFinRow                 As Long
    Dim sCelda                  As String
    Dim nCuenta                 As Integer
    Dim nMaxCols                As Long
    Dim nMaxRows                As Long
    Dim sCeldaIni               As String
    Dim sCeldaFin               As String
    Dim sGlsValores             As String
    Dim sValor                  As String
    Dim dValorFecha             As Date
    
    Dim nTipoDatoSalida         As Integer
    Dim sFormatoCelda           As String
    Dim sIndSeparadorMiles      As String
    Dim sNumDecimales           As String
    Dim sFormatoIn              As String
    Dim sFormatoOut             As String
    
    Dim sArray                  As String
    Dim sDelimitador            As String
    Dim Ary(256)                As Long
    Dim sGlsError               As String
    
    Dim sNomArchivoParam        As String
    Dim sNomArchivoData         As String
    Dim sNomArchivoExcel        As String
    Dim sNomHojaExcel           As String
    Dim bGuardaPlanilla         As Boolean
    Dim sTitulo                 As String
    
    On Error GoTo er_ExportToExcelFromFile
    
    Screen.MousePointer = vbHourglass
    
    ' Valida que los archivos existan
    sNomArchivoParam = sFile & ".par"
    sNomArchivoData = sFile & ".dat"
    GrabaLog "Archivo de parametros : " & sNomArchivoParam
    GrabaLog "Archivo de datos : " & sNomArchivoData

    If Not Exist(sNomArchivoParam) Or Not Exist(sNomArchivoData) Then
        Screen.MousePointer = vbNormal
        ExportToExcelFromFile = False
        Exit Function
    End If
    
    sDelimitador = Chr(9)
    For Col = 1 To 256
        DoEvents
        Ary(Col) = 9 ' Saltar columna
    Next Col
    
    ' Encabezado de la consulta
    DoEvents
    Open sNomArchivoParam For Input As #1
    
    If XML_ExtraeNodo("Archivo") Then
        sNomArchivoExcel = XML_ValorElemento("Nombre")
        sNomHojaExcel = XML_ValorElemento("Hoja")
        bGuardaPlanilla = IIf(XML_ValorElemento("Guardar") = "S", True, False)
        nMaxRows = CLng(XML_ValorElemento("Filas"))
        sTitulo = XML_ValorElemento("Titulo")
    End If
    
    ReDim Preserve gaRegParametros(0) As rRegParametros
    DoEvents
    nItem = 0
    If XML_ExtraeNodo("Parametros") Then
        While XML_ExtraeNodoHijo("Parametro")
            DoEvents
            nItem = nItem + 1
            ReDim Preserve gaRegParametros(nItem) As rRegParametros
            gaRegParametros(nItem).Nombre = XML_ValorElemento("Nombre")
            gaRegParametros(nItem).Valor = XML_ValorElemento("Valor")
        Wend
    End If
    
    ' Columnas
    ReDim Preserve gaRegColumnas(0) As rRegColumnas
    DoEvents
    nItem = 0
    If XML_ExtraeNodo("Columnas") Then
        While XML_ExtraeNodoHijo("Columna")
            DoEvents
            nItem = nItem + 1
            ReDim Preserve gaRegColumnas(nItem) As rRegColumnas
            gaRegColumnas(nItem).Nombre = XML_ValorElemento("Nombre")
            gaRegColumnas(nItem).Titulo = XML_ValorElemento("Titulo")
            gaRegColumnas(nItem).Formato = XML_ValorElemento("Formato")
            gaRegColumnas(nItem).Ancho = XML_ValorElemento("Ancho")
            gaRegColumnas(nItem).AlineamientoHor = XML_ValorElemento("AlineamientoHor")
            gaRegColumnas(nItem).Tipo = XML_ValorElemento("Tipo")
        Wend
    End If
    Close #1
        
    nMaxCols = UBound(gaRegColumnas)
        
    GrabaLog "Abriendo excel ..."

    'Set ObjExcel = New Excel.Application
    Set ObjExcel = CreateObject("Excel.application")
        
    ' Se incorpora la posibilidad que el archivo a exportar no exista.
    If sNomArchivoExcel = "" Or Not Exist(sNomArchivoExcel) Then
        ObjExcel.Workbooks.Add
        DoEvents
        If sNomHojaExcel <> "" Then
            ObjExcel.Sheets(1).Select
            ObjExcel.Sheets(1).Name = sNomHojaExcel
        End If
    '</V1.3.0>
    Else
        GrabaLog "Abriendo archivo " & sNomArchivoExcel & " ..."
        
        ObjExcel.Workbooks.Open FileName:=sNomArchivoExcel
        DoEvents
        If sNomHojaExcel = "" Then
            ObjExcel.Sheets(1).Select
            ObjExcel.Sheets.Add
        Else
            bHojaFound = False
            For nHoja = 1 To ObjExcel.Sheets.Count
                DoEvents
                If ObjExcel.Sheets(nHoja).Name = sNomHojaExcel Then
                    ObjExcel.DisplayAlerts = False
                    ObjExcel.Sheets(nHoja).Select
                    If nHoja = ObjExcel.Sheets.Count Then
                        ObjExcel.Sheets.Add
                        ObjExcel.Sheets(nHoja + 1).Select
                        ObjExcel.ActiveWindow.SelectedSheets.Delete
                        ObjExcel.Sheets(nHoja).Name = sNomHojaExcel
                    Else
                        ObjExcel.ActiveWindow.SelectedSheets.Delete
                        ObjExcel.Sheets.Add
                        ObjExcel.Sheets(nHoja).Name = sNomHojaExcel
                    End If
                    bHojaFound = True
                    Exit For
                End If
            Next nHoja
            
            If Not bHojaFound Then
                ObjExcel.Sheets(1).Select
                ObjExcel.Sheets.Add
                DoEvents
                ' Si la hoja no existe se crea una nueva y se renombra la hoja
                If sNomHojaExcel <> "" Then
                    ObjExcel.Sheets(1).Select
                    ObjExcel.Sheets(1).Name = sNomHojaExcel
                End If
            End If
        End If
    End If
    
    GrabaLog "Creando encabezado ..."
    
    ' Crea encabezado del reporte
    DoEvents
    ObjExcel.Range("A1:A2").Font.Color = RGB(0, 0, 0)
    ObjExcel.Range("A1:A2").Font.Name = "Arial"
    ObjExcel.Range("A1:A2").Font.Bold = True
    ObjExcel.Range("A1:A2").Font.Italic = True
    ObjExcel.Range("A1").Font.Size = 14
    ObjExcel.Range("A1").Value = "Sistema de Consultas Satélite"
    ObjExcel.Range("A2").Font.Size = 12
    ObjExcel.Range("A2").Value = sTitulo
    ObjExcel.Range("A3:A3").Font.Size = 11
    ObjExcel.Range("A3:A3").Font.Bold = True
    ObjExcel.Range("A3").Value = "" 'msGlsHeader
    ObjExcel.Range("A1:J1").HorizontalAlignment = -4131
    ObjExcel.Range("A1:J1").WrapText = False
    ObjExcel.Range("A1:J1").Orientation = 0
    ObjExcel.Range("A1:J1").ShrinkToFit = False
    ObjExcel.Range("A1:J1").MergeCells = True
    ObjExcel.Range("A2:J2").HorizontalAlignment = -4131
    ObjExcel.Range("A2:J2").WrapText = False
    ObjExcel.Range("A2:J2").Orientation = 0
    ObjExcel.Range("A2:J2").ShrinkToFit = False
    ObjExcel.Range("A2:J2").MergeCells = True
    
    nLastRow = 4
    For nItem = 1 To UBound(gaRegParametros)
        DoEvents
        ObjExcel.Range("A" & CStr(nLastRow)).Value = gaRegParametros(nItem).Nombre & " = " & gaRegParametros(nItem).Valor
        ObjExcel.Range("A" & CStr(nLastRow) & ":J" & CStr(nLastRow)).HorizontalAlignment = -4131
        ObjExcel.Range("A" & CStr(nLastRow) & ":J" & CStr(nLastRow)).WrapText = False
        ObjExcel.Range("A" & CStr(nLastRow) & ":J" & CStr(nLastRow)).Orientation = 0
        ObjExcel.Range("A" & CStr(nLastRow) & ":J" & CStr(nLastRow)).ShrinkToFit = False
        ObjExcel.Range("A" & CStr(nLastRow) & ":J" & CStr(nLastRow)).MergeCells = True
        ObjExcel.Range("A" & CStr(nLastRow) & ":J" & CStr(nLastRow)).Font.Bold = True
    
        nLastRow = nLastRow + 1
    Next nItem
    
    nLastRow = nLastRow + 1
    nIniRow = nLastRow
    nCuenta = 0
    Dim IndPrimeravez As Integer
    IndPrimeravez = 0
    For Col = 1 To nMaxCols
        DoEvents
        nCuenta = nCuenta + 1
        ObjExcel.Range(IIf(IndPrimeravez = 0, "", Chr(65 + ((Col - 1) \ 26) - 1)) & Chr(65 + nCuenta - 1) & CStr(nLastRow)).Borders.Color = RGB(0, 0, 0)
        ObjExcel.Range(IIf(IndPrimeravez = 0, "", Chr(65 + ((Col - 1) \ 26) - 1)) & Chr(65 + nCuenta - 1) & CStr(nLastRow)).Interior.ColorIndex = 3
        ObjExcel.Range(IIf(IndPrimeravez = 0, "", Chr(65 + ((Col - 1) \ 26) - 1)) & Chr(65 + nCuenta - 1) & CStr(nLastRow)).Font.ColorIndex = 2
        ObjExcel.Range(IIf(IndPrimeravez = 0, "", Chr(65 + ((Col - 1) \ 26) - 1)) & Chr(65 + nCuenta - 1) & CStr(nLastRow)).NumberFormat = "@"
        ObjExcel.Range(IIf(IndPrimeravez = 0, "", Chr(65 + ((Col - 1) \ 26) - 1)) & Chr(65 + nCuenta - 1) & CStr(nLastRow)).Value = gaRegColumnas(Col).Titulo
        If (Col Mod 26) = 0 Then
            nCuenta = 0
            IndPrimeravez = 1
        End If
        
        Ary(Col) = gaRegColumnas(Col).Tipo
    Next
    nLastRow = nLastRow + 1
    nLastCol = nMaxCols
    
    ' Traspasa la data a la planilla excel
    GrabaLog "Exportando archivo " & sNomArchivoData & "..."
        
    Col = 1
    sCeldaIni = fsColNameExcel(65 + Col - 1) & CStr(nLastRow)
    
    DoEvents
    If nMaxRows > 0 Then
        With ObjExcel.ActiveSheet.QueryTables.Add(Connection:="TEXT;" & sNomArchivoData, Destination:=Range(sCeldaIni))
            .Name = "Data"
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = False
            .AdjustColumnWidth = False
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = xlWindows
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            
            .TextFileTabDelimiter = True
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = False
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(Ary(1), Ary(2), Ary(3), Ary(4), Ary(5), Ary(6), Ary(7), Ary(8), Ary(9), Ary(10), Ary(11), Ary(12), Ary(13), Ary(14), Ary(15), Ary(16), Ary(17), Ary(18), Ary(19), Ary(20), _
                Ary(21), Ary(22), Ary(23), Ary(24), Ary(25), Ary(26), Ary(27), Ary(28), Ary(29), Ary(30), Ary(31), Ary(32), Ary(33), Ary(34), Ary(35), Ary(36), Ary(37), Ary(38), Ary(39), Ary(40), _
                Ary(41), Ary(42), Ary(43), Ary(44), Ary(45), Ary(46), Ary(47), Ary(48), Ary(49), Ary(50), Ary(51), Ary(52), Ary(53), Ary(54), Ary(55), Ary(56), Ary(57), Ary(58), Ary(59), Ary(60), _
                Ary(61), Ary(62), Ary(63), Ary(64), Ary(65), Ary(66), Ary(67), Ary(68), Ary(69), Ary(70), Ary(71), Ary(72), Ary(73), Ary(74), Ary(75), Ary(76), Ary(77), Ary(78), Ary(79), Ary(80), _
                Ary(81), Ary(82), Ary(83), Ary(84), Ary(85), Ary(86), Ary(87), Ary(88), Ary(89), Ary(90), Ary(91), Ary(92), Ary(93), Ary(94), Ary(95), Ary(96), Ary(97), Ary(98), Ary(99), Ary(100), _
                Ary(101), Ary(102), Ary(103), Ary(104), Ary(105), Ary(106), Ary(107), Ary(108), Ary(109), Ary(110), Ary(111), Ary(112), Ary(113), Ary(114), Ary(115), Ary(116), Ary(117), Ary(118), Ary(119), Ary(120), _
                Ary(121), Ary(122), Ary(123), Ary(124), Ary(125), Ary(126), Ary(127), Ary(128), Ary(129), Ary(130), Ary(131), Ary(132), Ary(133), Ary(134), Ary(135), Ary(136), Ary(137), Ary(138), Ary(139), Ary(140), _
                Ary(141), Ary(142), Ary(143), Ary(144), Ary(145), Ary(146), Ary(147), Ary(148), Ary(149), Ary(150), Ary(151), Ary(152), Ary(153), Ary(154), Ary(155), Ary(156), Ary(157), Ary(158), Ary(159), Ary(160), _
                Ary(161), Ary(162), Ary(163), Ary(164), Ary(165), Ary(166), Ary(167), Ary(168), Ary(169), Ary(170), Ary(171), Ary(172), Ary(173), Ary(174), Ary(175), Ary(176), Ary(177), Ary(178), Ary(179), Ary(180), _
                Ary(181), Ary(182), Ary(183), Ary(184), Ary(185), Ary(186), Ary(187), Ary(188), Ary(189), Ary(190), Ary(191), Ary(192), Ary(193), Ary(194), Ary(195), Ary(196), Ary(197), Ary(198), Ary(199), Ary(200), _
                Ary(201), Ary(202), Ary(203), Ary(204), Ary(205), Ary(206), Ary(207), Ary(208), Ary(209), Ary(210), Ary(211), Ary(212), Ary(213), Ary(214), Ary(215), Ary(216), Ary(217), Ary(218), Ary(219), Ary(220), _
                Ary(221), Ary(222), Ary(223), Ary(224), Ary(225), Ary(226), Ary(227), Ary(228), Ary(229), Ary(230), Ary(231), Ary(232), Ary(233), Ary(234), Ary(235), Ary(236), Ary(237), Ary(238), Ary(239), Ary(240), _
                Ary(241), Ary(242), Ary(243), Ary(244), Ary(245), Ary(246), Ary(247), Ary(248), Ary(249), Ary(250), Ary(251), Ary(252), Ary(253), Ary(254), Ary(255), Ary(256))
        
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
        DoEvents
        End With
    End If
    
    ' Formatea columnas
    GrabaLog "Formateando columnas ..."
    For Col = 1 To nMaxCols
        DoEvents
        sCeldaIni = fsColNameExcel(65 + Col - 1) & CStr(nLastRow)
        If nMaxRows > 0 Then
            sCeldaFin = fsColNameExcel(65 + Col - 1) & CStr(nLastRow + nMaxRows - 1)
        Else
            sCeldaFin = sCeldaIni
        End If

        ObjExcel.Range(sCeldaIni & ":" & sCeldaFin).NumberFormat = gaRegColumnas(Col).Formato
        ObjExcel.Range(sCeldaIni & ":" & sCeldaFin).Font.Size = 9
        ObjExcel.Range(sCeldaIni & ":" & sCeldaFin).HorizontalAlignment = gaRegColumnas(Col).AlineamientoHor
        ObjExcel.Range(sCeldaIni & ":" & sCeldaFin).ColumnWidth = gaRegColumnas(Col).Ancho
    Next

    ' Ajusta columnas en caso que no existan grillas
    ObjExcel.Cells.Select
    ObjExcel.Cells.EntireColumn.AutoFit
    ObjExcel.Range("A1:A1").Select
    ObjExcel.ActiveSheet.PageSetup.Orientation = 2
    DoEvents
    
    '<V1.3.0>
    If Not bGuardaPlanilla Then
        ObjExcel.Visible = True
        DoEvents
    Else
        ' Se graba automáticamente la planilla y se cierra
        GrabaLog "Guardando archivo"
        If Exist(sNomArchivoExcel) Then
            ObjExcel.ActiveWorkbook.Save
            DoEvents
        Else
            ObjExcel.ActiveWorkbook.SaveAs FileName:=sNomArchivoExcel, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
            DoEvents
        End If
        ObjExcel.ActiveWorkbook.Close
    End If
    '</V1.3.0>
    
    GrabaLog "Finalizado sin errores"
    DoEvents
    Set ObjExcel = Nothing
    
    On Error Resume Next
    Kill sNomArchivoParam
    Kill sNomArchivoData

    Screen.MousePointer = vbNormal
    
    ExportToExcelFromFile = True
    Exit Function

er_ExportToExcelFromFile:
    If Err = 1004 Then
        Resume Next
    Else
        sGlsError = Format(Err, "&") & "-" & Error
        GrabaLog sGlsError
        On Error Resume Next
        Kill sNomArchivoParam
        Kill sNomArchivoData
        ObjExcel.ActiveWorkbook.Close
        Set ObjExcel = Nothing
        Screen.MousePointer = vbNormal
        
        ExportToExcelFromFile = False
        Exit Function
    End If
    '</V1.3.0>
End Function

Function Exist(sFile As String) As Boolean
    Dim nn As Long

    On Error GoTo noExistFile
    If Trim(sFile) <> "" Then
        nn = FileLen(sFile)
        Exist = True
    Else
        Exist = False
    End If
Exit Function

noExistFile:
    Exist = False
    Exit Function
End Function

Function XML_ExtraeNodo(sElemento As String) As Boolean
    Dim sLinea      As String
    Dim sSalida     As String
    Dim bFinBuscar  As Boolean
    
    sSalida = ""
    bFinBuscar = False
    
    While Not EOF(1) And Not bFinBuscar
        Line Input #1, sLinea
        
        If InStr(LCase(sLinea), "<" & LCase(sElemento) & ">") > 0 Then
            sSalida = sSalida & sLinea
            While Not EOF(1) And Not bFinBuscar
                Line Input #1, sLinea
                sSalida = sSalida & sLinea
                
                If InStr(LCase(sLinea), "</" & LCase(sElemento) & ">") > 0 Then
                   bFinBuscar = True
                End If
            Wend
        End If
    Wend
    
    gsNodo = sSalida
    gsNodoHijo = gsNodo
    If gsNodo = "" Then
        XML_ExtraeNodo = False
    Else
        XML_ExtraeNodo = True
    End If
End Function
Function XML_ExtraeNodoHijo(sElemento As String) As Boolean
    Dim sSalida     As String
    Dim bFinBuscar  As Boolean
    Dim nPos1       As Long
    Dim nPos2       As Long
    Dim sAux        As String
    
    sSalida = ""
    bFinBuscar = False
    
    nPos1 = InStr(LCase(gsNodo), "<" & LCase(sElemento) & ">")
    If nPos1 > 0 Then
        sAux = Mid(gsNodo, nPos1)
        nPos2 = InStr(LCase(sAux), "</" & LCase(sElemento) & ">")
        If nPos2 > 0 Then
            sAux = Left(sAux, nPos2 + Len("</" & LCase(sElemento) & ">") - 1)
            gsNodo = Mid(gsNodo, nPos1 + nPos2 + Len("</" & LCase(sElemento) & ">") - 1)
        Else
            gsNodo = ""
        End If
    End If
    
    gsNodoHijo = sAux
    If gsNodoHijo = "" Then
        XML_ExtraeNodoHijo = False
    Else
        XML_ExtraeNodoHijo = True
    End If
End Function


Function fsColNameExcel(ByVal nColumna As Long) As String
    Dim sLetra1 As String
    Dim sLetra2 As String
    Dim nLetra1 As Integer
    Dim nLetra2 As Integer
    
    nLetra1 = (nColumna - 65) \ 26
    nLetra2 = nColumna - (nLetra1 * 26)
    sLetra1 = ""
    If nLetra1 > 0 Then
        sLetra1 = Chr(65 + nLetra1 - 1)
    End If
    sLetra2 = Chr(nLetra2)
    
    If (sLetra1 > "I" And sLetra2 <> "") Or (sLetra1 = "I" And sLetra2 >= "U") Then
        fsColNameExcel = "IU"
    Else
        fsColNameExcel = sLetra1 & sLetra2
    End If
End Function

Sub GrabaLog(sTexto As String)
    'Open App.Path & "\Exportar.log" For Append As #3
    'Print #3, sTexto
    'Close #3
End Sub


Function XML_ValorElemento(sElemento As String) As String
    Dim nPos    As Long
    Dim sAux        As String
    
    sAux = ""
    nPos = InStr(LCase(gsNodoHijo), "<" & LCase(sElemento) & "=")
    If nPos > 0 Then
        sAux = Mid(gsNodoHijo, nPos)
        nPos = InStr(sAux, "> ")
        If nPos = 0 Then nPos = InStr(sAux, "><")
        If nPos = 0 Then nPos = InStr(sAux, ">")
        
        If nPos > 0 Then
            sAux = Left(sAux, nPos - 1)
        End If
        sAux = Mid(sAux, Len(sElemento & "=") + 2)
    End If
    
    XML_ValorElemento = sAux
End Function


