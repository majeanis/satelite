Attribute VB_Name = "Ind_Excel"
Option Explicit
Public mStr_NombrePestanas As String
Private mStr_Mensaje As String, mStr_Ruta_Imagen As String, mStr_RutaImagen As String
Private mLon_Filas As Long, mLon_Columnas As Long
Private maStr_Filas_Grillas() As String
Private maStr_Columnas_Grillas() As String
Private maStr_Pocision() As String
Private mInt_Grilla As Integer
Public s_TipoDatoExcel As String
Public gObj_Excel As Excel.Application

Public Type recColumnas
    TipoDatoSalida      As Long
    FormatoIn           As String
End Type

Dim cChar(276)  As String
Dim cEsp1(276)  As String
Dim cEsp2(276)  As String

Sub EsperaExcel(sArchivo As String)
    While Exist(sArchivo & ".dat")
    Wend
End Sub

Function ExportarAExcel(ByVal sTitulo As String, grdData As fpSpread, rsData As ADODB.Recordset, rsFormatos As ADODB.Recordset, ProgressBar1 As ProgressBar, StatusBar1 As StatusBar, bGuardaPlanilla As Boolean, sGlsError As String) As Boolean
    Dim Row                 As Long
    Dim Col                 As Long
    Dim nLastRow            As Long
    Dim ObjExcel            As Object
    Dim nHoja               As Integer
    Dim bHojaFound          As Boolean
    Dim lsControl           As Control
    
    Dim nIndice             As Integer
    Dim nItem               As Integer
    Dim nCampos             As Integer
    Dim sTextoOld           As String
    Dim nIniRow             As Long
    Dim nFinRow             As Long
    Dim sCelda              As String
    Dim nCuenta             As Integer
    
    Dim sCeldaIni           As String
    Dim sCeldaFin           As String
    Dim sGlsValores         As String
    Dim sValor              As String
    Dim dValorFecha         As Date
    Dim nValorDouble        As Double
    Dim nValorLong          As Long
    Dim fCampos             As Field
    
    Dim nTipoDatoSalida     As Integer
    Dim sFormatoCelda       As String
    Dim sIndSeparadorMiles  As String
    Dim sNumDecimales       As String
    Dim sFormatoIn          As String
    Dim sFormatoOut         As String
    
    On Error GoTo er_ExportarAExcel
    
    sTitulo = Replace(sTitulo, ".sql", "")
    sTitulo = Replace(sTitulo, ".SQL", "")
        
    Screen.MousePointer = vbHourglass
    StatusBar1.Panels(2).Text = "Preparando excel ..."

    Set ObjExcel = CreateObject("Excel.Application")
            
    '<V1.3.0>
    ' Se incorpora la posibilidad que el archivo a exportar no exista.
    If gsNomArchivoExportar = "" Or Not Exist(gsNomArchivoExportar) Then
        ObjExcel.Workbooks.Add
        If gsNomHojaExportar <> "" Then
            ObjExcel.Sheets(1).Select
            ObjExcel.Sheets(1).Name = gsNomHojaExportar
        End If
    '</V1.3.0>
    Else
        StatusBar1.Panels(2).Text = "Abriendo archivo " & gsNomArchivoExportar & " ..."
        GrabaLog "Abriendo archivo " & gsNomArchivoExportar & " ..."
        
        ObjExcel.Workbooks.Open FileName:=gsNomArchivoExportar
        If gsNomHojaExportar = "" Then
            ObjExcel.Sheets(1).Select
            ObjExcel.Sheets.Add
        Else
            bHojaFound = False
            For nHoja = 1 To ObjExcel.Sheets.Count
                If ObjExcel.Sheets(nHoja).Name = gsNomHojaExportar Then
                    ObjExcel.DisplayAlerts = False
                    ObjExcel.Sheets(nHoja).Select
                    If nHoja = ObjExcel.Sheets.Count Then
                        ObjExcel.Sheets.Add
                        ObjExcel.Sheets(nHoja + 1).Select
                        ObjExcel.ActiveWindow.SelectedSheets.Delete
                        ObjExcel.Sheets(nHoja).Name = gsNomHojaExportar
                    Else
                        ObjExcel.ActiveWindow.SelectedSheets.Delete
                        ObjExcel.Sheets.Add
                        ObjExcel.Sheets(nHoja).Name = gsNomHojaExportar
                    End If
                    bHojaFound = True
                    Exit For
                End If
            Next nHoja
            If Not bHojaFound Then
                ObjExcel.Sheets(1).Select
                ObjExcel.Sheets.Add
                '</V1.3.0>
                ' Si la hoja no existe se crea una nueva y se renombra la hoja
                If gsNomHojaExportar <> "" Then
                    ObjExcel.Sheets(1).Select
                    ObjExcel.Sheets(1).Name = gsNomHojaExportar
                End If
                '</V1.3.0>
            End If
        End If
    End If
    
    ' Crea encabezado del reporte
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
        ObjExcel.Range("A" & CStr(nLastRow)).Value = gaRegParametros(nItem).Nombre & " = " & gaRegParametros(nItem).valor
        ObjExcel.Range("A" & CStr(nLastRow) & ":J" & CStr(nLastRow)).HorizontalAlignment = -4131
        ObjExcel.Range("A" & CStr(nLastRow) & ":J" & CStr(nLastRow)).WrapText = False
        ObjExcel.Range("A" & CStr(nLastRow) & ":J" & CStr(nLastRow)).Orientation = 0
        ObjExcel.Range("A" & CStr(nLastRow) & ":J" & CStr(nLastRow)).ShrinkToFit = False
        ObjExcel.Range("A" & CStr(nLastRow) & ":J" & CStr(nLastRow)).MergeCells = True
        ObjExcel.Range("A" & CStr(nLastRow) & ":J" & CStr(nLastRow)).Font.Bold = True
    
        nLastRow = nLastRow + 1
    Next nItem
    
    StatusBar1.Panels(2).Text = "Preparando columnas ..."
    nLastRow = nLastRow + 1
    nIniRow = nLastRow
    nCuenta = 0
    Dim IndPrimeravez As Integer
    IndPrimeravez = 0
    For Col = 1 To grdData.MaxCols
        nCuenta = nCuenta + 1
        ObjExcel.Range(IIf(IndPrimeravez = 0, "", Chr(65 + ((Col - 1) \ 26) - 1)) & Chr(65 + nCuenta - 1) & CStr(nLastRow)).Borders.Color = RGB(0, 0, 0)
        ObjExcel.Range(IIf(IndPrimeravez = 0, "", Chr(65 + ((Col - 1) \ 26) - 1)) & Chr(65 + nCuenta - 1) & CStr(nLastRow)).Interior.ColorIndex = 3
        ObjExcel.Range(IIf(IndPrimeravez = 0, "", Chr(65 + ((Col - 1) \ 26) - 1)) & Chr(65 + nCuenta - 1) & CStr(nLastRow)).Font.ColorIndex = 2
        ObjExcel.Range(IIf(IndPrimeravez = 0, "", Chr(65 + ((Col - 1) \ 26) - 1)) & Chr(65 + nCuenta - 1) & CStr(nLastRow)).NumberFormat = "@"
        ObjExcel.Range(IIf(IndPrimeravez = 0, "", Chr(65 + ((Col - 1) \ 26) - 1)) & Chr(65 + nCuenta - 1) & CStr(nLastRow)).Value = fsGetGrilla(grdData, 0, Col)
        If (Col Mod 26) = 0 Then
            nCuenta = 0
            IndPrimeravez = 1
        End If
    Next
    nLastRow = nLastRow + 1
    
    ' Formatea columnas
    GrabaLog "Formateando columnas"
    grdData.Row = 1
    Col = 1
    For Each fCampos In rsData.Fields
        sCeldaIni = fsColNameExcel(65 + Col - 1) & CStr(nLastRow)
        sCeldaFin = fsColNameExcel(65 + Col - 1) & CStr(nLastRow + grdData.MaxRows - 1)

        grdData.Col = Col
        sFormatoCelda = "@"
        nTipoDatoSalida = fnTipoDatoRecordset(fCampos.Type)
        
        ' Busca formato de salida del campo
        sIndSeparadorMiles = "S"
        If nTipoDatoSalida = wc_tipo_dato_integer Then
            sNumDecimales = "0"
        Else
            If fCampos.NumericScale = 0 Then
                sNumDecimales = "0"
            Else
                sNumDecimales = "2"
            End If
        End If
        sFormatoOut = "dd/mm/yyyy"
        
        rsFormatos.Filter = "nom_columna='" & LCase(fCampos.Name) & "'"
        If Not rsFormatos.EOF Then
            nTipoDatoSalida = fnTipoDato("" & rsFormatos!cod_tipo_dato_salida)
            sIndSeparadorMiles = "" & rsFormatos!ind_separador_miles
            sNumDecimales = "" & rsFormatos!num_decimales
            sFormatoIn = "" & rsFormatos!gls_formato_entrada
            sFormatoOut = "" & rsFormatos!gls_formato_salida
            
            sFormatoIn = Replace(sFormatoIn, gsSignoMenor, "<")
            sFormatoIn = Replace(sFormatoIn, gsSignoComillas, """")
            sFormatoOut = Replace(sFormatoOut, gsSignoMenor, "<")
            sFormatoOut = Replace(sFormatoOut, gsSignoComillas, """")
        End If
        
        ' Carga valor formateado
        Select Case nTipoDatoSalida
        Case wc_tipo_dato_float, wc_tipo_dato_integer
            sFormatoCelda = fsFormatoNumerico(sIndSeparadorMiles, sNumDecimales) ' fsFormatoValorNumerico("" & rsData(nY - 1), sIndSeparadorMiles, sNumDecimales)
        '<INI SP1.2.1>
        Case wc_tipo_dato_fecha, wc_tipo_dato_hora
        '<FIN SP1.2.1>
            sFormatoCelda = sFormatoOut
        End Select
        
        ObjExcel.Range(sCeldaIni & ":" & sCeldaFin).NumberFormat = sFormatoCelda
        ObjExcel.Range(sCeldaIni & ":" & sCeldaFin).Font.Size = 9
        
        If grdData.TypeHAlign = 0 Then
            If nTipoDatoSalida = wc_tipo_dato_integer And Val(sNumDecimales) > 0 Then
                ObjExcel.Range(sCeldaIni & ":" & sCeldaFin).HorizontalAlignment = -4152
            Else
                ObjExcel.Range(sCeldaIni & ":" & sCeldaFin).HorizontalAlignment = -4131
            End If
        ElseIf grdData.TypeHAlign = 1 Then
            ObjExcel.Range(sCeldaIni & ":" & sCeldaFin).HorizontalAlignment = -4152
        End If
        ObjExcel.Range(sCeldaIni & ":" & sCeldaFin).ColumnWidth = IIf((grdData.ColWidth(Col) \ 100) + 1 <= 255, (grdData.ColWidth(Col) \ 100) + 1, 255)
        
        Col = Col + 1
    Next
    
    ' Traspasa la data a la planilla excel
    ProgressBar1.Max = grdData.MaxCols
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    
    rsData.Filter = ""
    Col = 1
    For Each fCampos In rsData.Fields
        grdData.Row = 1
        grdData.Col = Col
        GrabaLog "Subiendo columna " & fCampos.Name
        
        nTipoDatoSalida = fnTipoDatoRecordset(fCampos.Type)
        
        ' Busca formato de salida del campo
        sIndSeparadorMiles = "S"
        If nTipoDatoSalida = wc_tipo_dato_integer Then
            sNumDecimales = "0"
        Else
            If fCampos.NumericScale = 0 Then
                sNumDecimales = "0"
            Else
                sNumDecimales = "2"
            End If
        End If
        sFormatoIn = ""
        sFormatoOut = ""
        
        rsFormatos.Filter = "nom_columna='" & LCase(rsData(Col - 1).Name) & "'"
        If Not rsFormatos.EOF Then
            nTipoDatoSalida = fnTipoDato("" & rsFormatos!cod_tipo_dato_salida)
            sIndSeparadorMiles = "" & rsFormatos!ind_separador_miles
            sNumDecimales = "" & rsFormatos!num_decimales
            If nTipoDatoSalida = wc_tipo_dato_integer And Val(sNumDecimales) > 0 Then
                nTipoDatoSalida = wc_tipo_dato_float
            End If
            
            sFormatoIn = "" & rsFormatos!gls_formato_entrada
            sFormatoOut = "" & rsFormatos!gls_formato_salida
            
            sFormatoIn = Replace(sFormatoIn, gsSignoMenor, "<")
            sFormatoIn = Replace(sFormatoIn, gsSignoComillas, """")
            sFormatoOut = Replace(sFormatoOut, gsSignoMenor, "<")
            sFormatoOut = Replace(sFormatoOut, gsSignoComillas, """")
        End If
        
        Select Case nTipoDatoSalida
        '<INI SP1.2.1>
        Case wc_tipo_dato_fecha, wc_tipo_dato_hora
        '<FIN SP1.2.1>
            rsData.MoveFirst
            Row = 1
            While Not rsData.EOF
                sCeldaIni = fsColNameExcel(65 + Col - 1) & CStr(nLastRow + Row - 1)
                sValor = fsValorFechaExcel("" & rsData(Col - 1), sFormatoIn)
                dValorFecha = fdValorFecha(sValor)
                If dValorFecha <> gdNullDate Then
                    If dValorFecha < gdMinDateExcel Then
                        ObjExcel.Range(sCeldaIni).NumberFormat = "@"
                        ObjExcel.Range(sCeldaIni).Value = "" & dValorFecha
                    Else
                        ObjExcel.Range(sCeldaIni).Value = dValorFecha
                    End If
                End If
                
                rsData.MoveNext
                Row = Row + 1
            Wend
        
        Case wc_tipo_dato_float
            rsData.MoveFirst
            Row = 1
            While Not rsData.EOF
                sCeldaIni = fsColNameExcel(65 + Col - 1) & CStr(nLastRow + Row - 1)
                sValor = fsValorDobleExcel("" & rsData(Col - 1))
                nValorDouble = fnValorDoble(sValor)
                ObjExcel.Range(sCeldaIni).Value = nValorDouble
                
                rsData.MoveNext
                Row = Row + 1
            Wend

        Case wc_tipo_dato_integer
            rsData.MoveFirst
            Row = 1
            While Not rsData.EOF
                sCeldaIni = fsColNameExcel(65 + Col - 1) & CStr(nLastRow + Row - 1)
                sValor = fsValorDobleExcel("" & rsData(Col - 1))
                nValorLong = fnValorEntero(sValor)
                ObjExcel.Range(sCeldaIni).Value = nValorLong
                
                rsData.MoveNext
                Row = Row + 1
            Wend

        Case Else
            rsData.MoveFirst
            Row = 1
            While Not rsData.EOF
                sCeldaIni = fsColNameExcel(65 + Col - 1) & CStr(nLastRow + Row - 1)
                sValor = Replace(Replace("" & rsData(Col - 1), Chr(13), " "), Chr(10), "")
                ObjExcel.Range(sCeldaIni).Value = sValor
                
                rsData.MoveNext
                Row = Row + 1
            Wend
        End Select
        
        ProgressBar1.Value = Col
        Col = Col + 1
    Next
    nLastRow = nLastRow + grdData.MaxRows
    nFinRow = nLastRow - 1

    ' Ajusta columnas en caso que no existan grillas
    GrabaLog "Terminó de subir la data"
    
    ObjExcel.Cells.Select
    ObjExcel.Cells.EntireColumn.AutoFit
    ObjExcel.Range("A1:A1").Select
    ObjExcel.ActiveSheet.PageSetup.Orientation = 2
    
    '<V1.3.0>
    If Not bGuardaPlanilla Then
        ObjExcel.Visible = True
    Else
        ' Se graba automáticamente la planilla y se cierra
        GrabaLog "Guardando archivo"
        If Exist(gsNomArchivoExportar) Then
            ObjExcel.ActiveWorkbook.Save
        Else
            ObjExcel.ActiveWorkbook.SaveAs FileName:=gsNomArchivoExportar, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        End If
        ObjExcel.ActiveWorkbook.Close
    End If
    '</V1.3.0>
    
    GrabaLog "Fin exportar"
    StatusBar1.Panels(2).Text = ""
    ProgressBar1.Visible = False
    Set ObjExcel = Nothing
    Screen.MousePointer = vbNormal
    
    sGlsError = ""
    ExportarAExcel = True
    Exit Function

er_ExportarAExcel:
    If Err = 1004 Then
        Resume Next
    Else
        sGlsError = Error
        StatusBar1.Panels(2).Text = ""
        On Error Resume Next
        ObjExcel.ActiveWorkbook.Close
        Set ObjExcel = Nothing
        Screen.MousePointer = vbNormal
        
        ExportarAExcel = False
        Exit Function
    End If
End Function
Function ExportarToFile(ByVal sTitulo As String, grdData As fpSpread, rsData As ADODB.Recordset, rsFormatos As ADODB.Recordset, arrParametros() As rRegParametros, ProgressBar1 As ProgressBar, StatusBar1 As StatusBar, bGuardaPlanilla As Boolean, sGlsError As String) As Boolean
    '<V1.3.0>
    Dim nItem               As Long
    Dim Col                 As Long
    Dim Row                 As Long
    Dim fCampos             As Field
    Dim nTipoDatoSalida     As Long
    Dim sIndSeparadorMiles  As String
    Dim sNumDecimales       As String
    Dim sFormatoOut         As String
    Dim sFormatoIn          As String
    Dim sFormatoCelda       As String
    Dim sGlsValores         As String
    Dim sDelimitador        As String
    Dim nCampos             As Long
    Dim sAlignment          As String
    Dim nTipo               As Long
    Dim arrColumnas()       As recColumnas
    Dim sFile               As String
    Dim sGlsCampo           As String
    
    On Error GoTo er_ExportarToFile
    
    ' Carga caracteres HTML
    '<V1.3.1>
    CargaCaracteresEspeciales
    '</V1.3.1>
    
    If Not Exist(App.Path & "\Exportar.exe") Then
        MsgBox "Error en la instalación del sistema. Falta archivo para exportar a Excel", vbCritical, App.Title
        ExportarToFile = False
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    sDelimitador = Chr(9)
    
    nCampos = grdData.MaxCols
    ReDim arrColumnas(nCampos) As recColumnas
    
    sFile = fsNombreArchivo(sTitulo)
    
    ' Guarda informacion de la planilla
    Open sFile & ".par" For Output As #1
    Print #1, "<Archivo>"
    Print #1, " <Nombre=" & gsNomArchivoExportar & ">"
    Print #1, " <Hoja=" & gsNomHojaExportar & ">"
    Print #1, " <Guardar=" & IIf(bGuardaPlanilla, "S", "N") & ">"
    Print #1, " <Filas=" & Format(grdData.MaxRows, "&") & ">"
    Print #1, " <Titulo=" & sTitulo & ">"
    Print #1, "</Archivo>"
    
    Print #1, "<Parametros>"
    For nItem = 1 To UBound(arrParametros)
        Print #1, " <Parametro>"
        Print #1, "  <Nombre=" & arrParametros(nItem).Nombre & ">"
        Print #1, "  <Valor=" & arrParametros(nItem).valor & ">"
        Print #1, " </Parametro>"
    Next nItem
    Print #1, "</Parametros>"
    
    ' Guarda información de las columnas
    Print #1, "<Columnas>"
    Col = 1
    For Each fCampos In rsData.Fields
        grdData.Col = Col
        
        ' Busca formato de salida del campo
        nTipoDatoSalida = fnTipoDatoRecordset(fCampos.Type)
        sIndSeparadorMiles = "S"
        If nTipoDatoSalida = wc_tipo_dato_integer Then
            sNumDecimales = "0"
        Else
            If fCampos.NumericScale = 0 Then
                sNumDecimales = "0"
            Else
                sNumDecimales = "2"
            End If
        End If
                
        ' Busca formato de salida del campo
        sFormatoIn = ""
        sFormatoOut = "dd/mm/yyyy"
        
        rsFormatos.Filter = "nom_columna='" & LCase(fCampos.Name) & "'"
        If Not rsFormatos.EOF Then
            nTipoDatoSalida = fnTipoDato("" & rsFormatos!cod_tipo_dato_salida)
            sNumDecimales = "" & rsFormatos!num_decimales
            sIndSeparadorMiles = "" & rsFormatos!ind_separador_miles
            If nTipoDatoSalida = wc_tipo_dato_integer And Val(sNumDecimales) > 0 Then
                nTipoDatoSalida = wc_tipo_dato_float
            End If
            
            sFormatoIn = "" & rsFormatos!gls_formato_entrada
            sFormatoOut = "" & rsFormatos!gls_formato_salida
            
            sFormatoIn = Replace(sFormatoIn, gsSignoMenor, "<")
            sFormatoIn = Replace(sFormatoIn, gsSignoComillas, """")
            sFormatoOut = Replace(sFormatoOut, gsSignoMenor, "<")
            sFormatoOut = Replace(sFormatoOut, gsSignoComillas, """")
        End If
                
        arrColumnas(Col).TipoDatoSalida = nTipoDatoSalida
        arrColumnas(Col).FormatoIn = sFormatoIn
        
        ' Carga valor formateado
        sFormatoCelda = "@"
        Select Case nTipoDatoSalida
        Case wc_tipo_dato_float, wc_tipo_dato_integer
            sFormatoCelda = fsFormatoNumerico(sIndSeparadorMiles, sNumDecimales)
            nTipo = 1
        '<INI SP1.2.1>
        Case wc_tipo_dato_fecha, wc_tipo_dato_hora
        '<FIN SP1.2.1>
            sFormatoCelda = sFormatoOut
            nTipo = 5
        Case Else
            sFormatoCelda = "&"
            nTipo = 2
        End Select
        
        If grdData.TypeHAlign = 0 Then
            If nTipoDatoSalida = wc_tipo_dato_integer And Val(sNumDecimales) > 0 Then
                sAlignment = "-4152"
            Else
                sAlignment = "-4131"
            End If
        ElseIf grdData.TypeHAlign = 1 Then
            sAlignment = "-4152"
        End If
        
        Print #1, " <Columna>"
        Print #1, "  <Nombre=" & rsData(Col - 1).Name & ">"
        Print #1, "  <Titulo=" & fsGetGrilla(grdData, 0, Col) & ">"
        Print #1, "  <Formato=" & sFormatoCelda & ">"
        Print #1, "  <Ancho=" & IIf((grdData.ColWidth(Col) \ 100) + 1 <= 255, (grdData.ColWidth(Col) \ 100) + 1, 255) & ">"
        Print #1, "  <AlineamientoHor=" & sAlignment & ">"
        Print #1, "  <Tipo=" & Format(nTipo, "&") & ">"
        Print #1, " </Columna>"
                
        Col = Col + 1
    Next
    Print #1, "</Columnas>"
    Close #1
    
    ' Guarda la data en el archivo
    StatusBar1.Panels(2).Text = "Guardando información ..."
    If grdData.MaxRows > 0 Then
        ProgressBar1.Max = grdData.MaxRows
    Else
        ProgressBar1.Max = 1
    End If
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    
    Open sFile & ".dat" For Output As #1
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
    End If
    Row = 1
    While Not rsData.EOF
        grdData.Row = Row
        sGlsValores = ""
     
        Col = 1
        For Each fCampos In rsData.Fields
            nTipoDatoSalida = arrColumnas(Col).TipoDatoSalida
            sFormatoIn = arrColumnas(Col).FormatoIn
                        
            Select Case nTipoDatoSalida
            Case wc_tipo_dato_fecha, wc_tipo_dato_hora
                sGlsValores = sGlsValores & fsValorFechaExcel("" & rsData(Col - 1), sFormatoIn) & sDelimitador
                                
            Case wc_tipo_dato_float
                sGlsValores = sGlsValores & fsValorDobleExcel("" & rsData(Col - 1)) & sDelimitador
    
            Case wc_tipo_dato_integer
                sGlsValores = sGlsValores & fsValorIntegerExcel("" & rsData(Col - 1)) & sDelimitador
            
            Case Else
                sGlsCampo = "" & rsData(Col - 1)
                If InStr(sGlsCampo, Chr(9)) Then
                    sGlsCampo = sGlsCampo
                End If
                sGlsCampo = Replace(Replace(Replace(sGlsCampo, Chr(13), " "), Chr(10), ""), Chr(9), "")
                '<V1.3.1>
                sGlsCampo = fsReemplazaCarEspec(sGlsCampo)
                '</V1.3.1>
                sGlsValores = sGlsValores & sGlsCampo & sDelimitador
            
            End Select
            
            Col = Col + 1
        Next
        
        sGlsValores = Left(sGlsValores, Len(sGlsValores) - 1)
        Print #1, sGlsValores
        
        rsData.MoveNext
        Row = Row + 1
        ProgressBar1.Value = Row - 1
    Wend
    Close #1
    ProgressBar1.Visible = False
        
    Shell App.Path & "\Exportar.exe " & sFile, vbMinimizedNoFocus
    Call EsperaExcel(sFile)
    
    StatusBar1.Panels(2).Text = ""
    Screen.MousePointer = vbNormal
    
    ExportarToFile = True
    Exit Function

er_ExportarToFile:
    GrabaLog Error
    GrabaLog Format(Err, "&")
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    
    On Error Resume Next
    Kill sFile & ".par"
    Kill sFile & ".dat"

    ExportarToFile = False
    '</V1.3.0>
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

Function fsNombreArchivo(sTitulo As String) As String
    '<V1.3.0>
    Dim sDrive              As String
    Dim sPath               As String
    Dim sFile               As String
    Dim nPos                As Integer
    Dim bCarpetaTemporal    As Boolean
    Dim sCarpetaFinal       As String
    Dim sHora               As String
    
    On Error GoTo ErrNombreArchivo
    
    sHora = Format(Time(), "hhmmss")
    sFile = fsSoloLetras(sTitulo) & "_" & sHora
    
    ObtieneVersion
    If gbVersionLocal Then
        sCarpetaFinal = App.Path & "\Temp"
    Else
        sCarpetaFinal = "C:\Temp"
    End If
    
    If Not Exist(sCarpetaFinal) Then
        MkDir sCarpetaFinal
    End If
    
    fsNombreArchivo = sCarpetaFinal & "\" & sFile
    Exit Function
    
ErrNombreArchivo:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    fsNombreArchivo = ""
    '</V1.3.0>
End Function

Function CargaCaracteresEspeciales()
    '<V1.3.1>
    cEsp1(32) = "&#32;"
    cEsp1(33) = "&#33;"
    cEsp1(34) = "&#34;&quot;"
    cEsp1(35) = "&#35;"
    cEsp1(36) = "&#36;"
    cEsp1(37) = "&#37;"
    cEsp1(38) = "&#38;&amp;"
    cEsp1(39) = "&#39;"
    cEsp1(40) = "&#40;"
    cEsp1(41) = "&#41;"
    cEsp1(42) = "&#42;"
    cEsp1(43) = "&#43;"
    cEsp1(44) = "&#44;"
    cEsp1(45) = "&#45;"
    cEsp1(46) = "&#46;"
    cEsp1(47) = "&#47;"
    cEsp1(48) = "&#48;"
    cEsp1(49) = "&#49;"
    cEsp1(50) = "&#50;"
    cEsp1(51) = "&#51;"
    cEsp1(52) = "&#52;"
    cEsp1(53) = "&#53;"
    cEsp1(54) = "&#54;"
    cEsp1(55) = "&#55;"
    cEsp1(56) = "&#56;"
    cEsp1(57) = "&#57;"
    cEsp1(58) = "&#58;"
    cEsp1(59) = "&#59;"
    cEsp1(60) = "&#60;&lt;"
    cEsp1(61) = "&#61;"
    cEsp1(62) = "&#62;&gt;"
    cEsp1(63) = "&#63;"
    cEsp1(64) = "&#64;"
    cEsp1(65) = "&#65;"
    cEsp1(66) = "&#66;"
    cEsp1(67) = "&#67;"
    cEsp1(68) = "&#68;"
    cEsp1(69) = "&#69;"
    cEsp1(70) = "&#70;"
    cEsp1(71) = "&#71;"
    cEsp1(72) = "&#72;"
    cEsp1(73) = "&#73;"
    cEsp1(74) = "&#74;"
    cEsp1(75) = "&#75;"
    cEsp1(76) = "&#76;"
    cEsp1(77) = "&#77;"
    cEsp1(78) = "&#78;"
    cEsp1(79) = "&#79;"
    cEsp1(80) = "&#80;"
    cEsp1(81) = "&#81;"
    cEsp1(82) = "&#82;"
    cEsp1(83) = "&#83;"
    cEsp1(84) = "&#84;"
    cEsp1(85) = "&#85;"
    cEsp1(86) = "&#86;"
    cEsp1(87) = "&#87;"
    cEsp1(88) = "&#88;"
    cEsp1(89) = "&#89;"
    cEsp1(90) = "&#90;"
    cEsp1(91) = "&#91;"
    cEsp1(92) = "&#92;"
    cEsp1(93) = "&#93;"
    cEsp1(94) = "&#94;"
    cEsp1(95) = "&#95;"
    cEsp1(96) = "&#96;"
    cEsp1(97) = "&#97;"
    cEsp1(98) = "&#98;"
    cEsp1(99) = "&#99;"
    cEsp1(100) = "&#100;"
    cEsp1(101) = "&#101;"
    cEsp1(102) = "&#102;"
    cEsp1(103) = "&#103;"
    cEsp1(104) = "&#104;"
    cEsp1(105) = "&#105;"
    cEsp1(106) = "&#106;"
    cEsp1(107) = "&#107;"
    cEsp1(108) = "&#108;"
    cEsp1(109) = "&#109;"
    cEsp1(110) = "&#110;"
    cEsp1(111) = "&#111;"
    cEsp1(112) = "&#112;"
    cEsp1(113) = "&#113;"
    cEsp1(114) = "&#114;"
    cEsp1(115) = "&#115;"
    cEsp1(116) = "&#116;"
    cEsp1(117) = "&#117;"
    cEsp1(118) = "&#118;"
    cEsp1(119) = "&#119;"
    cEsp1(120) = "&#120;"
    cEsp1(121) = "&#121;"
    cEsp1(122) = "&#122;"
    cEsp1(123) = "&#123;"
    cEsp1(124) = "&#124;"
    cEsp1(125) = "&#125;"
    cEsp1(126) = "&#126;"
    cEsp1(160) = "&#160;&nbsp;"
    cEsp1(161) = "&#161;&iexcl;"
    cEsp1(162) = "&#162;&cent;"
    cEsp1(163) = "&#163;&pound;"
    cEsp1(164) = "&#164;&curren;"
    cEsp1(165) = "&#165;&yen;"
    cEsp1(166) = "&#166;&brvbar;"
    cEsp1(167) = "&#167;&sect;"
    cEsp1(168) = "&#168;&uml;"
    cEsp1(169) = "&#169;&copy;"
    cEsp1(170) = "&#170;&ordf;"
    cEsp1(171) = "&#171;&laquo;"
    cEsp1(172) = "&#172;&not;"
    cEsp1(173) = "&#173;&shy;"
    cEsp1(174) = "&#174;&reg;"
    cEsp1(175) = "&#175;&macr;"
    cEsp1(176) = "&#176;&deg;"
    cEsp1(177) = "&#177;&plusmn;"
    cEsp1(178) = "&#178;&sup2;"
    cEsp1(179) = "&#179;&sup3;"
    cEsp1(180) = "&#180;&acute;"
    cEsp1(181) = "&#181;&micro;"
    cEsp1(182) = "&#182;&para;"
    cEsp1(183) = "&#183;&middot;"
    cEsp1(184) = "&#184;&cedil;"
    cEsp1(185) = "&#185;&sup1;"
    cEsp1(186) = "&#186;&ordm;"
    cEsp1(187) = "&#187;&raquo;"
    cEsp1(188) = "&#188;&frac14;"
    cEsp1(189) = "&#189;&frac12;"
    cEsp1(190) = "&#190;&frac34;"
    cEsp1(191) = "&#191;&iquest;"
    cEsp1(192) = "&#192;&Agrave;"
    cEsp1(193) = "&#193;&Aacute;"
    cEsp1(194) = "&#194;&Acirc;"
    cEsp1(195) = "&#195;&Atilde;"
    cEsp1(196) = "&#196;&Auml;"
    cEsp1(197) = "&#197;&Aring;"
    cEsp1(198) = "&#198;&AElig;"
    cEsp1(199) = "&#199;&Ccedil;"
    cEsp1(200) = "&#200;&Egrave;"
    cEsp1(201) = "&#201;&Eacute;"
    cEsp1(202) = "&#202;&Ecirc;"
    cEsp1(203) = "&#203;&Euml;"
    cEsp1(204) = "&#204;&Igrave;"
    cEsp1(205) = "&#205;&Iacute;"
    cEsp1(206) = "&#206;&Icirc;"
    cEsp1(207) = "&#207;&Iuml;"
    cEsp1(208) = "&#208;&ETH;"
    cEsp1(209) = "&#209;&Ntilde;"
    cEsp1(210) = "&#210;&Ograve;"
    cEsp1(211) = "&#211;&Oacute;"
    cEsp1(212) = "&#212;&Ocirc;"
    cEsp1(213) = "&#213;&Otilde;"
    cEsp1(214) = "&#214;&Ouml;"
    cEsp1(215) = "&#215;&times;"
    cEsp1(216) = "&#216;&Oslash;"
    cEsp1(217) = "&#217;&Ugrave;"
    cEsp1(218) = "&#218;&Uacute;"
    cEsp1(219) = "&#219;&Ucirc;"
    cEsp1(220) = "&#220;&Uuml;"
    cEsp1(221) = "&#221;&Yacute;"
    cEsp1(222) = "&#222;&THORN;"
    cEsp1(223) = "&#223;&szlig;"
    cEsp1(224) = "&#224;&agrave;"
    cEsp1(225) = "&#225;&aacute;"
    cEsp1(226) = "&#226;&acirc;"
    cEsp1(227) = "&#227;&atilde;"
    cEsp1(228) = "&#228;&auml;"
    cEsp1(229) = "&#229;&aring;"
    cEsp1(230) = "&#230;&aelig;"
    cEsp1(231) = "&#231;&ccedil;"
    cEsp1(232) = "&#232;&egrave;"
    cEsp1(233) = "&#233;&eacute;"
    cEsp1(234) = "&#234;&ecirc;"
    cEsp1(235) = "&#235;&euml;"
    cEsp1(236) = "&#236;&igrave;"
    cEsp1(237) = "&#237;&iacute;"
    cEsp1(238) = "&#238;&icirc;"
    cEsp1(239) = "&#239;&iuml;"
    cEsp1(240) = "&#240;&eth;"
    cEsp1(241) = "&#241;&ntilde;"
    cEsp1(242) = "&#242;&ograve;"
    cEsp1(243) = "&#243;&oacute;"
    cEsp1(244) = "&#244;&ocirc;"
    cEsp1(245) = "&#245;&otilde;"
    cEsp1(246) = "&#246;&ouml;"
    cEsp1(247) = "&#247;&divide;"
    cEsp1(248) = "&#248;&oslash;"
    cEsp1(249) = "&#249;&ugrave;"
    cEsp1(250) = "&#250;&uacute;"
    cEsp1(251) = "&#251;&ucirc;"
    cEsp1(252) = "&#252;&uuml;"
    cEsp1(253) = "&#253;&yacute;"
    cEsp1(254) = "&#254;&thorn;"
    cEsp1(255) = "&#255;&yuml;"
    cEsp1(256) = "&#338;"
    cEsp1(257) = "&#339;"
    cEsp1(258) = "&#352;"
    cEsp1(259) = "&#353;"
    cEsp1(260) = "&#376;"
    cEsp1(261) = "&#402;"
    cEsp1(262) = "&#8211;"
    cEsp1(263) = "&#8212;"
    cEsp1(264) = "&#8216;"
    cEsp1(265) = "&#8217;"
    cEsp1(266) = "&#8218;"
    cEsp1(267) = "&#8220;"
    cEsp1(268) = "&#8221;"
    cEsp1(269) = "&#8222;"
    cEsp1(270) = "&#8224;"
    cEsp1(271) = "&#8225;"
    cEsp1(272) = "&#8226;"
    cEsp1(273) = "&#8230;"
    cEsp1(274) = "&#8240;"
    cEsp1(275) = "&#8364;&euro;"
    cEsp1(276) = "&#8482;"
    
    cEsp2(34) = "&quot;"
    cEsp2(38) = "&amp;"
    cEsp2(60) = "&lt;"
    cEsp2(62) = "&gt;"
    cEsp2(160) = "&nbsp;"
    cEsp2(161) = "&iexcl;"
    cEsp2(162) = "&cent;"
    cEsp2(163) = "&pound;"
    cEsp2(164) = "&curren;"
    cEsp2(165) = "&yen;"
    cEsp2(166) = "&brvbar;"
    cEsp2(167) = "&sect;"
    cEsp2(168) = "&uml;"
    cEsp2(169) = "&copy;"
    cEsp2(170) = "&ordf;"
    cEsp2(171) = "&laquo;"
    cEsp2(172) = "&not;"
    cEsp2(173) = "&shy;"
    cEsp2(174) = "&reg;"
    cEsp2(175) = "&macr;"
    cEsp2(176) = "&deg;"
    cEsp2(177) = "&plusmn;"
    cEsp2(178) = "&sup2;"
    cEsp2(179) = "&sup3;"
    cEsp2(180) = "&acute;"
    cEsp2(181) = "&micro;"
    cEsp2(182) = "&para;"
    cEsp2(183) = "&middot;"
    cEsp2(184) = "&cedil;"
    cEsp2(185) = "&sup1;"
    cEsp2(186) = "&ordm;"
    cEsp2(187) = "&raquo;"
    cEsp2(188) = "&frac14;"
    cEsp2(189) = "&frac12;"
    cEsp2(190) = "&frac34;"
    cEsp2(191) = "&iquest;"
    cEsp2(192) = "&Agrave;"
    cEsp2(193) = "&Aacute;"
    cEsp2(194) = "&Acirc;"
    cEsp2(195) = "&Atilde;"
    cEsp2(196) = "&Auml;"
    cEsp2(197) = "&Aring;"
    cEsp2(198) = "&AElig;"
    cEsp2(199) = "&Ccedil;"
    cEsp2(200) = "&Egrave;"
    cEsp2(201) = "&Eacute;"
    cEsp2(202) = "&Ecirc;"
    cEsp2(203) = "&Euml;"
    cEsp2(204) = "&Igrave;"
    cEsp2(205) = "&Iacute;"
    cEsp2(206) = "&Icirc;"
    cEsp2(207) = "&Iuml;"
    cEsp2(208) = "&ETH;"
    cEsp2(209) = "&Ntilde;"
    cEsp2(210) = "&Ograve;"
    cEsp2(211) = "&Oacute;"
    cEsp2(212) = "&Ocirc;"
    cEsp2(213) = "&Otilde;"
    cEsp2(214) = "&Ouml;"
    cEsp2(215) = "&times;"
    cEsp2(216) = "&Oslash;"
    cEsp2(217) = "&Ugrave;"
    cEsp2(218) = "&Uacute;"
    cEsp2(219) = "&Ucirc;"
    cEsp2(220) = "&Uuml;"
    cEsp2(221) = "&Yacute;"
    cEsp2(222) = "&THORN;"
    cEsp2(223) = "&szlig;"
    cEsp2(224) = "&agrave;"
    cEsp2(225) = "&aacute;"
    cEsp2(226) = "&acirc;"
    cEsp2(227) = "&atilde;"
    cEsp2(228) = "&auml;"
    cEsp2(229) = "&aring;"
    cEsp2(230) = "&aelig;"
    cEsp2(231) = "&ccedil;"
    cEsp2(232) = "&egrave;"
    cEsp2(233) = "&eacute;"
    cEsp2(234) = "&ecirc;"
    cEsp2(235) = "&euml;"
    cEsp2(236) = "&igrave;"
    cEsp2(237) = "&iacute;"
    cEsp2(238) = "&icirc;"
    cEsp2(239) = "&iuml;"
    cEsp2(240) = "&eth;"
    cEsp2(241) = "&ntilde;"
    cEsp2(242) = "&ograve;"
    cEsp2(243) = "&oacute;"
    cEsp2(244) = "&ocirc;"
    cEsp2(245) = "&otilde;"
    cEsp2(246) = "&ouml;"
    cEsp2(247) = "&divide;"
    cEsp2(248) = "&oslash;"
    cEsp2(249) = "&ugrave;"
    cEsp2(250) = "&uacute;"
    cEsp2(251) = "&ucirc;"
    cEsp2(252) = "&uuml;"
    cEsp2(253) = "&yacute;"
    cEsp2(254) = "&thorn;"
    cEsp2(255) = "&yuml;"
    cEsp2(275) = "&euro;"

    cChar(32) = " "
    cChar(33) = "!"
    cChar(34) = """"
    cChar(35) = "#"
    cChar(36) = "$"
    cChar(37) = "%"
    cChar(38) = "&"
    cChar(39) = "'"
    cChar(40) = "("
    cChar(41) = ")"
    cChar(42) = "*"
    cChar(43) = "+"
    cChar(44) = ","
    cChar(45) = "-"
    cChar(46) = "."
    cChar(47) = "/"
    cChar(48) = "0"
    cChar(49) = "1"
    cChar(50) = "2"
    cChar(51) = "3"
    cChar(52) = "4"
    cChar(53) = "5"
    cChar(54) = "6"
    cChar(55) = "7"
    cChar(56) = "8"
    cChar(57) = "9"
    cChar(58) = ":"
    cChar(59) = ";"
    cChar(60) = "<"
    cChar(61) = "="
    cChar(62) = ">"
    cChar(63) = "?"
    cChar(64) = "@"
    cChar(65) = "A"
    cChar(66) = "B"
    cChar(67) = "C"
    cChar(68) = "D"
    cChar(69) = "E"
    cChar(70) = "F"
    cChar(71) = "G"
    cChar(72) = "H"
    cChar(73) = "I"
    cChar(74) = "J"
    cChar(75) = "K"
    cChar(76) = "L"
    cChar(77) = "M"
    cChar(78) = "N"
    cChar(79) = "O"
    cChar(80) = "P"
    cChar(81) = "Q"
    cChar(82) = "R"
    cChar(83) = "S"
    cChar(84) = "T"
    cChar(85) = "U"
    cChar(86) = "V"
    cChar(87) = "W"
    cChar(88) = "X"
    cChar(89) = "Y"
    cChar(90) = "Z"
    cChar(91) = "["
    cChar(92) = "\"
    cChar(93) = "]"
    cChar(94) = "^"
    cChar(95) = "_"
    cChar(96) = "`"
    cChar(97) = "a"
    cChar(98) = "b"
    cChar(99) = "c"
    cChar(100) = "d"
    cChar(101) = "e"
    cChar(102) = "f"
    cChar(103) = "g"
    cChar(104) = "h"
    cChar(105) = "i"
    cChar(106) = "j"
    cChar(107) = "k"
    cChar(108) = "l"
    cChar(109) = "m"
    cChar(110) = "n"
    cChar(111) = "o"
    cChar(112) = "p"
    cChar(113) = "q"
    cChar(114) = "r"
    cChar(115) = "s"
    cChar(116) = "t"
    cChar(117) = "u"
    cChar(118) = "v"
    cChar(119) = "w"
    cChar(120) = "x"
    cChar(121) = "y"
    cChar(122) = "z"
    cChar(123) = "{"
    cChar(124) = "|"
    cChar(125) = "}"
    cChar(126) = "~"
    cChar(160) = " "
    cChar(161) = "¡"
    cChar(162) = "¢"
    cChar(163) = "£"
    cChar(164) = "¤"
    cChar(165) = "¥"
    cChar(166) = "¦"
    cChar(167) = "§"
    cChar(168) = "¨"
    cChar(169) = "©"
    cChar(170) = "ª"
    cChar(171) = "«"
    cChar(172) = "¬"
    cChar(173) = "­"
    cChar(174) = "®"
    cChar(175) = "¯"
    cChar(176) = "°"
    cChar(177) = "±"
    cChar(178) = "²"
    cChar(179) = "³"
    cChar(180) = "´"
    cChar(181) = "µ"
    cChar(182) = "¶"
    cChar(183) = "·"
    cChar(184) = "¸"
    cChar(185) = "¹"
    cChar(186) = "º"
    cChar(187) = "»"
    cChar(188) = "¼"
    cChar(189) = "½"
    cChar(190) = "¾"
    cChar(191) = "¿"
    cChar(192) = "À"
    cChar(193) = "Á"
    cChar(194) = "Â"
    cChar(195) = "Ã"
    cChar(196) = "Ä"
    cChar(197) = "Å"
    cChar(198) = "Æ"
    cChar(199) = "Ç"
    cChar(200) = "È"
    cChar(201) = "É"
    cChar(202) = "Ê"
    cChar(203) = "Ë"
    cChar(204) = "Ì"
    cChar(205) = "Í"
    cChar(206) = "Î"
    cChar(207) = "Ï"
    cChar(208) = "Ð"
    cChar(209) = "Ñ"
    cChar(210) = "Ò"
    cChar(211) = "Ó"
    cChar(212) = "Ô"
    cChar(213) = "Õ"
    cChar(214) = "Ö"
    cChar(215) = "×"
    cChar(216) = "Ø"
    cChar(217) = "Ù"
    cChar(218) = "Ú"
    cChar(219) = "Û"
    cChar(220) = "Ü"
    cChar(221) = "Ý"
    cChar(222) = "Þ"
    cChar(223) = "ß"
    cChar(224) = "à"
    cChar(225) = "á"
    cChar(226) = "â"
    cChar(227) = "ã"
    cChar(228) = "ä"
    cChar(229) = "å"
    cChar(230) = "æ"
    cChar(231) = "ç"
    cChar(232) = "è"
    cChar(233) = "é"
    cChar(234) = "ê"
    cChar(235) = "ë"
    cChar(236) = "ì"
    cChar(237) = "í"
    cChar(238) = "î"
    cChar(239) = "ï"
    cChar(240) = "ð"
    cChar(241) = "ñ"
    cChar(242) = "ò"
    cChar(243) = "ó"
    cChar(244) = "ô"
    cChar(245) = "õ"
    cChar(246) = "ö"
    cChar(247) = "÷"
    cChar(248) = "ø"
    cChar(249) = "ù"
    cChar(250) = "ú"
    cChar(251) = "û"
    cChar(252) = "ü"
    cChar(253) = "ý"
    cChar(254) = "þ"
    cChar(255) = "ÿ"
    cChar(256) = "Œ"
    cChar(257) = "œ"
    cChar(258) = "Š"
    cChar(259) = "š"
    cChar(260) = "Ÿ"
    cChar(261) = "ƒ"
    cChar(262) = "–"
    cChar(263) = "—"
    cChar(264) = "‘"
    cChar(265) = "’"
    cChar(266) = "‚"
    cChar(267) = "“"
    cChar(268) = "”"
    cChar(269) = "„"
    cChar(270) = "†"
    cChar(271) = "‡"
    cChar(272) = "•"
    cChar(273) = "…"
    cChar(274) = "‰"
    cChar(275) = "€"
    cChar(276) = "™"
    '<V1.3.1>
End Function

Function fsReemplazaCarEspec(sGlsCampo As String) As String
    '<V1.3.1>
    Dim i   As Integer
    
    For i = 1 To 267
        sGlsCampo = Replace(sGlsCampo, cEsp1(i), cChar(i))
        sGlsCampo = Replace(sGlsCampo, cEsp2(i), cChar(i))
    Next i
    
    fsReemplazaCarEspec = sGlsCampo
    '</V1.3.1>
End Function

Function fsSoloLetras(sTexto As String) As String
    Dim i       As Integer
    Dim sAux    As String
    
    sAux = ""
    For i = 1 To Len(sTexto)
        Select Case LCase(Mid(sTexto, i, 1))
        Case "a" To "z", " "
            sAux = sAux + Mid(sTexto, i, 1)
        End Select
    Next i
    fsSoloLetras = sAux
End Function


