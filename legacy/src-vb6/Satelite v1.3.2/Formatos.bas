Attribute VB_Name = "Formatos"
Option Explicit


Const gsGlsCeros = "0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" _
                 & "0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" _
                 & "0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"

Function fsValorFechaExcel(sValor As Variant, sFormatoIn As String) As String
    Dim sFechaAux       As String
    Dim nPosDia         As Integer
    Dim nPosMes         As Integer
    Dim nPosAño         As Integer
    Dim nNumDigDia      As Integer
    Dim nNumDigMes      As Integer
    Dim nNumDigAño      As Integer
    
    On Error GoTo Err_fsValorFechaExcel
    
    If sFormatoIn = "" Then
        '<INI SP1.2.1>
        'sFechaAux = Format(DateValue(sValor), "dd/mm/yyyy") & " " & Format(TimeValue(sValor), "hh:mm:ss")
        sFechaAux = Format(DateValue(sValor), "yyyy/mm/dd") & " " & Format(TimeValue(sValor), "hh:mm:ss")
        '<FIN SP1.2.1>
    Else
        nPosDia = InStr(LCase(sFormatoIn), "dd")
        nNumDigDia = 2
        If nPosDia = 0 Then
            nPosDia = InStr(LCase(sFormatoIn), "d")
            nNumDigDia = 1
        End If
        
        nPosMes = InStr(LCase(sFormatoIn), "mm")
        nNumDigMes = 2
        If nPosMes = 0 Then
            nPosMes = InStr(LCase(sFormatoIn), "m")
            nNumDigMes = 1
        End If
        
        nPosAño = InStr(LCase(sFormatoIn), "yyyy")
        nNumDigAño = 4
        If nPosAño = 0 Then
            nPosAño = InStr(LCase(sFormatoIn), "yy")
            nNumDigAño = 2
        End If
        
        sFechaAux = Mid(sValor, nPosDia, nNumDigDia) & "/" & Mid(sValor, nPosMes, nNumDigMes) & "/" & Mid(sValor, nPosAño, nNumDigAño)
    End If
    
    fsValorFechaExcel = sFechaAux
    Exit Function
    
Err_fsValorFechaExcel:
    fsValorFechaExcel = ""
    Exit Function
End Function

Function fsConvierteAFecha(sValor As Variant, sFormatoIn As String) As String

    fsConvierteAFecha = sValor

End Function

Function fsFormatoValorFecha(sValor As Variant, sFormatoIn As String, sFormatoOut As String) As String
    Dim sFechaAux       As String
    Dim nPosDia         As Integer
    Dim nPosMes         As Integer
    Dim nPosAño         As Integer
    Dim nNumDigDia      As Integer
    Dim nNumDigMes      As Integer
    Dim nNumDigAño      As Integer
    Dim dValorFecha     As Date
    
    On Error GoTo Err_fsFormatoValorFecha
    
    If sFormatoIn = "" Then
        fsFormatoValorFecha = Format(sValor, sFormatoOut)
    Else
        nPosDia = InStr(LCase(sFormatoIn), "dd")
        nNumDigDia = 2
        If nPosDia = 0 Then
            nPosDia = InStr(LCase(sFormatoIn), "d")
            nNumDigDia = 1
        End If
        
        nPosMes = InStr(LCase(sFormatoIn), "mm")
        nNumDigMes = 2
        If nPosMes = 0 Then
            nPosMes = InStr(LCase(sFormatoIn), "m")
            nNumDigMes = 1
        End If
        
        nPosAño = InStr(LCase(sFormatoIn), "yyyy")
        nNumDigAño = 4
        If nPosAño = 0 Then
            nPosAño = InStr(LCase(sFormatoIn), "yy")
            nNumDigAño = 2
        End If
        
        sFechaAux = Mid(sValor, nPosDia, nNumDigDia) & "/" & Mid(sValor, nPosMes, nNumDigMes) & "/" & Mid(sValor, nPosAño, nNumDigAño)
        dValorFecha = fdValorFecha(sFechaAux)
        If dValorFecha <> gdNullDate Then
            fsFormatoValorFecha = Format(dValorFecha, sFormatoOut)
        End If
    End If
    
    Exit Function
    
Err_fsFormatoValorFecha:
    fsFormatoValorFecha = ""
    Exit Function
End Function

Function fsFormatoValorNumerico(sNumero As Variant, sIndSeparadorMiles As String, sNumDecimales As String) As String
    Dim sFormato    As String
    
    On Error GoTo Err_fsFormatoValorNumerico
        
    sFormato = fsFormatoNumerico(sIndSeparadorMiles, sNumDecimales)

    fsFormatoValorNumerico = Format(CDbl(sNumero), sFormato)
    
    Exit Function
    
Err_fsFormatoValorNumerico:
    fsFormatoValorNumerico = ""
    Exit Function
End Function

Function fsFormatoNumerico(sSeparadorMiles As String, sNumDecimales As String) As String
    Dim nNumDecimales   As Integer
    
    nNumDecimales = Val(sNumDecimales)
    If nNumDecimales < 0 Then nNumDecimales = 0
    
    If sSeparadorMiles = "S" Then
        fsFormatoNumerico = "#,##0" & IIf(nNumDecimales > 0, "." & Left(gsGlsCeros, nNumDecimales), "")
    Else
        fsFormatoNumerico = "###0" & IIf(nNumDecimales > 0, "." & Left(gsGlsCeros, nNumDecimales), "")
    End If
End Function

Function fsValorDobleExcel(sValor As String) As String
    On Error GoTo ErrValorDoble
    
    fsValorDobleExcel = Replace(CDbl(sValor), ",", ".")
    Exit Function
    
ErrValorDoble:
    fsValorDobleExcel = ""
    Exit Function
End Function

Function fsValorIntegerExcel(sValor As String) As String
    On Error GoTo ErrValorInteger
    
    fsValorIntegerExcel = Replace(CLng(sValor), ",", ".")
    Exit Function
    
ErrValorInteger:
    fsValorIntegerExcel = ""
    Exit Function
End Function

