Attribute VB_Name = "Funciones_02"
'*** Módulo estándar que contiene procedimientos para trabajar con archivos.   ***
'*** Forma parte de la aplicación de ejemplo Bloc de notas MDI.                ***
'*********************************************************************************
Option Explicit

Sub FileOpenProc()
    Dim intRetVal
    On Error Resume Next
    Dim strOpenFileName As String
    frmMdiPadre.CMDialog1.Filename = ""
    frmMdiPadre.CMDialog1.ShowOpen
    If Err <> 32755 Then    ' El usuario eligió Cancelar.
        strOpenFileName = frmMdiPadre.CMDialog1.Filename
        ' Si el archivo es mayor de 65K, no se puede
        ' abrir, de modo que se cancela la operación.
        If FileLen(strOpenFileName) > 65000 Then
            MsgBox "El archivo es demasiado grande para abrirlo."
            Exit Sub
        End If
        
        OpenFile (strOpenFileName)
        UpdateFileMenu (strOpenFileName)
        ' Muestra las barras de herramientas si no son visibles.
        If gToolsHidden Then
            frmMdiPadre.imgCutButton.Visible = True
            frmMdiPadre.imgCopyButton.Visible = True
            frmMdiPadre.imgPasteButton.Visible = True
            gToolsHidden = False
        End If
    End If
End Sub

Function GetFileName(Filename As Variant)
    ' Muestra un cuadro de diálogo Guardar como y devuelve un nombre de archivo.
    ' Si el usuario elige Cancelar, devuelve una cadena vacía.
    On Error Resume Next
    frmMdiPadre.CMDialog1.Filename = Filename
    frmMdiPadre.CMDialog1.ShowSave
    If Err <> 32755 Then    ' El usuario eligió Cancelar.
        GetFileName = frmMdiPadre.CMDialog1.Filename
    Else
        GetFileName = ""
    End If
End Function

Function OnRecentFilesList(Filename) As Integer
  Dim i     ' Variable contador.

  For i = 1 To 4
    If frmMdiPadre.mnuRecentFile(i).Caption = Filename Then
      OnRecentFilesList = True
      Exit Function
    End If
  Next i
    OnRecentFilesList = False
End Function

Sub OpenFile(Filename)
    Dim fIndex As Integer
    
    On Error Resume Next
    ' Abre el archivo seleccionado.
    Open Filename For Input As #1
    If Err Then
        MsgBox "No se puede abrir el archivo: " + Filename
        Exit Sub
    End If
    ' Cambia el puntero del mouse a reloj de arena.
    Screen.MousePointer = 11
    
    ' Modifica el título del formulario y lo muestra.
    fIndex = FindFreeIndex()
    Document(fIndex).Tag = fIndex
    Document(fIndex).Caption = UCase(Filename)
    Document(fIndex).Text1.Text = StrConv(InputB(LOF(1), 1), vbUnicode)
    FState(fIndex).Dirty = False
    Document(fIndex).Show
    Close #1
    ' Restablece el puntero del mouse.
    Screen.MousePointer = 0
End Sub

Sub SaveFileAs(Filename)
    On Error Resume Next
    Dim strContents As String

    ' Abre el archivo.
    Open Filename For Output As #1
    ' Coloca el contenido del bloc de notas en una variable.
    strContents = frmMdiPadre.ActiveForm.Text1.Text
    ' Muestra el puntero reloj de arena.
    Screen.MousePointer = 11
    ' Escribe el contenido de la variable en un archivo.
    Print #1, strContents
    Close #1
    ' Restablece el puntero del mouse.
    Screen.MousePointer = 0
    ' Establece el título del formulario.
    If Err Then
        MsgBox Error, 48, App.Title
    Else
        frmMdiPadre.ActiveForm.Caption = UCase(Filename)
        ' Restablece el indicador Dirty.
        FState(frmMdiPadre.ActiveForm.Tag).Dirty = False
    End If
End Sub

Sub UpdateFileMenu(Filename)
        Dim intRetVal As Integer
        ' Comprueba si el archivo abierto se encuentra en la matriz de controles del menú Archivo.
        intRetVal = OnRecentFilesList(Filename)
        If Not intRetVal Then
            ' Escribe el nombre del archivo abierto en el Registro del sistema.
            WriteRecentFiles (Filename)
        End If
        ' Actualiza la lista de los archivos abiertos recientemente en la matriz de controles
        ' del menú Archivo.
        GetRecentFiles
End Sub


