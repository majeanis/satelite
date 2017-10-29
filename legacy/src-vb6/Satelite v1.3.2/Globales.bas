Attribute VB_Name = "Globales"
Public Type rRegParametros
   Nombre       As String
   Tipo         As String
   Opcional     As Boolean
   valor        As String
   Ayuda        As String
   TipoAyuda    As String
End Type
Public gaRegParametros() As rRegParametros

Public Type rRegColumnas
   Nombre       As String
   Tipo         As String
   Opcional     As Boolean
   valor        As String
   Ayuda        As String
   TipoAyuda    As String
End Type
Public gaRegColumnas() As rRegColumnas

