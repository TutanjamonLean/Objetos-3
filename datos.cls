VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "datos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private materia As String
Private nombre As String
Private Nacimiento As String
Private Telefono As Long
Private nota As String
Public Sub constructor(thisNombre As String, thisNacimiento As String, ThisNota As String, thisTelefono As Long, thisMateria As String)
    materia = thisMateria
    nombre = thisNombre
    Naciiento = thisNaciiento
    Telefono = thisTelefono
    nota = ThisNota
End Sub
'Public Function GetApellido() As String
'    GetApellido = Apellido
'End Function
'Public Function GetNombre() As String
'    GetNombre
'End Function
'Public Function GetNacimiento() As String
'    GetNacimiento
'End Function
'Public Function GetTelefono() As Long
'    GetTelefono
'End Function
'Public Function GetEmail() As String
'    GetEmail
'End Function
'Public Sub SetNombre(thisNombre As String)
'
'    Nombre = thisNombre
'
'End Sub
'Public Sub SetApellido(thisApellido As String)
'
'    Apellido = thisApellido
'
'End Sub
'
'Public Sub SetNacimiento(thisNacimiento As String)
'
'    Nacimiento = thisNacimiento
'
'End Sub
'Public Sub SetTelefono(thisTelefono As Long)
'
'    Telefono = thisTelefono
'
'End Sub
'Public Sub SetEmail(thisEmail As String)
'
'    Email = thisEmail
'
'End Sub
