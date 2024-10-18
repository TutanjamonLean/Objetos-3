VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   21675
   LinkTopic       =   "Form1"
   ScaleHeight     =   9765
   ScaleWidth      =   21675
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   3960
      TabIndex        =   5
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14040
      TabIndex        =   10
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      TabIndex        =   9
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   8
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   7
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "E-mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14040
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "N. de Telefono"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "F. de Nacimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Nombres"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Apellido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   0
      Top             =   1200
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim alumnos() As New datos
Dim linea As String
Private Sub Command1_Click()
Dim subA As Integer
Dim cont As Integer
Dim alum As Integer

    
    Open App.Path & "/alumnos.txt" For Input As #1
        
        Do Until EOF(1)
        
        ReDim Preserve alumnos(alum)
        
        Input #1, linea
        
            For subA = 0 To Len(linea)
                
                Select Case cont
                Case Is = 0
                    
                    If Mid(linea, subA, 1) <> ";" Then
                        
                        alumnos(alumn).SetApellido = alumnos(alum).SetApellido & Mid(linea, subA, 1)
                        
                    Else
                        
                        cont = cont + 1
                        
                    End If
                
                Case Is = 1
                    
                    If Mid(linea, subA, 1) <> ";" Then
                        
                        alumnos(alum).SetNombre = alumnos(alum).SetNombre & Mid(linea, subA, 1)
                        
                    Else
                        
                        cont = cont + 1
                        
                    End If
                    
                End Select
                
                alum = alum + 1
                
            Next subA
        
        
        
        Print GetApellido
        Print GetNombre
        
        Loop
        
    Close #1
    
    
    
    
    
    
    
    
    
    
End Sub
