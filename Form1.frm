VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
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
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   14760
      TabIndex        =   13
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1335
      Left            =   14520
      TabIndex        =   12
      Top             =   6480
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid Grilla 
      Height          =   3615
      Left            =   600
      TabIndex        =   11
      Top             =   5160
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   8880
      TabIndex        =   5
      Top             =   3360
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
Dim ApellidoA As String
Dim Nombre As String
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
                        
                        ApellidoA = alumnos(alum).SetApellido & Mid(linea, subA, 1)
                        cont = cont + 1
                        
                    End If
                
'                Case Is = 1
'
'                    If Mid(linea, subA, 1) <> ";" Then
'
'                        Nombre = alumnos(alum).SetNombre & Mid(linea, subA, 1)
'                        Nombre = alumnos(alum).SetNombre
'                    Else
'
'                        cont = cont + 1
'
'                    End If
'
'                End Select
                
                alum = alum + 1
                
                Print alumnos(alum).GetApellido
                
                
        Next subA
        
        Loop
        
    Close #1
    
    
    
    
    
    
    
    
    
    
End Sub
Private Sub Command2_Click()

    Grilla.AddItem "Leandro" & vbTab & "20" & vbTab & "2004" & vbTab & "45467967" & vbTab
    Grilla.AddItem "Darius" & vbTab & "30" & vbTab & "2009" & vbTab & "Noxus" & vbTab

    Grilla.RemoveItem (1)



End Sub
Private Sub Command3_Click()
Dim Linea2 As String
Dim A As Integer
Dim comparar As String
    
    Open App.Path & "/alumnos.txt" For Input As #1
        
        Do Until EOF(1)
        
        Input #1, Linea2
        
        For A = 0 To Len(Linea2)
            comparar = Mid(Linea2, A, 1)
            
            if



    Close #1






End Sub
