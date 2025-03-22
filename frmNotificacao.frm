VERSION 5.00
Begin VB.Form frmNotificacao 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5205
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrEncolher 
      Left            =   3360
      Top             =   720
   End
   Begin VB.Timer tmrFechar 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3840
      Top             =   720
   End
   Begin VB.Label lblFechar 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblMensagem 
      BackStyle       =   0  'Transparent
      Caption         =   "Sua mensagem aqui..."
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.Image imgIcone 
      Height          =   375
      Left            =   240
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblTitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo Aqui"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   300
      Width           =   3615
   End
End
Attribute VB_Name = "frmNotificacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblFechar_Click()
   Unload Me
End Sub

Private Sub tmrFechar_Timer()
    Unload Me
    tmrFechar.Enabled = False
End Sub

Private Sub Form_Move()
    frmSombra.Left = Me.Left - 8
    frmSombra.Top = Me.Top - 8
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim I As Integer
    Dim newHeight As Integer
    Dim newTop As Integer

    For I = Me.Height To 0 Step -50
        newHeight = I
        newTop = Me.Top + (Me.Height - newHeight)
        frmSombra.Height = newHeight
        Me.Height = newHeight
        frmSombra.Top = newTop
        Me.Top = newTop
        DoEvents
        Sleep 10
    Next I
    
End Sub
