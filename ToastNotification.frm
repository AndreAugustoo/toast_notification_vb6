VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Alert"
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Error"
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Success"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MostrarNotificacao Success, "Produto cadastrado com sucesso!", "Você cadastrou o produto 12345."
End Sub

Private Sub Command2_Click()
   MostrarNotificacao Error, "Falha ao cadastrar produto!", "Já existe um produto com o código 12345."
End Sub

Private Sub Command3_Click()
   MostrarNotificacao Alert, "Atenção, dados incompletos!", "Nome incompleto para o produto com o código 12345."
End Sub
