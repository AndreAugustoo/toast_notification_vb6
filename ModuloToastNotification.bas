Attribute VB_Name = "ModuloToastNotification"
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As T_Rect) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Type T_Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Enum E_TipoNotificacao
   Success = 1
   Error = 2
   Alert = 3
End Enum

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

Public Sub LoadPNG(ByRef P_ComponenteImagem As Image, P_CaminhoImagem As String)
    Dim StdPictureExInstance As New StdPictureEx
    
    Set P_ComponenteImagem.Picture = StdPictureExInstance.LoadPicture(P_CaminhoImagem)
End Sub

Public Sub Sleep(ByVal dwMilliseconds As Long)
    Dim StartTime As Long
    StartTime = GetTickCount
    Do While GetTickCount - StartTime < dwMilliseconds
        DoEvents
    Loop
End Sub

Private Sub AplicarBordasArredondadas(Form As Form)
    Dim hRgn As Long

    hRgn = CreateRoundRectRgn(0, 0, Form.Width \ Screen.TwipsPerPixelX, Form.Height \ Screen.TwipsPerPixelY, 20, 20)

    If hRgn <> 0 Then
        SetWindowRgn Form.hWnd, hRgn, True
        DeleteObject hRgn
    End If
End Sub

Private Function GetTaskbarSize() As T_Rect
    Dim hWnd As Long
    Dim rect As T_Rect

    hWnd = FindWindow("Shell_TrayWnd", vbNullString)

    If hWnd <> 0 Then
        GetWindowRect hWnd, rect
    End If
    
    GetTaskbarSize = rect
    
End Function

Private Sub AjustarTamanhoFormularioNotificacao()
   Dim screenWidth As Long
   Dim screenHeight As Long
   Dim rect As T_Rect

   rect = GetTaskbarSize

   screenWidth = GetSystemMetrics(SM_CXSCREEN)
   screenHeight = GetSystemMetrics(SM_CYSCREEN)

   frmNotificacao.Left = ((screenWidth * 15) - frmNotificacao.Width) - 100
   frmNotificacao.Top = ((screenHeight * 15) - frmNotificacao.Height) - ((rect.Bottom - rect.Top) * 15) - 100
   
   AplicarBordasArredondadas frmNotificacao
   AjustarTamanhoFormularioSombra
End Sub

Private Sub AjustarTamanhoFormularioSombra()
   With frmSombra
        .Width = frmNotificacao.Width + 30
        .Height = frmNotificacao.Height + 30
        .Left = frmNotificacao.Left - 15
        .Top = frmNotificacao.Top - 15
        .ZOrder 0
        .Show
    End With
    AplicarBordasArredondadas frmSombra
End Sub

Public Sub MostrarNotificacao(ByVal Tipo As E_TipoNotificacao, ByVal Titulo As String, ByVal Mensagem As String)
   
   With frmNotificacao
      AjustarTamanhoFormularioNotificacao

      .lblTitulo.Caption = Titulo
      .lblMensagem.Caption = Mensagem
      
      DefinirIcone Tipo
   
      .Show vbModeless
      
      .tmrFechar.Enabled = False
      .tmrFechar.Interval = 4000
      .tmrFechar.Enabled = True
   End With
   
End Sub

Private Sub DefinirIcone(ByVal Tipo As E_TipoNotificacao)
    With frmNotificacao
      Select Case Tipo
         Case Alert
            LoadPNG .imgIcone, "C:\Projects\VB6\ToastNotification\img\alert.png"
         Case Error
            LoadPNG .imgIcone, "C:\Projects\VB6\ToastNotification\img\error.png"
         Case Success
            LoadPNG .imgIcone, "C:\Projects\VB6\ToastNotification\img\success.png"
      End Select
   End With
End Sub
