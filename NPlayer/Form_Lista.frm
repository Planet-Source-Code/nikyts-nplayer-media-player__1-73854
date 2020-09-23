VERSION 5.00
Begin VB.Form Form_Lista 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin NPlayer.McListBox List1 
      Height          =   1215
      Left            =   840
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2143
      Picture         =   "Form_Lista.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      SelColor        =   16542723
      BorderStyle     =   0
      RowHeight       =   18
      ShowIcon        =   -1  'True
      SelectionStyle  =   0
      Path            =   "C:\NPlayer v2\"
   End
   Begin NPlayer.N_Button Botao_Fechar 
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      Picture         =   "Form_Lista.frx":001C
      PictureHover    =   "Form_Lista.frx":09A2
      PictureDown     =   "Form_Lista.frx":1328
   End
   Begin VB.Image Botao_Redimensionar 
      Height          =   240
      Left            =   3600
      Picture         =   "Form_Lista.frx":1CAE
      Top             =   6960
      Width           =   225
   End
   Begin VB.Image Skin_Down_Centro 
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      Picture         =   "Form_Lista.frx":1FF0
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label Label_Titulo 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "   Lista"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   4695
   End
   Begin VB.Image Skin_Top_Direita 
      Enabled         =   0   'False
      Height          =   405
      Left            =   960
      Picture         =   "Form_Lista.frx":4182
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Skin_Lateral_Esquerda 
      Enabled         =   0   'False
      Height          =   540
      Left            =   0
      Picture         =   "Form_Lista.frx":444C
      Top             =   400
      Width           =   105
   End
   Begin VB.Image Skin_Lateral_Direita 
      Enabled         =   0   'False
      Height          =   630
      Left            =   960
      Picture         =   "Form_Lista.frx":47EE
      Top             =   360
      Width           =   105
   End
   Begin VB.Image Skin_Down_Esquerda 
      Enabled         =   0   'False
      Height          =   615
      Left            =   0
      Picture         =   "Form_Lista.frx":4C20
      Top             =   840
      Width           =   105
   End
   Begin VB.Image Skin_Down_Direita 
      Enabled         =   0   'False
      Height          =   615
      Left            =   960
      Picture         =   "Form_Lista.frx":503A
      Top             =   840
      Width           =   105
   End
   Begin VB.Image Skin_Top_Esquerda 
      Enabled         =   0   'False
      Height          =   405
      Left            =   0
      Picture         =   "Form_Lista.frx":5454
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Skin_Top_Centro 
      Height          =   405
      Left            =   110
      Picture         =   "Form_Lista.frx":571E
      Top             =   0
      Width           =   870
   End
End
Attribute VB_Name = "Form_Lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'MOVER FORMULÁRIO
Dim H, v As Long

'Redimensionar formulário
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const WM_NCLBUTTONDOWN = &HA1
Const HTBOTTOMRIGHT = 17
Const HTCAPTION = 2

Private Sub Botao_Fechar_Click()
    'Fechar formulário
    'Me.Hide
    Unload Me
End Sub

Private Sub Botao_Redimensionar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Redimensionar o formulário conforme as dimensões pretendidas
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
    End If
End Sub

Private Sub Botao_Redimensionar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Redimensionar.MousePointer = 8
End Sub

Private Sub Form_Activate()
    Form_Wmp.Wmp.settings.mute = True
End Sub

Private Sub Form_Load()
    'Propriedades do formulário
    With Me
        .BackColor = vbWhite
    End With
    'Carregar icons na grelha
    Set List1.ImageList = Form_Principal.McImageList1
    
    'Definir cor da linha de selecao da Lista
    List1.SelColor = &HFC6C03
End Sub

Private Sub Form_Resize()
    On Error GoTo CORRIGIR_ERRO
    'Skin_Top_Esquerda
    Skin_Top_Esquerda.Top = 0
    Skin_Top_Esquerda.Left = 0
    
    'Skin_Top_Centro
    Skin_Top_Centro.Top = 0
    Skin_Top_Centro.Stretch = True
    Skin_Top_Centro.Width = Me.Width - Skin_Top_Esquerda.Width - Skin_Top_Direita.Width
    Skin_Top_Centro.Left = Skin_Top_Esquerda.Width
    
    'Skin_Top_Direita
    Skin_Top_Direita.Top = 0
    Skin_Top_Direita.Left = Skin_Top_Esquerda.Width + Skin_Top_Centro.Width
       
    'Skin_Lateral_Esquerda
    Skin_Lateral_Esquerda.Left = 0
    Skin_Lateral_Esquerda.Stretch = True
    Skin_Lateral_Esquerda.Height = Me.Height - Skin_Top_Esquerda.Height - Skin_Down_Esquerda.Height
    Skin_Lateral_Esquerda.Top = Skin_Top_Esquerda.Height
    
    'Skin_Lateral_Direita
    Skin_Lateral_Direita.Left = Me.Width - Skin_Lateral_Direita.Width
    Skin_Lateral_Direita.Stretch = True
    Skin_Lateral_Direita.Height = Me.Height - Skin_Top_Direita.Height - Skin_Down_Direita.Height
    Skin_Lateral_Direita.Top = Skin_Top_Direita.Height
    
     'Skin_Down_Esquerda
    Skin_Down_Esquerda.Top = Me.Height - Skin_Down_Esquerda.Height
    Skin_Down_Esquerda.Left = 0
    
    'Skin_Down_Centro
    Skin_Down_Centro.Top = Me.Height - Skin_Down_Centro.Height
    Skin_Down_Centro.Stretch = True
    Skin_Down_Centro.Width = Me.Width - Skin_Down_Esquerda.Width - Skin_Down_Direita.Width
    Skin_Down_Centro.Left = Skin_Top_Esquerda.Width
    
    'Skin_Down_Direita
    Skin_Down_Direita.Top = Me.Height - Skin_Down_Direita.Height
    Skin_Down_Direita.Left = Skin_Down_Esquerda.Width + Skin_Down_Centro.Width
    
    'Label_Titulo
    Label_Titulo.Width = Me.Width - Skin_Top_Direita.Width - Botao_Fechar.Width

    'Botao_Fechar boxs
    Botao_Fechar.Top = 0
    Botao_Fechar.Left = Me.Width - Botao_Fechar.Width - 80
    
    'List1
    List1.Height = Me.Height - Skin_Top_Centro.Height - Skin_Down_Centro.Height
    List1.Top = Skin_Top_Centro.Top + Skin_Top_Centro.Height
    List1.Width = Me.Width - Skin_Lateral_Esquerda.Width - Skin_Lateral_Direita.Width
    List1.Left = Skin_Lateral_Esquerda.Left + Skin_Lateral_Esquerda.Width
    
    'Botao_Redimensionar
    If Me.Height <> Skin_Top_Centro.Height Then
        Botao_Redimensionar.Top = Me.Height - Botao_Redimensionar.Height
        Botao_Redimensionar.Left = Me.Width - Botao_Redimensionar.Width
    End If
Exit Sub
CORRIGIR_ERRO:
    Me.Height = 4300
    Me.Width = 3500
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover formuário
    If Me.WindowState = 0 Then
        If Button = vbLeftButton Then
           Me.Left = Me.Left - (H - X)
           Me.Top = Me.Top - (v - Y)
        End If
    End If
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Valor de x e y
    H = X
    v = Y
End Sub
'
'Private Sub List1_Click()
'    'On Error Resume Next
'    Form_Principal.Grelha.ListIndex = List1.ListIndex
'    Form_Principal.Lista_Directorios.ListIndex = List1.ListIndex
'End Sub

Private Sub List1_DbClick()
    With Form_Principal
        .Grelha.ListIndex = List1.ListIndex
        '.Grelha.ListIndex = List1.ListIndex
        .Tocar_Media
    End With
End Sub

'Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyDelete Then
'        'Verificar se a lista contem ficheiros
'        If List1.ListCount = 0 Then Exit Sub
'        Dim Temp As String
'        Temp = Form_Principal.Lista_Directorios.Text
'        With Form_Principal
'            .Grelha.ListIndex = List1.ListIndex
'            .Lista_Directorios.ListIndex = List1.ListIndex
'            .Grelha.Remove Form_Principal.Grelha.ListIndex
'            .Lista_Directorios.RemoveItem Form_Principal.Lista_Directorios.ListIndex
'        End With
'        List1.Remove List1.ListIndex
'        If Form_Principal.Wmp.URL = Temp Then Form_Principal.Wmp.Controls.stop: Form_Principal.Timer_Duracao.Enabled = False
'    End If
'End Sub

Private Sub Skin_Top_Centro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Valor de x e y
    H = X
    v = Y
End Sub

Private Sub Skin_Top_Centro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover formuário
    If Me.WindowState = 0 Then
        If Button = vbLeftButton Then
           Me.Left = Me.Left - (H - X)
           Me.Top = Me.Top - (v - Y)
        End If
    End If
End Sub
