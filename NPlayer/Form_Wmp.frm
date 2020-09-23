VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form_Wmp 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Video"
   ClientHeight    =   5505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7650
   Icon            =   "Form_Wmp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Barra_Slider 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   860
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   7935
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4440
      Width           =   7935
      Begin VB.PictureBox Picture_Slide_Som 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   5400
         ScaleHeight     =   150
         ScaleWidth      =   1500
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   60
         Width           =   1500
         Begin NPlayer.N_Button Slide_Som 
            Height          =   90
            Left            =   0
            TabIndex        =   9
            Top             =   30
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   159
            Picture         =   "Form_Wmp.frx":57E2
            PictureHover    =   "Form_Wmp.frx":59E4
            PictureDown     =   "Form_Wmp.frx":5BE6
         End
         Begin VB.Image Image_Barra_Slide_Som 
            Enabled         =   0   'False
            Height          =   150
            Left            =   0
            Picture         =   "Form_Wmp.frx":5DE8
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1485
         End
      End
      Begin VB.PictureBox SliderBar_Mascara 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   150
         Left            =   140
         ScaleHeight     =   150
         ScaleWidth      =   4695
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   50
         Width           =   4695
         Begin NPlayer.N_Button Slide_Mascara 
            Height          =   90
            Left            =   0
            TabIndex        =   7
            Top             =   30
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   159
            Picture         =   "Form_Wmp.frx":69E2
            PictureHover    =   "Form_Wmp.frx":6BE4
            PictureDown     =   "Form_Wmp.frx":6DE6
         End
         Begin VB.Image Image_Barra_Slide_Mascara 
            Enabled         =   0   'False
            Height          =   150
            Left            =   0
            Picture         =   "Form_Wmp.frx":6FE8
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4695
         End
      End
      Begin NPlayer.N_Button Botao_Antes 
         Height          =   360
         Left            =   480
         TabIndex        =   10
         Top             =   405
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         Picture         =   "Form_Wmp.frx":A202
         PictureHover    =   "Form_Wmp.frx":A914
         PictureDown     =   "Form_Wmp.frx":B026
      End
      Begin NPlayer.N_Button Botao_Primeiro 
         Height          =   360
         Left            =   120
         TabIndex        =   11
         Top             =   405
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         Picture         =   "Form_Wmp.frx":B738
         PictureHover    =   "Form_Wmp.frx":BE4A
         PictureDown     =   "Form_Wmp.frx":C55C
      End
      Begin NPlayer.N_Button Botao_Pause 
         Height          =   360
         Left            =   1380
         TabIndex        =   12
         Top             =   405
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   635
         Picture         =   "Form_Wmp.frx":CC6E
         PictureHover    =   "Form_Wmp.frx":D380
         PictureDown     =   "Form_Wmp.frx":DA92
      End
      Begin NPlayer.N_Button Botao_Pasta 
         Height          =   360
         Left            =   2085
         TabIndex        =   13
         Top             =   405
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         Picture         =   "Form_Wmp.frx":E1A4
         PictureHover    =   "Form_Wmp.frx":E916
         PictureDown     =   "Form_Wmp.frx":F088
      End
      Begin NPlayer.N_Button Botao_Play 
         Height          =   360
         Left            =   1035
         TabIndex        =   14
         Top             =   405
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   635
         Picture         =   "Form_Wmp.frx":F7FA
         PictureHover    =   "Form_Wmp.frx":FF0C
         PictureDown     =   "Form_Wmp.frx":1061E
      End
      Begin NPlayer.N_Button Botao_Stop 
         Height          =   360
         Left            =   1725
         TabIndex        =   15
         Top             =   405
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         Picture         =   "Form_Wmp.frx":10D30
         PictureHover    =   "Form_Wmp.frx":11442
         PictureDown     =   "Form_Wmp.frx":11B54
      End
      Begin NPlayer.N_Button Botao_Seguinte 
         Height          =   360
         Left            =   2685
         TabIndex        =   16
         Top             =   405
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         Picture         =   "Form_Wmp.frx":12266
         PictureHover    =   "Form_Wmp.frx":12978
         PictureDown     =   "Form_Wmp.frx":1308A
      End
      Begin NPlayer.N_Button Botao_Ultimo 
         Height          =   360
         Left            =   3045
         TabIndex        =   17
         Top             =   405
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         Picture         =   "Form_Wmp.frx":1379C
         PictureHover    =   "Form_Wmp.frx":13EAE
         PictureDown     =   "Form_Wmp.frx":145C0
      End
      Begin VB.Image Botao_Redimensionar 
         Height          =   240
         Left            =   7560
         Picture         =   "Form_Wmp.frx":14CD2
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label_Faixa_Actual 
         AutoSize        =   -1  'True
         BackColor       =   &H001E1F1D&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   3840
         TabIndex        =   19
         Top             =   495
         Width           =   105
      End
      Begin VB.Label Tempo_Estimado_Top 
         AutoSize        =   -1  'True
         BackColor       =   &H001E1F1D&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   5340
         TabIndex        =   18
         Top             =   495
         Width           =   750
      End
      Begin VB.Image Botao_Mudo 
         Height          =   165
         Left            =   5040
         Picture         =   "Form_Wmp.frx":14ED4
         Top             =   60
         Width           =   165
      End
      Begin VB.Image Image5 
         Enabled         =   0   'False
         Height          =   360
         Left            =   3600
         Picture         =   "Form_Wmp.frx":150A2
         Top             =   405
         Width           =   2715
      End
      Begin VB.Image Skin_Down_Centro 
         Enabled         =   0   'False
         Height          =   615
         Left            =   0
         Picture         =   "Form_Wmp.frx":183E4
         Stretch         =   -1  'True
         Top             =   240
         Width           =   9060
      End
   End
   Begin NPlayer.N_Button Botao_Restaurar 
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   503
      Picture         =   "Form_Wmp.frx":1A576
      PictureHover    =   "Form_Wmp.frx":1AB20
      PictureDown     =   "Form_Wmp.frx":1B0CA
   End
   Begin NPlayer.N_Button Botao_Maximizar 
      Height          =   285
      Left            =   4320
      TabIndex        =   3
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   503
      Picture         =   "Form_Wmp.frx":1B674
      PictureHover    =   "Form_Wmp.frx":1BC1E
      PictureDown     =   "Form_Wmp.frx":1C1C8
   End
   Begin NPlayer.N_Button Botao_Fechar 
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      Picture         =   "Form_Wmp.frx":1C772
      PictureHover    =   "Form_Wmp.frx":1D0F8
      PictureDown     =   "Form_Wmp.frx":1DA7E
   End
   Begin VB.Image Botao_On_Top 
      Height          =   300
      Left            =   120
      Picture         =   "Form_Wmp.frx":1E404
      ToolTipText     =   "Por cima"
      Top             =   0
      Width           =   405
   End
   Begin WMPLibCtl.WindowsMediaPlayer Wmp 
      Height          =   2640
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   3975
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   0   'False
      baseURL         =   ""
      volume          =   25
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   7011
      _cy             =   4657
   End
   Begin VB.Label Label_Titulo 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "              Video"
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
      Width           =   2535
   End
   Begin VB.Image Skin_Top_Direita 
      Enabled         =   0   'False
      Height          =   405
      Left            =   960
      Picture         =   "Form_Wmp.frx":1EAD6
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Skin_Lateral_Esquerda 
      Enabled         =   0   'False
      Height          =   540
      Left            =   0
      Picture         =   "Form_Wmp.frx":1EDA0
      Top             =   400
      Width           =   105
   End
   Begin VB.Image Skin_Lateral_Direita 
      Enabled         =   0   'False
      Height          =   630
      Left            =   960
      Picture         =   "Form_Wmp.frx":1F142
      Top             =   360
      Width           =   105
   End
   Begin VB.Image Skin_Down_Esquerda 
      Enabled         =   0   'False
      Height          =   615
      Left            =   0
      Picture         =   "Form_Wmp.frx":1F574
      Top             =   840
      Width           =   105
   End
   Begin VB.Image Skin_Down_Direita 
      Enabled         =   0   'False
      Height          =   615
      Left            =   960
      Picture         =   "Form_Wmp.frx":1F98E
      Top             =   840
      Width           =   105
   End
   Begin VB.Image Skin_Top_Esquerda 
      Enabled         =   0   'False
      Height          =   405
      Left            =   0
      Picture         =   "Form_Wmp.frx":1FDA8
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Skin_Top_Centro 
      Height          =   405
      Left            =   110
      Picture         =   "Form_Wmp.frx":20072
      Top             =   0
      Width           =   870
   End
End
Attribute VB_Name = "Form_Wmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   NPlayer
'   COPYRIGHT © 2010 ELECTRIC NIKYTS ™  -  INFORMÁTICA & TECNOLOGIA
'   WWW.NIKYTS.COM.SAPO.PT
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Declaração de variáveis
'VARIÁVERIS DO SLIDER VIDEO MASCARA
Dim Tx_2 As Integer, Ty_2 As Integer, DN_2 As Boolean
Dim Txa_2 As Integer, DNa_2 As Boolean
Dim Tyb_2, DNb_2 As Boolean
Dim NewLeft_2 As Integer

'VARIÁVERIS DO SLIDER SOM
Dim TX_Som As Integer, Ty_Som As Integer, DN_Som As Boolean
Dim Txa_Som As Integer, DNa_Som As Boolean
Dim Tyb_Som, Dnb_Som As Boolean
Dim NewLeft_Som As Integer

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

'Variavel on top
Dim on_top As Boolean

Private Sub Botao_Antes_Click()
    Form_Principal.Botao_Antes_Click
End Sub

Private Sub Botao_Fechar_Click()
    'Fechar formulário
    Me.Hide
    With Form_Principal
        If .Barra_Biblioteca.Visible = False Then Exit Sub
            .Frame_Lista.Height = .Frame_Centro.Height - .Barra_Biblioteca.Height '- .Frame_Video.Height
            '.Frame_Video.Visible = True
    End With
End Sub

Private Sub Botao_Maximizar_Click()
    'Maximizar formulário
    Me.WindowState = 2
    
    Botao_Maximizar.Visible = False
    Botao_Restaurar.Visible = True
End Sub


Private Sub Botao_Mudo_Click()
    Form_Principal.Botao_Mudo_Click
End Sub

Private Sub Botao_On_Top_Click()
    'Colocar form on top ou não
    If on_top = True Then
        AlwaysOnTop Me, -2
        on_top = False
        Botao_On_Top.Picture = Form_Imagens.On_Top_Normal.Picture
    Else
        AlwaysOnTop Me, -1
        on_top = True
        Botao_On_Top.Picture = Form_Imagens.On_Top_Over.Picture
    End If
End Sub

Private Sub Botao_Pasta_Click()
    Form_Principal.Botao_Pasta_Click
End Sub

Private Sub Botao_Pause_Click()
    Form_Principal.Botao_Pause_Click
End Sub

Private Sub Botao_Play_Click()
    Form_Principal.Botao_Play_Click
End Sub

Private Sub Botao_Primeiro_Click()
    Form_Principal.Botao_Primeiro_Click
End Sub

Private Sub Botao_Restaurar_Click()
    'Restaurar formulário
    Me.WindowState = 0
    
    Botao_Maximizar.Visible = True
    Botao_Restaurar.Visible = False
End Sub

Private Sub Botao_Seguinte_Click()
    Form_Principal.Botao_Seguinte_Click
End Sub

Private Sub Botao_Stop_Click()
    Form_Principal.Botao_Stop_Click
End Sub

Private Sub Botao_Ultimo_Click()
    Form_Principal.Botao_Ultimo_Click
End Sub

Private Sub Form_Activate()
    Wmp.settings.mute = True
End Sub

Private Sub Form_Load()
    'Chamar o procedimento para contruir o formulário
    Desenhar_Formulario
    
    'Volume do media player
    Wmp.settings.volume = 0
    
    'Colocar o formulário por cima dos outros
    AlwaysOnTop Me, -1
    on_top = True
    Botao_On_Top.Picture = Form_Imagens.On_Top_Over.Picture
End Sub

Private Sub Form_Resize()
    'Chamar o procedimento para contruir o formulário
    Desenhar_Formulario
End Sub

Private Sub Label_Titulo_DblClick()
    'Maximizar/ Restaurar formulário
    If Me.WindowState = 0 Then
        Me.WindowState = 2
    Else
        Me.WindowState = 0
    End If
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

Private Sub Skin_Top_Centro_DblClick()
    Label_Titulo_DblClick
End Sub

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

Private Sub Slide_Mascara_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DNa_2 = True
    Txa_2 = X
End Sub

Private Sub Slide_Mascara_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DNa_2 Then
        NewLeft_2 = Slide_Mascara.Left + X - Txa_2
        If NewLeft_2 < Image_Barra_Slide_Mascara.Left + 5 Then
            NewLeft_2 = Image_Barra_Slide_Mascara.Left + 5
        End If
        If NewLeft_2 > Image_Barra_Slide_Mascara.Width + Image_Barra_Slide_Mascara.Left - 8 - Slide_Mascara.Width Then
            NewLeft_2 = Image_Barra_Slide_Mascara.Width + Image_Barra_Slide_Mascara.Left - 8 - Slide_Mascara.Width
        End If
        Slide_Mascara.Left = NewLeft_2
    End If
End Sub

Private Sub Slide_Mascara_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'On Error Resume Next
    Dim offseti As Single
    DNa_2 = False
    offseti = (Slide_Mascara.Left - Image_Barra_Slide_Mascara.Left - 3) / (Image_Barra_Slide_Mascara.Width - 10 - Slide_Mascara.Width)
    Form_Principal.Wmp.Controls.CurrentPosition = Int(Form_Principal.Wmp.currentMedia.Duration * offseti)
    Wmp.Controls.CurrentPosition = Int(Wmp.currentMedia.Duration * offseti)
    
    'Wmp.Controls.CurrentPosition = Int(Wmp.currentMedia.Duration * offseti)
    'Wmp.Controls.CurrentPosition = Int(Wmp.currentMedia.Duration * offseti)
End Sub

Private Sub Slide_Som_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DNa_Som = True
    Txa_Som = X
End Sub

Private Sub Slide_Som_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DNa_Som Then
        NewLeft_Som = Slide_Som.Left + X - Txa
        If NewLeft_Som < Image_Barra_Slide_Som.Left + 1 Then
            NewLeft_Som = Image_Barra_Slide_Som.Left + 1
        End If
        If NewLeft_Som > Image_Barra_Slide_Som.Width + Image_Barra_Slide_Som.Left - 7 - Slide_Som.Width Then
            NewLeft_Som = Image_Barra_Slide_Som.Width + Image_Barra_Slide_Som.Left - 7 - Slide_Som.Width
        End If
        Slide_Som.Left = NewLeft_Som
        Form_Principal.Slide_Som.Left = NewLeft_Som
    End If
End Sub

Private Sub Slide_Som_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim offseti As Single
    DNa_Som = False
    'offseti = (Slide_Som.Left - Image_Barra_Slide.Left - 3) / (Image_Barra_Slide.Width - 10 - Slide_Som.Width)

    'Verificar a posiçãp do slider do volume
    If Slide_Som.Left >= 0 And Slide_Som.Left <= 100 Then
        Form_Principal.Label_Volume.Caption = "1"
    ElseIf Slide_Som.Left > 100 And Slide_Som.Left <= 150 Then
        Form_Principal.Label_Volume.Caption = "2"
    ElseIf Slide_Som.Left > 150 And Slide_Som.Left <= 200 Then
        Form_Principal.Label_Volume.Caption = "3"
    ElseIf Slide_Som.Left > 200 And Slide_Som.Left <= 250 Then
        Form_Principal.Label_Volume.Caption = "4"
    ElseIf Slide_Som.Left > 250 And Slide_Som.Left <= 300 Then
        Form_Principal.Label_Volume.Caption = "5"
    ElseIf Slide_Som.Left > 300 And Slide_Som.Left <= 350 Then
        Form_Principal.Label_Volume.Caption = "6"
    ElseIf Slide_Som.Left > 350 And Slide_Som.Left <= 400 Then
        Form_Principal.Label_Volume.Caption = "7"
    ElseIf Slide_Som.Left > 400 And Slide_Som.Left <= 450 Then
        Form_Principal.Label_Volume.Caption = "8"
    ElseIf Slide_Som.Left > 450 And Slide_Som.Left <= 500 Then
        Form_Principal.Label_Volume.Caption = "9"
    ElseIf Slide_Som.Left > 500 And Slide_Som.Left <= 550 Then
        Form_Principal.Label_Volume.Caption = "10"
    ElseIf Slide_Som.Left > 550 And Slide_Som.Left <= 600 Then
        Form_Principal.Label_Volume.Caption = "11"
    ElseIf Slide_Som.Left > 600 And Slide_Som.Left <= 650 Then
        Form_Principal.Label_Volume.Caption = "12"
    ElseIf Slide_Som.Left > 650 And Slide_Som.Left <= 700 Then
        Form_Principal.Label_Volume.Caption = "13"
    ElseIf Slide_Som.Left > 700 And Slide_Som.Left <= 750 Then
        Form_Principal.Label_Volume.Caption = "14"
    ElseIf Slide_Som.Left > 750 And Slide_Som.Left <= 800 Then
        Form_Principal.Label_Volume.Caption = "15"
    ElseIf Slide_Som.Left > 800 And Slide_Som.Left <= 850 Then
        Form_Principal.Label_Volume.Caption = "16"
    ElseIf Slide_Som.Left > 850 And Slide_Som.Left <= 900 Then
        Form_Principal.Label_Volume.Caption = "17"
    ElseIf Slide_Som.Left > 900 And Slide_Som.Left <= 960 Then
        Form_Principal.Label_Volume.Caption = "18"
    ElseIf Slide_Som.Left > 960 And Slide_Som.Left <= 1040 Then
        Form_Principal.Label_Volume.Caption = "19"
    ElseIf Slide_Som.Left > 1040 And Slide_Som.Left <= 1110 Then
        Form_Principal.Label_Volume.Caption = "20"
    End If

    Form_Principal.Verificar_Volume
End Sub

Public Sub Desenhar_Formulario()
    'Desenhar formulário
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
    
    'Barra_Slider
    Barra_Slider.Top = Me.Height - Barra_Slider.Height
    Barra_Slider.Width = Me.Width - Skin_Lateral_Esquerda.Width - Skin_Lateral_Direita.Width
    Barra_Slider.Left = Skin_Lateral_Esquerda.Left + Skin_Lateral_Esquerda.Width
    
     'Skin_Down_Esquerda
    Skin_Down_Esquerda.Top = Me.Height - Skin_Down_Esquerda.Height
    Skin_Down_Esquerda.Left = 0
    
    'Skin_Down_Centro
    'Skin_Down_Centro.Top = Me.Height - Skin_Down_Centro.Height
    Skin_Down_Centro.Stretch = True
    Skin_Down_Centro.Width = Barra_Slider.Width ' - Skin_Down_Esquerda.Width - Skin_Down_Direita.Width
    Skin_Down_Centro.Left = 0
    
    'Skin_Down_Direita
    Skin_Down_Direita.Top = Me.Height - Skin_Down_Direita.Height
    Skin_Down_Direita.Left = Skin_Down_Esquerda.Width + Skin_Down_Centro.Width
    
    'Label_Titulo
    Label_Titulo.Width = Me.Width - Skin_Top_Direita.Width - Botao_Fechar.Width

    'Botao_Fechar boxs
    Botao_Fechar.Top = 0
    Botao_Fechar.Left = Me.Width - Botao_Fechar.Width - 80
    
    'Botao_Maximizar
    Botao_Maximizar.Top = 0
    Botao_Maximizar.Left = Botao_Fechar.Left - Botao_Maximizar.Width
       
    'Botao_Restaurar
    Botao_Restaurar.Top = 0
    Botao_Restaurar.Left = Botao_Maximizar.Left
    
    'Wmp
    Wmp.Height = Me.Height - Skin_Top_Centro.Height - Barra_Slider.Height
    Wmp.Top = Skin_Top_Centro.Top + Skin_Top_Centro.Height
    Wmp.Width = Me.Width - Skin_Lateral_Esquerda.Width - Skin_Lateral_Direita.Width
    Wmp.Left = Skin_Lateral_Esquerda.Left + Skin_Lateral_Esquerda.Width
    
    'Botao_Redimensionar
    If Me.Height <> Skin_Top_Centro.Height Then
        Botao_Redimensionar.Top = Barra_Slider.Height - Botao_Redimensionar.Height
        Botao_Redimensionar.Left = Barra_Slider.Width - Botao_Redimensionar.Width
    End If
Exit Sub
CORRIGIR_ERRO:
    Me.Height = 1200
    Me.Width = 3000
End Sub
