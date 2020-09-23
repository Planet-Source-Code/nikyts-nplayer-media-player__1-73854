VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form_Youtube 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Youtube downloader"
   ClientHeight    =   6990
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   7740
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Form_Youtube.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tranferir tudo"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4320
      TabIndex        =   23
      Top             =   3645
      Width           =   1500
   End
   Begin NPlayer.McImageList McImageList1 
      Left            =   6480
      Top             =   2280
      _ExtentX        =   450
      _ExtentY        =   873
      Images0         =   "Form_Youtube.frx":57E2
      ImageCount      =   1
   End
   Begin NPlayer.McListBox List1 
      Height          =   1815
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   3201
      Picture         =   "Form_Youtube.frx":5B40
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
      BorderStyle     =   131072
      IconFocus       =   0   'False
      RowHeight       =   18
      ShowIcon        =   -1  'True
      SelectionStyle  =   0
      Path            =   "C:\NPlayer v2\"
      IconExtractSize =   1
   End
   Begin VB.TextBox Text_Clip 
      BackColor       =   &H00FC6C03&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1800
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Informações"
      Height          =   2475
      Left            =   480
      TabIndex        =   3
      Top             =   4140
      Width           =   6825
      Begin ComctlLib.ProgressBar Pb_Progresso 
         Height          =   285
         Left            =   1530
         TabIndex        =   12
         Top             =   2040
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   503
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Video:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   380
         Width           =   555
      End
      Begin VB.Label Label_Nome_do_Video 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FC6C03&
         Height          =   200
         Left            =   1560
         TabIndex        =   19
         Top             =   380
         Width           =   4995
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Progresso:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2070
         Width           =   930
      End
      Begin VB.Label Label_Percentagem 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1530
         TabIndex        =   11
         Top             =   1710
         Width           =   2565
      End
      Begin VB.Label Label123 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Percentagem:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1740
         Width           =   1200
      End
      Begin VB.Label Label_Guardado 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1530
         TabIndex        =   9
         Top             =   1410
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Guardado:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1410
         Width           =   915
      End
      Begin VB.Label Label_Restante 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1530
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Restante:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label_Tamanho 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1530
         TabIndex        =   5
         Top             =   750
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tamanho:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   750
         Width           =   870
      End
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   3360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2640
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Text_Endereco 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1620
      TabIndex        =   0
      Top             =   960
      Width           =   4305
   End
   Begin NPlayer.N_Button Botao_Fechar 
      Height          =   285
      Left            =   4800
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      Picture         =   "Form_Youtube.frx":5B5C
      PictureHover    =   "Form_Youtube.frx":64E2
      PictureDown     =   "Form_Youtube.frx":6E68
   End
   Begin NPlayer.N_Button Botao_Adicionar 
      Height          =   330
      Left            =   6045
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   960
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      Picture         =   "Form_Youtube.frx":77EE
      PictureHover    =   "Form_Youtube.frx":8CE0
      PictureDown     =   "Form_Youtube.frx":A1D2
   End
   Begin NPlayer.N_Button Botao_Baixar 
      Height          =   330
      Left            =   6000
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      Picture         =   "Form_Youtube.frx":B6C4
      PictureHover    =   "Form_Youtube.frx":CBB6
      PictureDown     =   "Form_Youtube.frx":E0A8
   End
   Begin NPlayer.N_Button Botao_Remover 
      Height          =   330
      Left            =   480
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      Picture         =   "Form_Youtube.frx":F59A
      PictureHover    =   "Form_Youtube.frx":10A8C
      PictureDown     =   "Form_Youtube.frx":11F7E
   End
   Begin NPlayer.N_Button Botao_Minimizar 
      Height          =   285
      Left            =   4200
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      Picture         =   "Form_Youtube.frx":13470
      PictureHover    =   "Form_Youtube.frx":13A66
      PictureDown     =   "Form_Youtube.frx":1405C
   End
   Begin NPlayer.N_Button Botao_Opcoes 
      Height          =   270
      Left            =   960
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   405
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   476
      Picture         =   "Form_Youtube.frx":14652
      PictureHover    =   "Form_Youtube.frx":15274
      PictureDown     =   "Form_Youtube.frx":15E96
   End
   Begin NPlayer.N_Button Botao_Procurar 
      Height          =   270
      Left            =   120
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   405
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   476
      Picture         =   "Form_Youtube.frx":16AB8
      PictureHover    =   "Form_Youtube.frx":176DA
      PictureDown     =   "Form_Youtube.frx":182FC
   End
   Begin VB.Image Skin_Top_Menu 
      Enabled         =   0   'False
      Height          =   300
      Left            =   110
      Picture         =   "Form_Youtube.frx":18F1E
      Top             =   410
      Width           =   450
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   405
      Left            =   100
      Picture         =   "Form_Youtube.frx":19690
      Top             =   0
      Width           =   315
   End
   Begin VB.Image Skin_Lateral_Esquerda 
      Enabled         =   0   'False
      Height          =   540
      Left            =   0
      Picture         =   "Form_Youtube.frx":19D92
      Top             =   400
      Width           =   105
   End
   Begin VB.Image Skin_Down_Esquerda 
      Enabled         =   0   'False
      Height          =   570
      Left            =   0
      Picture         =   "Form_Youtube.frx":1A134
      Top             =   840
      Width           =   105
   End
   Begin VB.Image Skin_Top_Esquerda 
      Enabled         =   0   'False
      Height          =   405
      Left            =   0
      Picture         =   "Form_Youtube.frx":1A506
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Skin_Down_Centro 
      Enabled         =   0   'False
      Height          =   120
      Left            =   105
      Picture         =   "Form_Youtube.frx":1A7D0
      Top             =   1290
      Width           =   990
   End
   Begin VB.Label Label_Titulo 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "           Youtube downloader - Electric nikyts"
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
      TabIndex        =   15
      Top             =   90
      Width           =   4695
   End
   Begin VB.Image Skin_Top_Direita 
      Enabled         =   0   'False
      Height          =   405
      Left            =   960
      Picture         =   "Form_Youtube.frx":1AE52
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Skin_Lateral_Direita 
      Enabled         =   0   'False
      Height          =   630
      Left            =   960
      Picture         =   "Form_Youtube.frx":1B11C
      Top             =   360
      Width           =   105
   End
   Begin VB.Image Skin_Down_Direita 
      Enabled         =   0   'False
      Height          =   570
      Left            =   960
      Picture         =   "Form_Youtube.frx":1B54E
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço:"
      Height          =   195
      Left            =   450
      TabIndex        =   2
      Top             =   990
      Width           =   870
   End
   Begin VB.Image Skin_Top_Centro 
      Height          =   405
      Left            =   110
      Picture         =   "Form_Youtube.frx":1B920
      Top             =   0
      Width           =   870
   End
End
Attribute VB_Name = "Form_Youtube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MOVER FORMULÁRIO
Dim H, v As Long

'Variáveis do progressbar
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_USER = &H400
Const CCM_FIRST = &H2000&
Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
Const PBM_SETBARCOLOR = (WM_USER + 9)

Private Function ChangePBForeColour(ByVal hwnd As Long, ByVal lColor As Long)
    'muda a cor da barra
    SendMessage hwnd, PBM_SETBARCOLOR, 0, ByVal lColor
End Function

Private Function ChangePBBackColour(ByVal hwnd As Long, ByVal lColor As Long)
    'muda a cor do fundo
    SendMessage hwnd, PBM_SETBKCOLOR, 0, ByVal lColor
End Function

Sub ResetControls()
    Text_Endereco.Text = ""
    Label_Nome_do_Video.Caption = ""
    Label_Tamanho.Caption = ""
    Label_Restante.Caption = ""
    Label_Guardado.Caption = ""
    Label_Percentagem.Caption = ""
    Pb_Progresso.Value = 0
    Text_Clip.Text = ""
End Sub

Private Sub Botao_Baixar_Click()
'    Dim i As Integer
'    For i = 0 To List1.ListCount - 1
    If List1.ListCount = 0 Then Exit Sub
    If Text_Clip.Text = "" Then
        MsgBox ("Selecione o video que pretende baixar.")
        Exit Sub
    Else
        DownloadVideo GetVideoInfo(Text_Clip.Text, Inet1), VideoName & ".flv" 'List1.List(i)
    End If
'    Next i
End Sub

Private Sub Botao_Adicionar_Click()
    Dim str1 As String
    If InStr(Text_Endereco.Text, "youtube.com/watch") Then
        str1 = Left(Text_Endereco.Text, 42)
        List1.AddItem str1, -1, 0
        Text_Endereco.Text = Empty
    End If
End Sub

Private Sub Botao_Fechar_Click()
    'Fechar o formulário
    Unload Me
End Sub

Private Sub Botao_Minimizar_Click()
    'Minimizar formulário
    Me.WindowState = 1
End Sub

Private Sub Botao_Procurar_Click()
    'Abrir o formulário do webbrowser
    Load Form_Browser
    With Form_Browser
        .Show
        .Combo_Site.Text = "www.youtube.com"
        .WebBrowser1.Navigate "www.youtube.com"
    End With
End Sub

Private Sub Botao_Remover_Click()
    'Remover linha selecionada
    If List1.ListCount = 0 Then Exit Sub
    List1.Remove List1.ListIndex
End Sub

Private Sub Form_Load()
    'Ajusta as cores do fundo e dos indicadores de progressbar
    ChangePBForeColour Pb_Progresso.hwnd, &HFC6C03
    ChangePBBackColour Pb_Progresso.hwnd, &HE0E0E0
    
    'Carregar icons na grelha
    Set List1.ImageList = McImageList1
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
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
       
    'Skin_Top_Menu
    Skin_Top_Menu.Stretch = True
    Skin_Top_Menu.Width = Me.Width - Skin_Top_Esquerda.Width - Skin_Top_Direita.Width
    Skin_Top_Menu.Left = Skin_Top_Esquerda.Left + Skin_Top_Esquerda.Width
    
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
    
    'Botao_Minimizar
    Botao_Minimizar.Top = 0
    Botao_Minimizar.Left = Botao_Fechar.Left - Botao_Minimizar.Width
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

Private Sub List1_Click()
    'Passar o conteudo da linha selecionada para a text clip
    If List1.ListCount = 0 Then Exit Sub
    Text_Clip.Text = List1.Text
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
'
'Private Sub mnuPaste_Click()
'    Dim strClipData As String
'
'    strClipData = Clipboard.GetText(vbCFText)
'
'    If InStr(strClipData, "youtube.com/watch") Then
'        'Text_Clip.Text = Clipboard.GetText(vbCFText)
'        'Text_Endereco.Text = Left(Text_Clip.Text, 42)
'        List1.AddItem Text_Endereco.Text
'    End If
'End Sub

Private Sub Text_Endereco_Change()
    'Text_Clip.Text = Clipboard.GetText(vbCFText)
    'Text_Endereco.Text = Left(Text_Clip.Text, 42)
End Sub
