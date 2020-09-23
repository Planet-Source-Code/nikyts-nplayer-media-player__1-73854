VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form_Browser_Opcoes 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame_Geral 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   4875
      Begin VB.Label Label_Guardar_Pagina_Em_Branco 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         Caption         =   "Nao"
         Height          =   195
         Left            =   3480
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Iniciar a página em branco"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Width           =   1875
      End
      Begin VB.Image Check_Abrir_Em_Branco 
         Height          =   195
         Left            =   0
         Picture         =   "Form_Browser_Opcoes.frx":0000
         Top             =   840
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Página ao iniciar o navegador:"
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   2160
      End
      Begin MSForms.TextBox Text_Pagina_Inicial 
         Height          =   285
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   4320
         VariousPropertyBits=   679495707
         BackColor       =   16777215
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "7620;503"
         Value           =   "http://www.nikyts.com.sapo.pt"
         BorderColor     =   12632256
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame_Propriedades 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   2400
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   4875
      Begin VB.Label Label_Guardar_Historico 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         Caption         =   "Sim"
         Height          =   195
         Left            =   3240
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label_Tela_Cheia 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         Caption         =   "Nao"
         Height          =   195
         Left            =   3240
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Iniciar o navegador com a tela cheia"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   960
         Width           =   2580
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Guardar histórico durante a navegação"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   2775
      End
      Begin VB.Image Check_Normal 
         Height          =   195
         Left            =   0
         Picture         =   "Form_Browser_Opcoes.frx":027E
         Top             =   960
         Width           =   210
      End
      Begin VB.Image Check_Over 
         Height          =   195
         Left            =   0
         Picture         =   "Form_Browser_Opcoes.frx":04FC
         Top             =   600
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Propriedades do navegador:"
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   2010
      End
   End
   Begin VB.Frame Frame_Favoritos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   2400
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   4875
      Begin NPlayer.McListBox Lista_Favoritos 
         Height          =   3135
         Left            =   0
         TabIndex        =   16
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   5530
         Picture         =   "Form_Browser_Opcoes.frx":077A
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
         IconFocus       =   0   'False
         RowHeight       =   18
         SelectionStyle  =   0
         AutoHideScrollBars=   -1  'True
         Path            =   "C:\NPlayer v2\"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione o endereço que pretende eliminar e prima ""Delete""."
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   4380
      End
   End
   Begin NPlayer.N_Button Botao_Fechar 
      Height          =   285
      Left            =   6480
      TabIndex        =   14
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      Picture         =   "Form_Browser_Opcoes.frx":0796
      PictureHover    =   "Form_Browser_Opcoes.frx":111C
      PictureDown     =   "Form_Browser_Opcoes.frx":1AA2
   End
   Begin NPlayer.McListBox Lista_Opcoes 
      Height          =   3735
      Left            =   360
      TabIndex        =   15
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   6588
      Picture         =   "Form_Browser_Opcoes.frx":2428
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
      IconFocus       =   0   'False
      RowHeight       =   18
      SelectionStyle  =   0
      AutoHideScrollBars=   -1  'True
      Path            =   "C:\NPlayer v2\"
   End
   Begin VB.Image Botao_Ok 
      Height          =   330
      Left            =   6120
      Picture         =   "Form_Browser_Opcoes.frx":2444
      Top             =   4920
      Width           =   1200
   End
   Begin VB.Image Botao_Cancelar 
      Height          =   330
      Left            =   4800
      Picture         =   "Form_Browser_Opcoes.frx":3926
      Top             =   4920
      Width           =   1200
   End
   Begin VB.Image Skin_Top_Direita 
      Enabled         =   0   'False
      Height          =   405
      Left            =   960
      Picture         =   "Form_Browser_Opcoes.frx":4E08
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Skin_Top_Esquerda 
      Enabled         =   0   'False
      Height          =   405
      Left            =   0
      Picture         =   "Form_Browser_Opcoes.frx":50D2
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Skin_Lateral_Esquerda 
      Enabled         =   0   'False
      Height          =   540
      Left            =   0
      Picture         =   "Form_Browser_Opcoes.frx":539C
      Top             =   400
      Width           =   105
   End
   Begin VB.Image Skin_Lateral_Direita 
      Enabled         =   0   'False
      Height          =   630
      Left            =   960
      Picture         =   "Form_Browser_Opcoes.frx":573E
      Top             =   360
      Width           =   105
   End
   Begin VB.Image Skin_Down_Esquerda 
      Enabled         =   0   'False
      Height          =   570
      Left            =   0
      Picture         =   "Form_Browser_Opcoes.frx":5B70
      Top             =   840
      Width           =   105
   End
   Begin VB.Image Skin_Down_Direita 
      Enabled         =   0   'False
      Height          =   570
      Left            =   960
      Picture         =   "Form_Browser_Opcoes.frx":5F42
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label_Titulo 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "   Opções do navegador"
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
      Top             =   80
      Width           =   2895
   End
   Begin VB.Image Skin_Down_Centro 
      Enabled         =   0   'False
      Height          =   105
      Left            =   120
      Picture         =   "Form_Browser_Opcoes.frx":6314
      Top             =   1320
      Width           =   990
   End
   Begin VB.Image Skin_Top_Centro 
      Height          =   405
      Left            =   0
      Picture         =   "Form_Browser_Opcoes.frx":68CE
      Top             =   0
      Width           =   870
   End
End
Attribute VB_Name = "Form_Browser_Opcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   NPLAYER
'   COPYRIGHT © 2010 ELECTRIC NIKYTS ™  -  INFORMÁTICA & TECNOLOGIA
'   WWW.NIKYTS.COM.SAPO.PT
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'MOVER FORMULÁRIO
Dim H, v As Long

Private Sub Botao_Cancelar_Click()
    'Cancelar operação
    Unload Me
End Sub

Private Sub Botao_Fechar_Click()
    'Fechar formulário
    Unload Me
End Sub

Private Sub Botao_Ok_Click()
    'Finalizar o acesso do formulário opcoes do browser
    Unload Me
End Sub

Private Sub Check_Abrir_Em_Branco_Click()
    'Selecionar a opcao "Iniciar a página em branco"
    If Text_Pagina_Inicial.Enabled = True Then
        Text_Pagina_Inicial.Enabled = False
        Check_Abrir_Em_Branco.Picture = Form_Imagens.Check_Over.Picture
    Else
        Text_Pagina_Inicial.Enabled = True
        Check_Abrir_Em_Branco.Picture = Form_Imagens.Check_Normal.Picture
    End If
End Sub

Private Sub Form_Activate()
    'Definir cor da linha de selecao da Lista
    Lista_Opcoes.SelColor = &HFC6C03
    Lista_Favoritos.SelColor = &HFC6C03
End Sub

Private Sub Form_Load()
    'Propriedades iniciais do formulário
    With Lista_Opcoes
        .AddItem "Geral"
        .AddItem "Favoritos"
        .AddItem "Propriedades"
    End With
    
    'Posicionar as frames
    Frame_Favoritos.Top = Frame_Geral.Top
    Frame_Favoritos.Left = Frame_Geral.Left
    Frame_Propriedades.Top = Frame_Geral.Top
    Frame_Propriedades.Left = Frame_Geral.Left
    
    'Selecionar 1º linha da lista de opcoes, a qual se refere á visualizacao da frame_geral
    Lista_Opcoes.ListIndex = 0
End Sub

Private Sub Form_Resize()
    'Desenhar formulário
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
    Label_Titulo.Top = 80
    Label_Titulo.Width = Me.Width - Skin_Top_Direita.Width - Botao_Fechar.Width

    'Botao_Fechar boxs
    Botao_Fechar.Top = 0
    Botao_Fechar.Left = Me.Width - Botao_Fechar.Width - 80
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.WindowState = 0 Then
        If Button = vbLeftButton Then
           Me.Left = Me.Left - (H - X)
           Me.Top = Me.Top - (v - Y)
        End If
    End If
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    H = X
    v = Y
End Sub

Private Sub Skin_Top_Centro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.WindowState = 0 Then
        If Button = vbLeftButton Then
           Me.Left = Me.Left - (H - X)
           Me.Top = Me.Top - (v - Y)
        End If
    End If
End Sub

Private Sub Skin_Top_Centro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    H = X
    v = Y
End Sub

Private Sub Lista_Opcoes_Click()
    'Selecionar opcoes pretendicas
    Select Case Lista_Opcoes.Text
        Case "Geral"
            Frame_Geral.Visible = True
            Frame_Favoritos.Visible = False
            Frame_Propriedades.Visible = False
        Exit Sub
        
        Case "Favoritos"
            Frame_Geral.Visible = False
            Frame_Favoritos.Visible = True
            Frame_Propriedades.Visible = False
        Exit Sub
        
        Case "Propriedades"
            Frame_Geral.Visible = False
            Frame_Favoritos.Visible = False
            Frame_Propriedades.Visible = True
        Exit Sub
    End Select
End Sub

