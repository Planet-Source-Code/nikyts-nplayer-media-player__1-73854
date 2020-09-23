VERSION 5.00
Begin VB.Form Form_Propriedades 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox Lista_Propriedades 
      Height          =   3765
      Left            =   480
      TabIndex        =   14
      Top             =   960
      Width           =   1935
   End
   Begin VB.Frame Frame_Tela 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   2640
      TabIndex        =   3
      Top             =   720
      Width           =   5775
      Begin VB.PictureBox Palete_de_Cores 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   1920
         ScaleHeight     =   510
         ScaleWidth      =   1995
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   1995
         Begin VB.Label Label_Skin 
            BackColor       =   &H007B7B7B&
            Caption         =   "Cinza"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   4
            Left            =   1560
            TabIndex        =   9
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label_Skin 
            BackColor       =   &H0003EE56&
            Caption         =   "Verde"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Index           =   3
            Left            =   1200
            TabIndex        =   8
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label_Skin 
            BackColor       =   &H0004F1E9&
            Caption         =   "Amarelo"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   7
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label_Skin 
            BackColor       =   &H00007DF2&
            Caption         =   "Laranja"
            ForeColor       =   &H000080FF&
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   6
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label_Skin 
            BackColor       =   &H00FF890A&
            Caption         =   "Azul"
            ForeColor       =   &H00FF890A&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.Label Label_Cor 
         BackColor       =   &H00FF890A&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Escolha a tela pretendida:"
         Height          =   195
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   1845
      End
      Begin VB.Image Image_Skin 
         Height          =   3405
         Left            =   0
         Picture         =   "Form_Propriedades.frx":0000
         Top             =   840
         Width           =   5175
      End
      Begin VB.Label Label_Skin_Escolhido 
         Alignment       =   2  'Center
         BackColor       =   &H00FF890A&
         Caption         =   "Azul"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.PictureBox Botao_Cancelar 
      Height          =   330
      Left            =   5520
      ScaleHeight     =   270
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   5160
      Width           =   1200
   End
   Begin VB.PictureBox Botao_Ok 
      Height          =   330
      Left            =   6840
      ScaleHeight     =   270
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   5160
      Width           =   1200
   End
   Begin NPlayer.N_Button Botao_Fechar 
      Height          =   285
      Left            =   7680
      TabIndex        =   13
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      Picture         =   "Form_Propriedades.frx":396E6
      PictureHover    =   "Form_Propriedades.frx":3A06C
      PictureDown     =   "Form_Propriedades.frx":3A9F2
   End
   Begin VB.Image Skin_Top_Direita 
      Enabled         =   0   'False
      Height          =   405
      Left            =   960
      Picture         =   "Form_Propriedades.frx":3B378
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Skin_Top_Esquerda 
      Enabled         =   0   'False
      Height          =   405
      Left            =   0
      Picture         =   "Form_Propriedades.frx":3B642
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Skin_Lateral_Esquerda 
      Enabled         =   0   'False
      Height          =   540
      Left            =   0
      Picture         =   "Form_Propriedades.frx":3B90C
      Top             =   400
      Width           =   105
   End
   Begin VB.Image Skin_Lateral_Direita 
      Enabled         =   0   'False
      Height          =   630
      Left            =   960
      Picture         =   "Form_Propriedades.frx":3BCAE
      Top             =   360
      Width           =   105
   End
   Begin VB.Image Skin_Down_Esquerda 
      Enabled         =   0   'False
      Height          =   570
      Left            =   0
      Picture         =   "Form_Propriedades.frx":3C0E0
      Top             =   840
      Width           =   105
   End
   Begin VB.Image Skin_Down_Direita 
      Enabled         =   0   'False
      Height          =   570
      Left            =   960
      Picture         =   "Form_Propriedades.frx":3C4B2
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label_Titulo 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "    Propriedades"
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
      TabIndex        =   2
      Top             =   80
      Width           =   2895
   End
   Begin VB.Image Skin_Down_Centro 
      Enabled         =   0   'False
      Height          =   105
      Left            =   120
      Picture         =   "Form_Propriedades.frx":3C884
      Top             =   1200
      Width           =   990
   End
   Begin VB.Image Skin_Top_Centro 
      Height          =   405
      Left            =   0
      Picture         =   "Form_Propriedades.frx":3CE3E
      Top             =   0
      Width           =   870
   End
End
Attribute VB_Name = "Form_Propriedades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   MEDIA PLAYER
'   COPYRIGHT © 2008 ELECTRIC NIKYTS ™  -  INFORMÁTICA & TECNOLOGIA
'   WWW.NIKYTS.COM.SAPO.PT
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'MOVER FORMULÁRIO
Dim H, v As Long

'Selecionar skin pretendido
Dim Skin_Escolhido As String

Private Sub Botao_Cancelar_Click()
    'Cancelar operação
    Me.Hide
End Sub

Private Sub Botao_Fechar_Click()
    'Fechar formulário
    Me.Hide
End Sub

Private Sub Botao_Ok_Click()
    'Verificar o skin escolhido
    Me.Hide
End Sub

Private Sub Form_Load()
    'Preencher a lista das propriedades
    With Lista_Propriedades
        .AddItem "Tela do programa"
        .ListIndex = 0
        .AddItem "Geral"
        .AddItem "Player"
        .AddItem "Opções"
    End With
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

Private Sub Label_Cor_Do_Skin_Click()

End Sub

Private Sub Label_Skin_Click(Index As Integer)
    'Selecionar o skin pretendido
    'Label_Skin_Escolhido.BackColor = Label_Skin(Index).BackColor
    Label_Cor.BackColor = Label_Skin(Index).BackColor
    Label_Skin_Escolhido.Caption = Label_Skin(Index).Caption
    
    'Pre-visualizar as imagens do skin selecionado
    Select Case Label_Skin(Index).Index
        Case 0
            Image_Skin.Picture = Form_Imagens.Skin_Azul.Picture
            Exit Sub
        Case 1
            Image_Skin.Picture = Form_Imagens.Skin_Laranja.Picture
            Exit Sub
        Case 2
            Image_Skin.Picture = Form_Imagens.Skin_Amarelo.Picture
            Exit Sub
        Case 3
            Image_Skin.Picture = Form_Imagens.Skin_Verde.Picture
            Exit Sub
        Case 4
            Image_Skin.Picture = Form_Imagens.Skin_Cinza.Picture
            Exit Sub
    End Select
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

