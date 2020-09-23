VERSION 5.00
Begin VB.Form Form_Pesquisa 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin NPlayer.McListBox Lista_Directorios 
      Height          =   2295
      Left            =   1560
      TabIndex        =   8
      Top             =   960
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4048
      Picture         =   "Form_Pesquisa.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   131072
      ShowIcon        =   -1  'True
      Mode            =   5
      Path            =   "C:\NPlayer v2\"
   End
   Begin VB.PictureBox Botao_Pesquisa 
      Height          =   225
      Left            =   6120
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   5
      Top             =   650
      Width           =   225
   End
   Begin VB.ListBox List1 
      BackColor       =   &H0000FF00&
      Height          =   1740
      Left            =   1800
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   1770
      Left            =   240
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin NPlayer.N_Button Botao_Fechar 
      Height          =   285
      Left            =   6720
      TabIndex        =   6
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      Picture         =   "Form_Pesquisa.frx":001C
      PictureHover    =   "Form_Pesquisa.frx":09A2
      PictureDown     =   "Form_Pesquisa.frx":1328
   End
   Begin NPlayer.McListBox Lista_Ficheiros 
      Height          =   3375
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5953
      Picture         =   "Form_Pesquisa.frx":1CAE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Path            =   "C:\NPlayer v2\"
   End
   Begin VB.Label Label_Directorio 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   320
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   4815
   End
   Begin VB.Image Skin_Down_Centro 
      Enabled         =   0   'False
      Height          =   120
      Left            =   105
      Picture         =   "Form_Pesquisa.frx":1CCA
      Top             =   1290
      Width           =   990
   End
   Begin VB.Image Skin_Top_Direita 
      Enabled         =   0   'False
      Height          =   405
      Left            =   960
      Picture         =   "Form_Pesquisa.frx":234C
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Skin_Top_Esquerda 
      Enabled         =   0   'False
      Height          =   405
      Left            =   0
      Picture         =   "Form_Pesquisa.frx":2616
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Skin_Lateral_Esquerda 
      Enabled         =   0   'False
      Height          =   540
      Left            =   0
      Picture         =   "Form_Pesquisa.frx":28E0
      Top             =   400
      Width           =   105
   End
   Begin VB.Image Skin_Lateral_Direita 
      Enabled         =   0   'False
      Height          =   630
      Left            =   960
      Picture         =   "Form_Pesquisa.frx":2C82
      Top             =   360
      Width           =   105
   End
   Begin VB.Image Skin_Down_Esquerda 
      Enabled         =   0   'False
      Height          =   570
      Left            =   0
      Picture         =   "Form_Pesquisa.frx":30B4
      Top             =   840
      Width           =   105
   End
   Begin VB.Image Skin_Down_Direita 
      Enabled         =   0   'False
      Height          =   570
      Left            =   960
      Picture         =   "Form_Pesquisa.frx":3486
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label_Titulo 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "   Abrir"
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
   Begin VB.Image Botao_Tudo 
      Height          =   330
      Left            =   240
      Picture         =   "Form_Pesquisa.frx":3858
      Top             =   4800
      Width           =   1200
   End
   Begin VB.Image Botao_Cancelar 
      Height          =   330
      Left            =   5040
      Picture         =   "Form_Pesquisa.frx":4D3A
      Top             =   4815
      Width           =   1200
   End
   Begin VB.Image Botao_Ok 
      Height          =   330
      Left            =   6360
      Picture         =   "Form_Pesquisa.frx":621C
      Top             =   4815
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisar em:"
      Height          =   210
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1020
   End
   Begin VB.Image Skin_Top_Centro 
      Height          =   405
      Left            =   0
      Picture         =   "Form_Pesquisa.frx":76FE
      Top             =   0
      Width           =   870
   End
End
Attribute VB_Name = "Form_Pesquisa"
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

Private Sub Botao_Cancelar_Click()
    'cancelar operação
    'Ocultar a lista dos directorios
    Lista_Directorios.Visible = False
    
    Unload Me
End Sub

Private Sub Botao_Cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animação do botão cancelar
    Botao_Ok.Picture = Form_Imagens.Ok_Normal.Picture
    Botao_Cancelar.Picture = Form_Imagens.Cancelar_Over.Picture
    Botao_Tudo.Picture = Form_Imagens.Tudo_Normal.Picture
End Sub

Private Sub Botao_Fechar_Click()
    Unload Me
End Sub

Private Sub Botao_Pesquisa_Click()
    'Efectuar pesquisa pelos directorios
    If Lista_Directorios.Visible = False Then
        Lista_Directorios.Visible = True
    Else
        Lista_Directorios.Visible = False
    End If
End Sub

Private Sub Botao_Tudo_Click()
    'Adicionar todos os ficheiros do file1 na lista do form janela
    'Ocultar a lista dos directorios
    Lista_Directorios.Visible = False
    
    'On Error Resume Next
    Dim i As Integer
    For i = 0 To File1.ListCount - 1
        List1.ListIndex = Lista_Ficheiros.ListIndex
        With Form_Principal
            .Grelha.AddItem File1.List(i), -1, 0
            .Lista_Directorios.AddItem File1.Path & "\" & List1.List(i)
        End With
    Next i
End Sub

Private Sub Botao_Tudo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animação do botão tudo
    Botao_Ok.Picture = Form_Imagens.Ok_Normal.Picture
    Botao_Cancelar.Picture = Form_Imagens.Cancelar_Normal.Picture
    Botao_Tudo.Picture = Form_Imagens.Tudo_Over.Picture
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Combo1_Change()
    'Efectuar pesquisa pelos directorios
    If Lista_Directorios.Visible = False Then
        Lista_Directorios.Visible = True
    Else
        Lista_Directorios.Visible = False
    End If
End Sub

Private Sub File1_DblClick()
    'Adicionar ficheiro á lista
    Dim i As Integer
    For i = 0 To File1.ListCount - 1
        If File1.Selected(i) Then
            Form_Principal.Grelha.AddItem DirPasta.Path & "\" & File1.List(i), -1, 0
            'Form_Principal.Grelha.AddItem File1.List(i)
            'Form_Principal.Grelha.AddItem File1.List(i)
        End If
    Next i
End Sub

Private Sub Form_Activate()
    'Definir cor da linha de selecao da Lista
    Lista_Directorios.SelColor = &HFC6C03
    Lista_Ficheiros.SelColor = &HFC6C03
End Sub

Private Sub Form_Click()
    'Ocultar a lista dos directorios
    Lista_Directorios.Visible = False
End Sub

Private Sub Form_Load()
    'Indicar a pasta actual
    Label_Directorio.Caption = Lista_Directorios.Text
    
    With Me
        .BackColor = vbWhite
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
End Sub

Private Sub Botao_Ok_Click()
    'Ocultar a lista dos directorios
    Lista_Directorios.Visible = False

    Unload Me
End Sub

Private Sub Botao_Ok_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animação do botao ok
    Botao_Ok.Picture = Form_Imagens.Ok_Over.Picture
    Botao_Cancelar.Picture = Form_Imagens.Cancelar_Normal.Picture
    Botao_Tudo.Picture = Form_Imagens.Tudo_Normal.Picture
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

    'Control boxs
    Botao_Fechar.Top = 0
    Botao_Fechar.Left = Me.Width - Botao_Fechar.Width - 80
End Sub

Private Sub Lista_Directorios_Click()
    List1.Clear
    Lista_Ficheiros.Clear
    File1.FileName = Lista_Directorios.Text
    Dim i As Integer
    For i = 0 To File1.ListCount - 1
        Lista_Ficheiros.AddItem File1.List(i)  'Lista_Directorios.Text & "\" &
        List1.AddItem File1.List(i)
    Next i
    
    'Indicar a pasta actual
    Label_Directorio.Caption = Lista_Directorios.Text
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

Private Sub Skin_Top_Centro_DblClick()
    'Maximizar\ Restaurar formulário
    If Me.WindowState = 0 Then
        'Maximizar formulário
        Me.WindowState = 2
        Botao_Maximizar.Visible = False
        Botao_Restaurar.Visible = True
    ElseIf Me.WindowState = 2 Then
        'Restaurar formulário
        Me.WindowState = 0
        Botao_Maximizar.Visible = True
        Botao_Restaurar.Visible = False
    End If
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

Private Sub Lista_Ficheiros_Click()
    'Ocultar a lista dos directorios
    List1.ListIndex = Lista_Ficheiros.ListIndex
    Lista_Directorios.Visible = False
End Sub

Private Sub Lista_Ficheiros_DbClick()
    'Carregar a grelha com o ficheiro pretendido
    With Form_Principal
        .Grelha.AddItem Lista_Ficheiros.Text, -1, 0
        .Lista_Directorios.AddItem File1.Path & "\" & Lista_Ficheiros.Text
    End With
End Sub

Private Sub Label_Directorio_Click()
    'Ocultar a lista dos directorios
    Lista_Directorios.Visible = False
End Sub
