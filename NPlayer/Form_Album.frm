VERSION 5.00
Begin VB.Form Form_Album 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   7305
   ClientLeft      =   975
   ClientTop       =   1035
   ClientWidth     =   9270
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
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7305
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text_Caminho 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Text            =   "c:\users\nikyts\music\"
      Top             =   720
      Width           =   6135
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   4935
      Left            =   4200
      TabIndex        =   2
      Top             =   1560
      Width           =   3375
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   4350
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   3405
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pesquisar no C:\"
      Height          =   405
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   3405
   End
   Begin NPlayer.N_Button Botao_Fechar 
      Height          =   285
      Left            =   7680
      TabIndex        =   5
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      Picture         =   "Form_Album.frx":0000
      PictureHover    =   "Form_Album.frx":0986
      PictureDown     =   "Form_Album.frx":130C
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Caminho:"
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   720
      Width           =   840
   End
   Begin VB.Image Skin_Down_Centro 
      Enabled         =   0   'False
      Height          =   105
      Left            =   120
      Picture         =   "Form_Album.frx":1C92
      Top             =   1200
      Width           =   990
   End
   Begin VB.Label Label_Titulo 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "    Pesquisar albuns/ pastas"
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
      TabIndex        =   6
      Top             =   80
      Width           =   2895
   End
   Begin VB.Image Skin_Down_Direita 
      Enabled         =   0   'False
      Height          =   570
      Left            =   960
      Picture         =   "Form_Album.frx":224C
      Top             =   840
      Width           =   105
   End
   Begin VB.Image Skin_Down_Esquerda 
      Enabled         =   0   'False
      Height          =   570
      Left            =   0
      Picture         =   "Form_Album.frx":261E
      Top             =   840
      Width           =   105
   End
   Begin VB.Image Skin_Lateral_Direita 
      Enabled         =   0   'False
      Height          =   630
      Left            =   960
      Picture         =   "Form_Album.frx":29F0
      Top             =   360
      Width           =   105
   End
   Begin VB.Image Skin_Lateral_Esquerda 
      Enabled         =   0   'False
      Height          =   540
      Left            =   0
      Picture         =   "Form_Album.frx":2E22
      Top             =   400
      Width           =   105
   End
   Begin VB.Image Skin_Top_Esquerda 
      Enabled         =   0   'False
      Height          =   405
      Left            =   0
      Picture         =   "Form_Album.frx":31C4
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Skin_Top_Direita 
      Enabled         =   0   'False
      Height          =   405
      Left            =   960
      Picture         =   "Form_Album.frx":348E
      Top             =   0
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Conteúdo das Patas:"
      Height          =   195
      Left            =   4200
      TabIndex        =   4
      Top             =   1320
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Diretório:"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   825
   End
   Begin VB.Image Skin_Top_Centro 
      Height          =   405
      Left            =   0
      Picture         =   "Form_Album.frx":3758
      Top             =   0
      Width           =   870
   End
End
Attribute VB_Name = "Form_Album"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   NPLAYER
'   COPYRIGHT © 2008 ELECTRIC NIKYTS ™  -  INFORMÁTICA & TECNOLOGIA
'   WWW.NIKYTS.COM.SAPO.PT
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'MOVER FORMULÁRIO
Dim H, v As Long
Option Explicit

Private Sub Botao_Fechar_Click()
    'Fechar o formulário
    Unload Me
End Sub

Private Sub Command1_Click()
    'Listar pasta existentes no caminho indicado
    ListSubDirs Text_Caminho.Text
End Sub

Private Sub List1_Click()
    ListFiles (List1.Text) + "\"
End Sub

Private Sub ListFiles(Path)
    List2.Clear
    On Error Resume Next
    Dim Count, D(), i, DirName
    DirName = Dir(Path, 6)
    Do While DirName <> ""
    If DirName <> "." And DirName <> ".." Then
    List2.AddItem DirName
    End If
    DirName = Dir
    Loop
End Sub

Private Sub ListSubDirs(Path)
    On Error Resume Next
    Dim Count, D(), i, DirName
    DirName = Dir(Path, 16)
    Do While DirName <> ""
    If DirName <> "." And DirName <> ".." Then
    If GetAttr(Path + DirName) = 16 Then
    If (Count Mod 10) = 0 Then
    ReDim Preserve D(Count + 10)
    End If
    Count = Count + 1
    D(Count) = DirName
    End If
    End If
    DirName = Dir
    Loop
    For i = 1 To Count
    List1.AddItem Path & D(i)
    ListSubDirs Path & D(i) & "\"
    Next i
    DoEvents
End Sub

Private Sub Form_Resize()
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
End Sub

Private Sub Label_Site_Click()
    'Abrir página
    Unload Me
    Load Form_Browser
    With Form_Browser
        .Combo_Site.Text = "http://www.nikyts.com.sapo.pt"
        .WebBrowser1.Navigate .Combo_Site.Text
        .Show
    End With
    'Call ShellExecute(0, "open", "http://www.nikyts.com.sapo.pt", vbNullString, vbNullString, SW_NORMAL)
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

