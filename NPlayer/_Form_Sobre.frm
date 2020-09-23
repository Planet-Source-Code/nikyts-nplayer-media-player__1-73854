VERSION 5.00
Begin VB.Form Form_Sobre 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Pesquisa autom·tica"
   ClientHeight    =   4605
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ùForm_Sobre.frx":0000
   ScaleHeight     =   4605
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin NPlayer.N_Button Botao_Fechar 
      Height          =   285
      Left            =   5700
      TabIndex        =   9
      Top             =   100
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      Picture         =   "ùForm_Sobre.frx":6040E
      PictureHover    =   "ùForm_Sobre.frx":60D94
      PictureDown     =   "ùForm_Sobre.frx":6171A
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Web:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   4260
      Width           =   450
   End
   Begin VB.Label Label_Companhia 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Companhia"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label_Comentario 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Comentario"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1680
      TabIndex        =   6
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Label Label_Direitos 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Direitos"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1680
      TabIndex        =   5
      Top             =   2880
      Width           =   660
   End
   Begin VB.Label Label_Site 
      AutoSize        =   -1  'True
      BackColor       =   &H00FC6C03&
      BackStyle       =   0  'Transparent
      Caption         =   "www.nikyts.com.sapo.pt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   4260
      Width           =   2085
   End
   Begin VB.Label Label_Autor 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Autor"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   2640
      Width           =   465
   End
   Begin VB.Label Label_Versao 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Vers„o:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1680
      TabIndex        =   2
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nplayer - Electric Nikyts"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   2610
   End
   Begin VB.Label Label_Titulo 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "   Sobre"
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
      Height          =   210
      Left            =   0
      TabIndex        =   0
      Top             =   130
      Width           =   630
   End
End
Attribute VB_Name = "Form_Sobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'MOVER FORMUL¡RIO
Dim H, v As Long

'API para abrir web
Private Const SW_NORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Botao_Fechar_Click()
    'Fechar formul·rio
    Unload Me
End Sub

Private Sub Form_Load()
    'Preencher dados do sistema
    Label_Companhia.Caption = "Companhia: " & App.CompanyName
    Label_Versao.Caption = "Vers„o: " & App.Major & "." & App.Minor & "." & App.Revision
    Label_Autor.Caption = "Autor: " & App.LegalTrademarks
    Label_Comentario.Caption = App.FileDescription
    Label_Direitos.Caption = App.LegalCopyright
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Valor de x e y
    H = X
    v = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Retomar cores originais
    Label_Site.ForeColor = &H808080
    
    'Mover formu·rio
    If Me.WindowState = 0 Then
        If Button = vbLeftButton Then
           Me.Left = Me.Left - (H - X)
           Me.Top = Me.Top - (v - Y)
        End If
    End If
End Sub

Private Sub Label_Site_Click()
    'Abrir p·gina
    Unload Me
    Load Form_Browser
    With Form_Browser
        .Combo_Site.Text = "http://www.nikyts.com.sapo.pt"
        .WebBrowser1.Navigate .Combo_Site.Text
        .Show
    End With
    'Call ShellExecute(0, "open", "http://www.nikyts.com.sapo.pt", vbNullString, vbNullString, SW_NORMAL)
End Sub

Private Sub Label_Site_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Definir cor da selecao da label link
    Label_Site.ForeColor = vbWhite
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover formu·rio
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
