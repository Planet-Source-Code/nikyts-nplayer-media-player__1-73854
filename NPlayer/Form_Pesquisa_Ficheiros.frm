VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form_Pesquisa_Ficheiros 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6945
   ClientLeft      =   2700
   ClientTop       =   2655
   ClientWidth     =   8385
   ClipControls    =   0   'False
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6945
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Botao_Exportar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exportar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Extensão do ficheiro a pesquisar"
      Height          =   1935
      Left            =   360
      TabIndex        =   9
      Top             =   3960
      Width           =   7575
      Begin VB.TextBox Text_Extensao 
         BackColor       =   &H0000FF00&
         Height          =   285
         Left            =   5880
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton Opcao_Avi 
         BackColor       =   &H00FFFFFF&
         Caption         =   "*.avi"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Opcao_Mp3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "*.mp3"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Timer Timer_Progressbar 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   360
      Top             =   3120
   End
   Begin VB.CommandButton Botao_Iniciar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Iniciar"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Botao_Cancelar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   1335
   End
   Begin VB.ListBox Lista_Directorios 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   1665
      IntegralHeight  =   0   'False
      Left            =   5880
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin NPlayer.N_Button Botao_Fechar 
      Height          =   285
      Left            =   6480
      TabIndex        =   2
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      Picture         =   "Form_Pesquisa_Ficheiros.frx":0000
      PictureHover    =   "Form_Pesquisa_Ficheiros.frx":0986
      PictureDown     =   "Form_Pesquisa_Ficheiros.frx":130C
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin NPlayer.McListBox Grelha 
      Height          =   1815
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3201
      Picture         =   "Form_Pesquisa_Ficheiros.frx":1C92
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
      SelectionStyle  =   0
      Path            =   "C:\NPlayer v2\"
   End
   Begin NPlayer.McImageList McImageList1 
      Left            =   6360
      Top             =   360
      _ExtentX        =   661
      _ExtentY        =   873
      Images0         =   "Form_Pesquisa_Ficheiros.frx":1CAE
      ImageCount      =   1
   End
   Begin VB.Label Label_Contador 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Resultado da pesquisa"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   1920
   End
   Begin VB.Image Skin_Down_Centro 
      Enabled         =   0   'False
      Height          =   105
      Left            =   120
      Picture         =   "Form_Pesquisa_Ficheiros.frx":200C
      Top             =   1320
      Width           =   990
   End
   Begin VB.Image Skin_Down_Direita 
      Enabled         =   0   'False
      Height          =   570
      Left            =   960
      Picture         =   "Form_Pesquisa_Ficheiros.frx":25C6
      Top             =   840
      Width           =   105
   End
   Begin VB.Image Skin_Down_Esquerda 
      Enabled         =   0   'False
      Height          =   570
      Left            =   0
      Picture         =   "Form_Pesquisa_Ficheiros.frx":2998
      Top             =   840
      Width           =   105
   End
   Begin VB.Image Skin_Lateral_Direita 
      Enabled         =   0   'False
      Height          =   630
      Left            =   960
      Picture         =   "Form_Pesquisa_Ficheiros.frx":2D6A
      Top             =   360
      Width           =   105
   End
   Begin VB.Image Skin_Lateral_Esquerda 
      Enabled         =   0   'False
      Height          =   540
      Left            =   0
      Picture         =   "Form_Pesquisa_Ficheiros.frx":319C
      Top             =   400
      Width           =   105
   End
   Begin VB.Image Skin_Top_Esquerda 
      Enabled         =   0   'False
      Height          =   405
      Left            =   0
      Picture         =   "Form_Pesquisa_Ficheiros.frx":353E
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Skin_Top_Direita 
      Enabled         =   0   'False
      Height          =   405
      Left            =   960
      Picture         =   "Form_Pesquisa_Ficheiros.frx":3808
      Top             =   0
      Width           =   105
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Directórios"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   3480
      Width           =   7575
   End
   Begin VB.Label Label_Titulo 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "  Pesquisa automática"
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
      TabIndex        =   3
      Top             =   80
      Width           =   2895
   End
   Begin VB.Image Skin_Top_Centro 
      Height          =   405
      Left            =   0
      Picture         =   "Form_Pesquisa_Ficheiros.frx":3AD2
      Top             =   0
      Width           =   870
   End
End
Attribute VB_Name = "Form_Pesquisa_Ficheiros"
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

Option Explicit
'Variáveis para listar ficheiros automáticamente
Dim picHeight%, hLB&, FileSpec$, UseFileSpec%
Dim TotalDirs%, TotalFiles%, Running%
Dim WFD As WIN32_FIND_DATA, hItem&, hFile&
Const vbBackslash = "\"
Const vbAllFiles = "*.*"
Const vbKeyDot = 46

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

Private Sub Botao_Exportar_Click()
    'Carregar as listas com os ficheiros encontrados
    Label_Titulo.Caption = "  Pesquisa automática - Aguarde..."
    Dim Musica() As String
    Dim Linha As Integer
    
    'Contar as linhas da Lista_Directorios para depois remover as //
    With Form_Principal
        .Lista_Directorios.Clear
        .Grelha.Clear
        For Linha = 0 To Lista_Directorios.ListCount - 1
            .Lista_Directorios.AddItem Lista_Directorios.List(Linha)
            Musica = Split(Lista_Directorios.List(Linha), "\")
            .Grelha.AddItem Musica(UBound(Musica)), -1, 0 'e depois adicona na Grelha
            .Lista_Directorios.ListIndex = 0
            .Grelha.ListIndex = 0
            .Tocar_Media
        Next Linha
    End With
    Botao_Fechar_Click
End Sub

Private Sub Grelha_DbClick()
    'Adicionar o ficheiro selecionado
    Dim Musica() As String
    Dim Linha As Integer
    
    'Igualar o index
    Lista_Directorios.ListIndex = Grelha.ListIndex
    With Form_Principal
        .Lista_Directorios.AddItem Lista_Directorios.Text
        Musica = Split(Lista_Directorios.Text, "\")
        .Grelha.AddItem Musica(UBound(Musica)), -1, 0
    End With
End Sub

Private Sub Opcao_Avi_Click()
    'Selecionar a extensão do ficheiro a procurar
    Text_Extensao.Text = Opcao_Avi.Caption
End Sub

Private Sub Opcao_Mp3_Click()
    'Selecionar a extensão do ficheiro a procurar
    Text_Extensao.Text = Opcao_Mp3.Caption
End Sub

Private Sub Timer_Progressbar_Timer()
Static i As Integer
    'ativa a progress bar
    If i = 100 Then i = 0
    ProgressBar1.Value = i
    i = i + 1
    'If i = 100 Then Unload Me
End Sub

Private Sub Botao_Cancelar_Click()
    Botao_Fechar_Click
End Sub

Private Sub Botao_Fechar_Click()
    Set Form_Pesquisa_Ficheiros = Nothing
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Running% Then Running% = False
End Sub

Private Sub Form_Load()
    'Informações iniciais do formulário
    hLB& = Lista_Directorios.hwnd
    SendMessage hLB&, LB_INITSTORAGE, 30000&, ByVal 30000& * 200
    
    'Ajusta as cores do fundo e dos indicadores de progressbar
    ChangePBForeColour ProgressBar1.hwnd, &HFC6C03
    ChangePBBackColour ProgressBar1.hwnd, &HE0E0E0
    
    'Carregar icons na grelha
    Set Grelha.ImageList = McImageList1
    ProgressBar1.Value = 100
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

Private Sub Botao_Iniciar_Click()
    'Verificar se a extensao do ficheiro foi escolhida
    If Text_Extensao.Text = "" Then
        MsgBox ("Indique a extensão do ficherio que pretende procurar.")
        Exit Sub
    Else
    
    Dim Arquivos() As String
    Dim F As Integer
    Timer_Progressbar.Enabled = True
    Lista_Directorios.Clear
    'Arquivos = Split("*.mp3;*.wav;*.snd;*.au;*.aif;*.aifc;*.aiff;*.wma;*.mid;*.rmi;*.midi;*.AU;*.avi;*.wmv;*.mpg;*.mpeg;*.mp2;*.m1v;*.mpe", ";")
    Arquivos = Split(Text_Extensao.Text)
    
    For F = 0 To UBound(Arquivos)
        If Running% Then: Running% = False: Exit Sub
        Dim drvbitmask&, maxpwr%, pwr%
        On Error Resume Next
        FileSpec$ = Arquivos(F)
        If Len(FileSpec$) = 0 Then Exit Sub
        Running% = True
        UseFileSpec% = True
        drvbitmask& = GetLogicalDrives()
        If drvbitmask& Then
            maxpwr% = Int(Log(drvbitmask&) / Log(2))
            For pwr% = 0 To maxpwr%
                If Running% And (2 ^ pwr% And drvbitmask&) Then Call SearchDirs(Chr$(vbKeyA + pwr%) & ":\")
            Next
        End If
        
        'Carregar as listas com os ficheiros encontrados
        Dim Musica() As String
        Dim Linha As Integer
        'Contar as linhas da Lista_Directorios para depois remover as //
        For Linha = 0 To Lista_Directorios.ListCount - 1
            Musica = Split(Lista_Directorios.List(Linha), "\")
            Grelha.AddItem Musica(UBound(Musica)), -1, 0 'e depois adicona na Grelha
        Next Linha
        
        Running% = False
        UseFileSpec% = False
        Label_Contador = "Resultado da pesquisa (" & Lista_Directorios.ListCount & " ficheiros encontrados)"
        Label1.Caption = "Concluído"
        Timer_Progressbar.Enabled = False
        ProgressBar1.Value = 100
        
        'Verificar o resultado da pesquisa
        If Lista_Directorios.ListCount <> 0 Then
            Botao_Exportar.Enabled = True
        End If
    Next F
    End If
End Sub

Private Sub SearchDirs(curpath$)  ' curpath$ is passed w/ trailing "\"
    Dim dirs%, dirbuf$(), i%
    Label1.Caption = curpath$
    DoEvents
    If Not Running% Then Exit Sub
    
    hItem& = FindFirstFile(curpath$ & vbAllFiles, WFD)
    If hItem& <> INVALID_HANDLE_VALUE Then
        
        Do
            If (WFD.dwFileAttributes And vbDirectory) Then
                
                If Asc(WFD.cFileName) <> vbKeyDot Then
                    TotalDirs% = TotalDirs% + 1
                    If (dirs% Mod 10) = 0 Then ReDim Preserve dirbuf$(dirs% + 10)
                    dirs% = dirs% + 1
                    dirbuf$(dirs%) = Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
                End If
            ElseIf Not UseFileSpec% Then
                TotalFiles% = TotalFiles% + 1
            End If
        Loop While FindNextFile(hItem&, WFD)
        Call FindClose(hItem&)
    End If
    
    If UseFileSpec% Then
        SendMessage hLB&, WM_SETREDRAW, 0, 0
        Call SearchFileSpec(curpath$)
        SendMessage hLB&, WM_VSCROLL, SB_BOTTOM, 0
        SendMessage hLB&, WM_SETREDRAW, 1, 0
    End If
    For i% = 1 To dirs%: SearchDirs curpath$ & dirbuf$(i%) & vbBackslash: Next i%
End Sub

Private Sub SearchFileSpec(curpath$)   ' curpath$ is passed w/ trailing "\"
    hFile& = FindFirstFile(curpath$ & FileSpec$, WFD)
    If hFile& <> INVALID_HANDLE_VALUE Then
        Do
            DoEvents
            If Not Running% Then Exit Sub
            SendMessage hLB&, LB_ADDSTRING, 0, _
                ByVal curpath$ & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
        Loop While FindNextFile(hFile&, WFD)
        Call FindClose(hFile&)
    End If
End Sub
