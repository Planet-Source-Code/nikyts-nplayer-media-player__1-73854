VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form_Radio 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Radio"
   ClientHeight    =   2820
   ClientLeft      =   -60
   ClientTop       =   0
   ClientWidth     =   7335
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
   Icon            =   "Form_Radio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "Form_Radio.frx":57E2
   ScaleHeight     =   2820
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin NPlayer.N_Button Botao_Next 
      Height          =   330
      Left            =   180
      TabIndex        =   13
      Top             =   550
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      Picture         =   "Form_Radio.frx":48E34
      PictureHover    =   "Form_Radio.frx":4945E
      PictureDown     =   "Form_Radio.frx":49A88
   End
   Begin NPlayer.N_Button Botao_Stop 
      Height          =   555
      Left            =   180
      TabIndex        =   10
      Top             =   1500
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   979
      Picture         =   "Form_Radio.frx":4A0B2
      PictureHover    =   "Form_Radio.frx":4AB6C
      PictureDown     =   "Form_Radio.frx":4B592
   End
   Begin NPlayer.N_Button Botao_Play 
      Height          =   585
      Left            =   180
      TabIndex        =   8
      Top             =   900
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   1032
      Picture         =   "Form_Radio.frx":4BFB8
      PictureHover    =   "Form_Radio.frx":4CA66
      PictureDown     =   "Form_Radio.frx":4D514
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8640
      Top             =   0
   End
   Begin VB.Timer Timer_Artista 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4080
      Top             =   0
   End
   Begin VB.Timer Timer_Musica 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3720
      Top             =   0
   End
   Begin VB.Timer Timer_Progress 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3000
      Top             =   0
   End
   Begin VB.ListBox List1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      IntegralHeight  =   0   'False
      Left            =   7680
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin NPlayer.McListBox Grelha 
      Height          =   1815
      Left            =   3000
      TabIndex        =   5
      Top             =   540
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   3201
      Picture         =   "Form_Radio.frx":4DFC2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   328965
      ForeColor       =   6316128
      SelColor        =   328965
      SelForeColor    =   16542723
      BorderStyle     =   0
      IconFocus       =   0   'False
      RowHeight       =   18
      SelectionStyle  =   0
      AutoHideScrollBars=   -1  'True
      Path            =   "C:\NPlayer v2\"
   End
   Begin NPlayer.N_Button Botao_Pause 
      Height          =   585
      Left            =   180
      TabIndex        =   9
      Top             =   900
      Visible         =   0   'False
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   1032
      Picture         =   "Form_Radio.frx":4DFDE
      PictureHover    =   "Form_Radio.frx":4EA8C
      PictureDown     =   "Form_Radio.frx":4F53A
   End
   Begin VB.Label Label_Minimizar 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   2400
      TabIndex        =   15
      Top             =   120
      Width           =   120
   End
   Begin VB.Label Label_Fechar 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   2600
      TabIndex        =   14
      Top             =   120
      Width           =   120
   End
   Begin VB.Label Label_Mudo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FC6C03&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   330
      TabIndex        =   12
      ToolTipText     =   "Mudo"
      Top             =   2115
      Width           =   165
   End
   Begin VB.Label Label_Lista 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00050505&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   180
      TabIndex        =   11
      ToolTipText     =   "Ver lista"
      Top             =   2115
      Width           =   135
   End
   Begin VB.Image Image_Progress 
      Enabled         =   0   'False
      Height          =   120
      Index           =   1
      Left            =   1230
      Picture         =   "Form_Radio.frx":4FFE8
      Top             =   2550
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image_Progress 
      Enabled         =   0   'False
      Height          =   120
      Index           =   2
      Left            =   1230
      Picture         =   "Form_Radio.frx":504CA
      Top             =   2550
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image_Progress 
      Enabled         =   0   'False
      Height          =   120
      Index           =   3
      Left            =   1230
      Picture         =   "Form_Radio.frx":509AC
      Top             =   2550
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image_Progress 
      Enabled         =   0   'False
      Height          =   120
      Index           =   4
      Left            =   1230
      Picture         =   "Form_Radio.frx":50E8E
      Top             =   2550
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image_Progress 
      Enabled         =   0   'False
      Height          =   120
      Index           =   5
      Left            =   1230
      Picture         =   "Form_Radio.frx":51370
      Top             =   2550
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image_Progress 
      Enabled         =   0   'False
      Height          =   120
      Index           =   0
      Left            =   1230
      Picture         =   "Form_Radio.frx":51852
      Top             =   2550
      Width           =   735
   End
   Begin VB.Label Label_Titulo 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Rádio"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   150
      TabIndex        =   7
      Top             =   120
      Width           =   2025
   End
   Begin VB.Image Image_Capa_Zoom 
      Height          =   1785
      Left            =   720
      Stretch         =   -1  'True
      Top             =   570
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image Image_Capa 
      Height          =   1125
      Left            =   1065
      Stretch         =   -1  'True
      Top             =   870
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Alterar item atual"
      ForeColor       =   &H003E3F3F&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   2140
      Width           =   1800
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblinfo 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   9000
      TabIndex        =   3
      Top             =   840
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin WMPLibCtl.WindowsMediaPlayer MPlayer 
      Height          =   1215
      Left            =   7680
      TabIndex        =   1
      Top             =   1680
      Width           =   855
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1508
      _cy             =   2143
   End
   Begin VB.Label Label_Sem_Capa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FC6C03&
      BackStyle       =   0  'Transparent
      Caption         =   "Sem capa"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      TabIndex        =   6
      Top             =   1320
      Width           =   870
   End
End
Attribute VB_Name = "Form_Radio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AID As Long
Dim albumID As Integer
Dim Artist$
Dim Author$
Dim Label$
Dim Title$
Dim adType$
Dim oldTitle$
Dim Album$
Dim SmallCover$
Dim MedCover$
Dim LargeCover$
Private Type TGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare Function OleLoadPicturePath Lib "oleaut32.dll" (ByVal szURLorPath As Long, ByVal punkCaller As Long, ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long
Dim Capt$
Dim sCapt$
Dim LeftChar
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Const LB_ITEMFROMPOINT = &H1A9
Dim i As Integer
Dim eTitle$
Dim EMess$
Dim mError As Long
''Private WithEvents Tray As Class_Radio
Private SaveState As Integer

'MOVER FORMULÁRIO
Dim H, v As Long

'Variavel para ouvir/ mudo
Dim Mudo As Boolean

Private Sub Botao_Next_Click()
    MPlayer.Controls.Next
End Sub

Private Sub Botao_Pause_Click()
    MPlayer.Controls.pause
    Botao_Play.Visible = True
    Botao_Pause.Visible = False
End Sub

Private Sub Botao_Play_Click()
    MPlayer.Controls.play
    Botao_Pause.Visible = True
    Botao_Play.Visible = False
End Sub

Private Sub Botao_Stop_Click()
    MPlayer.Controls.stop
    Botao_Pause.Visible = True
    Botao_Play.Visible = False
End Sub

Private Sub Form_Load()
    'Propriedades iniciais do formulário
    Mudo = False
    
''    'Setup how we want the task tray to work and display
''    Set Tray = New Class_Radio
''    ' Initialize settings here
''    Tray.Initialize Me
''    Tray.AutoRefresh = True
''    Tray.Tooltip = "Rádio"
''    Tray.AddIcon
''    Tray.Refresh
''    Tray.ShowBalloonTip Tray.Tooltip, "Rádio", NIIF_INFO + NIIF_NOSOUND, 1000
''    SaveState = Me.WindowState
''    'hook the keyboard for usage of the media keys
''    Hook (Me.hwnd)
    
    'Preencher a lista com os canais disponiveis
    List1.Clear
    List1.AddItem "Adult Alternative"
    List1.ItemData(List1.NewIndex) = 12
    List1.AddItem "Adult Contemporary"
    List1.ItemData(List1.NewIndex) = 14
    List1.AddItem "Alternative Rock"
    List1.ItemData(List1.NewIndex) = 47
    List1.AddItem "Big Band and Swing"
    List1.ItemData(List1.NewIndex) = 54
    List1.AddItem "Bluegrass"
    List1.ItemData(List1.NewIndex) = 15
    List1.AddItem "Blues"
    List1.ItemData(List1.NewIndex) = 16
    List1.AddItem "Celtic"
    List1.ItemData(List1.NewIndex) = 46
    List1.AddItem "Christian Contemporary"
    List1.ItemData(List1.NewIndex) = 17
    List1.AddItem "Christmas Celebration"
    List1.ItemData(List1.NewIndex) = 61
    List1.AddItem "Classic 60s"
    List1.ItemData(List1.NewIndex) = 71
    List1.AddItem "Classic Country"
    List1.ItemData(List1.NewIndex) = 70
    List1.AddItem "Classic Hits"
    List1.ItemData(List1.NewIndex) = 19
    List1.AddItem "Classic Rock"
    List1.ItemData(List1.NewIndex) = 22
    List1.AddItem "Classical"
    List1.ItemData(List1.NewIndex) = 21
    List1.AddItem "Country"
    List1.ItemData(List1.NewIndex) = 23
    List1.AddItem "Dance"
    List1.ItemData(List1.NewIndex) = 24
    List1.AddItem "Disco"
    List1.ItemData(List1.NewIndex) = 25
    List1.AddItem "Electronica"
    List1.ItemData(List1.NewIndex) = 26
    List1.AddItem "Folk"
    List1.ItemData(List1.NewIndex) = 27
    List1.AddItem "Forever Fifties"
    List1.ItemData(List1.NewIndex) = 53
    List1.AddItem "Halloween Rock"
    List1.ItemData(List1.NewIndex) = 59
    List1.AddItem "Hip Hop"
    List1.ItemData(List1.NewIndex) = 28
    List1.AddItem "Hot Hits"
    List1.ItemData(List1.NewIndex) = 73
    List1.AddItem "Indie Rock"
    List1.ItemData(List1.NewIndex) = 30
    List1.AddItem "Jazz"
    List1.ItemData(List1.NewIndex) = 31
    List1.AddItem "Mash-Ups"
    List1.ItemData(List1.NewIndex) = 63
    List1.AddItem "Metal Rock"
    List1.ItemData(List1.NewIndex) = 34
    List1.AddItem "Musical Magic"
    List1.ItemData(List1.NewIndex) = 55
    List1.AddItem "Native American"
    List1.ItemData(List1.NewIndex) = 36
    List1.AddItem "New Age"
    List1.ItemData(List1.NewIndex) = 35
    List1.AddItem "R&B Classics"
    List1.ItemData(List1.NewIndex) = 37
    List1.AddItem "Reggae"
    List1.ItemData(List1.NewIndex) = 38
    List1.AddItem "Retro Radio"
    List1.ItemData(List1.NewIndex) = 29
    List1.AddItem "Rock"
    List1.ItemData(List1.NewIndex) = 39
    List1.AddItem "Rockin 80's"
    List1.ItemData(List1.NewIndex) = 40
    List1.AddItem "Smooth Jazz"
    List1.ItemData(List1.NewIndex) = 41
    List1.AddItem "Soundtracks"
    List1.ItemData(List1.NewIndex) = 64
    List1.AddItem "Top 40"
    List1.ItemData(List1.NewIndex) = 20
    List1.AddItem "Top Alternative 2003"
    List1.ItemData(List1.NewIndex) = 52
    List1.AddItem "Top Hits 2003"
    List1.ItemData(List1.NewIndex) = 51
    List1.AddItem "Top Hits 2004"
    List1.ItemData(List1.NewIndex) = 62
    List1.AddItem "Top Hits 2005"
    List1.ItemData(List1.NewIndex) = 72
    List1.AddItem "Urban"
    List1.ItemData(List1.NewIndex) = 42
    List1.AddItem "Vintage Vault"
    List1.ItemData(List1.NewIndex) = 57
    List1.AddItem "Women's Alternative"
    List1.ItemData(List1.NewIndex) = 43
    List1.AddItem "World"
    List1.ItemData(List1.NewIndex) = 44
    'set the listindex to the last channel selected
    List1.ListIndex = Val(GetSetting(App.EXEName, "Settings", "Channel"))
    
    'Preencher grelha de canais
    Dim i As Integer
    For i = 0 To List1.ListCount - 1
        Grelha.AddItem List1.List(i), -1, 0
    Next i
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Valor de x e y
    H = X
    v = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover formuário
    If Me.WindowState = 0 Then
        If Button = vbLeftButton Then
           Me.Left = Me.Left - (H - X)
           Me.Top = Me.Top - (v - Y)
        End If
    End If
    
    Image_Capa.Visible = True
    Image_Capa_Zoom.Visible = False
    
    'Repor as cores das labels
    Label_Fechar.ForeColor = &H808080
    Label_Minimizar.ForeColor = &H808080
End Sub

Private Sub Form_Resize()
''    On Error Resume Next
''    ' This will place an icon in the task tray when the user minimizes this form
''    If Me.WindowState = 1 Then
''        ' Form was minimized
''        Me.Hide
''        Exit Sub
''    End If
''    SaveState = Me.WindowState
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Hook(hwnd, False) 'Unhook
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyRight
            'get the next song
            MPlayer.Controls.Next
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case UCase$(Chr$(KeyAscii))
        'use the keyboard to control the player
        Case Is = "P"
            'Pause/Play
            If MPlayer.playState = wmppsPlaying Then
                Form_Radio.MPlayer.Controls.pause
                Debug.Print "Pausa..."
            Else
                MPlayer.Controls.play
                Debug.Print "Reproduzir..."
            End If
            KeyAscii = 0
        Case Is = "S"
            'Stop
            MPlayer.Controls.stop
        Case Is = "B"
            'Back
            MPlayer.Controls.previous
        Case Is = "N"
            'Next
            MPlayer.Controls.Next
        Case Is = "F"
            'FastForward
            MPlayer.Controls.fastForward
        Case Is = "R"
            'Rewind
            MPlayer.Controls.fastReverse
        Case Is = "Q"
            'let's quit...
            MPlayer.Controls.stop
            Unload Me
    End Select
End Sub

Private Sub Grelha_DbClick()
    'Igualar o listindex de ambas as listboxs
    List1.ListIndex = Grelha.ListIndex
    Timer_Musica.Enabled = False
    Timer_Artista.Enabled = False
    List1_DblClick
    Botao_Play.Visible = True
    Botao_Pause.Visible = False
End Sub

Private Sub Image_Capa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image_Capa.Visible = False
    Image_Capa_Zoom.Visible = True
End Sub

Private Sub Label_Fechar_Click()
    Botao_Stop_Click
    Unload Me
End Sub

Private Sub Label_Fechar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animacao da label
    Label_Fechar.ForeColor = vbWhite
    Label_Minimizar.ForeColor = &H808080
End Sub

Private Sub Label_Lista_Click()
    'Ver/ ocultar lista
    If Me.Width = 2850 Then
        Me.Width = 7335
        Label_Lista.BackColor = &HFC6C03
        Label_Lista.ForeColor = vbWhite
        Label_Lista.ToolTipText = "Ocultar lista"
    Else
        Me.Width = 2850
        Label_Lista.BackColor = &H50505
        Label_Lista.ForeColor = &H808080
        Label_Lista.ToolTipText = "Ver lista"
    End If
End Sub

Private Sub Label_Minimizar_Click()
    'Minimizar formulário
    Me.WindowState = 1
End Sub

Private Sub Label_Minimizar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animacao da label
    Label_Fechar.ForeColor = &H808080
    Label_Minimizar.ForeColor = vbWhite
End Sub

Private Sub Label_Mudo_Click()
    'Ver/ ocultar lista
    If Mudo = False Then
        Label_Mudo.BackColor = &H50505
        Label_Mudo.ForeColor = &H808080
        Label_Mudo.ToolTipText = "Ouvir"
        Mudo = True
        MPlayer.settings.mute = True
    Else
        Label_Mudo.BackColor = &HFC6C03
        Label_Mudo.ForeColor = vbWhite
        Label_Mudo.ToolTipText = "Mudo"
        Mudo = False
        MPlayer.settings.mute = False
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

Private Sub List1_DblClick()
    Dim lIndex As Integer
    Timer_Progress.Enabled = True
    'double clicking selects the channel to be played
    lIndex = List1.ListIndex
    'display the format
    'save the channel
    SaveSetting App.EXEName, "Settings", "Channel", lIndex
    'launch the channel in the media player
    MPlayer.URL = "http://www.gotradio.com/player/launch.asp?refer=web&id=" & List1.ItemData(lIndex)
    MPlayer.SetFocus
''    Tray.Tooltip = "Conectado a " & lblFormat
''    Tray.Refresh
''    Tray.ShowBalloonTip Tray.Tooltip, "Rádio", NIIF_INFO + NIIF_NOSOUND, 1000
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
    'the return key launches the channel
    If KeyAscii = 13 Then
        List1_DblClick
    End If
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    List1.ZOrder 0
    ' present related tip message
    Dim lXPoint As Long
    Dim lYPoint As Long
    Dim lIndex As Long
    If Button = 0 Then ' if no button was pressed
        lXPoint = CLng(X / Screen.TwipsPerPixelX)
        lYPoint = CLng(Y / Screen.TwipsPerPixelY)
        With List1
            ' get selected item from list
            lIndex = SendMessage(.hwnd, LB_ITEMFROMPOINT, 0, ByVal ((lYPoint * 65536) + lXPoint))
            ' show tip or clear last one
            If (lIndex >= 0) And (lIndex <= .ListCount) Then
                .ToolTipText = "Select " & .List(lIndex)
            Else
                .ToolTipText = ""
            End If
        End With '(List1)
    End If '(button=0)
End Sub

Private Sub MPlayer_Buffering(ByVal Start As Boolean)
    lblStatus = "Processando..."
    Debug.Print lblStatus
End Sub

Private Sub MPlayer_CurrentItemChange(ByVal pdispMedia As Object)
    lblStatus = "Alterar item atual"
    Debug.Print lblStatus
End Sub

Private Sub MPlayer_CurrentMediaItemAvailable(ByVal bstrItemName As String)
    lblStatus = "Item disponivel"
    'Debug.Print lblStatus
    GetInfo
End Sub

Private Sub MPlayer_DeviceConnect(ByVal pDevice As WMPLibCtl.IWMPSyncDevice)
    lblStatus = "Conectado"
    Debug.Print lblStatus
End Sub

Private Sub MPlayer_DeviceStatusChange(ByVal pDevice As WMPLibCtl.IWMPSyncDevice, ByVal NewStatus As WMPLibCtl.WMPDeviceStatus)
    lblStatus = "Mudando estado"
    Debug.Print lblStatus
End Sub

Private Sub MPlayer_MediaChange(ByVal Item As Object)
    lblStatus = "Alterar media " '; Item
    Debug.Print lblStatus
End Sub

Private Sub MPlayer_NewStream()
    lblStatus = "Novo processo"
    Debug.Print lblStatus
End Sub

Private Sub MPlayer_PlayStateChange(ByVal NewState As Long)
    If NewState = wmppsPlaying Then '3
        'if the player is playing, fill in the song info
        GetInfo
    ElseIf NewState = wmppsTransitioning Then '9
        'clear the song info if the song is done
        ClearInfo
    End If
    lblStatus = "Tocar - " & NewState
    Debug.Print lblStatus
End Sub

Private Sub MPlayer_StatusChange()
    lblStatus = "Mudança de estado"
    'Debug.Print lblStatus
End Sub

Sub ClearInfo()
    lblinfo = ""
    Label_Titulo.Caption = "Rádio"
    Picture1.Visible = False
    Image_Capa.Visible = False
    Picture1.Picture = LoadPicture("")
    Label_Sem_Capa.Visible = True
End Sub

Sub GetInfo()
    On Error GoTo Oops
    Me.MousePointer = 11
    With MPlayer.currentMedia
        'retrieve the info on the currently playing song
        AID = Val(.getItemInfo("AID"))
        Author$ = .getItemInfo("AUTHOR")
        Artist$ = .getItemInfo("Artist")
        Label$ = .getItemInfo("COPYRIGHT")
        Title$ = .getItemInfo("TITLE")
        adType$ = .getItemInfo("ADTYPE")
        'check if the current media being played is an ad
        If LCase(adType$) <> "none" Then
            'we have an ad, so let's skip it
            Me.Caption = "Skipping Ad..."
            MPlayer.Controls.Next
        End If
        If MPlayer.Controls.isAvailable("FastForward") = False Then
            'we can't skip this one and I don't think there's any way around it
        End If
        If oldTitle$ <> Title$ Then
            'this is a new song... change info
            albumID = Val(.getItemInfo("ALBUMID"))
            'clear the album cover picture
            Picture1.Picture = LoadPicture("")
            Image_Capa.Picture = Picture1.Picture
            Image_Capa_Zoom.Picture = Picture1.Picture
            'clear our data
            Album$ = ""
            SmallCover$ = ""
            MedCover$ = ""
            LargeCover$ = ""
            'check if we have a valid albumID
            If albumID <> 0 Then
                'it is a valid song, so now let's get the rest of the info
                Album$ = .getItemInfo("ALBUM")
                'get the 3 pictures that are supplied for the album art
                SmallCover$ = .getItemInfo("SCOVER")
                MedCover$ = Replace(.getItemInfo("MCOVER"), "LZ", "MZ", , , vbTextCompare)
                LargeCover$ = .getItemInfo("LCOVER")
                'hide the picture so it loads nicer
                Picture1.Visible = False
                Image_Capa.Visible = False
                Label_Sem_Capa.Visible = True
                
                'get the large album cover picture
                Picture1.Picture = OLELoadPicture(LargeCover$)
                Image_Capa.Picture = Picture1.Picture
                Image_Capa_Zoom.Picture = Picture1.Picture
                'check picture size and if the picture is too big, get a smaller one
                If Picture1.Top + Picture1.Height > Me.Height / Screen.TwipsPerPixelY Then
                    'try getting the next size down...
                    Picture1.Picture = OLELoadPicture(MedCover$)
                    Image_Capa.Picture = Picture1.Picture
                    Image_Capa_Zoom.Picture = Picture1.Picture
                End If
                'show the picture
                Picture1.Visible = True
                Image_Capa.Visible = True
                Label_Sem_Capa.Visible = False
                'put the picture at the top of the z order in case it overlaps the list box
                Picture1.ZOrder 0
            End If
            'put together the info into the label
            lblinfo = "Title: " & Title$ & vbCrLf
            lblinfo = lblinfo & "Artist: " & Artist$ & vbCrLf
            lblinfo = lblinfo & "Album: " & Album$ & vbCrLf
            lblinfo = lblinfo & "Author: " & Author$ & vbCrLf
''            Tray.Tooltip = lblinfo
            'i don't think the timeout works...
''            Tray.Refresh
            'instead of showing the balloon, which does not go away,  just set the tip
''            'Tray.ShowBalloonTip Tray.Tooltip, "Rádio", NIIF_INFO + NIIF_NOSOUND, 1000
''            Tray.Refresh
            lblinfo = Replace(lblinfo, "&", "&&")
            
            'now put together the info we want on the window caption
            Capt$ = Title$ & " ; Artist = " & Artist & "; Album = " & Album
            'now display that info in the text box
''            Text1.Text = lblinfo & "----------------------------------------" & vbCrLf
            'find out how many items there are to read and set up a loop to get them all
''            For i = 0 To .attributeCount - 1
''                'add each attribute's name to the text box
''                Text1.Text = Text1.Text & i & " " & .getAttributeName(i) & " = "
''                'now get the data for that attribute and put it in the text box also
''                Text1.Text = Text1.Text & .getItemInfo(.getAttributeName(i)) & vbCrLf
''            Next i
            oldTitle$ = Title$
            
            'Ocultar o progress bar
            Timer_Progress.Enabled = False
            Image_Progress(0).Visible = True
            Dim X As Integer
            For X = 1 To 5
                Image_Progress(X).Visible = False
            Next X
                End If
            End With
            
            'Activar o timer do titulo que mostra o titulo e artista da musica
            Timer_Musica.Enabled = True
            GoTo Exit_GetInfo
Oops:
    'Abort=3,Retry=4,Ignore=5
    eTitle$ = App.Title & ": Error in Subroutine GetInfo "
    EMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
    EMess$ = EMess$ & "Occurred in GetInfo"
    EMess$ = EMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
    mError = MsgBox(EMess$, vbAbortRetryIgnore, eTitle$)
    If mError = vbRetry Then Resume
    If mError = vbIgnore Then Resume Next
Exit_GetInfo:
    Me.MousePointer = 0
End Sub

Public Function OLELoadPicture(ByVal strFilename As String) As Picture
    On Error GoTo Oops
    'This function gets a picture from a url path
    Dim myTGUID As TGUID
    myTGUID.Data1 = &H7BF80980
    myTGUID.Data2 = &HBF32
    myTGUID.Data3 = &H101A
    myTGUID.Data4(0) = &H8B
    myTGUID.Data4(1) = &HBB
    myTGUID.Data4(2) = &H0
    myTGUID.Data4(3) = &HAA
    myTGUID.Data4(4) = &H0
    myTGUID.Data4(5) = &H30
    myTGUID.Data4(6) = &HC
    myTGUID.Data4(7) = &HAB
    OleLoadPicturePath StrPtr(strFilename), 0, 0, 0, myTGUID, OLELoadPicture
    GoTo Exit_LoadPicture
LblError:
    Set OLELoadPicture = VB.LoadPicture(strFilename)
    GoTo Exit_LoadPicture
Oops:
    'Abort=3,Retry=4,Ignore=5
    eTitle$ = App.Title & ": Error in Subroutine LoadPicture "
    EMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
    EMess$ = EMess$ & "Occurred in LoadPicture"
    EMess$ = EMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
    mError = MsgBox(EMess$, vbAbortRetryIgnore, eTitle$)
    If mError = vbRetry Then Resume
    If mError = vbIgnore Then Resume Next
Exit_LoadPicture:
End Function

Private Sub Picture1_DblClick()
    Picture1.Picture = OLELoadPicture(MedCover$)
    Image_Capa.Picture = Picture1.Picture
    Image_Capa_Zoom.Picture = Picture1.Picture
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.ZOrder 0
End Sub

Private Sub Timer_Artista_Timer()
    Label_Titulo.Caption = Artist$
    Timer_Musica.Enabled = True
    Timer_Artista.Enabled = False
End Sub

Private Sub Timer_Musica_Timer()
    Label_Titulo.Caption = Title$ 'Artist$
    Timer_Artista.Enabled = True
    Timer_Musica.Enabled = False
End Sub

Private Sub Timer_Progress_Timer()
    If Image_Progress(0).Visible = True Then
        Image_Progress(0).Visible = False
        Image_Progress(1).Visible = True
    ElseIf Image_Progress(1).Visible = True Then
        Image_Progress(1).Visible = False
        Image_Progress(2).Visible = True
    ElseIf Image_Progress(2).Visible = True Then
        Image_Progress(2).Visible = False
        Image_Progress(3).Visible = True
    ElseIf Image_Progress(3).Visible = True Then
        Image_Progress(3).Visible = False
        Image_Progress(4).Visible = True
    ElseIf Image_Progress(4).Visible = True Then
        Image_Progress(4).Visible = False
        Image_Progress(5).Visible = True
    ElseIf Image_Progress(5).Visible = True Then
        Image_Progress(5).Visible = False
        Image_Progress(0).Visible = True
    End If
End Sub

Private Sub Timer1_Timer()
    If Capt$ = "" Then Exit Sub
    'scroll the info on the minimized window caption
    If Me.WindowState = vbMinimized Then
        LeftChar = LeftChar + 1
        sCapt$ = Mid$(Capt$, LeftChar)
        If Len(sCapt$) = 0 Then
            sCapt$ = Capt$
            LeftChar = 0
        End If
        Me.Caption = sCapt$
    Else
        Me.Caption = Capt$
    End If
    Me.Refresh
    DoEvents
End Sub

''Private Sub Tray_DoubleClick(Button As Integer)
''    ' If they double click with the left mouse
''    ' button we will simply show the form
''    If Button = 0 Then
''        Me.WindowState = SaveState
''        Me.Show
''    End If
''End Sub

''Private Sub Tray_MouseDown(Button As Integer)
''    ' If they right click on the task tray then
''    ' we will simply show them a popup menu
''    If Button = 1 Then
''        ' And popup the menu
''        PopupMenu mnuPopup, , , , mnuShow
''    End If
''End Sub

