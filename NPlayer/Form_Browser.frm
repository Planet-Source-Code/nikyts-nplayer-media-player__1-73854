VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form_Browser 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Nplayer - Navegador - Electric Nikyts"
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin NPlayer.N_Button Botao_Anterior 
      Height          =   360
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Página anterior"
      Top             =   500
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      Picture         =   "Form_Browser.frx":0000
      PictureHover    =   "Form_Browser.frx":0712
      PictureDown     =   "Form_Browser.frx":0E24
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   7095
      ExtentX         =   12515
      ExtentY         =   5318
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.ComboBox Combo_Site 
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Top             =   500
      Width           =   4935
   End
   Begin NPlayer.N_Button Botao_Restaurar 
      Height          =   285
      Left            =   4920
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   503
      Picture         =   "Form_Browser.frx":1536
      PictureHover    =   "Form_Browser.frx":1AE0
      PictureDown     =   "Form_Browser.frx":208A
   End
   Begin NPlayer.N_Button Botao_Maximizar 
      Height          =   285
      Left            =   6120
      TabIndex        =   3
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   503
      Picture         =   "Form_Browser.frx":2634
      PictureHover    =   "Form_Browser.frx":2BDE
      PictureDown     =   "Form_Browser.frx":3188
   End
   Begin NPlayer.N_Button Botao_Fechar 
      Height          =   285
      Left            =   6600
      TabIndex        =   4
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      Picture         =   "Form_Browser.frx":3732
      PictureHover    =   "Form_Browser.frx":40B8
      PictureDown     =   "Form_Browser.frx":4A3E
   End
   Begin NPlayer.N_Button Botao_Minimizar 
      Height          =   285
      Left            =   5520
      TabIndex        =   5
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      Picture         =   "Form_Browser.frx":53C4
      PictureHover    =   "Form_Browser.frx":59BA
      PictureDown     =   "Form_Browser.frx":5FB0
   End
   Begin NPlayer.N_Button Botao_Seguinte 
      Height          =   360
      Left            =   480
      TabIndex        =   8
      ToolTipText     =   "Página seguinte"
      Top             =   500
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      Picture         =   "Form_Browser.frx":65A6
      PictureHover    =   "Form_Browser.frx":6CB8
      PictureDown     =   "Form_Browser.frx":73CA
   End
   Begin NPlayer.N_Button Botao_Actualizar 
      Height          =   360
      Left            =   960
      TabIndex        =   9
      ToolTipText     =   "Actualizar"
      Top             =   500
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      Picture         =   "Form_Browser.frx":7ADC
      PictureHover    =   "Form_Browser.frx":81EE
      PictureDown     =   "Form_Browser.frx":8900
   End
   Begin NPlayer.N_Button Botao_Opcoes 
      Height          =   360
      Left            =   2160
      TabIndex        =   10
      ToolTipText     =   "Opções"
      Top             =   500
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      Picture         =   "Form_Browser.frx":9012
      PictureHover    =   "Form_Browser.frx":9724
      PictureDown     =   "Form_Browser.frx":9E36
   End
   Begin NPlayer.N_Button Botao_Favoritos 
      Height          =   360
      Left            =   1800
      TabIndex        =   11
      ToolTipText     =   "Adicionar aos favoritos"
      Top             =   500
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      Picture         =   "Form_Browser.frx":A548
      PictureHover    =   "Form_Browser.frx":AC5A
      PictureDown     =   "Form_Browser.frx":B36C
   End
   Begin NPlayer.N_Button Botao_Parar 
      Height          =   360
      Left            =   1320
      TabIndex        =   12
      ToolTipText     =   "Parar"
      Top             =   500
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      Picture         =   "Form_Browser.frx":BA7E
      PictureHover    =   "Form_Browser.frx":C190
      PictureDown     =   "Form_Browser.frx":C8A2
   End
   Begin VB.Image Skin_Top_Menu 
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      Picture         =   "Form_Browser.frx":CFB4
      Top             =   360
      Width           =   1035
   End
   Begin VB.Image Botao_Redimensionar 
      Height          =   240
      Left            =   7800
      Picture         =   "Form_Browser.frx":F146
      Top             =   6120
      Width           =   225
   End
   Begin VB.Image Skin_Top_Esquerda 
      Enabled         =   0   'False
      Height          =   405
      Left            =   0
      Picture         =   "Form_Browser.frx":F488
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Skin_Down_Direita 
      Enabled         =   0   'False
      Height          =   615
      Left            =   960
      Picture         =   "Form_Browser.frx":F752
      Top             =   840
      Width           =   105
   End
   Begin VB.Image Skin_Down_Esquerda 
      Enabled         =   0   'False
      Height          =   615
      Left            =   0
      Picture         =   "Form_Browser.frx":FB6C
      Top             =   840
      Width           =   105
   End
   Begin VB.Image Skin_Lateral_Direita 
      Enabled         =   0   'False
      Height          =   630
      Left            =   960
      Picture         =   "Form_Browser.frx":FF86
      Top             =   360
      Width           =   105
   End
   Begin VB.Image Skin_Lateral_Esquerda 
      Enabled         =   0   'False
      Height          =   540
      Left            =   0
      Picture         =   "Form_Browser.frx":103B8
      Top             =   400
      Width           =   105
   End
   Begin VB.Image Skin_Top_Direita 
      Enabled         =   0   'False
      Height          =   405
      Left            =   960
      Picture         =   "Form_Browser.frx":1075A
      Top             =   0
      Width           =   105
   End
   Begin VB.Label Label_Titulo 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "   Navegador"
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
      Top             =   90
      Width           =   2535
   End
   Begin VB.Image Skin_Top_Centro 
      Height          =   405
      Left            =   110
      Picture         =   "Form_Browser.frx":10A24
      Top             =   0
      Width           =   870
   End
   Begin VB.Image Skin_Down_Centro 
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      Picture         =   "Form_Browser.frx":11CF6
      Top             =   960
      Width           =   1035
   End
End
Attribute VB_Name = "Form_Browser"
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

Private Sub Botao_Actualizar_Click()
    'Actualizar a página em visualização
    WebBrowser1.Refresh
End Sub

Private Sub Botao_Anterior_Click()
    On Error Resume Next
    'Ver a página anterior
    WebBrowser1.GoBack
End Sub

Private Sub Botao_Favoritos_Click()
    'Acionar site aos favoritos
    If Combo_Site.Text = "" Then Exit Sub
    Load Form_Browser_Opcoes
    With Form_Browser_Opcoes
        .Lista_Favoritos.AddItem Combo_Site.Text
    End With
End Sub

Private Sub Botao_Fechar_Click()
    'Fechar formulário
    'Verificar o numero de janelas abertas
    If Numero_de_Janelas = 1 Then
        Unload Form_Browser
        Exit Sub
    End If
    Unload Me
    Numero_de_Janelas = Numero_de_Janelas - 1
End Sub

Private Sub Botao_Maximizar_Click()
    'Maximizar formulário
    Me.WindowState = 2
    
    Botao_Maximizar.Visible = False
    Botao_Restaurar.Visible = True
End Sub

Private Sub Botao_Minimizar_Click()
    'Minimizar formulário
    Me.WindowState = 1
End Sub

Private Sub Botao_Parar_Click()
    'Para a página em execução
    WebBrowser1.stop
End Sub

Private Sub Botao_Opcoes_Click()
    'Ver o formuário das opcoes do browser
    Form_Browser_Opcoes.Show ' vbModal
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

Private Sub Botao_Restaurar_Click()
    'Restaurar formulário
    Me.WindowState = 0
    
    Botao_Maximizar.Visible = True
    Botao_Restaurar.Visible = False
End Sub

Private Sub Botao_Seguinte_Click()
    On Error Resume Next
    'Ver a página seguinte
    WebBrowser1.GoForward
End Sub

Private Sub Combo_Site_KeyDown(KeyCode As Integer, Shift As Integer)
    'Reproduzir som atraves do enter
    If KeyCode = vbKeyReturn Then
        Abrir_Site
    End If
End Sub

Public Sub Abrir_Site()
    'Abrir o site digitalizado na combo_site
    WebBrowser1.Navigate (Combo_Site.Text)
    Combo_Site.AddItem Combo_Site.Text 'WebBrowser1.LocationURL, 0
    'Combo_Site.ListIndex = 0
End Sub

Private Sub Form_Load()
    'Propriedades iniciais do formulário
    WebBrowser1.Navigate Form_Browser_Opcoes.Text_Pagina_Inicial.Text
    
    'Adicionar "1" ao numero de janelas abertas
    Numero_de_Janelas = Numero_de_Janelas + 1
End Sub

Private Sub Form_Resize()
    On Error GoTo CORRIGIR_ERRO
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
    
    'Skin_Top_Menu
    Skin_Top_Menu.Top = Skin_Top_Centro.Top + Skin_Top_Centro.Height
    Skin_Top_Menu.Stretch = True
    Skin_Top_Menu.Width = Me.Width - Skin_Top_Esquerda.Width - Skin_Top_Direita.Width
    Skin_Top_Menu.Left = Skin_Top_Esquerda.Left + Skin_Top_Esquerda.Width
    
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
    
    'Botao_Maximizar
    Botao_Maximizar.Top = 0
    Botao_Maximizar.Left = Botao_Fechar.Left - Botao_Maximizar.Width
    
    'Botao_Minimizar
    Botao_Minimizar.Top = 0
    Botao_Minimizar.Left = Botao_Maximizar.Left - Botao_Minimizar.Width
    
    'Botao_Restaurar
    Botao_Restaurar.Top = 0
    Botao_Restaurar.Left = Botao_Maximizar.Left
       
    'WebBrowser1
    With WebBrowser1
        .Height = Me.Height - Skin_Top_Centro.Height - Skin_Top_Menu.Height - Skin_Down_Centro.Height
        .Top = Skin_Top_Menu.Top + Skin_Top_Menu.Height
        .Width = Me.Width - Skin_Lateral_Esquerda.Width - Skin_Lateral_Direita.Width
        .Left = Skin_Lateral_Esquerda.Left + Skin_Lateral_Esquerda.Width
    End With
    
    'Botao_Anterior
    With Botao_Anterior
        .Top = 550
        .Left = 220
    End With
    
    'Botao_Seguinte
    With Botao_Seguinte
        .Top = 550
        .Left = Botao_Anterior.Left + Botao_Anterior.Width
    End With
    
    'Botao_Actualizar
    With Botao_Actualizar
        .Top = 550
        .Left = Botao_Seguinte.Left + Botao_Seguinte.Width + 120
    End With
    
    'Botao_Parar
    With Botao_Parar
        .Top = 550
        .Left = Botao_Actualizar.Left + Botao_Actualizar.Width
    End With
    
    'Botao_Favoritos
    With Botao_Favoritos
        .Top = 550
        .Left = Botao_Parar.Left + Botao_Parar.Width + 120
    End With
    
    'Botao_Opcoes
    With Botao_Opcoes
        .Top = 550
        .Left = Botao_Favoritos.Left + Botao_Favoritos.Width
    End With
    
    'Combo_Site
    With Combo_Site
        .Top = 570
        .Width = Me.Width - Botao_Redimensionar.Width - 2800
        .Left = Botao_Opcoes.Left + Botao_Opcoes.Width + 120
    End With
    
    'Botao_Redimensionar
    If Me.Height <> Skin_Top_Centro.Height Then
        Botao_Redimensionar.Top = Me.Height - Botao_Redimensionar.Height
        Botao_Redimensionar.Left = Me.Width - Botao_Redimensionar.Width
    End If
Exit Sub
CORRIGIR_ERRO:
    Me.Height = 6450
    Me.Width = 8115
End Sub

Private Sub Lista_Directorios_Click()
    Lista_Ficheiros.Clear
    File1.FileName = Lista_Directorios.Text
    Dim i As Integer
    For i = 0 To File1.ListCount - 1
        Lista_Ficheiros.AddItem File1.List(i)  'Lista_Directorios.Text & "\" &
    Next i
    
    'Indicar a pasta actual
    Text1.Text = Lista_Directorios.Text
    
    'Ocultar a lista dos directorios
    'Lista_Directorios.Visible = False
End Sub

Private Sub Label_Titulo_DblClick()
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

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Combo_Site.Text = Right(WebBrowser1.LocationURL, Len(WebBrowser1.LocationURL) - Len("http://"))
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    If Not Found Then
        Combo_Site.AddItem Right(WebBrowser1.LocationURL, Len(WebBrowser1.LocationURL) - Len("http://")), 0
    Else
        'Delete the item and add item as index 0
        Combo_Site.RemoveItem i - 1
        Combo_Site.Text = Right(WebBrowser1.LocationURL, Len(WebBrowser1.LocationURL) - Len("http://"))
        Combo_Site.AddItem Right(WebBrowser1.LocationURL, Len(WebBrowser1.LocationURL) - Len("http://")), 0
    End If
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
    'abrir a nova página em uma nova janela
    Dim Nova_Janela As Form_Browser
    Set Nova_Janela = New Form_Browser
    Set ppDisp = Nova_Janela.WebBrowser1.object
    With Nova_Janela
        .Show
        .WindowState = 0
    End With
End Sub
