VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form_Principal 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   9720
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   14235
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
   Icon            =   "Form_Principal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9720
   ScaleWidth      =   14235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Frame_Mascara 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5510
      Left            =   6480
      Picture         =   "Form_Principal.frx":57E2
      ScaleHeight     =   5505
      ScaleWidth      =   7650
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   7650
      Begin VB.PictureBox Imagem_Grafico_Mascara 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FC6C03&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFAE0F&
         Height          =   420
         Left            =   1080
         ScaleHeight     =   420
         ScaleWidth      =   3705
         TabIndex        =   102
         Top             =   2520
         Width           =   3705
      End
      Begin NPlayer.N_Button Botao_Sobre_Mascara 
         Height          =   390
         Left            =   7000
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   570
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   688
         Picture         =   "Form_Principal.frx":F8400
         PictureHover    =   "Form_Principal.frx":F8B3A
         PictureDown     =   "Form_Principal.frx":F9274
      End
      Begin NPlayer.N_Button Aumentar_Volume 
         Height          =   135
         Left            =   7280
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   5115
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   238
         Picture         =   "Form_Principal.frx":F99AE
         PictureHover    =   "Form_Principal.frx":F9AD8
         PictureDown     =   "Form_Principal.frx":F9C02
      End
      Begin NPlayer.N_Button Diminuir_Volume 
         Height          =   60
         Left            =   5720
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   5180
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   106
         Picture         =   "Form_Principal.frx":F9D2C
         PictureHover    =   "Form_Principal.frx":F9DDE
         PictureDown     =   "Form_Principal.frx":F9E90
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00131313&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1365
         Left            =   5880
         ScaleHeight     =   1365
         ScaleWidth      =   1335
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1335
         Begin VB.Image Image_Volume 
            Height          =   1365
            Left            =   0
            Picture         =   "Form_Principal.frx":F9F42
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.PictureBox SliderBar_Mascara 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   300
         ScaleHeight     =   135
         ScaleWidth      =   4695
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   4200
         Width           =   4700
         Begin NPlayer.N_Button Slide_Mascara 
            Height          =   105
            Left            =   0
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   10
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   185
            Picture         =   "Form_Principal.frx":FFEC8
            PictureHover    =   "Form_Principal.frx":100112
            PictureDown     =   "Form_Principal.frx":10035C
         End
         Begin VB.Image Image_Barra_Slide_Mascara 
            Enabled         =   0   'False
            Height          =   135
            Left            =   0
            Picture         =   "Form_Principal.frx":1005A6
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4695
         End
      End
      Begin NPlayer.N_Button Botao_Mascara_Primeiro 
         Height          =   420
         Left            =   300
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   4680
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   741
         Picture         =   "Form_Principal.frx":101B6C
         PictureHover    =   "Form_Principal.frx":1027FE
         PictureDown     =   "Form_Principal.frx":103490
      End
      Begin NPlayer.N_Button Botao_Mascara_Antes 
         Height          =   420
         Left            =   860
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   4680
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   741
         Picture         =   "Form_Principal.frx":104122
         PictureHover    =   "Form_Principal.frx":104D44
         PictureDown     =   "Form_Principal.frx":105966
      End
      Begin NPlayer.N_Button Botao_Mascara_Pause 
         Height          =   420
         Left            =   2135
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   4680
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   741
         Picture         =   "Form_Principal.frx":106588
         PictureHover    =   "Form_Principal.frx":1071AA
         PictureDown     =   "Form_Principal.frx":107DCC
      End
      Begin NPlayer.N_Button Botao_Mascara_Seguinte 
         Height          =   420
         Left            =   3960
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   4680
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   741
         Picture         =   "Form_Principal.frx":1089EE
         PictureHover    =   "Form_Principal.frx":109680
         PictureDown     =   "Form_Principal.frx":10A312
      End
      Begin NPlayer.N_Button Botao_Mascara_Pasta 
         Height          =   420
         Left            =   3210
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   4680
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   741
         Picture         =   "Form_Principal.frx":10AFA4
         PictureHover    =   "Form_Principal.frx":10BBC6
         PictureDown     =   "Form_Principal.frx":10C7E8
      End
      Begin NPlayer.N_Button Botao_Mascara_Stop 
         Height          =   420
         Left            =   2660
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   4680
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   741
         Picture         =   "Form_Principal.frx":10D40A
         PictureHover    =   "Form_Principal.frx":10E09C
         PictureDown     =   "Form_Principal.frx":10ED2E
      End
      Begin NPlayer.N_Button Botao_Mascara_Play 
         Height          =   420
         Left            =   1610
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   4680
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   741
         Picture         =   "Form_Principal.frx":10F9C0
         PictureHover    =   "Form_Principal.frx":1105E2
         PictureDown     =   "Form_Principal.frx":111204
      End
      Begin NPlayer.N_Button Botao_Mascara_Ultimo 
         Height          =   420
         Left            =   4515
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   4680
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   741
         Picture         =   "Form_Principal.frx":111E26
         PictureHover    =   "Form_Principal.frx":112A48
         PictureDown     =   "Form_Principal.frx":11366A
      End
      Begin NPlayer.N_Button Botao_Minimizar_Mascara 
         Height          =   135
         Left            =   6840
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   150
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   238
         Picture         =   "Form_Principal.frx":11428C
         PictureHover    =   "Form_Principal.frx":1143FE
         PictureDown     =   "Form_Principal.frx":114570
      End
      Begin NPlayer.N_Button Botao_Maximizar_Mascara 
         Height          =   150
         Left            =   7080
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   120
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   265
         Picture         =   "Form_Principal.frx":1146E2
         PictureHover    =   "Form_Principal.frx":11484C
         PictureDown     =   "Form_Principal.frx":1149B6
      End
      Begin NPlayer.N_Button Botao_Fechar_Mascara 
         Height          =   150
         Left            =   7320
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   120
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   265
         Picture         =   "Form_Principal.frx":114B20
         PictureHover    =   "Form_Principal.frx":114C8A
         PictureDown     =   "Form_Principal.frx":114DF4
      End
      Begin NPlayer.McListBox McListBox1 
         Height          =   3735
         Left            =   240
         TabIndex        =   103
         Top             =   5760
         Width           =   7180
         _ExtentX        =   12674
         _ExtentY        =   6588
         Picture         =   "Form_Principal.frx":114F5E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   0
         ForeColor       =   16777215
         SelColor        =   16542723
         BorderStyle     =   0
         IconFocus       =   0   'False
         RowHeight       =   18
         SelectionStyle  =   0
         Path            =   "C:\NPlayer v2\"
      End
      Begin VB.Label Label_Percentagem_Volume_2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E23701&
         BackStyle       =   0  'Transparent
         Caption         =   "Volume 50%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001E1F1D&
         Height          =   180
         Left            =   4245
         TabIndex        =   101
         Top             =   3520
         Width           =   990
      End
      Begin VB.Image Botao_Ordenar_Mascara 
         Height          =   165
         Left            =   600
         Picture         =   "Form_Principal.frx":114F7A
         Top             =   3120
         Width           =   480
      End
      Begin VB.Image Botao_Repetir_Faixa_Mascara 
         Height          =   165
         Left            =   600
         Picture         =   "Form_Principal.frx":1153DC
         Top             =   3315
         Width           =   480
      End
      Begin VB.Image Botao_Repetir_Album_Mascara 
         Height          =   165
         Left            =   600
         Picture         =   "Form_Principal.frx":11583E
         Top             =   3525
         Width           =   480
      End
      Begin VB.Image Image14 
         Enabled         =   0   'False
         Height          =   270
         Left            =   560
         Picture         =   "Form_Principal.frx":115CA0
         Top             =   1600
         Width           =   315
      End
      Begin VB.Label Label_Conexao_Mascara 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Serviço de internet: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00201D1F&
         Height          =   180
         Left            =   920
         TabIndex        =   100
         ToolTipText     =   "Verificar a conexão á internet"
         Top             =   1630
         Width           =   1530
      End
      Begin VB.Label Tempo_Estimado_Top_Mascara 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   5520
         TabIndex        =   99
         Top             =   1920
         Width           =   1620
      End
      Begin VB.Label Label_Data 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data "
         ForeColor       =   &H00E03100&
         Height          =   180
         Left            =   5730
         TabIndex        =   98
         Top             =   2640
         Width           =   1200
      End
      Begin VB.Label Label_Hora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hora"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E03100&
         Height          =   180
         Left            =   6210
         TabIndex        =   97
         Top             =   3000
         Width           =   840
      End
      Begin VB.Image Image10 
         Enabled         =   0   'False
         Height          =   465
         Left            =   5280
         Picture         =   "Form_Principal.frx":116162
         Top             =   525
         Width           =   45
      End
      Begin VB.Label Label_Ver_Propriedades 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Propriedades"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   5580
         TabIndex        =   93
         Top             =   675
         Width           =   1125
      End
      Begin VB.Image Image8 
         Enabled         =   0   'False
         Height          =   465
         Left            =   4320
         Picture         =   "Form_Principal.frx":116318
         Top             =   525
         Width           =   45
      End
      Begin VB.Label Label_Ver_Tv 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tv"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   4680
         TabIndex        =   92
         Top             =   675
         Width           =   210
      End
      Begin VB.Label Label_Placa_Mascara 
         AutoSize        =   -1  'True
         BackColor       =   &H00E23701&
         BackStyle       =   0  'Transparent
         Caption         =   "Placa de som não detectada"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   1200
         TabIndex        =   78
         Top             =   3520
         Width           =   2115
      End
      Begin VB.Label Label_Mudo 
         AutoSize        =   -1  'True
         BackColor       =   &H00131313&
         BackStyle       =   0  'Transparent
         Caption         =   "Mudo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   5280
         TabIndex        =   77
         Top             =   4200
         Width           =   390
      End
      Begin VB.Label Label_Ver_Radio 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Rádio"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   3480
         TabIndex        =   76
         Top             =   675
         Width           =   480
      End
      Begin VB.Image Image7 
         Enabled         =   0   'False
         Height          =   465
         Left            =   6980
         Picture         =   "Form_Principal.frx":1164CE
         Top             =   525
         Width           =   45
      End
      Begin VB.Label Label_Ver_Navegador 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Navegador"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   1920
         TabIndex        =   75
         Top             =   675
         Width           =   930
      End
      Begin VB.Image Image6 
         Enabled         =   0   'False
         Height          =   465
         Left            =   3120
         Picture         =   "Form_Principal.frx":116684
         Top             =   525
         Width           =   45
      End
      Begin VB.Label Label_Ver_Biblioteca 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Biblioteca"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   615
         TabIndex        =   74
         Top             =   675
         Width           =   825
      End
      Begin VB.Image Image9 
         Enabled         =   0   'False
         Height          =   465
         Left            =   1680
         Picture         =   "Form_Principal.frx":11683A
         Top             =   525
         Width           =   45
      End
      Begin VB.Label Label_Faixa_Mascara 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Utilitários"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   480
         TabIndex        =   73
         Top             =   1125
         Width           =   645
      End
      Begin VB.Label Label_Faixa_Actual_Mascara 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "0 de 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   165
         Left            =   5640
         TabIndex        =   72
         Top             =   1410
         Width           =   1500
      End
      Begin VB.Label Label_Titulo_Mascara 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H004E4E4E&
         BackStyle       =   0  'Transparent
         Caption         =   "NPlayer - Electric Nikyts"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   600
         TabIndex        =   71
         Top             =   150
         Width           =   6255
      End
      Begin VB.Label Label_Ver_Video 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Video"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   5520
         TabIndex        =   96
         Top             =   3510
         Width           =   810
      End
      Begin VB.Label Label_Ver_Lista 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Lista"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   6420
         TabIndex        =   95
         Top             =   3510
         Width           =   810
      End
      Begin VB.Image Image12 
         Enabled         =   0   'False
         Height          =   345
         Left            =   6420
         Picture         =   "Form_Principal.frx":1169F0
         Top             =   3450
         Width           =   810
      End
      Begin VB.Image Image11 
         Enabled         =   0   'False
         Height          =   345
         Left            =   5520
         Picture         =   "Form_Principal.frx":1178EE
         Top             =   3450
         Width           =   810
      End
      Begin VB.Image Image13 
         Height          =   690
         Left            =   5610
         Picture         =   "Form_Principal.frx":1187EC
         Top             =   2520
         Width           =   1485
      End
   End
   Begin VB.Frame Frame_Centro 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   13935
      Begin VB.ListBox Lista_Directorios 
         BackColor       =   &H0000FF00&
         Height          =   1035
         Left            =   4560
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   901
      End
      Begin NPlayer.McImageList McImageList1 
         Left            =   4680
         Top             =   3000
         _ExtentX        =   661
         _ExtentY        =   873
         Images0         =   "Form_Principal.frx":11BE16
         ImageCount      =   1
      End
      Begin NPlayer.McListBox Grelha 
         Height          =   1215
         Left            =   4560
         TabIndex        =   61
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         Picture         =   "Form_Principal.frx":11C174
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
      End
      Begin VB.PictureBox Barra_Biblioteca 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   120
         Picture         =   "Form_Principal.frx":11C190
         ScaleHeight     =   360
         ScaleWidth      =   3675
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   0
         Width           =   3675
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Utilitários"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   60
            Top             =   100
            Width           =   3675
         End
      End
      Begin VB.Frame Frame_Lista 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   3675
         Begin VB.Frame Frame_Utilitarios 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1305
            Left            =   240
            TabIndex        =   104
            Top             =   1470
            Width           =   3615
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   36
               X1              =   480
               X2              =   840
               Y1              =   1080
               Y2              =   1080
            End
            Begin VB.Label Label_Utilitarios_Agenda 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Agenda de contactos"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   109
               Top             =   960
               Width           =   1785
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   180
               Index           =   24
               Left            =   840
               Picture         =   "Form_Principal.frx":120D92
               Top             =   960
               Width           =   240
            End
            Begin VB.Image No_Utilitarios 
               Height          =   135
               Left            =   0
               Picture         =   "Form_Principal.frx":121014
               Top             =   40
               Width           =   135
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   34
               X1              =   480
               X2              =   840
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   180
               Index           =   14
               Left            =   360
               Picture         =   "Form_Principal.frx":121152
               Top             =   0
               Width           =   240
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   20
               X1              =   480
               X2              =   480
               Y1              =   240
               Y2              =   1080
            End
            Begin VB.Label Label_Utilitarios 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Utilitários"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   720
               TabIndex        =   108
               Top             =   0
               Width           =   795
            End
            Begin VB.Label Label_Utilitarios_Youtube 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Youtube downloader"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   107
               Top             =   480
               Width           =   1740
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   180
               Index           =   13
               Left            =   840
               Picture         =   "Form_Principal.frx":1213D4
               Top             =   480
               Width           =   240
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   19
               X1              =   480
               X2              =   840
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Label Label_Utilitarios_Tag 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Tag editor"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   106
               Top             =   240
               Width           =   870
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   180
               Index           =   12
               Left            =   840
               Picture         =   "Form_Principal.frx":121656
               Top             =   240
               Width           =   240
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00404040&
               BorderStyle     =   3  'Dot
               Index           =   18
               X1              =   120
               X2              =   350
               Y1              =   120
               Y2              =   120
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   17
               X1              =   480
               X2              =   840
               Y1              =   840
               Y2              =   840
            End
            Begin VB.Label Label_Utilitarios_Gestor 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Gestor de filmes"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   105
               Top             =   720
               Width           =   1410
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   180
               Index           =   11
               Left            =   840
               Picture         =   "Form_Principal.frx":1218D8
               Top             =   720
               Width           =   240
            End
         End
         Begin VB.Frame Frame_Tv 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   240
            TabIndex        =   52
            Top             =   1250
            Width           =   3615
            Begin VB.Image No_Tv 
               Height          =   135
               Left            =   0
               Picture         =   "Form_Principal.frx":121B5A
               Top             =   40
               Width           =   135
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   180
               Index           =   23
               Left            =   840
               Picture         =   "Form_Principal.frx":121C98
               Top             =   720
               Width           =   165
            End
            Begin VB.Label Label_Tv_Tuga 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Tv Tuga"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   56
               Top             =   720
               Width           =   690
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   35
               X1              =   480
               X2              =   840
               Y1              =   840
               Y2              =   840
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00404040&
               BorderStyle     =   3  'Dot
               Index           =   33
               X1              =   120
               X2              =   350
               Y1              =   120
               Y2              =   120
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   180
               Index           =   22
               Left            =   840
               Picture         =   "Form_Principal.frx":121E8A
               Top             =   240
               Width           =   165
            End
            Begin VB.Label Label_Tv_Iniciar 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Iniciar ligação"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   55
               Top             =   240
               Width           =   1200
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   32
               X1              =   480
               X2              =   840
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   180
               Index           =   21
               Left            =   840
               Picture         =   "Form_Principal.frx":12207C
               Top             =   480
               Width           =   165
            End
            Begin VB.Label Label_Tv_Ver 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Ver canais disponiveis"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   54
               Top             =   480
               Width           =   1905
            End
            Begin VB.Label Label_Tv 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Televisão"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   720
               TabIndex        =   53
               Top             =   0
               Width           =   810
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   31
               X1              =   480
               X2              =   480
               Y1              =   240
               Y2              =   840
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   180
               Index           =   20
               Left            =   360
               Picture         =   "Form_Principal.frx":12226E
               Top             =   0
               Width           =   240
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   30
               X1              =   480
               X2              =   840
               Y1              =   600
               Y2              =   600
            End
         End
         Begin VB.Frame Frame_Radio 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   240
            TabIndex        =   48
            Top             =   1020
            Width           =   3615
            Begin VB.Image No_Radio 
               Height          =   135
               Left            =   0
               Picture         =   "Form_Principal.frx":1224F0
               Top             =   40
               Width           =   135
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00404040&
               BorderStyle     =   3  'Dot
               Index           =   25
               X1              =   60
               X2              =   60
               Y1              =   120
               Y2              =   3240
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00404040&
               BorderStyle     =   3  'Dot
               Index           =   26
               X1              =   120
               X2              =   350
               Y1              =   120
               Y2              =   120
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   28
               X1              =   480
               X2              =   840
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   195
               Index           =   19
               Left            =   360
               Picture         =   "Form_Principal.frx":12262E
               Top             =   0
               Width           =   225
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   29
               X1              =   480
               X2              =   480
               Y1              =   240
               Y2              =   600
            End
            Begin VB.Label Label_Radio 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Radio"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   720
               TabIndex        =   51
               Top             =   0
               Width           =   480
            End
            Begin VB.Label Label_Radio_Ver 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Ver canais disponiveis"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   50
               Top             =   480
               Width           =   1905
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   210
               Index           =   18
               Left            =   840
               Picture         =   "Form_Principal.frx":1228E0
               Top             =   480
               Width           =   210
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   27
               X1              =   480
               X2              =   840
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Label Label_Radio_Iniciar 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Iniciar ligação"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   49
               Top             =   240
               Width           =   1200
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   210
               Index           =   17
               Left            =   840
               Picture         =   "Form_Principal.frx":122B8A
               Top             =   240
               Width           =   210
            End
         End
         Begin VB.Frame Frame_Navegador 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   240
            TabIndex        =   44
            Top             =   800
            Width           =   3615
            Begin VB.Image No_Navegador 
               Height          =   135
               Left            =   0
               Picture         =   "Form_Principal.frx":122E34
               Top             =   40
               Width           =   135
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00404040&
               BorderStyle     =   3  'Dot
               Index           =   24
               X1              =   60
               X2              =   60
               Y1              =   120
               Y2              =   3240
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00404040&
               BorderStyle     =   3  'Dot
               Index           =   23
               X1              =   120
               X2              =   350
               Y1              =   120
               Y2              =   120
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   150
               Index           =   16
               Left            =   870
               Picture         =   "Form_Principal.frx":122F72
               Top             =   240
               Width           =   165
            End
            Begin VB.Label Label_Navegador_Abrir 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Abrir janela"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   47
               Top             =   240
               Width           =   1005
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   22
               X1              =   480
               X2              =   840
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   150
               Index           =   15
               Left            =   870
               Picture         =   "Form_Principal.frx":12311C
               Top             =   480
               Width           =   165
            End
            Begin VB.Label Label_Navegador_Ver 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Ver favoritos"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   46
               Top             =   480
               Width           =   1110
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   21
               X1              =   480
               X2              =   840
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Label Label_Navegador 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Navegador"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   720
               TabIndex        =   45
               Top             =   0
               Width           =   930
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   12
               X1              =   480
               X2              =   480
               Y1              =   240
               Y2              =   600
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   195
               Index           =   8
               Left            =   360
               Picture         =   "Form_Principal.frx":1232C6
               Top             =   0
               Width           =   225
            End
         End
         Begin VB.Frame Frame_Listas 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   240
            TabIndex        =   41
            Top             =   570
            Width           =   3615
            Begin VB.Image No_Listas 
               Height          =   135
               Left            =   0
               Picture         =   "Form_Principal.frx":123578
               Top             =   40
               Width           =   135
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00404040&
               BorderStyle     =   3  'Dot
               Index           =   14
               X1              =   120
               X2              =   350
               Y1              =   120
               Y2              =   120
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00404040&
               BorderStyle     =   3  'Dot
               Index           =   15
               X1              =   60
               X2              =   60
               Y1              =   120
               Y2              =   3240
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   195
               Index           =   5
               Left            =   360
               Picture         =   "Form_Principal.frx":1236B6
               Top             =   0
               Width           =   240
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   9
               X1              =   480
               X2              =   480
               Y1              =   240
               Y2              =   360
            End
            Begin VB.Label Label_Listas 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Coleção de listas"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   720
               TabIndex        =   43
               Top             =   0
               Width           =   1455
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   13
               X1              =   480
               X2              =   840
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Label Label_Listas_Abrir 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Procurar listas"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   42
               Top             =   240
               Width           =   1230
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   225
               Index           =   9
               Left            =   840
               Picture         =   "Form_Principal.frx":123968
               Top             =   240
               Width           =   195
            End
         End
         Begin VB.Frame Frame_Loja 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   240
            TabIndex        =   37
            Top             =   345
            Width           =   3615
            Begin VB.Image No_Loja 
               Height          =   135
               Left            =   0
               Picture         =   "Form_Principal.frx":123C02
               Top             =   40
               Width           =   135
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00404040&
               BorderStyle     =   3  'Dot
               Index           =   4
               X1              =   60
               X2              =   60
               Y1              =   120
               Y2              =   3240
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00404040&
               BorderStyle     =   3  'Dot
               Index           =   7
               X1              =   120
               X2              =   350
               Y1              =   120
               Y2              =   120
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   225
               Index           =   7
               Left            =   360
               Picture         =   "Form_Principal.frx":123D40
               Top             =   0
               Width           =   240
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   11
               X1              =   480
               X2              =   480
               Y1              =   240
               Y2              =   600
            End
            Begin VB.Label Label_Loja 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Loja de música"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   720
               TabIndex        =   40
               Top             =   0
               Width           =   1290
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   10
               X1              =   480
               X2              =   840
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Label Label_Loja_Dilandau 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Dilandau"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   39
               Top             =   480
               Width           =   750
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   210
               Index           =   6
               Left            =   840
               Picture         =   "Form_Principal.frx":124052
               Top             =   480
               Width           =   225
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   8
               X1              =   480
               X2              =   840
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Label Label_Loja_Nikyts 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Nikyts"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   38
               Top             =   240
               Width           =   525
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   210
               Index           =   4
               Left            =   840
               Picture         =   "Form_Principal.frx":124334
               Top             =   240
               Width           =   225
            End
         End
         Begin VB.Frame Frame_Pesquisa 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   240
            TabIndex        =   31
            Top             =   120
            Width           =   3615
            Begin VB.Image No_Pesquisa 
               Height          =   135
               Left            =   0
               Picture         =   "Form_Principal.frx":124616
               Top             =   40
               Width           =   135
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00404040&
               BorderStyle     =   3  'Dot
               Index           =   0
               X1              =   60
               X2              =   60
               Y1              =   120
               Y2              =   3240
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00404040&
               BorderStyle     =   3  'Dot
               Index           =   2
               X1              =   120
               X2              =   350
               Y1              =   120
               Y2              =   120
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   225
               Index           =   3
               Left            =   840
               Picture         =   "Form_Principal.frx":124754
               Top             =   480
               Width           =   225
            End
            Begin VB.Label Label_Pesquisa_Todos 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Pesquisa automática"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   36
               Top             =   480
               Width           =   1755
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   3
               X1              =   480
               X2              =   840
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   210
               Index           =   2
               Left            =   840
               Picture         =   "Form_Principal.frx":124A66
               Top             =   960
               Width           =   195
            End
            Begin VB.Label Label_Pesquisa_Filmes 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Filmes"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   35
               Top             =   960
               Width           =   540
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   6
               X1              =   480
               X2              =   840
               Y1              =   1080
               Y2              =   1080
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   225
               Index           =   1
               Left            =   840
               Picture         =   "Form_Principal.frx":124CD8
               Top             =   720
               Width           =   225
            End
            Begin VB.Label Label_Pesquisa_Musicas 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Músicas"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   34
               Top             =   720
               Width           =   660
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   5
               X1              =   480
               X2              =   840
               Y1              =   840
               Y2              =   840
            End
            Begin VB.Label Label_Biblioteca 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Biblioteca"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   720
               TabIndex        =   33
               Top             =   0
               Width           =   825
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   1
               X1              =   480
               X2              =   480
               Y1              =   240
               Y2              =   1080
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   225
               Index           =   0
               Left            =   360
               Picture         =   "Form_Principal.frx":124FEA
               Top             =   0
               Width           =   225
            End
            Begin VB.Line Line_Arvore 
               BorderColor     =   &H00808080&
               BorderStyle     =   3  'Dot
               Index           =   16
               X1              =   480
               X2              =   840
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Label Label_Em_Reproducao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Lista em reprodução"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1200
               TabIndex        =   32
               Top             =   240
               Width           =   1755
            End
            Begin VB.Image Icon_Arvore 
               Enabled         =   0   'False
               Height          =   225
               Index           =   10
               Left            =   840
               Picture         =   "Form_Principal.frx":1252FC
               Top             =   240
               Width           =   225
            End
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            BorderWidth     =   2
            X1              =   4020
            X2              =   4020
            Y1              =   0
            Y2              =   4920
         End
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   735
         Left            =   240
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   5160
         Visible         =   0   'False
         Width           =   3975
         ExtentX         =   7011
         ExtentY         =   1296
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
      Begin VB.Image Skin_Lateral_Esquerda 
         Enabled         =   0   'False
         Height          =   1275
         Left            =   0
         Picture         =   "Form_Principal.frx":12560E
         Top             =   0
         Width           =   120
      End
      Begin VB.Image Skin_Lateral_Direita 
         Enabled         =   0   'False
         Height          =   1110
         Left            =   13200
         Picture         =   "Form_Principal.frx":125E48
         Top             =   0
         Width           =   120
      End
   End
   Begin VB.Frame Frame_Down 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      TabIndex        =   16
      Top             =   9120
      Width           =   14175
      Begin VB.PictureBox Skin_Down_Botoes 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   840
         Picture         =   "Form_Principal.frx":12657A
         ScaleHeight     =   600
         ScaleWidth      =   11415
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   11415
         Begin VB.Timer Timer_Duracao 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   4440
            Top             =   120
         End
         Begin VB.Timer Timer_Slider_Video 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   8280
            Top             =   60
         End
         Begin VB.PictureBox SliderBar 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Height          =   135
            Left            =   6720
            ScaleHeight     =   135
            ScaleWidth      =   2100
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   180
            Width           =   2100
            Begin NPlayer.N_Button Slide 
               Height          =   105
               Left            =   0
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   10
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   185
               Picture         =   "Form_Principal.frx":13CB3C
               PictureHover    =   "Form_Principal.frx":13CD86
               PictureDown     =   "Form_Principal.frx":13CFD0
            End
            Begin VB.Image Image_Barra_Slide 
               Enabled         =   0   'False
               Height          =   135
               Left            =   -1080
               Picture         =   "Form_Principal.frx":13D21A
               Stretch         =   -1  'True
               Top             =   0
               Width           =   2100
            End
         End
         Begin VB.PictureBox Picture_Slide_Som 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Left            =   9720
            ScaleHeight     =   135
            ScaleWidth      =   1500
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   180
            Width           =   1500
            Begin NPlayer.N_Button Slide_Som 
               Height          =   105
               Left            =   0
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   10
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   185
               Picture         =   "Form_Principal.frx":13E7E0
               PictureHover    =   "Form_Principal.frx":13EA2A
               PictureDown     =   "Form_Principal.frx":13EC74
            End
            Begin VB.Image Image_Barra_Slide_Som 
               Enabled         =   0   'False
               Height          =   135
               Left            =   0
               Picture         =   "Form_Principal.frx":13EEBE
               Top             =   0
               Width           =   1485
            End
         End
         Begin NPlayer.N_Button Botao_Antes 
            Height          =   360
            Left            =   370
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   60
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            Picture         =   "Form_Principal.frx":13F98C
            PictureHover    =   "Form_Principal.frx":14009E
            PictureDown     =   "Form_Principal.frx":1407B0
         End
         Begin NPlayer.N_Button Botao_Primeiro 
            Height          =   360
            Left            =   0
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   60
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            Picture         =   "Form_Principal.frx":140EC2
            PictureHover    =   "Form_Principal.frx":1415D4
            PictureDown     =   "Form_Principal.frx":141CE6
         End
         Begin NPlayer.N_Button Botao_Pause 
            Height          =   360
            Left            =   1260
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   60
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            Picture         =   "Form_Principal.frx":1423F8
            PictureHover    =   "Form_Principal.frx":142B0A
            PictureDown     =   "Form_Principal.frx":14321C
         End
         Begin NPlayer.N_Button Botao_Pasta 
            Height          =   360
            Left            =   1970
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   60
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
            Picture         =   "Form_Principal.frx":14392E
            PictureHover    =   "Form_Principal.frx":1440A0
            PictureDown     =   "Form_Principal.frx":144812
         End
         Begin NPlayer.N_Button Botao_Play 
            Height          =   360
            Left            =   920
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   60
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            Picture         =   "Form_Principal.frx":144F84
            PictureHover    =   "Form_Principal.frx":145696
            PictureDown     =   "Form_Principal.frx":145DA8
         End
         Begin NPlayer.N_Button Botao_Stop 
            Height          =   360
            Left            =   1610
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   60
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            Picture         =   "Form_Principal.frx":1464BA
            PictureHover    =   "Form_Principal.frx":146BCC
            PictureDown     =   "Form_Principal.frx":1472DE
         End
         Begin NPlayer.N_Button Botao_Seguinte 
            Height          =   360
            Left            =   2570
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   60
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            Picture         =   "Form_Principal.frx":1479F0
            PictureHover    =   "Form_Principal.frx":148102
            PictureDown     =   "Form_Principal.frx":148814
         End
         Begin NPlayer.N_Button Botao_Ultimo 
            Height          =   360
            Left            =   2930
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   60
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            Picture         =   "Form_Principal.frx":148F26
            PictureHover    =   "Form_Principal.frx":149638
            PictureDown     =   "Form_Principal.frx":149D4A
         End
         Begin VB.Label Label_Faixa_Actual 
            AutoSize        =   -1  'True
            BackColor       =   &H001E1F1D&
            BackStyle       =   0  'Transparent
            Caption         =   "0 de 0"
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
            TabIndex        =   65
            Top             =   150
            Width           =   480
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
            TabIndex        =   64
            Top             =   150
            Width           =   750
         End
         Begin VB.Image Image5 
            Enabled         =   0   'False
            Height          =   360
            Left            =   3600
            Picture         =   "Form_Principal.frx":14A45C
            Top             =   60
            Width           =   2715
         End
         Begin VB.Image Botao_Mudo 
            Height          =   165
            Left            =   9360
            Picture         =   "Form_Principal.frx":14D79E
            Top             =   165
            Width           =   165
         End
      End
      Begin VB.Image Skin_Down_Esquerda 
         Enabled         =   0   'False
         Height          =   600
         Left            =   0
         Picture         =   "Form_Principal.frx":14D96C
         Top             =   0
         Width           =   120
      End
      Begin VB.Image Skin_Down_Direita 
         Enabled         =   0   'False
         Height          =   600
         Left            =   13560
         Picture         =   "Form_Principal.frx":14DD6E
         Top             =   0
         Width           =   120
      End
      Begin VB.Image Skin_Down_Centro 
         Enabled         =   0   'False
         Height          =   600
         Left            =   120
         Picture         =   "Form_Principal.frx":14E170
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.Frame Frame_Top 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1890
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13935
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   375
         Begin WMPLibCtl.WindowsMediaPlayer Wmp 
            Height          =   240
            Left            =   0
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   300
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
            uiMode          =   "none"
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
            _cx             =   529
            _cy             =   423
         End
      End
      Begin NPlayer.N_Button Setas_Selecao 
         Height          =   180
         Left            =   180
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1600
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   318
         Picture         =   "Form_Principal.frx":14F832
         PictureHover    =   "Form_Principal.frx":14FAC4
         PictureDown     =   "Form_Principal.frx":14FD56
      End
      Begin NPlayer.N_Button Botao_Minimizar 
         Height          =   135
         Left            =   13200
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   150
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   238
         Picture         =   "Form_Principal.frx":14FFE8
         PictureHover    =   "Form_Principal.frx":15015A
         PictureDown     =   "Form_Principal.frx":1502CC
      End
      Begin NPlayer.N_Button Botao_Maximizar 
         Height          =   150
         Left            =   13440
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   265
         Picture         =   "Form_Principal.frx":15043E
         PictureHover    =   "Form_Principal.frx":1505A8
         PictureDown     =   "Form_Principal.frx":150712
      End
      Begin NPlayer.N_Button Botao_Fechar 
         Height          =   150
         Left            =   13680
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   265
         Picture         =   "Form_Principal.frx":15087C
         PictureHover    =   "Form_Principal.frx":1509E6
         PictureDown     =   "Form_Principal.frx":150B50
      End
      Begin VB.PictureBox pichook 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   120
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Skin_Top_Player 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1890
         Left            =   960
         Picture         =   "Form_Principal.frx":150CBA
         ScaleHeight     =   1890
         ScaleWidth      =   11415
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   11415
         Begin VB.Timer Timer_Actualiza 
            Interval        =   10
            Left            =   7560
            Top             =   360
         End
         Begin VB.Timer Timer_Conexao 
            Interval        =   10
            Left            =   8640
            Top             =   840
         End
         Begin VB.PictureBox Imagem_Grafico 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FC6C03&
            BorderStyle     =   0  'None
            FillColor       =   &H00FFFFFF&
            ForeColor       =   &H00FFAE0F&
            Height          =   420
            Left            =   3480
            ScaleHeight     =   420
            ScaleWidth      =   3945
            TabIndex        =   11
            Top             =   960
            Width           =   3945
            Begin VB.TextBox Text1 
               BackColor       =   &H0000FF00&
               Height          =   285
               Left            =   480
               TabIndex        =   12
               Text            =   "50"
               Top             =   120
               Visible         =   0   'False
               Width           =   345
            End
            Begin VB.Timer Timer_Grafico 
               Enabled         =   0   'False
               Interval        =   100
               Left            =   0
               Top             =   0
            End
         End
         Begin NPlayer.N_Button Botao_Sobre 
            Height          =   165
            Left            =   7240
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   720
            Width           =   180
            _ExtentX        =   318
            _ExtentY        =   291
            Picture         =   "Form_Principal.frx":197124
            PictureHover    =   "Form_Principal.frx":197302
            PictureDown     =   "Form_Principal.frx":1974E0
         End
         Begin NPlayer.N_Button Botao_Ver_Ecra_Visualizacao 
            Height          =   450
            Left            =   1200
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Ver ecrã de visualização"
            Top             =   1200
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   794
            Picture         =   "Form_Principal.frx":1976BE
            PictureHover    =   "Form_Principal.frx":198430
            PictureDown     =   "Form_Principal.frx":1991A2
         End
         Begin NPlayer.N_Button Botao_Mascara 
            Height          =   450
            Left            =   1920
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Ver em modo mascara"
            Top             =   1200
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   794
            Picture         =   "Form_Principal.frx":199F14
            PictureHover    =   "Form_Principal.frx":19AC86
            PictureDown     =   "Form_Principal.frx":19B9F8
         End
         Begin VB.Label Label_Propriedades 
            Alignment       =   2  'Center
            BackColor       =   &H00E53100&
            Caption         =   "Propriedades"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF890A&
            Height          =   165
            Left            =   7560
            TabIndex        =   66
            Top             =   720
            Width           =   1035
         End
         Begin VB.Label Label_Percentagem_Volume 
            AutoSize        =   -1  'True
            BackColor       =   &H00E23701&
            BackStyle       =   0  'Transparent
            Caption         =   "Volume 50%"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H001E1F1D&
            Height          =   180
            Left            =   7560
            TabIndex        =   63
            Top             =   1425
            Width           =   990
         End
         Begin VB.Label Label_Volume 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   8280
            TabIndex        =   62
            Top             =   960
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label Label_Conexao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Serviço de internet: "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00201D1F&
            Height          =   180
            Left            =   3240
            TabIndex        =   15
            ToolTipText     =   "Verificar a conexão á internet"
            Top             =   720
            Width           =   1530
         End
         Begin VB.Label Label_Placa 
            AutoSize        =   -1  'True
            BackColor       =   &H00FC6C03&
            BackStyle       =   0  'Transparent
            Caption         =   "Placa de som não detectada"
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
            Left            =   3480
            TabIndex        =   13
            Top             =   1425
            Width           =   2115
         End
         Begin VB.Image Botao_Repetir_Album 
            Height          =   165
            Left            =   2880
            Picture         =   "Form_Principal.frx":19C76A
            Top             =   1420
            Width           =   480
         End
         Begin VB.Image Botao_Repetir_Faixa 
            Height          =   165
            Left            =   2880
            Picture         =   "Form_Principal.frx":19CBCC
            Top             =   1220
            Width           =   480
         End
         Begin VB.Image Botao_Ordenar 
            Height          =   165
            Left            =   2880
            Picture         =   "Form_Principal.frx":19D02E
            Top             =   1020
            Width           =   480
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   270
            Left            =   2880
            Picture         =   "Form_Principal.frx":19D490
            Top             =   690
            Width           =   315
         End
         Begin VB.Label Label_Titulo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H004E4E4E&
            BackStyle       =   0  'Transparent
            Caption         =   "NPlayer - Electric Nikyts"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   3
            Top             =   120
            Width           =   11415
         End
         Begin VB.Image Botao_Duas_Frames 
            Height          =   450
            Left            =   1200
            Picture         =   "Form_Principal.frx":19D952
            ToolTipText     =   "Ver tela em duas frames"
            Top             =   600
            Width           =   555
         End
         Begin VB.Image Botao_Uma_Frame 
            Height          =   450
            Left            =   1920
            Picture         =   "Form_Principal.frx":19E6B4
            ToolTipText     =   "Ver tela em uma frame"
            Top             =   600
            Width           =   555
         End
         Begin VB.Image Botao_Repetir 
            Height          =   450
            Left            =   9720
            Picture         =   "Form_Principal.frx":19F416
            ToolTipText     =   "Repetir lista"
            Top             =   600
            Width           =   555
         End
         Begin VB.Image Botao_Ocultar 
            Height          =   450
            Left            =   9000
            Picture         =   "Form_Principal.frx":1A0178
            ToolTipText     =   "Ocultar player"
            Top             =   600
            Width           =   555
         End
         Begin VB.Image Botao_Listas 
            Height          =   450
            Left            =   9720
            Picture         =   "Form_Principal.frx":1A0EDA
            ToolTipText     =   "Listas"
            Top             =   1200
            Width           =   555
         End
         Begin VB.Image Botao_Rename 
            Height          =   450
            Left            =   9000
            Picture         =   "Form_Principal.frx":1A1C3C
            ToolTipText     =   "Editor de ficheiros"
            Top             =   1200
            Width           =   555
         End
      End
      Begin VB.Image Skin_Top_Esquerda 
         Enabled         =   0   'False
         Height          =   1890
         Left            =   0
         Picture         =   "Form_Principal.frx":1A299E
         Top             =   0
         Width           =   120
      End
      Begin VB.Image Skin_Top_Direita 
         Enabled         =   0   'False
         Height          =   1890
         Left            =   12960
         Picture         =   "Form_Principal.frx":1A35B0
         Top             =   0
         Width           =   120
      End
      Begin VB.Image Skin_Top_Centro 
         Height          =   1890
         Left            =   120
         Picture         =   "Form_Principal.frx":1A41C2
         Top             =   0
         Width           =   885
      End
   End
   Begin VB.Image Imagem_Fundo_Menu 
      Height          =   4845
      Left            =   14520
      Picture         =   "Form_Principal.frx":1A9A9C
      Top             =   0
      Visible         =   0   'False
      Width           =   5190
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Adicionar 
         Caption         =   "Adicionar"
         Begin VB.Menu Adicionar_Ficheiro 
            Caption         =   "Adicionar ficheiro"
         End
         Begin VB.Menu Adicionar_Pasta 
            Caption         =   "Adicionar pasta"
         End
         Begin VB.Menu Pesquisar_Ficheiros_No_Computador 
            Caption         =   "Pesquisar ficheiros no computador"
         End
      End
      Begin VB.Menu Lina3 
         Caption         =   "-"
      End
      Begin VB.Menu Manipular_Lista 
         Caption         =   "Manipular lista"
         Begin VB.Menu Remover_Linha 
            Caption         =   "Remover linha"
         End
         Begin VB.Menu Limpar_Tudo 
            Caption         =   "Limpar tudo"
         End
      End
      Begin VB.Menu Ir_Para 
         Caption         =   "Ir para"
         Begin VB.Menu Primeira_Linha 
            Caption         =   "Primeira linha"
         End
         Begin VB.Menu Linha_Anterior 
            Caption         =   "Linha anterior"
         End
         Begin VB.Menu Linha_Seguinte 
            Caption         =   "Linha seguinte"
         End
         Begin VB.Menu Ultima_Linha 
            Caption         =   "Ultima linha"
         End
      End
      Begin VB.Menu Linha1 
         Caption         =   "-"
      End
      Begin VB.Menu Procurar 
         Caption         =   "Procurar"
         Begin VB.Menu Procurar_Texto_NA_Lista 
            Caption         =   "Procurar texto na lista"
         End
         Begin VB.Menu Procurar_Lista_Existente 
            Caption         =   "Procurar lista existente"
         End
      End
      Begin VB.Menu Loja_De_Musica 
         Caption         =   "Loja de música"
         Begin VB.Menu Nikyts 
            Caption         =   "Nikyts"
         End
         Begin VB.Menu Dilandau 
            Caption         =   "Dilandau"
         End
      End
      Begin VB.Menu Traco2 
         Caption         =   "-"
      End
      Begin VB.Menu Guardar_Lista 
         Caption         =   "Guardar lista"
      End
      Begin VB.Menu Controles_Do_Player 
         Caption         =   "Controles do player"
         Begin VB.Menu Parar 
            Caption         =   "Parar"
         End
         Begin VB.Menu Reproduzir 
            Caption         =   "Reproduzir"
         End
         Begin VB.Menu Paausa 
            Caption         =   "Paausa"
         End
         Begin VB.Menu Menu_Controles_Mudo 
            Caption         =   "Mudo"
         End
      End
      Begin VB.Menu Traco4 
         Caption         =   "-"
      End
      Begin VB.Menu Ver 
         Caption         =   "Ver"
         Begin VB.Menu Tela_Em_Cheio 
            Caption         =   "Tela em cheio"
         End
         Begin VB.Menu Ver_Duas_Frames 
            Caption         =   "Ver em duas Frames"
         End
      End
   End
   Begin VB.Menu Menu_Icon 
      Caption         =   "Menu Icon"
      Visible         =   0   'False
      Begin VB.Menu Icon_Restaurar 
         Caption         =   "Restaurar"
      End
      Begin VB.Menu Traco_Icon1 
         Caption         =   "-"
      End
      Begin VB.Menu Icon_Parar 
         Caption         =   "Parar"
      End
      Begin VB.Menu Icon_Reproduzir 
         Caption         =   "Reproduzir"
      End
      Begin VB.Menu Icon_Pausa 
         Caption         =   "Pausa"
      End
      Begin VB.Menu Menu_Icon_Mudo 
         Caption         =   "Mudo"
      End
      Begin VB.Menu Traco_Icon2 
         Caption         =   "-"
      End
      Begin VB.Menu Icon_Fechar 
         Caption         =   "Fechar"
      End
   End
End
Attribute VB_Name = "Form_Principal"
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
'VARIÁVERIS DO SLIDER VIDEO
Dim tx As Integer, Ty As Integer, DN As Boolean
Dim Txa As Integer, DNa As Boolean
Dim Tyb, DNb As Boolean
Dim NewLeft As Integer

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

'Faixa em reproducao
Public Faixa_em_Reproducao As String

'Com/ Sem som
Public Mudo As Boolean

'Tipo_De_Ficheiro_Escolhido
Public Tipo_De_Ficheiro_Escolhido As String

''VARIAVEIS LISTAR ARQUIVOS
'Dim PicHeight%, hLB&, FileSpec$, UseFileSpec%
'Dim TotalDirs%, TotalFiles%, Running%
'Dim WFD As WIN32_FIND_DATA, hItem&, hFile&
'Const vbBackslash = "\"
'Const vbAllFiles = "*.*"
'Const vbKeyDot = 46

'VARIAVEIS DRIVE
Private m_strPasta As String

'VARIAVEIS AUTOMATICA
Dim Arquivos() As String
Dim F As Integer

'Variavel para gardar o conteudo da Grelha
Dim i As Integer

'Variavel para ver a duracao do ficheiro a reproduzir
Public VideoDuration As Double

'MOVER FORMULÁRIO
Dim H, v As Long

'play ou pause
Public Musica_Play As Boolean

'tray icon
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim t As NOTIFYICONDATA

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_FINDSTRING = &H18F

'DETECTAR PLACA DE SOM
Private Declare Function waveOutGetNumDevs Lib "winmm" () As Long

'PROGRESS BAR EDITADO
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_USER = &H400
Const CCM_FIRST = &H2000&
Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
Const PBM_SETBARCOLOR = (WM_USER + 9)

'API par verificar a ligação á net
Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Private Const CONNECT_LAN As Long = &H2
Private Const CONNECT_MODEM As Long = &H1
Private Const CONNECT_PROXY As Long = &H4
Private Const CONNECT_RAS As Long = &H10
Private Const CONNECT_OFFLINE As Long = &H20
Private Const CONNECT_CONFIGURED As Long = &H40

'Variável para verificar se está em modo hide (stray icon)
Dim Modo_Tray As Boolean

'Variáveis para criar o gráfico da barra top centro
Dim Numbars
Dim Step As Single
Dim BWidth  As Single
Dim X As Integer
Dim HArray(0 To 99) As Single

'Ajusta o Form para sempre exibir a barra de tarefas do windows, full screen
Private Const SPI_GETWORKAREA = 48
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Declare Function SystemParametersInfo Lib "user32" _
Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Variavel para verificar a janela do formulário
Dim Tela_Cheia As Boolean

Public Function PosFormRelativeTaskBar(F As Form)
    'Função para ao maximizar o form seja visivel a barra do windows iniciar
    'Colocar o WindowsState=0 normal
    Dim WindowRect As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
    
    SetWindowPos hwnd, 0, WindowRect.Left, WindowRect.Top, WindowRect.Right - WindowRect.Left, WindowRect.Bottom - WindowRect.Top, 0
    F.Top = WindowRect.Bottom * Screen.TwipsPerPixelY - F.Height
    F.Left = WindowRect.Right * Screen.TwipsPerPixelX - F.Width
End Function

Public Function IsWebConnected(Optional ByRef ConnType As String) As Boolean
    'Função par verificar a ligação á net
    Dim dwFlags As Long
    Dim WebTest As Boolean
    ConnType = ""
    WebTest = InternetGetConnectedState(dwFlags, 0&)
    Select Case WebTest
        Case dwFlags And CONNECT_LAN: ConnType = "LAN"
        Case dwFlags And CONNECT_MODEM: ConnType = "Modem"
        Case dwFlags And CONNECT_PROXY: ConnType = "Proxy"
        Case dwFlags And CONNECT_OFFLINE: ConnType = "Offline"
        Case dwFlags And CONNECT_CONFIGURED: ConnType = "Config"
        Case dwFlags And CONNECT_RAS: ConnType = "Remota"
    End Select
    IsWebConnected = WebTest
End Function

Private Sub Botao_Sobre_Click()
    'Ver direitos de autor
    Form_Sobre.Show vbModal
End Sub

Private Sub Botao_Fechar_Click()
    'Fechar programa
    'Cancelar media player
    Botao_Stop_Click
    'Salvar List1
    For i = 0 To Grelha.ListCount - 1
        SaveSetting "NPlayer", "Valores_Gravados", "Conteudo_Grelha" & i, Grelha.List(i)
    Next i
    SaveSetting "NPlayer", "Valores_Gravados", "Conteudo_Grelha", Grelha.ListCount

    'Salvar Lista_Directorios
    For j = 0 To Lista_Directorios.ListCount - 1
        SaveSetting "NPlayer", "Valores_Gravados", "Conteudo_Lista_Directorios" & j, Lista_Directorios.List(j)
    Next j
    SaveSetting "NPlayer", "Valores_Gravados", "Conteudo_Lista_Directorios", Lista_Directorios.ListCount

    'Salvar url do media player
    SaveSetting "NPlayer", "Valores_Gravados", "Conteudo_Wmp", Wmp.URL
    SaveSetting "NPlayer", "Valores_Gravados", "Conteudo_Linha", Grelha.ListIndex

    'Remover do sistema o icon do programa
    t.cbSize = Len(t)
    t.hwnd = pichook.hwnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t  'Remove o ícone da barra de tarefas.

    'Fechar programa
    Unload Me
    'Unload Form_Lista
    Unload Form_Wmp
    End
End Sub

Private Sub Botao_Maximizar_Click()
    'Maximixar/ Restaurar Formulários
    If Tela_Cheia = True Then
        With Me
            .Height = 9720
            .Width = 14340
        End With
        Tela_Cheia = False
    Else
        PosFormRelativeTaskBar Me
        Tela_Cheia = True
    End If
End Sub

Private Sub Botao_Minimizar_Click()
    'Minimizar programa
    'On Error Resume Next
    'Me.WindowState = 1
    Me.Hide
    Modo_Tray = True
    'Form_Lista.Hide
    'Form_Wmp.Hide
End Sub
Sub Drawbox(boxLeft As Single, BoxWidth As Single, boxColor As ColorConstants, Value As Single, Index As Integer)
    'Procedimento para criar o gráfico
    Dim BoxTop As Single
    Dim C As Integer
      Imagem_Grafico.DrawWidth = 1
      Imagem_Grafico_Mascara.DrawWidth = 1
      Imagem_Grafico.ForeColor = boxColor
      Imagem_Grafico_Mascara.ForeColor = boxColor
      For C = 1 To BWidth
          Imagem_Grafico.Line (boxLeft + C, (Value * (Imagem_Grafico.Height / 100)))-(boxLeft + C, Imagem_Grafico.Height)
          Imagem_Grafico_Mascara.Line (boxLeft + C, (Value * (Imagem_Grafico_Mascara.Height / 100)))-(boxLeft + C, Imagem_Grafico_Mascara.Height)
      Next C
End Sub

Private Sub Botao_Sobre_Mascara_Click()
    Botao_Sobre_Click
End Sub

Private Sub Frame_Mascara_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    H = X
    v = Y
End Sub

Private Sub Frame_Mascara_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.WindowState = 0 Then
        If Button = vbLeftButton Then
           Me.Left = Me.Left - (H - X)
           Me.Top = Me.Top - (v - Y)
        End If
    End If
    'Chamar o procedimento para repor as imagens originais
    Repor_Imagens
End Sub

Private Sub Label_Mudo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label_Mudo.ForeColor = vbWhite
End Sub

Private Sub Label_Propriedades_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label_Propriedades.ForeColor = vbWhite
End Sub

Private Sub Label_Utilitarios_Agenda_Click()
    Remover_Selecao
    Label_Utilitarios_Agenda.BackColor = &HFC6C03
    Label_Utilitarios_Agenda.ForeColor = vbWhite
End Sub

Private Sub Label_Utilitarios_Click()
    'Selecionar item
    Remover_Selecao
    Label_Utilitarios.BackColor = &HFC6C03
    Label_Utilitarios.ForeColor = vbWhite
End Sub

Private Sub Label_Utilitarios_DblClick()
    'Atalho para
    No_Utilitarios_Click
End Sub

Private Sub Label_Utilitarios_Gestor_Click()
    Remover_Selecao
    Label_Utilitarios_Gestor.BackColor = &HFC6C03
    Label_Utilitarios_Gestor.ForeColor = vbWhite
End Sub

Private Sub Label_Utilitarios_Tag_Click()
    Remover_Selecao
    Label_Utilitarios_Tag.BackColor = &HFC6C03
    Label_Utilitarios_Tag.ForeColor = vbWhite
    Form_Tag.Show
End Sub

Private Sub Label_Utilitarios_Youtube_Click()
    'Abrir o programa youtube downloader
    Remover_Selecao
    Label_Utilitarios_Youtube.BackColor = &HFC6C03
    Label_Utilitarios_Youtube.ForeColor = vbWhite
    Form_Youtube.Show
End Sub

Private Sub Label_Ver_Biblioteca_Click()
    Botao_Maximizar_Mascara_Click
End Sub

Private Sub Label_Ver_Biblioteca_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animação da label
    Label_Ver_Biblioteca.ForeColor = vbWhite
End Sub

Private Sub Label_Ver_Lista_Click()
    'Ver formulário da lista
    Load Form_Lista
    With Form_Lista
        .List1.Clear
        Dim i As Integer
        For i = 0 To Grelha.ListCount - 1
            .List1.AddItem Grelha.List(i), -1, 0
            .List1.ListIndex = Grelha.ListIndex
        Next i
        .Show
    End With
End Sub

Private Sub Label_Ver_Lista_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animação da label
    Label_Ver_Lista.ForeColor = vbWhite
End Sub

Private Sub Label_Ver_Navegador_Click()
    Botao_Maximizar_Mascara_Click
End Sub

Private Sub Label_Ver_Navegador_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animação da label
    Label_Ver_Navegador.ForeColor = vbWhite
End Sub

Private Sub Label_Ver_Propriedades_Click()
    Label_Propriedades_Click
End Sub

Private Sub Label_Ver_Propriedades_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animação da label
    Label_Ver_Propriedades.ForeColor = vbWhite
End Sub

Private Sub Label_Ver_Radio_Click()
    Botao_Maximizar_Mascara_Click
    Label_Radio_Iniciar_Click
End Sub

Private Sub Label_Ver_Radio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animação da label
    Label_Ver_Radio.ForeColor = vbWhite
End Sub

Private Sub Label_Ver_Tv_Click()
    Botao_Maximizar_Mascara_Click
End Sub

Private Sub Label_Ver_Tv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animação da label
    Label_Ver_Tv.ForeColor = vbWhite
End Sub

Private Sub Label_Ver_Video_Click()
    'Ver o formulário video
    Form_Wmp.Show
End Sub

Private Sub Label_Ver_Video_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animação da label
    Label_Ver_Video.ForeColor = vbWhite
End Sub

Private Sub No_Utilitarios_Click()
    'Abrir/ fechar a pasta "Utilitários"
    If Frame_Utilitarios.Height = 225 Then
        Frame_Utilitarios.Height = 1305
        No_Utilitarios.Picture = Form_Imagens.No_Menos.Picture
    Else
        Frame_Utilitarios.Height = 225
        No_Utilitarios.Picture = Form_Imagens.No_Mais.Picture
    End If
    
    'Posicionar restantes frames
    Posicionar_Nos
End Sub

Private Sub Timer_Conexao_Timer()
    'Verificar a conexão á intertnet
    Dim msg As String
    If IsWebConnected(msg) Then
        msg = "Serviço de internet: " & msg
    Else
        msg = "Serviço de internet: Desligada"
    End If
    Label_Conexao.Caption = msg
    Label_Conexao_Mascara.Caption = msg
End Sub

Private Sub Timer_Grafico_Timer()
    'Criar o gráfico
    Static offset As Integer
    On Error Resume Next
    Dim ColorVal As Single
    Dim R As Integer
      Step = 255 / (Numbars + 5)
      Imagem_Grafico.Cls
      Imagem_Grafico_Mascara.Cls
      
      offset = offset + 1
      If offset > Numbars Then offset = 0
      For X = 0 To Numbars - 1
          Randomize Timer
          R = (Int(Rnd * 10) + 1)
          If R / 2 = R \ 2 Then
              If HArray(X) - R < 0 Then
                  HArray(X) = 0
              Else
                  HArray(X) = HArray(X) - R
              End If
          Else
              If HArray(X) + R > Imagem_Grafico.Height Then
                  HArray(X) = Imagem_Grafico.Height
              Else
                  HArray(X) = HArray(X) + R
              End If
          End If
          ColorVal = Step * (X + 5)
          Drawbox BWidth * X, BWidth, RGB(ColorVal, ColorVal, ColorVal), HArray(X), X
      Next X
      
      Imagem_Grafico.ForeColor = Imagem_Grafico.BackColor
      Imagem_Grafico_Mascara.ForeColor = Imagem_Grafico_Mascara.BackColor
      For X = 1 To Imagem_Grafico.Height Step 75
          Imagem_Grafico.Line (0, X)-(Imagem_Grafico.Width, X)
          Imagem_Grafico_Mascara.Line (0, X)-(Imagem_Grafico_Mascara.Width, X)
      Next X
End Sub

Private Sub Posicionar_Nos()
    'Posicionar as frames do "Treeview" cunstomizado
    Frame_Loja.Top = Frame_Pesquisa.Top + Frame_Pesquisa.Height
    Frame_Listas.Top = Frame_Loja.Top + Frame_Loja.Height
    Frame_Navegador.Top = Frame_Listas.Top + Frame_Listas.Height
    Frame_Radio.Top = Frame_Navegador.Top + Frame_Navegador.Height
    Frame_Tv.Top = Frame_Radio.Top + Frame_Radio.Height
    Frame_Utilitarios.Top = Frame_Tv.Top + Frame_Tv.Height
End Sub


Private Sub Adicionar_Ficheiro_Click()
    'Atalho para
    Botao_Pasta_Click
End Sub

Private Sub Adicionar_Pasta_Click()
    'Procuar texto existente na lista
    Y = MsgBox("Opção ainda não disponivel", vbOKOnly + vbInformation, "Uppsss")
End Sub

Private Sub Botao_Duas_Frames_Click()
    'Ver as duas frames
    Barra_Biblioteca.Visible = True
    Botao_Duas_Frames.Picture = Form_Imagens.Duas_Frames_Over.Picture
    Botao_Uma_Frame.Picture = Form_Imagens.Uma_Frame_Normal.Picture
    Frame_Lista.Visible = True
    'Frame_Video.Visible = True
    Line1.Visible = True
    
    'Grelha
    With Grelha
        .Height = Frame_Centro.Height
        .Top = 0
        .Width = Frame_Centro.Width - Skin_Lateral_Esquerda.Width - Skin_Lateral_Direita.Width - Barra_Biblioteca.Width - Line1.BorderWidth
        .Left = Frame_Lista.Left + Frame_Lista.Width
    End With
    
            
    'Ajustar WebBrowser1 consoante as medidas e posicoes da Grelha
    With WebBrowser1
        .Height = Grelha.Height
        .Top = 0
        .Width = Grelha.Width
        .Left = Grelha.Left
    End With
End Sub
Private Sub Botao_Fechar_Mascara_Click()
    Botao_Fechar_Click
End Sub

Private Sub Botao_Listas_Click()
    'Procurar por listas guardadas
    Form_Listas_Guardadas.Show vbModal
End Sub

Private Sub Botao_Mascara_Antes_Click()
    Botao_Antes_Click
    Form_Lista.List1.ListIndex = Grelha.ListIndex
End Sub

Private Sub Botao_Mascara_Click()
    'Ver program em modo mascara
    Modo_Mascara = True
    
    'Frame_Mascara
    With Frame_Mascara
        .Top = 0
        .Width = 7650
        .Left = 0
    End With

    With Me
        .WindowState = 0
        .Height = Frame_Mascara.Height
        .Width = Frame_Mascara.Width
    End With
    Frame_Mascara.Visible = True
    Verificar_Volume
End Sub

Private Sub Botao_Mascara_Pasta_Click()
    Botao_Pasta_Click
End Sub

Private Sub Botao_Mascara_Pause_Click()
    Botao_Pause_Click
End Sub

Private Sub Botao_Mascara_Play_Click()
    Botao_Play_Click
End Sub

Private Sub Botao_Mascara_Primeiro_Click()
    Botao_Primeiro_Click
    Form_Lista.List1.ListIndex = Grelha.ListIndex
End Sub

Private Sub Botao_Mascara_Seguinte_Click()
    Botao_Seguinte_Click
    Form_Lista.List1.ListIndex = Grelha.ListIndex
End Sub

Private Sub Botao_Mascara_Stop_Click()
    Botao_Stop_Click
End Sub

Private Sub Botao_Mascara_Ultimo_Click()
    Botao_Ultimo_Click
    Form_Lista.List1.ListIndex = Grelha.ListIndex
End Sub

Private Sub Botao_Maximizar_Mascara_Click()
    'Ver program em modo full screen
    Modo_Mascara = False
    With Me
        .Height = 10260
        .Width = 14550
        .WindowState = 0
    End With

    'Frame_Mascara
    With Frame_Mascara
        .Visible = False
    End With
    
    'Ocultar os restantes formulários que pertencem ao modo mascara
    Form_Wmp.Hide
    Form_Lista.Hide
End Sub

Private Sub Botao_Minimizar_Mascara_Click()
    Botao_Minimizar_Click
End Sub

Public Sub Botao_Mudo_Click()
    If Mudo = False Then
        Wmp.settings.mute = True
        'Form_Wmp.Wmp.settings.mute = True
        Mudo = True
        Botao_Mudo.Picture = Form_Imagens.Mudo_Over.Picture
        Form_Wmp.Botao_Mudo.Picture = Form_Imagens.Mudo_Over.Picture
        Label_Mudo.Caption = "Ouvir"
        'Menu_Controles_Mudo.Caption = "Ouvir"
        'Menu_Icon_Mudo.Caption = "Ouvir"
    Else
        Wmp.settings.mute = False
        'Form_Wmp.Wmp.settings.mute = False
        Mudo = False
        Botao_Mudo.Picture = Form_Imagens.Mudo_Normal.Picture
        Form_Wmp.Botao_Mudo.Picture = Form_Imagens.Mudo_Normal.Picture
        Label_Mudo.Caption = "Mudo"
        'Menu_Controles_Mudo.Caption = "Mudo"
        'Menu_Icon_Mudo.Caption = "Mudo"
    End If
End Sub

Private Sub Botao_Ocultar_Click()
    'Ocultar formulário e coloca-lo ao lado do clock
    Botao_Minimizar_Click
End Sub

Private Sub Botao_Rename_Click()
    'Ver formulário de edição de tags de ficheiros
    Form_Tag_Editar.Show
End Sub

Private Sub Botao_Repetir_Click()
    MsgBox ("Ainda não desponivel")
End Sub

Private Sub Botao_Ver_Ecra_Visualizacao_Click()
    'Ver formulário de Video
    Form_Wmp.Show
    Frame_Lista.Height = Frame_Centro.Height - Barra_Biblioteca.Height
    'Frame_Video.Visible = False
    Form_Principal.Slide_Som.Left = Slide_Som.Left
End Sub

Public Sub Botao_Pasta_Click()
    'Abrir exploradoe para carregar ficheiros
    Form_Pesquisa.Show vbModal
End Sub

Private Sub Botao_Uma_Frame_Click()
    'Ver apenas uma frame
    Barra_Biblioteca.Visible = False
    Botao_Duas_Frames.Picture = Form_Imagens.Duas_Frames_Normal.Picture
    Botao_Uma_Frame.Picture = Form_Imagens.Uma_Frame_Over.Picture
    Frame_Lista.Visible = False
    'Frame_Video.Visible = False
    Line1.Visible = False
    'Grelha
    With Grelha
        .Height = Frame_Centro.Height
        .Top = 0
        .Width = Frame_Centro.Width - Skin_Lateral_Esquerda.Width - Skin_Lateral_Direita.Width
        .Left = Skin_Lateral_Esquerda.Left + Skin_Lateral_Esquerda.Width
    End With
            
    'Ajustar WebBrowser1 consoante as medidas e posicoes da Grelha
    With WebBrowser1
        .Height = Grelha.Height
        .Top = 0
        .Width = Grelha.Width
        .Left = Grelha.Left
    End With

End Sub

Private Sub Dilandau_Click()
    'Atalho para
    Label_Loja_Dilandau_Click
End Sub

Private Sub Guardar_Lista_Click()
    'Guardar playlist
    Form_Guardar.Show vbModal
End Sub

Private Sub Label_Radio_Iniciar_Click()
    'Iniciar Radio
    Remover_Selecao
    Label_Radio_Iniciar.BackColor = &HFC6C03
    Label_Radio_Iniciar.ForeColor = vbWhite
    Form_Radio.Show
End Sub

Private Sub Label_Radio_Ver_Click()
    'Ver canais de radio disponiveis
    Remover_Selecao
    Label_Radio_Ver.BackColor = &HFC6C03
    Label_Radio_Ver.ForeColor = vbWhite
End Sub

Private Sub Label_Tv_Iniciar_Click()
    'Iniciar programa de tv online
    Remover_Selecao
    Label_Tv_Iniciar.BackColor = &HFC6C03
    Label_Tv_Iniciar.ForeColor = vbWhite
End Sub

Private Sub Label_Tv_Tuga_Click()
    'Abrir site "TV Tuga" para ver tv online
    Remover_Selecao
    Label_Tv_Tuga.BackColor = &HFC6C03
    Label_Tv_Tuga.ForeColor = vbWhite
End Sub

Private Sub Label_Tv_Ver_Click()
    'Ver canais disponiveis
    Remover_Selecao
    Label_Tv_Ver.BackColor = &HFC6C03
    Label_Tv_Ver.ForeColor = vbWhite
End Sub

Private Sub Limpar_Tudo_Click()
    'Verificar se a lista contem ficheiros, caso tenha, limpa a grelha toda
    If Grelha.ListCount = 0 Then Exit Sub
    Dim Temp As String
    Grelha.Clear
    Lista_Directorios.Clear
    If Wmp.URL = Temp Then Wmp.Controls.stop: Timer_Duracao.Enabled = False
End Sub

Private Sub Linha_Anterior_Click()
    'Atalho para
    Botao_Antes_Click
End Sub

Private Sub Linha_Seguinte_Click()
    'Atalho para
    Botao_Seguinte_Click
End Sub

Private Sub Nikyts_Click()
    'Atalho para
    Label_Loja_Nikyts_Click
End Sub

Private Sub Parar_Click()
    'Atalho para
    Botao_Stop_Click
End Sub

Private Sub Pausa_Click()
    'Atalho para
    Botao_Pause_Click
End Sub

Private Sub Pesquisar_Ficheiros_No_Computador_Click()
    'Atalho para
    Label_Pesquisa_Todos_Click
End Sub

Private Sub Primeira_Lina_Click()
    'Atalho para
    Botao_Primeiro_Click
End Sub

Private Sub Procurar_Lista_Existente_Click()
    'Atalho para
    Label_Listas_Abrir_Click
End Sub

Private Sub Procurar_Texto_NA_Lista_Click()
    'Procuar texto existente na lista
    Y = MsgBox("Opção ainda não disponivel", vbOKOnly + vbInformation, "Uppsss")
End Sub

Private Sub Remover_Linha_Click()
    'Verificar se a lista contem ficheiros
    If Grelha.ListCount = 0 Then Exit Sub
    Dim Temp As String
    Temp = Lista_Directorios.Text
    Grelha.Remove Grelha.ListIndex
    Lista_Directorios.RemoveItem Lista_Directorios.ListIndex
    If Wmp.URL = Temp Then Wmp.Controls.stop: Timer_Duracao.Enabled = False
End Sub

Private Sub Reproduzir_Click()
    'Atalho para
    Botao_Play_Click
End Sub

Private Sub Ultima_Linha_Click()
    'Atalho para
    Botao_Ultimo_Click
End Sub

Private Sub Ver_Duas_Frames_Click()
    'Atalho para
    Botao_Duas_Frames_Click
End Sub

Private Sub Form_Activate()
    Form_Wmp.Wmp.settings.mute = True
    
    'Definir cor da linha de selecao da Lista
    Grelha.SelColor = &HFC6C03
End Sub

Private Sub Form_Load()
    'Chamar procedimento pra desenhar o formulário
    Desenhar_Formulario
    
    'Verificar tamnho da janela
    Tela_Cheia = False
    
    'Carregar icons na grelha
    Set Grelha.ImageList = McImageList1
    
    'Estado do progress bar
    Atual = 0
       
    'Som
    Mudo = False

    'wmp inicialmente em pause
    Musica_Play = False
    
    'Proproedades do player
    VideoDuration = 0
    Slide_Som.Left = 550
    Label_Volume.Caption = "10"
    Wmp.settings.volume = 50

    'Detectar placa de som
    Dim Placa As Long
    Placa = waveOutGetNumDevs()
    If Placa > 0 Then
       Label_Placa.Caption = "Placa de som detectada"
    Else
       Label_Placa.Caption = "Placa de som não detectada"
    End If



'    'PODER CARREGAR OS ARQUIVOS NA Lista_Directorios
'    hLB& = Lista_Directorios.hwnd
'    SendMessage hLB&, LB_INITSTORAGE, 30000&, ByVal 30000& * 200

    'Carregar valores guardados do sistema, caso existem
    On Error Resume Next 'GoTo Erro_Valores_Guardados
    Dim i, j As Integer
    'Carregar Grelha
    For i = 0 To GetSetting("NPlayer", "Valores_Gravados", "Conteudo_Grelha") - 1
        Grelha.AddItem GetSetting("NPlayer", "Valores_Gravados", "Conteudo_Grelha" & i), -1, 0
        'Form_Lista.Grelha.AddItem GetSetting("NPlayer", "Valores_Gravados", "Conteudo_Grelha" & i)
    Next

    'Carregar Lista_Directorios
    For j = 0 To GetSetting("NPlayer", "Valores_Gravados", "Conteudo_Lista_Directorios") - 1
        Lista_Directorios.AddItem GetSetting("NPlayer", "Valores_Gravados", "Conteudo_Lista_Directorios" & j)
    Next

    'Carregar url guardado
    Wmp.URL = GetSetting("NPlayer", "Valores_Gravados", "Conteudo_Wmp")
    Wmp.Controls.stop
    Grelha.ListIndex = GetSetting("NPlayer", "Valores_Gravados", "Conteudo_Linha")
    Lista_Directorios.ListIndex = Grelha.ListIndex
    'Form_Lista.Grelha.ListIndex = Grelha.ListIndex
    
    'Mensagem no icon do projecto/ coloca-lo ao lado do clock
    t.cbSize = Len(t)
    t.hwnd = pichook.hwnd
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = Me.Icon
    t.szTip = "NPlayer - Electric Nikyts" & Chr$(10) 'Texto a ser exibido no icon
    Shell_NotifyIcon NIM_ADD, t
    App.TaskVisible = False
        
    'Modo_Tray
    Modo_Tray = False
    
    'Grafico da barra top centro
    Numbars = 30
    BWidth = Imagem_Grafico.Width / Numbars
    For X = 0 To 99
        HArray(X) = 50
    Next X
    
    'Auto-play
    Tocar_Media
    
    'Criar popup menu personalizado
    SetCustomMenus
End Sub

Private Sub Form_Resize()
    Desenhar_Formulario
End Sub

Private Sub Grelha_DbClick()
    'Reproduzir o ficheiro selecionado
    'Caso se verifique de que se trata da ultima musica da lista
    If Grelha.ListIndex = Grelha.ListCount Then
        Botao_Stop_Click
        Grelha.ListIndex = 0
        Lista_Directorios.ListIndex = Grelha.ListIndex
        Exit Sub
    Else
        Tocar_Media
    End If
End Sub

Public Sub Tocar_Media()
    'Procedimento para reproduzir os ficheiros
    'On Error Resume Next
    Unload Form_PopUp
    
    'Verificar se a extensão do ficheiro é de video
    Lista_Directorios.ListIndex = Grelha.ListIndex
    
    Dim v As Variant, v2 As Variant, S As String, Extensao As String, Extensoes As String, ExtensaoCerta As Boolean, Arquivo As String
    Arquivo = Lista_Directorios.Text
    v = Split(Arquivo, ".")
    'Armazena a extensÃƒÂ£o atual do arquivo
    Extensao = v(UBound(v))
    'Extensões relacionadas a videos
    Extensoes = "avi wmv flv mpeg cam dvdrip mpg mp4 vob dvd xvid vcd"
    v2 = Split(Extensoes, " ")
    For i = 0 To UBound(v2)
    If Extensao = v2(i) Then ExtensaoCerta = True: Exit For 'verifica em cada extensÃƒÂ£o da array a extensÃƒÂ£o do arquivo selecionado.
    Next
    If ExtensaoCerta = True Then
        'Caso se verifique então mostra o form video
        Form_Wmp.Show
    End If
    
    'Reproduzir o som
    Label_Titulo.Caption = "NPlayer - Electric Nikyts" & " [" & Grelha.Text & "]"
    Label_Titulo_Mascara.Caption = "NPlayer - Electric Nikyts" & " [" & Grelha.Text & "]"
    'Label_Faixa_Mascara.Caption = Grelha.Text
    Slide.Left = 0
    VideoDuration = 0
    Faixa_em_Reproducao = Lista_Directorios.Text
    Wmp.URL = Faixa_em_Reproducao
    Form_Wmp.Wmp.URL = Faixa_em_Reproducao
    Timer_Slider_Video.Enabled = True
    
    Tempo_Estimado_Top.Caption = "00:00"
    Form_Wmp.Tempo_Estimado_Top.Caption = "00:00"
    Tempo_Estimado_Top_Mascara.Caption = "00:00"
    Wmp.Controls.stop
    Form_Wmp.Wmp.Controls.stop
    Wmp.URL = Lista_Directorios.Text
    Form_Wmp.Wmp.URL = Lista_Directorios.Text
    Timer_Duracao.Enabled = True
    Wmp.Controls.play
    Form_Wmp.Wmp.Controls.play
    Musica_Play = True
    Timer_Grafico.Enabled = True
    
    If Mudo = True Then Wmp.settings.mute = True Else Wmp.settings.mute = False
    
    'ver form popup
    If Modo_Tray = True Then
        Form_PopUp.lblSong.Caption = Grelha.Text
        Form_PopUp.Show
    End If
End Sub

Private Sub Grelha_KeyDown(KeyCode As Integer, Shift As Integer)
    'Teclas de atalho
    Lista_Directorios.ListIndex = Grelha.ListIndex
    If KeyCode = vbKeyDelete Then
        'Verificar se a lista contem ficheiros
        If Grelha.ListCount = 0 Then Exit Sub
        Dim Temp As String
        Temp = Lista_Directorios.Text
        Grelha.Remove Grelha.ListIndex
        Lista_Directorios.RemoveItem Lista_Directorios.ListIndex
        If Wmp.URL = Temp Then Wmp.Controls.stop: Timer_Duracao.Enabled = False
    End If
    
    'Tocar ou pausar a musica atraves do backspace
    If KeyCode = vbKeySpace Then
        If Timer_Duracao.Enabled = False Then
            Botao_Play_Click
        Else
            Botao_Pause_Click
        End If
    End If
    
    'Reproduzir som atraves do enter
    If KeyCode = vbKeyReturn Then
        Tocar_Media
    End If
    
    'Ver formulário sobre
    If KeyCode = vbKeyF1 Then
        Form_Sobre.Show vbModal
    End If
    
    'Atalho para aumentar volume
    If KeyCode = vbKeyF3 Then
        Aumentar_Volume_Click
    End If
    
    'Atalho para diminuri volume
    If KeyCode = vbKeyF2 Then
        Diminuir_Volume_Click
    End If
    
    'Tecla de atalho do som
    If KeyCode = vbKeyF4 Then
        Botao_Mudo_Click
    End If
End Sub

Private Sub Grelha_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ver popup menu
    If Button = 2 Then PopupMenu Menu
End Sub

Private Sub Icon_Fechar_Click()
    'Atalho para
    Botao_Fechar_Click
End Sub

Private Sub Icon_Parar_Click()
    'Atalho para
    Botao_Stop_Click
End Sub

Private Sub Icon_Pausa_Click()
    'Atalho para
    Botao_Pause_Click
End Sub

Private Sub Icon_Reproduzir_Click()
    'Atalho para
    Botao_Play_Click
End Sub

Private Sub Icon_Restaurar_Click()
    'Restaurar formulário
    Me.Show
    Modo_Tray = False
End Sub

Private Sub Label_Em_Reproducao_Click()
    'Ver a lista em reproducao
    Remover_Selecao
    Label_Em_Reproducao.BackColor = &HFC6C03
    Label_Em_Reproducao.ForeColor = vbWhite
    
    Grelha.Visible = True
    WebBrowser1.Visible = False
End Sub

Private Sub Remover_Selecao()
    'Procedimento para remover a selecao dos Nos da treeview
    Label_Biblioteca.BackColor = vbWhite
    Label_Biblioteca.ForeColor = vbBlack
    Label_Em_Reproducao.BackColor = vbWhite
    Label_Em_Reproducao.ForeColor = vbBlack
    Label_Pesquisa_Todos.BackColor = vbWhite
    Label_Pesquisa_Todos.ForeColor = vbBlack
    Label_Pesquisa_Musicas.BackColor = vbWhite
    Label_Pesquisa_Musicas.ForeColor = vbBlack
    Label_Pesquisa_Filmes.BackColor = vbWhite
    Label_Pesquisa_Filmes.ForeColor = vbBlack
    Label_Loja.BackColor = vbWhite
    Label_Loja.ForeColor = vbBlack
    Label_Loja_Nikyts.BackColor = vbWhite
    Label_Loja_Nikyts.ForeColor = vbBlack
    Label_Loja_Dilandau.BackColor = vbWhite
    Label_Loja_Dilandau.ForeColor = vbBlack
    Label_Listas.BackColor = vbWhite
    Label_Listas.ForeColor = vbBlack
    Label_Listas_Abrir.BackColor = vbWhite
    Label_Listas_Abrir.ForeColor = vbBlack
    Label_Navegador.BackColor = vbWhite
    Label_Navegador.ForeColor = vbBlack
    Label_Navegador_Abrir.BackColor = vbWhite
    Label_Navegador_Abrir.ForeColor = vbBlack
    Label_Navegador_Ver.BackColor = vbWhite
    Label_Navegador_Ver.ForeColor = vbBlack
    Label_Radio.BackColor = vbWhite
    Label_Radio.ForeColor = vbBlack
    Label_Radio_Iniciar.BackColor = vbWhite
    Label_Radio_Iniciar.ForeColor = vbBlack
    Label_Radio_Ver.BackColor = vbWhite
    Label_Radio_Ver.ForeColor = vbBlack
    Label_Tv.BackColor = vbWhite
    Label_Tv.ForeColor = vbBlack
    Label_Tv_Iniciar.BackColor = vbWhite
    Label_Tv_Iniciar.ForeColor = vbBlack
    Label_Tv_Ver.BackColor = vbWhite
    Label_Tv_Ver.ForeColor = vbBlack
    Label_Tv_Tuga.BackColor = vbWhite
    Label_Tv_Tuga.ForeColor = vbBlack
    Label_Utilitarios.BackColor = vbWhite
    Label_Utilitarios.ForeColor = vbBlack
    Label_Utilitarios_Tag.BackColor = vbWhite
    Label_Utilitarios_Tag.ForeColor = vbBlack
    Label_Utilitarios_Youtube.BackColor = vbWhite
    Label_Utilitarios_Youtube.ForeColor = vbBlack
    Label_Utilitarios_Gestor.BackColor = vbWhite
    Label_Utilitarios_Gestor.ForeColor = vbBlack
    Label_Utilitarios_Agenda.BackColor = vbWhite
    Label_Utilitarios_Agenda.ForeColor = vbBlack
End Sub

Private Sub Label_Faixa_Mascara_Click()
    'Atalho para
    Botao_Maximizar_Mascara_Click
End Sub

Private Sub Label_Listas_Abrir_Click()
    'Abrir playlists
    Remover_Selecao
    Label_Listas_Abrir.BackColor = &HFC6C03
    Label_Listas_Abrir.ForeColor = vbWhite
    
    Form_Listas_Guardadas.Show vbModal
End Sub

Private Sub Label_Listas_Click()
    'Selecionar item
    Remover_Selecao
    Label_Listas.BackColor = &HFC6C03
    Label_Listas.ForeColor = vbWhite
End Sub

Private Sub Label_Listas_DblClick()
    'atalho para
    No_Listas_Click
End Sub

Private Sub label_loja_click()
    'Selecionar item
    Remover_Selecao
    Label_Loja.BackColor = &HFC6C03
    Label_Loja.ForeColor = vbWhite
End Sub

Private Sub Label_Loja_DblClick()
    'Abrir/ fechar a pasta "Loja de música"
    No_Loja_Click
End Sub

Private Sub Label_Loja_Dilandau_Click()
    'Abrir site "Dilandau"
    Remover_Selecao
    Label_Loja_Dilandau.BackColor = &HFC6C03
    Label_Loja_Dilandau.ForeColor = vbWhite
    
    With WebBrowser1
        .Navigate "www.dilandau.com"
        .Visible = True
    End With
    
    'Oculta a grelha
    Grelha.Visible = False
End Sub

Private Sub Label_Loja_Nikyts_Click()
    'Abrir site "Nikyts"
    Remover_Selecao
    Label_Loja_Nikyts.BackColor = &HFC6C03
    Label_Loja_Nikyts.ForeColor = vbWhite
    
    With WebBrowser1
        .Navigate "www.nikyts.no.sapo.pt"
        .Visible = True
    End With
    
    'Oculta a grelha
    Grelha.Visible = False
End Sub

Private Sub Label_Mudo_Click()
    'Atalho para
    Botao_Mudo_Click
End Sub

Private Sub Label_Navegador_Abrir_Click()
    'Iniciar navegador
    Remover_Selecao
    Label_Navegador_Abrir.BackColor = &HFC6C03
    Label_Navegador_Abrir.ForeColor = vbWhite
    
    Form_Browser.Show
End Sub

Private Sub Label_Navegador_Click()
    'Selecionar item
    Remover_Selecao
    Label_Navegador.BackColor = &HFC6C03
    Label_Navegador.ForeColor = vbWhite
End Sub

Private Sub Label_Navegador_DblClick()
    'Abrir/ fechar a pasta "Coleção de listas"
    No_Navegador_Click
End Sub

Private Sub Label_Navegador_Ver_Click()
    'Ver o formulário dos favoritos
    Remover_Selecao
    Label_Navegador_Ver.BackColor = &HFC6C03
    Label_Navegador_Ver.ForeColor = vbWhite
    
    With Form_Browser_Opcoes
        .Show vbModal
        .Lista_Favoritos.ListIndex = 2
        .Frame_Favoritos.Visible = True
        .Frame_Geral.Visible = False
        .Frame_Propriedades.Visible = False
    End With
End Sub

Private Sub Label_Biblioteca_Click()
    'Selecionar item
    Remover_Selecao
    Label_Biblioteca.BackColor = &HFC6C03
    Label_Biblioteca.ForeColor = vbWhite
End Sub

Private Sub Label_Biblioteca_DblClick()
    'Abrir/ fechar a pasta "Pesquisa"
    No_Pesquisa_Click
End Sub

Private Sub Label_Pesquisa_Filmes_Click()
    'Efectuar pesquisa
    Remover_Selecao
    Label_Pesquisa_Filmes.BackColor = &HFC6C03
    Label_Pesquisa_Filmes.ForeColor = vbWhite
    
    Form_Pesquisa.Show vbModal
End Sub

Private Sub Label_Pesquisa_Musicas_Click()
    'Efectuar pesquisa
    Remover_Selecao
    Label_Pesquisa_Musicas.BackColor = &HFC6C03
    Label_Pesquisa_Musicas.ForeColor = vbWhite
    
    Form_Pesquisa.Show vbModal
End Sub

Private Sub Label_Pesquisa_Todos_Click()
    'Efectuar pesquisa de ficheiros
    Remover_Selecao
    Label_Pesquisa_Todos.BackColor = &HFC6C03
    Label_Pesquisa_Todos.ForeColor = vbWhite
    
    Form_Pesquisa_Ficheiros.Show 'vbModal
End Sub

Private Sub Label_Propriedades_Click()
    'Ver palete de cores para escolher o skin pretendido
    Form_Propriedades.Show vbModal
End Sub

Private Sub Label_Radio_Click()
    'Selecionar item
    Remover_Selecao
    Label_Radio.BackColor = &HFC6C03
    Label_Radio.ForeColor = vbWhite
End Sub

Private Sub Label_Radio_DblClick()
    'Atalho para
    No_Radio_Click
End Sub

Private Sub Label_Titulo_Mascara_DblClick()
    Botao_Maximizar_Mascara_Click
End Sub

Private Sub Label_Titulo_Mascara_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    H = X
    v = Y
End Sub

Private Sub Label_Titulo_Mascara_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.WindowState = 0 Then
        If Button = vbLeftButton Then
           Me.Left = Me.Left - (H - X)
           Me.Top = Me.Top - (v - Y)
        End If
    End If
End Sub

Private Sub Label_Titulo_DblClick()
    Botao_Maximizar_Click
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    H = X
    v = Y
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.WindowState = 0 Then
        If Button = vbLeftButton Then
           Me.Left = Me.Left - (H - X)
           Me.Top = Me.Top - (v - Y)
        End If
    End If
End Sub

Private Sub Label_Tv_Click()
    'Selecionar item
    Remover_Selecao
    Label_Tv.BackColor = &HFC6C03
    Label_Tv.ForeColor = vbWhite
End Sub

Private Sub Label_Tv_DblClick()
    'Atalho para
    No_Tv_Click
End Sub

Private Sub Botao_Lista_Click()
    'Ver formulário da lista
    Load Form_Lista
    With Form_Lista
        .List1.Clear
        Dim i As Integer
        For i = 0 To Grelha.ListCount - 1
            .List1.AddItem Grelha.List(i), -1, 0
            .List1.ListIndex = Grelha.ListIndex
        Next i
        .Show
    End With
End Sub

Private Sub No_Listas_Click()
    'Abrir/ fechar a pasta "Coleção de listas"
    If Frame_Listas.Height = 225 Then
        Frame_Listas.Height = 705
        No_Listas.Picture = Form_Imagens.No_Menos.Picture
    Else
        Frame_Listas.Height = 225
        No_Listas.Picture = Form_Imagens.No_Mais.Picture
    End If
    
    'Posicionar restantes frames
    Posicionar_Nos
End Sub

Private Sub No_Loja_Click()
    'Abrir/ fechar a pasta "Loja de música"
    If Frame_Loja.Height = 225 Then
        Frame_Loja.Height = 975
        No_Loja.Picture = Form_Imagens.No_Menos.Picture
    Else
        Frame_Loja.Height = 225
        No_Loja.Picture = Form_Imagens.No_Mais.Picture
    End If
    
    'Posicionar restantes frames
    Posicionar_Nos
End Sub

Private Sub No_Navegador_Click()
    'Abrir/ fechar a pasta "Loja de música"
    If Frame_Navegador.Height = 225 Then
        Frame_Navegador.Height = 975
        No_Navegador.Picture = Form_Imagens.No_Menos.Picture
    Else
        Frame_Navegador.Height = 225
        No_Navegador.Picture = Form_Imagens.No_Mais.Picture
    End If
    
    'Posicionar restantes frames
    Posicionar_Nos
End Sub

Private Sub No_Pesquisa_Click()
    'Abrir/ fechar a pasta "Pesquisa"
    If Frame_Pesquisa.Height = 225 Then
        Frame_Pesquisa.Height = 1455
        No_Pesquisa.Picture = Form_Imagens.No_Menos.Picture
    Else
        Frame_Pesquisa.Height = 225
        No_Pesquisa.Picture = Form_Imagens.No_Mais.Picture
    End If
    
    'Posicionar restantes frames
    Posicionar_Nos
End Sub

Private Sub No_Radio_Click()
    'Abrir/ fechar a pasta "Radio"
    If Frame_Radio.Height = 225 Then
        Frame_Radio.Height = 975
        No_Radio.Picture = Form_Imagens.No_Menos.Picture
    Else
        Frame_Radio.Height = 225
        No_Radio.Picture = Form_Imagens.No_Mais.Picture
    End If
    
    'Posicionar restantes frames
    Posicionar_Nos
End Sub

Private Sub No_Tv_Click()
    'Abrir/ fechar a pasta "Televisão"
    If Frame_Tv.Height = 225 Then
        Frame_Tv.Height = 1185
        No_Tv.Picture = Form_Imagens.No_Menos.Picture
    Else
        Frame_Tv.Height = 225
        No_Tv.Picture = Form_Imagens.No_Mais.Picture
    End If
    
    'Posicionar restantes frames
    Posicionar_Nos
End Sub

Private Sub Setas_Selecao_Click()
    'Ver/ ocultar frames
    If Barra_Biblioteca.Visible = True Then
        Botao_Uma_Frame_Click
    Else
        Botao_Duas_Frames_Click
    End If
End Sub

Private Sub Skin_Top_Centro_DblClick()
    Botao_Maximizar_Click
End Sub

Private Sub Skin_Top_Centro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub Skin_Top_Player_DblClick()
    'Maximixar/ Restaurar Formulários
    Botao_Maximizar_Click
End Sub

Private Sub Skin_Top_Player_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    H = X
    v = Y
End Sub

Private Sub Skin_Top_Player_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.WindowState = 0 Then
        If Button = vbLeftButton Then
           Me.Left = Me.Left - (H - X)
           Me.Top = Me.Top - (v - Y)
        End If
    End If
    Label_Propriedades.ForeColor = &HFF890A
End Sub

Private Sub Tela_Em_Cheio_Click()
    'Atalho para
    Botao_Uma_Frame_Click
End Sub

Private Sub Timer_Actualiza_Timer()
    If Grelha.ListCount = 0 Then
        Exit Sub
    Else
        Label_Faixa_Actual.Caption = Grelha.ListIndex + 1 & " de " & Grelha.ListCount
        Label_Faixa_Actual_Mascara.Caption = Grelha.ListIndex + 1 & " de " & Grelha.ListCount
        Form_Wmp.Label_Faixa_Actual.Caption = Grelha.ListIndex + 1 & " de " & Grelha.ListCount
    End If
End Sub

Private Sub Timer_Duracao_Timer()
    'On Error Resume Next
    'Dim Duration
    Tempo_Estimado_Top.Caption = Duration(Wmp.Controls.CurrentPosition)
    Form_Wmp.Tempo_Estimado_Top.Caption = Duration(Wmp.Controls.CurrentPosition)
    Tempo_Estimado_Top_Mascara.Caption = Duration(Wmp.Controls.CurrentPosition)
    'Tempo_Estimado.Caption = "Tempo estimado: " & Duration(Wmp.Controls.CurrentPosition)
    'Tempo_Estimado_Mascara.Caption = "Tempo estimado: " & Duration(Wmp.Controls.CurrentPosition)
    Wmp.Controls.play
    Form_Wmp.Wmp.Controls.play
    Wmp_PositionChange 0, 1
    If VideoDuration >= 1 Then
        'Wmp.Controls.stop
        Wmp.Controls.play
        Form_Wmp.Wmp.Controls.play
    End If
End Sub

Private Sub Wmp_PositionChange(ByVal oldPosition As Double, ByVal newPosition As Double)
    On Error Resume Next
    Dim ComputeDuration
    VideoDuration = ComputeDuration(Trim(Wmp.Controls.currentItem.durationString))
End Sub

Private Sub Timer_Slider_Video_Timer()
    On Error Resume Next
    Dim tm As Integer, tt As Integer, tp As Single, offset As Integer
    Dim tm_2 As Integer, tt_2 As Integer, tp_2 As Single, offset_2 As Integer
    Dim tm_3 As Integer, tt_3 As Integer, tp_3 As Single, offset_3 As Integer
    
    tm = Int(Wmp.Controls.CurrentPosition)
    tt = Int(Wmp.currentMedia.Duration)
    tm_2 = Int(Wmp.Controls.CurrentPosition)
    tt_2 = Int(Wmp.currentMedia.Duration)
    tm_3 = Int(Wmp.Controls.CurrentPosition)
    tt_3 = Int(Wmp.currentMedia.Duration)
    
    If tm <> -1 Then
        tp = tm / tt
        tp_2 = tm_2 / tt_2
        tp_3 = tm_3 / tt_3
        
        offset = Int((Image_Barra_Slide.Width - 5 - Slide.Width) * tp)
        offset_2 = Int((Image_Barra_Slide_Mascara.Width - 5 - Slide_Mascara.Width) * tp_2)
        offset_3 = Int((Form_Wmp.Image_Barra_Slide_Mascara.Width - 5 - Form_Wmp.Slide_Mascara.Width) * tp_3)
        
        If Not DNa Then
            Slide.Left = offset + Image_Barra_Slide.Left + 3
            Slide_Mascara.Left = offset_2 + Image_Barra_Slide_Mascara.Left + 3
            Form_Wmp.Slide_Mascara.Left = offset_3 + Form_Wmp.Image_Barra_Slide_Mascara.Left + 3
        End If
        If Slide.Left >= 1720 Or Slide_Mascara.Left >= 4270 Then
            Botao_Seguinte_Click
        End If
    Else
    End If
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
    On Error Resume Next
    Dim offseti As Single
    DNa_2 = False
    offseti = (Slide_Mascara.Left - Image_Barra_Slide_Mascara.Left - 3) / (Image_Barra_Slide_Mascara.Width - 10 - Slide_Mascara.Width)
    Wmp.Controls.CurrentPosition = Int(Wmp.currentMedia.Duration * offseti)
    Form_Wmp.Wmp.Controls.CurrentPosition = Int(Form_Wmp.Wmp.currentMedia.Duration * offseti)
End Sub

Private Sub Slide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DNa = True
    Txa = X
End Sub

Private Sub Slide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DNa Then
        NewLeft = Slide.Left + X - Txa
        If NewLeft < Image_Barra_Slide.Left + 3 Then
            NewLeft = Image_Barra_Slide.Left + 3
        End If
        If NewLeft > Image_Barra_Slide.Width + Image_Barra_Slide.Left - 7 - Slide.Width Then
            NewLeft = Image_Barra_Slide.Width + Image_Barra_Slide.Left - 7 - Slide.Width
        End If
        Slide.Left = NewLeft
    End If
End Sub

Private Sub Slide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim offseti As Single
    DNa = False
    offseti = (Slide.Left - Image_Barra_Slide.Left - 3) / (Image_Barra_Slide.Width - 10 - Slide.Width)
    Wmp.Controls.CurrentPosition = Int(Wmp.currentMedia.Duration * offseti)
    Form_Wmp.Wmp.Controls.CurrentPosition = Int(Form_Wmp.Wmp.currentMedia.Duration * offseti)
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
        Form_Wmp.Slide_Som.Left = NewLeft_Som
    End If
End Sub

Private Sub Slide_Som_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim offseti As Single
    DNa_Som = False
    'offseti = (Slide_Som.Left - Image_Barra_Slide.Left - 3) / (Image_Barra_Slide.Width - 10 - Slide_Som.Width)

    'Verificar a posiçãp do slider do volume
    If Slide_Som.Left >= 0 And Slide_Som.Left <= 100 Then
        Label_Volume.Caption = "1"
    ElseIf Slide_Som.Left > 100 And Slide_Som.Left <= 150 Then
        Label_Volume.Caption = "2"
    ElseIf Slide_Som.Left > 150 And Slide_Som.Left <= 200 Then
        Label_Volume.Caption = "3"
    ElseIf Slide_Som.Left > 200 And Slide_Som.Left <= 250 Then
        Label_Volume.Caption = "4"
    ElseIf Slide_Som.Left > 250 And Slide_Som.Left <= 300 Then
        Label_Volume.Caption = "5"
    ElseIf Slide_Som.Left > 300 And Slide_Som.Left <= 350 Then
        Label_Volume.Caption = "6"
    ElseIf Slide_Som.Left > 350 And Slide_Som.Left <= 400 Then
        Label_Volume.Caption = "7"
    ElseIf Slide_Som.Left > 400 And Slide_Som.Left <= 450 Then
        Label_Volume.Caption = "8"
    ElseIf Slide_Som.Left > 450 And Slide_Som.Left <= 500 Then
        Label_Volume.Caption = "9"
    ElseIf Slide_Som.Left > 500 And Slide_Som.Left <= 550 Then
        Label_Volume.Caption = "10"
    ElseIf Slide_Som.Left > 550 And Slide_Som.Left <= 600 Then
        Label_Volume.Caption = "11"
    ElseIf Slide_Som.Left > 600 And Slide_Som.Left <= 650 Then
        Label_Volume.Caption = "12"
    ElseIf Slide_Som.Left > 650 And Slide_Som.Left <= 700 Then
        Label_Volume.Caption = "13"
    ElseIf Slide_Som.Left > 700 And Slide_Som.Left <= 750 Then
        Label_Volume.Caption = "14"
    ElseIf Slide_Som.Left > 750 And Slide_Som.Left <= 800 Then
        Label_Volume.Caption = "15"
    ElseIf Slide_Som.Left > 800 And Slide_Som.Left <= 850 Then
        Label_Volume.Caption = "16"
    ElseIf Slide_Som.Left > 850 And Slide_Som.Left <= 900 Then
        Label_Volume.Caption = "17"
    ElseIf Slide_Som.Left > 900 And Slide_Som.Left <= 960 Then
        Label_Volume.Caption = "18"
    ElseIf Slide_Som.Left > 960 And Slide_Som.Left <= 1040 Then
        Label_Volume.Caption = "19"
    ElseIf Slide_Som.Left > 1040 And Slide_Som.Left <= 1110 Then
        Label_Volume.Caption = "20"
    End If

    Verificar_Volume
End Sub

Public Sub Verificar_Volume()
    'Procedimento para verificar o estado do volume/ slider do player
    If Label_Volume.Caption = "1" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_1.Picture
        Wmp.settings.volume = 5
        'Form_Wmp.Wmp.settings.volume = 5
        Slide_Som.Left = 0
        Form_Wmp.Slide_Som.Left = 0
        
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "2" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_2.Picture
        Wmp.settings.volume = 10
        'Form_Wmp.Wmp.settings.volume = 10
        Slide_Som.Left = 100
        Form_Wmp.Slide_Som.Left = 100
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "3" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_3.Picture
        Wmp.settings.volume = 15
        'Form_Wmp.Wmp.settings.volume = 15
        Slide_Som.Left = 200
        Form_Wmp.Slide_Som.Left = 200
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "4" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_4.Picture
        Wmp.settings.volume = 20
        'Form_Wmp.Wmp.settings.volume = 20
        Slide_Som.Left = 250
        Form_Wmp.Slide_Som.Left = 250
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "5" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_5.Picture
        Wmp.settings.volume = 25
        'Form_Wmp.Wmp.settings.volume = 25
        Slide_Som.Left = 300
        Form_Wmp.Slide_Som.Left = 300
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "6" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_6.Picture
        Wmp.settings.volume = 30
        'Form_Wmp.Wmp.settings.volume = 30
        Slide_Som.Left = 350
        Form_Wmp.Slide_Som.Left = 350
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "7" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_7.Picture
        Wmp.settings.volume = 35
        'Form_Wmp.Wmp.settings.volume = 35
        Slide_Som.Left = 400
        Form_Wmp.Slide_Som.Left = 400
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "8" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_8.Picture
        Wmp.settings.volume = 40
        'Form_Wmp.Wmp.settings.volume = 40
        Slide_Som.Left = 450
        Form_Wmp.Slide_Som.Left = 450
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "9" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_9.Picture
        Wmp.settings.volume = 45
        'Form_Wmp.Wmp.settings.volume = 45
        Slide_Som.Left = 500
        Form_Wmp.Slide_Som.Left = 500
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "10" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_10.Picture
        Wmp.settings.volume = 50
        'Form_Wmp.Wmp.settings.volume = 50
        Slide_Som.Left = 550
        Form_Wmp.Slide_Som.Left = 550
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "11" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_11.Picture
        Wmp.settings.volume = 55
        'Form_Wmp.Wmp.settings.volume = 55
        Slide_Som.Left = 600
        Form_Wmp.Slide_Som.Left = 600
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "12" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_12.Picture
        Wmp.settings.volume = 60
        'Form_Wmp.Wmp.settings.volume = 60
        Slide_Som.Left = 650
        Form_Wmp.Slide_Som.Left = 650
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "13" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_13.Picture
        Wmp.settings.volume = 65
        'Form_Wmp.Wmp.settings.volume = 65
        Slide_Som.Left = 700
        Form_Wmp.Slide_Som.Left = 700
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "14" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_14.Picture
        Wmp.settings.volume = 70
        'Form_Wmp.Wmp.settings.volume = 70
        Slide_Som.Left = 750
        Form_Wmp.Slide_Som.Left = 750
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "15" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_15.Picture
        Wmp.settings.volume = 75
        'Form_Wmp.Wmp.settings.volume = 75
        Slide_Som.Left = 800
        Form_Wmp.Slide_Som.Left = 800
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "16" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_16.Picture
        Wmp.settings.volume = 80
        'Form_Wmp.Wmp.settings.volume = 80
        Slide_Som.Left = 850
        Form_Wmp.Slide_Som.Left = 850
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "17" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_17.Picture
        Wmp.settings.volume = 85
        'Form_Wmp.Wmp.settings.volume = 85
        Slide_Som.Left = 900
        Form_Wmp.Slide_Som.Left = 900
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "18" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_18.Picture
        Wmp.settings.volume = 90
        'Form_Wmp.Wmp.settings.volume = 90
        Slide_Som.Left = 960
        Form_Wmp.Slide_Som.Left = 960
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "19" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_19.Picture
        Wmp.settings.volume = 95
        'Form_Wmp.Wmp.settings.volume = 95
        Slide_Som.Left = 1040
        Form_Wmp.Slide_Som.Left = 1040
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    ElseIf Label_Volume.Caption = "20" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_20.Picture
        Wmp.settings.volume = 100
        'Form_Wmp.Wmp.settings.volume = 100
        Slide_Som.Left = 1110
        Form_Wmp.Slide_Som.Left = 1110
        If Mudo = True Then Wmp.settings.mute = True: 'Form_Wmp.Wmp.settings.mute = True
    End If

    'Percentagem da wmp
    If Label_Volume.Caption = "1" Then
        Label_Percentagem_Volume.Caption = "Volume: 1%"
        Label_Percentagem_Volume_2.Caption = "Volume: 1%"
    ElseIf Label_Volume.Caption = "2" Then
        Label_Percentagem_Volume.Caption = "Volume: 5%"
        Label_Percentagem_Volume_2.Caption = "Volume: 5%"
    ElseIf Label_Volume.Caption = "3" Then
        Label_Percentagem_Volume.Caption = "Volume: 10%"
        Label_Percentagem_Volume_2.Caption = "Volume: 10%"
    ElseIf Label_Volume.Caption = "4" Then
        Label_Percentagem_Volume.Caption = "Volume: 15%"
        Label_Percentagem_Volume_2.Caption = "Volume: 15%"
    ElseIf Label_Volume.Caption = "5" Then
        Label_Percentagem_Volume.Caption = "Volume: 20%"
        Label_Percentagem_Volume_2.Caption = "Volume: 20%"
    ElseIf Label_Volume.Caption = "6" Then
        Label_Percentagem_Volume.Caption = "Volume: 25%"
        Label_Percentagem_Volume_2.Caption = "Volume: 25%"
    ElseIf Label_Volume.Caption = "7" Then
        Label_Percentagem_Volume.Caption = "Volume: 30%"
        Label_Percentagem_Volume_2.Caption = "Volume: 30%"
    ElseIf Label_Volume.Caption = "8" Then
        Label_Percentagem_Volume.Caption = "Volume: 35"
        Label_Percentagem_Volume_2.Caption = "Volume: 35"
    ElseIf Label_Volume.Caption = "9" Then
        Label_Percentagem_Volume.Caption = "Volume: 45%"
        Label_Percentagem_Volume_2.Caption = "Volume: 45%"
    ElseIf Label_Volume.Caption = "10" Then
        Label_Percentagem_Volume.Caption = "Volume: 50%"
        Label_Percentagem_Volume_2.Caption = "Volume: 50%"
    ElseIf Label_Volume.Caption = "11" Then
        Label_Percentagem_Volume.Caption = "Volume: 55%"
        Label_Percentagem_Volume_2.Caption = "Volume: 55%"
    ElseIf Label_Volume.Caption = "12" Then
        Label_Percentagem_Volume.Caption = "Volume: 60%"
        Label_Percentagem_Volume_2.Caption = "Volume: 60%"
    ElseIf Label_Volume.Caption = "13" Then
        Label_Percentagem_Volume.Caption = "Volume: 65%"
        Label_Percentagem_Volume_2.Caption = "Volume: 65%"
    ElseIf Label_Volume.Caption = "14" Then
        Label_Percentagem_Volume.Caption = "Volume: 70%"
        Label_Percentagem_Volume_2.Caption = "Volume: 70%"
    ElseIf Label_Volume.Caption = "15" Then
        Label_Percentagem_Volume.Caption = "Volume: 75%"
        Label_Percentagem_Volume_2.Caption = "Volume: 75%"
    ElseIf Label_Volume.Caption = "16" Then
        Label_Percentagem_Volume.Caption = "Volume: 80%"
        Label_Percentagem_Volume_2.Caption = "Volume: 80%"
    ElseIf Label_Volume.Caption = "17" Then
        Label_Percentagem_Volume.Caption = "Volume: 85%"
        Label_Percentagem_Volume_2.Caption = "Volume: 85%"
    ElseIf Label_Volume.Caption = "18" Then
        Label_Percentagem_Volume.Caption = "Volume: 90%"
        Label_Percentagem_Volume_2.Caption = "Volume: 90%"
    ElseIf Label_Volume.Caption = "19" Then
        Label_Percentagem_Volume.Caption = "Volume: 95%"
        Label_Percentagem_Volume_2.Caption = "Volume: 95%"
    ElseIf Label_Volume.Caption = "20" Then
        Label_Percentagem_Volume.Caption = "Volume: 100%"
        Label_Percentagem_Volume_2.Caption = "Volume: 100%"
    End If
End Sub

Public Sub Botao_Pause_Click()
    Wmp.Controls.pause
    Form_Wmp.Wmp.Controls.pause
    Timer_Duracao.Enabled = False
    Timer_Grafico.Enabled = False
End Sub

Public Sub Botao_Play_Click()
    'If Faixa_em_Reproducao = "" Then
    '    MsgBox ("Selecione o ficheiro que pretende reproduzir")
    'Else
        'Wmp.URL = Faixa_em_Reproducao
        Wmp.Controls.play
        Form_Wmp.Wmp.Controls.play
        Timer_Duracao.Enabled = True
        Timer_Grafico.Enabled = True
    'End If
End Sub

Public Sub Botao_Primeiro_Click()
    Grelha.ListIndex = 0
    Lista_Directorios.ListIndex = Grelha.ListIndex
    'Form_Lista.List1.ListIndex = 0
    Tocar_Media
    Timer_Duracao.Enabled = True
End Sub

Public Sub Botao_Seguinte_Click()
    If Grelha.ListIndex = Grelha.ListCount - 1 Then Exit Sub
    Grelha.ListIndex = Grelha.ListIndex + 1
    Tocar_Media
    Timer_Duracao.Enabled = True
End Sub

Public Sub Botao_Stop_Click()
    Wmp.Controls.stop
    Form_Wmp.Wmp.Controls.stop
    Faixa_em_Reproducao = ""
    Timer_Duracao.Enabled = False
    Tempo_Estimado_Top.Caption = "00:00"
    Form_Wmp.Tempo_Estimado_Top.Caption = "00:00"
    Tempo_Estimado_Top_Mascara.Caption = "00:00"
    Timer_Grafico.Enabled = False
End Sub

Public Sub Botao_Ultimo_Click()
    'Ir para as ultima linha
    Dim contador As Long
    contador = Grelha.ListCount
    Grelha.ListIndex = contador - 1
    Lista_Directorios.ListIndex = Grelha.ListIndex
    'Form_Lista.List1.ListIndex = contador - 1
    Tocar_Media
    Timer_Duracao.Enabled = True
End Sub

Private Sub Diminuir_Volume_Click()
    'Diminuir volume do player
    If Label_Volume.Caption = "1" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_1.Picture
        Exit Sub
    End If
    Label_Volume = Label_Volume - 1
    Verificar_Volume
End Sub

Private Sub Aumentar_Volume_Click()
    'Diminuir volume do player
    If Label_Volume.Caption = "20" Then
        Image_Volume.Picture = Form_Imagens.Image_Volume_20.Picture
        Exit Sub
    End If
    Label_Volume = Label_Volume + 1
    Verificar_Volume
End Sub

Public Sub Botao_Antes_Click()
    If Grelha.ListIndex <= 0 Then Exit Sub
    Grelha.ListIndex = Grelha.ListIndex - 1
    Tocar_Media
    Timer_Duracao.Enabled = True
End Sub

'Public Sub Pesquisar_Ficheiros()
'    'Efecuar pesquisa automática consoante o formato do ficheiro escolhido
'    Dim Arquivos() As String
'    Dim F As Integer
'    'Botao_Stop_Click
'    Lista_Directorios.Clear
'    'Grelha.Clear
''    Form_Carregar.Show vbModal
'    Arquivos = Split("*.avi") 'Split(Tipo_De_Ficheiro_Escolhido)
'    'Arquivos = Split("*.mp3;*.wav;*.avi;*.wmv;*.mpg;*.mpeg", ";")
'
'    For F = 0 To UBound(Arquivos)
'        If Running% Then: Running% = False: Exit Sub
'        Dim drvbitmask&, maxpwr%, pwr%
'        '5On Error Resume Next
'        FileSpec$ = Arquivos(F)
'        If Len(FileSpec$) = 0 Then Exit Sub
'        Running% = True
'        UseFileSpec% = True
'        drvbitmask& = GetLogicalDrives()
'        If drvbitmask& Then
'            maxpwr% = Int(Log(drvbitmask&) / Log(2))
'            For pwr% = 0 To maxpwr%
'                If Running% And (2 ^ pwr% And drvbitmask&) Then Call SearchDirs(Chr$(vbKeyA + pwr%) & ":\")
'            Next
'       End If
'
'        Running% = False
'        UseFileSpec% = False
'        Label12.Caption = "Find File(s): " & Lista_Directorios.ListCount & " items found matching " & """" & FileSpec$ & """"
'        Beep
'    Next F
'    Botao_Stop_Click
'    Lista_Directorios.Clear
'    Grelha.Clear
'    Form_Carregar.Show vbModal
'    Arquivos = Split(Tipo_De_Ficheiro_Escolhido)
'    'Arquivos = Split("*.mp3;*.wav;*.avi;*.wmv;*.mpg;*.mpeg", ";")
'
'    For F = 0 To UBound(Arquivos)
'        If Running% Then: Running% = False: Exit Sub
'        Dim drvbitmask&, maxpwr%, pwr%
'
'        On Error Resume Next
'        FileSpec$ = Arquivos(F)
'        If Len(FileSpec$) = 0 Then Exit Sub
'        Running% = True
'        UseFileSpec% = True
'        drvbitmask& = GetLogicalDrives()
'
'        If drvbitmask& Then
'            maxpwr% = Int(Log(drvbitmask&) / Log(2))
'            For pwr% = 0 To maxpwr%
'                If Running% And (2 ^ pwr% And drvbitmask&) Then Call SearchDirs(Chr$(vbKeyA + pwr%) & ":\")
'            Next
'        End If
'
'        'Carregar as listas com os ficheiros encontrados
'        Dim Musica() As String
'        Dim Linha As Integer
'        'Contar as linhas da Lista_Directorios para depois remover as //
'        For Linha = 0 To Lista_Directorios.ListCount - 1
'            Musica = Split(Lista_Directorios.List(Linha), "\")
'            Grelha.AddItem Musica(UBound(Musica)), -1, 0 'e depois adicona na Grelha
'        Next Linha
'
'        'Repõe os objectos
'        Running% = False
'        UseFileSpec% = False
'        'Label_Titulo.Caption = "Procurando: " & Lista_Directorios.ListCount & " Ficheiros encontrados " & """" & FileSpec$ & """"
'        Unload Form_Carregar
'        Botao_Primeiro_Click
'    Next F
'End Sub

Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'pichook é uma picture box, utilizada pelo Windows para
'reconhecer o ícone na barra de tarefas.
    Static rec As Boolean, msg As Long
    msg = X / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
        Select Case msg
            Case WM_LBUTTONDBLCLK:
                Me.Show
                Modo_Tray = False
            Case WM_LBUTTONDOWN:
            Case WM_LBUTTONUP:
            Case WM_RBUTTONDBLCLK:
            Case WM_RBUTTONDOWN:
            Case WM_RBUTTONUP:
                'Ver o menu icon se for pressionado o botão direito
                Me.PopupMenu Menu_Icon
        End Select
        rec = False
    End If
End Sub

'Private Sub SearchDirs(curpath$)  ' curpath$ is passed w/ trailing "\"
'    'Procedimento para pesquisar ficheiros automáticamente
'    Dim dirs%, dirbuf$(), i%
'    Form_Carregar.Label1.Caption = "Procurando..." & curpath$
'    DoEvents
'    If Not Running% Then Exit Sub
'    hItem& = FindFirstFile(curpath$ & vbAllFiles, WFD)
'    If hItem& <> INVALID_HANDLE_VALUE Then
'        Do
'            If (WFD.dwFileAttributes And vbDirectory) Then
'
'                If Asc(WFD.cFileName) <> vbKeyDot Then
'                    TotalDirs% = TotalDirs% + 1
'                    If (dirs% Mod 10) = 0 Then ReDim Preserve dirbuf$(dirs% + 10)
'                    dirs% = dirs% + 1
'                    dirbuf$(dirs%) = Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
'                End If
'            ElseIf Not UseFileSpec% Then
'                TotalFiles% = TotalFiles% + 1
'            End If
'        Loop While FindNextFile(hItem&, WFD)
'        Call FindClose(hItem&)
'    End If
'    If UseFileSpec% Then
'        SendMessage hLB&, WM_SETREDRAW, 0, 0
'        Call SearchFileSpec(curpath$)
'        SendMessage hLB&, WM_VSCROLL, SB_BOTTOM, 0
'        SendMessage hLB&, WM_SETREDRAW, 1, 0
'    End If
'    For i% = 1 To dirs%: SearchDirs curpath$ & dirbuf$(i%) & vbBackslash: Next i%
'End Sub
'
'Private Sub SearchFileSpec(curpath$)   ' curpath$ is passed w/ trailing "\"
'    hFile& = FindFirstFile(curpath$ & FileSpec$, WFD)
'    If hFile& <> INVALID_HANDLE_VALUE Then
'        Do
'            DoEvents
'            If Not Running% Then Exit Sub
'            SendMessage hLB&, LB_ADDSTRING, 0, ByVal curpath$ & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
'            Form_Carregar.Label1.Caption = curpath$ & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
'        Loop While FindNextFile(hFile&, WFD)
'        Call FindClose(hFile&)
'    End If
'End Sub

Private Sub SetCustomMenus()
    'Procedimento para personalizar o popupmenu
    startODMenus Me, True
    With CustomMenu
        .Texture = True
        Set .Picture = Imagem_Fundo_Menu.Picture
        .UseCustomFonts = False
         .PosX = 20
    End With
    With CustomColor
        .ForeColor = vbBlack
        .DefTextColor = vbBlack ' vbBlack
        .HilightColor = &HFF890A
        .NormalColor = vbWhite
        '.BackColor = vbRed
        .SelectedTextColor = vbWhite
        .MenuTextColor = vbBlack
        '.BorderColor = RGB(240, 72, 72)
        '.RECTColor = vbGreen
    End With
  End Sub

Public Sub Desenhar_Formulario()
    'Ajustar os objectos contidos no formuário
    If Me.WindowState = 1 Then Exit Sub
    If Modo_Mascara = True Then Exit Sub
    
    '*********************************************************************** Frame Top ***********************************************************************
    'Frame_Top
    With Frame_Top
        .Top = 0
        .Width = Me.Width
        .Left = 0
    End With
    
    'Skin_Top_Esquerda
    With Skin_Top_Esquerda
        .Enabled = False
        .Top = 0
        .Left = 0
    End With
    
    'Skin_Top_Centro
    With Skin_Top_Centro
        '.Enabled = False
        .Top = 0
        .Stretch = True
        .Width = Frame_Top.Width - Skin_Top_Esquerda.Width - Skin_Top_Direita.Width
        .Left = Skin_Top_Esquerda.Left + Skin_Top_Esquerda.Width
    End With
    
    'Skin_Top_Direita
    With Skin_Top_Direita
        .Enabled = False
        .Top = 0
        .Left = Frame_Top.Width - .Width
    End With
    
    'Setas_Selecao
    With Setas_Selecao
        .Top = 1600
        .Left = 180
    End With
    
    'Botao_Fechar
    With Botao_Fechar
        .Top = 120
        .Left = Frame_Top.Width - 250
    End With
    
    'Botao_Maximizar
    With Botao_Maximizar
        .Top = 120
        .Left = Botao_Fechar.Left - 250
    End With
    
    'Botao_Minimizar
    With Botao_Minimizar
        .Top = 150
        .Left = Botao_Maximizar.Left - 250
    End With
    
    'Skin_Top_Player
    With Skin_Top_Player
        .Top = 0
        .Left = (Frame_Top.Width - .Width) / 2
    End With
    
    '********************************************************************** Frame Centro *******************************************************************
    'Frame_Centro
    With Frame_Centro
        .Height = Me.Height - Frame_Top.Height - Frame_Down.Height
        .Top = Frame_Top.Top + Frame_Top.Height
        .Width = Me.Width
        .Left = 0
    End With

    'Skin_Lateral_Esquerda
    With Skin_Lateral_Esquerda
        .Enabled = False
        .Stretch = True
        .Height = Frame_Centro.Height
        .Top = 0
        .Left = 0
    End With

    'Skin_Lateral_Direita
    With Skin_Lateral_Direita
        .Enabled = False
        .Stretch = True
        .Height = Frame_Centro.Height
        .Top = 0
        .Left = Frame_Centro.Width - .Width
    End With

    'Barra_Biblioteca
    With Barra_Biblioteca
        .Top = 0
        .Left = Skin_Lateral_Esquerda.Left + Skin_Lateral_Esquerda.Width
    End With

    'Line1
    With Line1
        '.X1 = Barra_Biblioteca.Left + Barra_Biblioteca.Width + 10
        '.X2 = Barra_Biblioteca.Left + Barra_Biblioteca.Width + 10
        .Y2 = Frame_Centro.Height
        .Y1 = 0
    End With

    'Frame_Lista
    With Frame_Lista
        .Height = Frame_Centro.Height - Barra_Biblioteca.Height ' - Frame_Video.Height
        .Top = Barra_Biblioteca.Top + Barra_Biblioteca.Height
    End With

    'Grelha
    With Grelha
        .Height = Frame_Centro.Height
        .Top = 0
        .Width = Frame_Centro.Width - Skin_Lateral_Esquerda.Width - Skin_Lateral_Direita.Width - Barra_Biblioteca.Width - Line1.BorderWidth
        .Left = Frame_Lista.Left + Frame_Lista.Width
    End With

    'WebBrowser1
    With WebBrowser1
        .Height = Grelha.Height
        .Top = 0
        .Width = Grelha.Width
        .Left = Grelha.Left
    End With
    '*********************************************************************** Frame Down **********************************************************************
    'Frame_Down
    With Frame_Down
        .Top = Me.Height - .Height
        .Width = Me.Width
        .Left = 0
    End With

    'Skin_Down_Esquerda
    With Skin_Down_Esquerda
        .Enabled = False
        .Top = 0
        .Left = 0
    End With

    'Skin_Down_Centro
    With Skin_Down_Centro
        .Enabled = False
        .Top = 0
        .Stretch = True
        .Width = Frame_Down.Width - Skin_Down_Esquerda.Width - Skin_Down_Direita.Width
        .Left = Skin_Down_Esquerda.Left + Skin_Down_Esquerda.Width
    End With

    'Skin_Down_Direita
    With Skin_Down_Direita
        .Enabled = False
        .Top = 0
        .Left = Frame_Down.Width - Skin_Down_Direita.Width
    End With

    'Skin_Down_Botoes
    With Skin_Down_Botoes
        .Top = 0
        .Left = (Frame_Down.Width - .Width) / 2
    End With

    'Posicionar frames da Treeview
    Posicionar_Nos
End Sub

Public Sub Repor_Imagens()
    Label_Mudo.ForeColor = &H808080
    Label_Ver_Biblioteca.ForeColor = &H808080
    Label_Ver_Navegador.ForeColor = &H808080
    Label_Ver_Radio.ForeColor = &H808080
    Label_Ver_Tv.ForeColor = &H808080
    Label_Ver_Propriedades.ForeColor = &H808080
    Label_Ver_Video.ForeColor = &H808080
    Label_Ver_Lista.ForeColor = &H808080
End Sub
