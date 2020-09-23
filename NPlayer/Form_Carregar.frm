VERSION 5.00
Begin VB.Form Form_Carregar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture_Carregar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   960
      ScaleHeight     =   1455
      ScaleWidth      =   2895
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2895
      Begin VB.Timer Timer_Carregar 
         Interval        =   200
         Left            =   2160
         Top             =   120
      End
      Begin VB.Line Linha 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   7
         Index           =   0
         X1              =   1455
         X2              =   1455
         Y1              =   415
         Y2              =   240
      End
      Begin VB.Line Linha 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   7
         Index           =   1
         X1              =   1575
         X2              =   1695
         Y1              =   525
         Y2              =   405
      End
      Begin VB.Line Linha 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   7
         Index           =   3
         X1              =   1575
         X2              =   1695
         Y1              =   765
         Y2              =   885
      End
      Begin VB.Line Linha 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   7
         Index           =   5
         X1              =   1215
         X2              =   1335
         Y1              =   885
         Y2              =   765
      End
      Begin VB.Line Linha 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   7
         Index           =   7
         X1              =   1335
         X2              =   1215
         Y1              =   510
         Y2              =   390
      End
      Begin VB.Line Linha 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   7
         Index           =   6
         X1              =   1080
         X2              =   1210
         Y1              =   645
         Y2              =   645
      End
      Begin VB.Line Linha 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   7
         Index           =   4
         X1              =   1455
         X2              =   1455
         Y1              =   1045
         Y2              =   885
      End
      Begin VB.Line Linha 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   7
         Index           =   2
         X1              =   1695
         X2              =   1825
         Y1              =   645
         Y2              =   645
      End
      Begin VB.Label Label_Aguarde 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Aguarde o Processamento..."
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   480
         TabIndex        =   1
         Top             =   1200
         Width           =   2025
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
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
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   2895
   End
End
Attribute VB_Name = "Form_Carregar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VARIÁVEIS DO PROGRESSBAR
'Const Escuro = &HFF890A 'Azul
Const Claro = &HC0C0C0 ' Cinza
Dim Atual As Integer

Private Sub Form_Load()
    'Estado do progress bar
    Atual = 0
End Sub

Private Sub Timer_Carregar_Timer()
    'ACTIVAR PROGRESSBAR
    DoEvents
    If Atual > 7 Then
        Atual = 0
        Linha(7).BorderColor = Claro
    End If
    Linha(Atual).BorderColor = &HFC6C03    'Escuro
    If Atual <> 0 Then Linha(Atual - 1).BorderColor = Claro
    Atual = Atual + 1
End Sub
