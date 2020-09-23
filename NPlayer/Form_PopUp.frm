VERSION 5.00
Begin VB.Form Form_PopUp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3300
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrPopUp 
      Interval        =   500
      Left            =   360
      Top             =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H001E1F1D&
      BackStyle       =   0  'Transparent
      Caption         =   "Media player"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblArtist 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label lblSong 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   1695
      Left            =   0
      Top             =   0
      Width           =   3300
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H001E1F1D&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   15
      TabIndex        =   3
      Top             =   15
      Width           =   3270
   End
End
Attribute VB_Name = "Form_PopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Dim DirectionIsUp As Boolean

Private Sub SetOnTop(Optional ByVal on_top As Boolean = True)
    If on_top Then
        SetWindowPos hwnd, _
            HWND_TOPMOST, 0, 0, 0, 0, _
            SWP_NOMOVE + SWP_NOSIZE
    Else
        SetWindowPos hwnd, _
            HWND_NOTOPMOST, 0, 0, 0, 0, _
            SWP_NOMOVE + SWP_NOSIZE
    End If
End Sub

Private Sub Form_Click()
    Form_Principal.Show
End Sub

Private Sub Form_Load()
    Me.Top = Screen.Height + 50
    Me.left = Screen.Width - (Me.Width + 100)
    DirectionIsUp = True
    SetOnTop
End Sub

Private Sub tmrPopUp_Timer()
    tmrPopUp.Interval = 10
        If DirectionIsUp Then
            Me.Top = Me.Top - 50
            If (Me.Top <= Screen.Height - (Me.Height - 10)) Then
                tmrPopUp.Interval = 3000
                DirectionIsUp = False
            End If
        Else
            Me.Top = Me.Top + 50
            If Me.Top >= Screen.Height + 10 Then
                tmrPopUp.Enabled = False
                Unload Me
            End If
     End If
End Sub
