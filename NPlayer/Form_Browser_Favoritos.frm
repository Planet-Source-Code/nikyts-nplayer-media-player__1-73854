VERSION 5.00
Begin VB.Form Form_Browser_Favoritos 
   Caption         =   "Form Browser Favoritos"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin NPlayer.McListBox Lista_Favoritos 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   6588
      Picture         =   "Form_Browser_Favoritos.frx":0000
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
      BorderStyle     =   0
      IconFocus       =   0   'False
      RowHeight       =   18
      SelectionStyle  =   0
      AutoHideScrollBars=   -1  'True
      Path            =   "C:\NPlayer v2\"
   End
End
Attribute VB_Name = "Form_Browser_Favoritos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
