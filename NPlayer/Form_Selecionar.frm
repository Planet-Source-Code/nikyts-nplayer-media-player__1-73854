VERSION 5.00
Begin VB.Form Form_Selecionar 
   Caption         =   "Selecionar ficheiro"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   ScaleHeight     =   284
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   409
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox fileMP3 
      Height          =   1845
      Left            =   2400
      Pattern         =   "*.mp3"
      TabIndex        =   2
      Top             =   0
      Width           =   2295
   End
   Begin VB.DirListBox dirMP3 
      Height          =   1665
      Left            =   0
      TabIndex        =   1
      Top             =   330
      Width           =   2415
   End
   Begin VB.DriveListBox drvMP3 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "Form_Selecionar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
