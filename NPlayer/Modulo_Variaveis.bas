Attribute VB_Name = "Modulo_Variaveis"
'Declaração das variáveis
'Cores utilizadas pelo sistema
Option Explicit
Public Numero_de_Janelas As Integer

Public Modo_Mascara As Boolean

'API para o procedimento alway's on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Colocar o formulário por cima dos outros
Sub AlwaysOnTop(FrmID As Form, OnTop As Integer)
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
    If OnTop = -1 Then
        OnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        OnTop = SetWindowPos(FrmID.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Sub

Sub Main()
    'Iniciar o numero inicial de janelas
    Numero_de_Janelas = 0
    Modo_Mascara = False
End Sub
