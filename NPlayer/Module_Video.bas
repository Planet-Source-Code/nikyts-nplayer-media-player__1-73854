Attribute VB_Name = "Module_Video"
Option Explicit
' Declarations for Move Borderless Form
Public Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As Any) As Long

Public Declare Sub ReleaseCapture Lib "user32" ()
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2


' Declarations for Form On Top
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal CX As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long

' Declarations for Rounded Form
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Function FileExist(FileNameIn As String) As Boolean
  
  On Error GoTo ErrRtn
  Dim i As Integer
  
  On Error Resume Next
  
  i = Len(Dir$(FileNameIn))
  
  If Err Or i = 0 Then
    FileExist = False
  Else
    FileExist = True
  End If
    
ProcExit:
  Exit Function

ErrRtn:
  Beep
  Resume ProcExit:
   
End Function

Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) As Long
  If Topmost = True Then
    SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
  Else
    SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
 End If

End Function

Public Function OnTop(Form As Form, Top As Boolean)
  SetTopMostWindow Form.hwnd, Top
End Function

Public Function ParamValue(ParseCharacter As String, _
                               tString As Variant, _
                               Index As Integer) As String
  Dim CurrentPosition As Integer
  Dim ParseToPosition As Integer
  Dim CurrentToken As Integer
  Dim TempString As String
  TempString = Trim(tString) + ParseCharacter
  If Len(TempString) = 1 Then Exit Function
  CurrentPosition = 1
  CurrentToken = 1
  Do
    ParseToPosition = InStr(CurrentPosition, TempString, _
                            ParseCharacter)
    If Index = CurrentToken Then
      ParamValue = Mid$(TempString, CurrentPosition, _
                        ParseToPosition - CurrentPosition)
      Exit Function
    End If
    CurrentToken = CurrentToken + 1
    CurrentPosition = ParseToPosition + 1
  Loop Until (CurrentPosition >= Len(TempString))
  
End Function

Public Function ParamCount(ParseCharacter As String, _
                               tString As Variant) As Integer
  Dim CurrentPosition As Integer
  Dim ParseToPosition As Integer
  Dim CurrentToken As Integer
  Dim TempString As String
  TempString = Trim(tString) + ParseCharacter
  If Len(TempString) = 1 Then Exit Function
  CurrentPosition = 1
  CurrentToken = 1
  Do
      ParseToPosition = InStr(CurrentPosition, TempString, _
                              ParseCharacter)
      CurrentToken = CurrentToken + 1
      CurrentPosition = ParseToPosition + 1
  Loop Until (CurrentPosition >= Len(TempString))
  ParamCount = CurrentToken - 1
  If Right(Trim(tString), 1) = ParseCharacter Then
    ParamCount = ParamCount + 1
  End If
  
End Function

Public Function Duration(TotalSeconds As Double) As String
  Dim Seconds
  Dim Minutes
  Dim Hours
  Dim Days
  Dim DayString As String
  Dim HourString As String
  Dim MinuteString As String
  Dim SecondString As String
  
  Seconds = Int(TotalSeconds Mod 60)
  Minutes = Int(TotalSeconds \ 60 Mod 60)
  Hours = Int(TotalSeconds \ 3600 Mod 24)
  Days = Int(TotalSeconds \ 3600 \ 24)

  If Days = 1 Then DayString = " day, " _
  Else: DayString = " days, "
  HourString = ":"
  MinuteString = ":"
  SecondString = ""
  
  Select Case Days
    Case 0
      Duration = Format(Hours, "#") & IIf(Hours = 0, "", HourString) & _
            Format(Minutes, "00") & MinuteString & _
             Format(Seconds, "00") & SecondString
    Case Else
      Duration = Days & DayString & _
          Format(Hours, "#") & IIf(Hours = 0, "", HourString) & Format _
          (Minutes, "00") & MinuteString & _
           Format(Seconds, "00") & SecondString
  End Select
                 
End Function

Public Function ComputeDuration(TimeHours As String) As Double
  If ParamCount(":", TimeHours) = 3 Then
    ComputeDuration = Val(ParamValue(":", TimeHours, 1)) * 3600 + _
                      Val(ParamValue(":", TimeHours, 2)) * 60 + _
                      Val(ParamValue(":", TimeHours, 3))
    Exit Function
  End If
  
  If ParamCount(":", TimeHours) = 2 Then
    ComputeDuration = Val(ParamValue(":", TimeHours, 1)) * 60 + _
                      Val(ParamValue(":", TimeHours, 2))
    Exit Function
  End If
  
  If ParamCount(":", TimeHours) = 1 Then
    ComputeDuration = Val(ParamValue(":", TimeHours, 1))
    Exit Function
  End If
End Function


