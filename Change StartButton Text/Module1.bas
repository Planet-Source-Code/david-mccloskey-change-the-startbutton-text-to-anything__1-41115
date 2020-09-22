Attribute VB_Name = "Module1"
Public Const WM_SETTEXT = &HC

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Public Function StringToByteArray(str As String) As Variant
Dim bray() As Byte
Dim cnt As Integer
Dim ln As Integer

ln = Len(str)

ReDim bray(ln)

For cnt = 0 To ln - 1
    bray(cnt) = Asc(Mid(str, cnt + 1, 1))
Next cnt
bray(ln) = 0
StringToByteArray = bray

End Function
