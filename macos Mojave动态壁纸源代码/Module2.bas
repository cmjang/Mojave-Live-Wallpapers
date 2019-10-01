Attribute VB_Name = "Module2"
Option Explicit
'·ÀÖ¹³ÌÐò¶à¿ª
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Function GetWinClass(Hwd As Long) As String
Dim retvalue As Long, TempStr As String * 254
retvalue = GetClassName(Hwd, TempStr, 254)
GetWinClass = StrConv(LeftB(StrConv(TempStr, vbFromUnicode), retvalue), vbUnicode)
End Function
Public Function GetWinText(ByVal hwnd As Long) As String
Dim i As Long, Title As String
Title = String(255, Chr(0))
i = GetWindowText(hwnd, Title, Len(Title) - 1)
GetWinText = StrConv(LeftB(StrConv(Title, vbFromUnicode), i), vbUnicode)
End Function


