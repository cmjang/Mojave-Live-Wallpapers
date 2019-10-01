VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   435
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   420
   ScaleWidth      =   435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'------------------------------------------------------------------------------------------------------------
Dim ts As Variant

Private Sub Form_Load()
Call change '修改壁纸
Form1.Hide
ts = Hour(Now())
'-----------------------------------------------------------------------------------------
'防止程序多开
If App.PrevInstance Then '如果重复启动
Dim Hwd As Long, mhwnd As Long, HwdOld As Long
Dim i As Integer, Title As String, ClassName As String
mhwnd = Me.hwnd
ClassName = GetWinClass(mhwnd)
Title = GetWinText(mhwnd)
HwdOld = mhwnd
mhwnd = 0
Hwd = FindWindowEx(mhwnd, 0, ClassName, Title)
Do Until Hwd = 0
If Hwd <> HwdOld Then '找到先前启动的本程序
'ShowWindow Hwd, 1
'SetForegroundWindow Hwd
Exit Do
End If
Hwd = FindWindowEx(mhwnd, Hwd, ClassName, Title)
Loop
End
End If

End Sub

Private Sub Timer1_Timer()
If ts <> Hour(Now()) Then
Call change
ts = Hour(Now())
Else

End If
End Sub
