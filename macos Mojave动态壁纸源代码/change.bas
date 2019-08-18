Attribute VB_Name = "Module1"
Option Explicit
'此修改桌面壁纸为指定图片函数
'==========================================================
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_SETDESKWALLPAPER = 20
Private Const SPIF_UPDATEINIFILE = &H1

Public Function change()
    Dim time As Variant
    Dim img As String
    Dim i As Long
    Dim t As Long
    time = Hour(Now()) '获取系统小时数
    '判断过程
    If time = 0 Then i = 16
    If time = 1 Then i = 16
    If time = 2 Then i = 16
    If time = 3 Then i = 16
    If time = 4 Then i = 1
    If time = 5 Then i = 2
    If time = 6 Then i = 3
    If time = 7 Then i = 4
    If time = 8 Then i = 5
    If time = 9 Then i = 6
    If time = 10 Then i = 6
    If time = 11 Then i = 7
    If time = 12 Then i = 7
    If time = 13 Then i = 7
    If time = 14 Then i = 8
    If time = 15 Then i = 9
    If time = 16 Then i = 10
    If time = 17 Then i = 11
    If time = 18 Then i = 12
    If time = 19 Then i = 13
    If time = 20 Then i = 14
    If time = 21 Then i = 15
    If time = 22 Then i = 16
    If time = 23 Then i = 16
'======================================================
    img = App.Path & "/img/" & CStr(i) & ".jpeg"
    t = SystemParametersInfo(ByVal SPI_SETDESKWALLPAPER, True, ByVal img, SPIF_UPDATEINIFILE)
End Function
