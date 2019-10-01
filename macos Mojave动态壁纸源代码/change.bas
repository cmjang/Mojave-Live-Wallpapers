Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
'此修改桌面壁纸为指定图片函数
'==========================================================
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_SETDESKWALLPAPER = 20
Private Const SPIF_UPDATEINIFILE = &H1

Public Function change()
    Dim time As Variant
    Dim img As String, img1 As String, img2 As String
    Dim i As Long
    Dim t As Long
    Dim d As Long
    Dim mon As Long
    Dim down As Variant
    Dim picscr As String, picscr2 As String
    time = Hour(Now()) '获取系统小时数
    img1 = App.Path & "\img" '取得img的路径
    mon = Month(Date)  '获取月份
    Dim a As Long
    '判断过程
    If mon <= 4 Then
    '1~4月份的
        If time = 0 Then i = 16
        If time = 1 Then i = 16
        If time = 2 Then i = 16
        If time = 3 Then i = 16
        If time = 4 Then i = 1
        If time = 5 Then i = 2
        If time = 6 Then i = 2
        If time = 7 Then i = 3
        If time = 8 Then i = 4
        If time = 9 Then i = 5
        If time = 10 Then i = 6
        If time = 11 Then i = 7
        If time = 12 Then i = 7
        If time = 13 Then i = 8
        If time = 14 Then i = 9
        If time = 15 Then i = 10
        If time = 16 Then i = 11
        If time = 17 Then i = 12
        If time = 18 Then i = 13
        If time = 19 Then i = 14
        If time = 20 Then i = 15
        If time = 21 Then i = 16
        If time = 22 Then i = 16
        If time = 23 Then i = 16
        ElseIf mon <= 6 Then
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
    ElseIf mon <= 9 Then
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
    ElseIf mon <= 12 Then
        If time = 0 Then i = 16
        If time = 1 Then i = 16
        If time = 2 Then i = 16
        If time = 3 Then i = 16
        If time = 4 Then i = 1
        If time = 5 Then i = 2
        If time = 6 Then i = 2
        If time = 7 Then i = 3
        If time = 8 Then i = 4
        If time = 9 Then i = 5
        If time = 10 Then i = 6
        If time = 11 Then i = 7
        If time = 12 Then i = 7
        If time = 13 Then i = 8
        If time = 14 Then i = 9
        If time = 15 Then i = 10
        If time = 16 Then i = 11
        If time = 17 Then i = 12
        If time = 18 Then i = 13
        If time = 19 Then i = 14
        If time = 20 Then i = 15
        If time = 21 Then i = 16
        If time = 22 Then i = 16
        If time = 23 Then i = 16
    End If
    '======================================================
    img = App.Path & "/img/" & CStr(i) & ".jpeg"
    picscr = "http://tan90.pw/img/" & CStr(i) & ".jpeg"
    If Dir$(img1, vbDirectory) <> "" Then
        'img存在
        If Dir(img) <> "" Then
                t = SystemParametersInfo(ByVal SPI_SETDESKWALLPAPER, True, ByVal img, SPIF_UPDATEINIFILE)
            Else
            d = URLDownloadToFile(0, picscr, img, 0, 0)
            t = SystemParametersInfo(ByVal SPI_SETDESKWALLPAPER, True, ByVal img, SPIF_UPDATEINIFILE)
            '下载全部图片
            For a = 1 To 16
            picscr2 = "http://tan90.pw/img/" & CStr(a) & ".jpeg"
            img2 = App.Path & "/img/" & CStr(a) & ".jpeg"
            If Dir(img2) = "" Then
            d = URLDownloadToFile(0, picscr2, img2, 0, 0)
            End If
        Next a
        End If
    Else
        'img不存在
        MkDir (img1) '创建文件夹
        d = URLDownloadToFile(0, picscr, img, 0, 0)
        t = SystemParametersInfo(ByVal SPI_SETDESKWALLPAPER, True, ByVal img, SPIF_UPDATEINIFILE)
        '下载全部图片
        For a = 1 To 16
            picscr2 = "http://tan90.pw/img/" & CStr(a) & ".jpeg"
            img2 = App.Path & "/img/" & CStr(a) & ".jpeg"
            If Dir(img2) = "" Then
            d = URLDownloadToFile(0, picscr2, img2, 0, 0)
            End If
        Next a
    End If

End Function
