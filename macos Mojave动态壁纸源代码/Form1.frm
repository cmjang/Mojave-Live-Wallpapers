VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   645
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   1800
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   645
   ScaleWidth      =   1800
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ts As Variant

Private Sub Form_Load()
Call change 'ÐÞ¸Ä±ÚÖ½
Form1.Hide
ts = Hour(Now())
End Sub

Private Sub Timer1_Timer()
If ts <> Hour(Now()) Then
Call change
ts = Hour(Now())
Else

End If
End Sub
