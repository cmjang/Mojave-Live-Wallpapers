VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mojave��̬ɳĮ��ֽ����"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   4935
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "ȡ����������"
      Height          =   735
      Left            =   3600
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���ó���Ϊ��������"
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�رն�̬��ֽ"
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub InitCommonControls _
Lib "comctl32.dll" ()

Private Sub Command2_Click()
 Call start
 Command2.Enabled = False
 Command3.Enabled = True
End Sub

Private Sub Command3_Click()
Call moend
Command3.Enabled = False
Command2.Enabled = True
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Command1_Click()
Shell "taskkill /im Mojave.exe /f", vbHide
Dim w As Object
Dim a As String
If exitproc("Mojave.exe") Then  ''�ж��Ƿ�Mojave.exe�Ľ��̴���
    Label1.Caption = "���"
    Command2.Enabled = False
    Command1.Enabled = False
Else
    Label1.Caption = "ERROR ж�ش���"
    Command1.Enabled = True
    End If
    
End Sub

Private Sub Form_Load()
If exitproc("Mojave.exe") Then  ''�ж��Ƿ�qq.exe�Ľ��̴���
    Label1.Caption = "��̬��ֽ��������,���ж�ؼ���"
Else
    Label1.Caption = "����δ����,����ж��"
    Command1.Enabled = False
    End If
End Sub
