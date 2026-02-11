VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "『へぇ～』システムへようこそ"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows の既定値
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Conf ID："
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Your ID："
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public user_id As Integer
Public conf_id As Integer
Public username As String

Private Sub Command1_Click()
'とりあえずプロトタイプ
    Select Case Text1.Text
        Case "michiru":
            user_id = 1
        Case "teshiga":
            user_id = 2
        Case "takahashi":
            user_id = 13
        Case "guille":
            user_id = 14
        Case "kuwa":
            user_id = 21
        Case "katayama":
            user_id = 22
        Case "teru":
            user_id = 23
        Case "daishiro":
            user_id = 25
        Case "j-square":
            user_id = 24
        Case "Dolphin":
            user_id = 27
        Case "minamida":
            user_id = 26
        Case "furukawa":
            user_id = 35
        Case "syasuda":
            user_id = 36
        Case "hsekiguc":
            user_id = 32
        Case "mtakagi":
            user_id = 33
        Case "tienaga":
            user_id = 41
        Case "mitsuo":
            user_id = 43
        Case "seantokyo":
            user_id = 54
        Case "mtoyoda":
            user_id = 52
        Case "kengo":
            user_id = 44
        Case "yhoshino":
            user_id = 46
        Case "tienaga":
            user_id = 41
        Case "abe":
            user_id = 48
        Case "koichik":
            user_id = 51
        Case "brueslee":
            user_id = 53
        Case "hiphoper":
            user_id = 42
        Case "tatsuaki":
            user_id = 55
        Case "mmikami":
            user_id = 45
        Case "yao":
            user_id = 56
        Case "nodaken":
            user_id = 11
        Case "kinage":
            user_id = 47
        Case "namal":
            user_id = 37
        Case "minutes":
            user_id = 100

        Case Else:
            MsgBox "Error"
            Text1.Text = ""
            Beep
            End
    End Select
    
    If user_id > 0 Then
    username = Text1.Text
    conf_id = Val(Text2.Text)
    frmMain.Show
    Unload Me
    End If

End Sub


'エンターが入ったときには入力が終わったと判断
Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Text1.Text = "" Then Exit Sub '入力が空白なら何もせずに抜ける
'
'    If KeyCode = vbKeyReturn Then
'    Command1_Click
'    End If
End Sub
