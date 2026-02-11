VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   7080
      Top             =   2280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2115
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   0
      Width           =   7455
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   2400
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI ' カーソル位置情報取得用変数
 x As Long
 y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long



Private Sub Command1_Click()
    Dim x1, x2, y1, y2
    Dim cn  As ADODB.Connection
    Dim rs  As ADODB.Recordset
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    Picture1.Line (0, Picture1.Height - 100)-(Picture1.Width, Picture1.Height - 100), QBColor(0)
    Picture1.Line (10, 0)-(10, Picture1.Height), QBColor(0)
    
    x1 = 10
    y1 = Picture1.Height - 100
    
    ' 接続文字列を設定
    cn.ConnectionString = "Provider=MSDASQL.1;Persisit Security Info=False;User ID=Dolphin;Data Source=PostgreSQL"
    cn.Open
    ' テーブルを指定し、レコードセットをオープン
    rs.Open "Select * From t", cn, adOpenDynamic, adLockBatchOptimistic, adCmdText

    With rs
        Do While Not rs.EOF
            x2 = .Fields("x")
            y2 = .Fields("y")
            Picture1.Line (x1, Picture1.Height - y1 - 100)-(x2, Picture1.Height - y2 - 100), QBColor(1)
            x1 = x2
            y1 = y2
            .MoveNext
        Loop
    End With
    
    rs.Close
    cn.Close
End Sub

Private Sub Picture1_Click()
    Dim Rect As POINTAPI
    Dim lngRet As Long
    
    ' マウスカーソル座標を取得
    lngRet = GetCursorPos(Rect)
    If lngRet <> 0 Then
        Me.Label1 = "X=" & Format$(Rect.x, "@@") & " Y=" & Format$(Rect.y, "@@")
    End If
End Sub
