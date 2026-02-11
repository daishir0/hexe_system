VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton Command2 
      Caption         =   "更新"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "抽出"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim cn  As ADODB.Connection
    Dim rs  As ADODB.Recordset
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    ' 接続文字列を設定
    cn.ConnectionString = "Provider=MSDASQL.1;Persisit Security Info=False;User ID=Dolphin;Data Source=PostgreSQL"
    cn.Open
    ' テーブルを指定し、レコードセットをオープン
    rs.Open "Select * From conf_detail_typet", cn, adOpenDynamic, adLockBatchOptimistic, adCmdText

    With rs
        Label1.Caption = .Fields("detail")
    End With
    
    rs.Close
    cn.Close
End Sub

Private Sub Command2_Click()
    Dim cn  As ADODB.Connection
    Dim rs  As ADODB.Recordset
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    ' 接続文字列を設定
    cn.ConnectionString = "Provider=MSDASQL.1;Persisit Security Info=False;User ID=Dolphin;Data Source=PostgreSQL"
    cn.Open
    ' テーブルを指定し、レコードセットをオープン
    rs.Open "Select * From conf_detail_typet", cn, adOpenDynamic, adLockOptimistic, adCmdText
    
    With rs
    '.Fields("id") = "1"
    .Fields("detail") = "雑談"
    .Update
    End With
    
    rs.Close
    cn.Close



End Sub

