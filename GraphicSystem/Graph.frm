VERSION 5.00
Begin VB.Form frmGraph 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "Graph"
   ClientHeight    =   3465
   ClientLeft      =   5445
   ClientTop       =   4755
   ClientWidth     =   8160
   Icon            =   "Graph.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   8160
   Begin VB.Timer Timer3 
      Interval        =   10000
      Left            =   6120
      Top             =   2640
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6600
      Top             =   2640
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7080
      Top             =   2640
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2100
      Left            =   120
      ScaleHeight     =   2040
      ScaleWidth      =   7815
      TabIndex        =   0
      Top             =   0
      Width           =   7875
      Begin VB.Line Line7 
         X1              =   7560
         X2              =   7560
         Y1              =   10
         Y2              =   2000
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   7750
         Y1              =   995
         Y2              =   995
      End
      Begin VB.Line Line5 
         X1              =   6048
         X2              =   6048
         Y1              =   10
         Y2              =   2000
      End
      Begin VB.Line Line4 
         X1              =   4536
         X2              =   4536
         Y1              =   10
         Y2              =   2000
      End
      Begin VB.Line Line3 
         X1              =   3024
         X2              =   3024
         Y1              =   10
         Y2              =   2000
      End
      Begin VB.Line Line2 
         X1              =   1512
         X2              =   1512
         Y1              =   10
         Y2              =   2000
      End
      Begin VB.Line Line1 
         X1              =   10
         X2              =   7750
         Y1              =   2000
         Y2              =   2000
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblDate 
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   2760
      Width           =   5175
   End
   Begin VB.Label lblPoint 
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim Start_Date As Date
Dim SpendTime As Integer
Dim x1, x2, y1, y2
Dim xx As Integer
Dim cn  As ADODB.Connection
Dim rs  As ADODB.Recordset

Private Sub Command1_Click()
    Call draw
End Sub

Private Sub Form_Activate()
    Call draw
End Sub

Private Sub Form_Load()
    SpendTime = 30
       
    ' 時間軸ラベル初期化
    Label1.Caption = Format(frmMain.Start_Time, "hh:nn")
    Label2.Caption = Format(DateAdd("n", 6, frmMain.Start_Time), "hh:nn")
    Label3.Caption = Format(DateAdd("n", 12, frmMain.Start_Time), "hh:nn")
    Label4.Caption = Format(DateAdd("n", 18, frmMain.Start_Time), "hh:nn")
    Label5.Caption = Format(DateAdd("n", 24, frmMain.Start_Time), "hh:nn")
    
    '表示位置をfrmMainの下に
    Me.Top = frmMain.Top + frmMain.Height
    Me.Left = frmMain.Left
    lblDate.Caption = "SD:" & frmMain.Start_Date & " ST:" & frmMain.Start_Time & " Now:" & Now & " SpT:" & SpendTime
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.ConnectionString = "Provider=MSDASQL.1;Persisit Security Info=False;User ID=postgres;Data Source=PostgreSQL"
    cn.Open
    
    Call draw
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Command1.Enabled = True
    cn.Close
    Set cn = Nothing
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblPoint.Caption = "X:" & X & " Y:" & Y & vbCrLf
End Sub

Private Sub Timer1_Timer()
    lblDate.Caption = "SD:" & frmMain.Start_Date & " ST:" & frmMain.Start_Time & " Now:" & Now & " SpT:" & SpendTime
End Sub

Private Sub Timer2_Timer()
    SpendTime = SpendTime + 30
    Label2.Caption = Format(DateAdd("n", SpendTime / 5, Start_Time), "hh:nn")
    Label3.Caption = Format(DateAdd("n", SpendTime / 5 * 2, Start_Time), "hh:nn")
    Label4.Caption = Format(DateAdd("n", SpendTime / 5 * 3, Start_Time), "hh:nn")
    Label5.Caption = Format(DateAdd("n", SpendTime / 5 * 4, Start_Time), "hh:nn")
End Sub

Private Sub Timer3_Timer()
    Dim Temp_Date As Date
    Temp_Date = frmMain.Start_Date
    
    Picture1.Cls
    ' テーブルを指定し、レコードセットをオープン
    rs.Open "SELECT * FROM hexe WHERE push_time > '" & Temp_Date & "' AND push_time < '" & DateAdd("s", 10, Temp_Date) & "'", cn, adOpenDynamic, adLockReadOnly, adCmdText
'    rs.Open "SELECT times FROM hexe WHERE push_time > '" & DateAdd("s", -60, Date) & "'", cn, adOpenDynamic, adLockReadOnly, adCmdText
    x1 = 10
    y1 = 0
    
    With rs
        Do While DateAdd("s", 10, Temp_Date) < Now
            x2 = x1 + 42
            If Not rs.EOF Then
                rs.Close
                rs.Open "SELECT sum(times) FROM hexe WHERE push_time > '" & Temp_Date & "' AND push_time < '" & DateAdd("s", 10, Temp_Date) & "'", cn, adOpenDynamic, adLockReadOnly, adCmdText
                y2 = rs.Fields("sum")
                Label6.Caption = "(" & x1 & "," & y1 & ")   (" & x2 & "," & y2 & ")"
            Else
                y2 = 0
            End If
            Picture1.Line (x1, Picture1.Height - y1 * 200 - 100)-(x2, Picture1.Height - y2 * 200 - 100), QBColor(1)
            y1 = y2
            x1 = x2
            Temp_Date = DateAdd("s", 10, Temp_Date)
            rs.Close
            rs.Open "SELECT * FROM hexe WHERE push_time > '" & Temp_Date & "' AND push_time < '" & DateAdd("s", 10, Temp_Date) & "'", cn, adOpenDynamic, adLockReadOnly, adCmdText
        Loop
    End With
    
    rs.Close
End Sub

Sub draw()
    Dim Temp_Date As Date
    Temp_Date = frmMain.Start_Date
    
    Picture1.Cls
    ' テーブルを指定し、レコードセットをオープン
    rs.Open "SELECT * FROM hexe WHERE push_time > '" & Temp_Date & "' AND push_time < '" & DateAdd("s", 10, Temp_Date) & "'", cn, adOpenDynamic, adLockReadOnly, adCmdText
'    rs.Open "SELECT times FROM hexe WHERE push_time > '" & DateAdd("s", -60, Date) & "'", cn, adOpenDynamic, adLockReadOnly, adCmdText
    x1 = 10
    y1 = 0
    
    With rs
        Do While DateAdd("s", 10, Temp_Date) < Now
            x2 = x1 + 42
            If Not rs.EOF Then
                rs.Close
                rs.Open "SELECT sum(times) FROM hexe WHERE push_time > '" & Temp_Date & "' AND push_time < '" & DateAdd("s", 10, Temp_Date) & "'", cn, adOpenDynamic, adLockReadOnly, adCmdText
                y2 = rs.Fields("sum")
                Label6.Caption = "(" & x1 & "," & y1 & ")   (" & x2 & "," & y2 & ")"
            Else
                y2 = 0
            End If
            Picture1.Line (x1, Picture1.Height - y1 * 50 - 100)-(x2, Picture1.Height - y2 * 50 - 100), QBColor(1)
            y1 = y2
            x1 = x2
            Temp_Date = DateAdd("s", 10, Temp_Date)
            rs.Close
            rs.Open "SELECT * FROM hexe WHERE push_time > '" & Temp_Date & "' AND push_time < '" & DateAdd("s", 10, Temp_Date) & "'", cn, adOpenDynamic, adLockReadOnly, adCmdText
        Loop
    End With
    
    rs.Close
End Sub
