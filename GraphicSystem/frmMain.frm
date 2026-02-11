VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "へぇ～システム"
   ClientHeight    =   4110
   ClientLeft      =   3165
   ClientTop       =   1230
   ClientWidth     =   8115
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   8115
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      Top             =   3720
      Width           =   615
   End
   Begin VB.ComboBox conf_detail 
      Height          =   300
      Left            =   1680
      TabIndex        =   14
      Text            =   "現在の状況を選んで下さい"
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "議事切り替え(&4)"
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   5
      Top             =   3240
      Width           =   1300
   End
   Begin VB.CheckBox bHexesound 
      Caption         =   "Check1"
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "来週のトピック(&6)"
      Height          =   375
      Index           =   5
      Left            =   6720
      TabIndex        =   7
      Top             =   3240
      Width           =   1300
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6240
      Top             =   1800
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "txt"
      DialogTitle     =   "保存先のファイルを選択してください"
      FileName        =   "untitled.txt"
      Filter          =   "テキストファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*"
   End
   Begin MSWinsockLib.Winsock sck2 
      Left            =   6960
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   8002
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   6960
      Tag             =   "5"
      Top             =   1800
   End
   Begin MSWinsockLib.Winsock sck1 
      Left            =   6960
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "150.37.222.255"
      RemotePort      =   8002
      LocalPort       =   8001
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "結論(&5)"
      Height          =   375
      Index           =   4
      Left            =   5400
      TabIndex        =   6
      Top             =   3240
      Width           =   1300
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "雑談(&3)"
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   4
      Top             =   3240
      Width           =   1300
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "意見・疑問(&2)"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   3240
      Width           =   1300
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "議事録(&1)"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1300
   End
   Begin VB.TextBox txtInput 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Text            =   "＞（ここにメッセージを打ち込んでください）"
      Top             =   2880
      Width           =   6135
   End
   Begin VB.TextBox txtOutput 
      Height          =   2655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7935
   End
   Begin VB.Label Label4 
      Caption         =   "へぇ～を鳴らさない"
      Height          =   375
      Left            =   6360
      TabIndex        =   12
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "へぇ"
      Height          =   255
      Left            =   7200
      TabIndex        =   9
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   6600
      TabIndex        =   8
      Top             =   2880
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "ファイル(&F)"
      Begin VB.Menu mnuSave 
         Caption         =   "名前を付けて保存(&S)..."
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "終了(&X)"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isInit As Boolean
Dim userName As String
Dim num As Integer
Dim hexe_sum As Integer
Dim SEND_TO_SERVER As Boolean
Dim conf_id As Integer
Public Start_Date As Date
Public Start_Time As Date

'WAVを鳴らすAPI関数
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
(ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
'PlaySoundの定数
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000
Const SND_LOOP = &H8
Const SND_SYNC = &H10

'始まった時間
Dim globalT As Date

Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()
    frmGraph.Show
    Command1.Enabled = False
End Sub

Private Sub Command2_Click()
    MsgBox conf_detail.List(conf_detail.ListIndex)
End Sub

Private Sub conf_detail_Change()
    MsgBox conf_detail.List(conf_detail.ListIndex)
End Sub

'フォーム読み込み
Private Sub Form_Load()

Start_Date = Now
Start_Time = Time

'（暫定）conf_idをfrmLoginから持ってくる
conf_id = frmLogin.conf_id

globalT = 0
' 二重起動チェック
  If App.PrevInstance Then
    MsgBox "既に起動されています。"
    End
  End If
  
' 変数初期化
  isInit = True
  SEND_TO_SERVER = True
    '（暫定）ユーザネームをfrmLoginから持ってくる
    userName = frmLogin.userName
    
    ' 送信側ソケット
    With sck1
      If .State = sckClosed Then
        'プロトコルをUDPにする。
        .Protocol = sckUDPProtocol
        'ローカルのポート番号を指定する。
        .Bind 8001
        '送信先の指定
        .RemoteHost = "255.255.255.255"
        'リモートポートの指定
        .RemotePort = 8002
      End If
    End With
    
    ' 受信側ソケット
    With sck2
        If .State = sckClosed Then
            .Protocol = sckUDPProtocol
            .Bind 8002
        End If
    End With

    'MsgBox sck.LocalIP & ":" & sck.LocalPort

End Sub

'ボタンを押した時の処理
Private Sub btnSend_Click(index As Integer)
    If txtInput = "" Then Exit Sub
    
    Dim prefix As String
    'メッセージの頭にプレフィックスを付加
    Select Case index
        Case 0:
            prefix = "(議事)"
        Case 1:
            prefix = "(意見)"
        Case 2:
            prefix = "(雑談)"
        Case 3:
            sendMessage "　（" & userName & "さんからの議事切替です。【" & txtInput & "】）"
            
            'DBへ登録
            If SEND_TO_SERVER Then
                put_MessageTB frmLogin.user_id, index, txtInput
            End If
    
            '入力ボックスの後処理
            With txtInput
                .Text = ""
                .SetFocus
            End With
            Exit Sub
        Case 4:
            prefix = "(結論)"
        Case 5:
            prefix = "(来週)"
        Case Else:
            MsgBox "Error"
            End
    End Select
    'メッセージ送信
    '（）と時間をとりあえずはずす
    'sendMessage Label3.Caption & str(Index) & ":" & prefix & " " & userName & " ＞" & txtInput
    sendMessage str(index) & ":" & userName & " ＞" & txtInput
    
    'DBへ登録
    If SEND_TO_SERVER Then
        put_MessageTB frmLogin.user_id, index, txtInput
    End If
    
    '入力ボックスの後処理
    With txtInput
        .Text = ""
        .SetFocus
    End With
End Sub

Private Sub mnuQuit_Click()

    If MsgBox("終了しますか", vbOKCancel + vbDefaultButton1 + vbQuestion, "終了") = vbCancel Then
    Exit Sub
    End If

End

End Sub

Private Sub mnuSave_Click()
On Error GoTo Err_Command1
Dim Contents As String
Dim FileName As String

CommonDialog1.ShowSave
' ファイル名を表示
FileName = CommonDialog1.FileName
Contents = txtOutput.Text

' ファイルに保存
Open FileName For Output As #1
Print #1, Contents;
Close #1

Exit Sub

Err_Command1:
' エラーの内容を表示
MsgBox Err.Description
End Sub

'メッセージ受信時の処理
Private Sub sck2_DataArrival(ByVal bytesTotal As Long)
    Dim str As String
    
    sck2.GetData str
    txtOutput.Text = txtOutput.Text & str & vbCrLf
    txtOutput.SelStart = Len(txtOutput)
End Sub

'メッセージ送信時の処理
Private Sub sendMessage(str As String)
    sck1.SendData str
End Sub

'へぇ～のhexe Tableへの登録．５回以上は５回としてカウント
Private Sub Timer1_Timer()
    
    If num <> 0 Then
    
    '５へぇストッパー
    If num > 5 Then num = 5
    
    ' DBへの登録
    If SEND_TO_SERVER Then
        put_hexeTB frmLogin.user_id, 0, num
    End If
    
    hexe_sum = hexe_sum + num
    num = 0
    Label1.Caption = 0 & " (" & hexe_sum & ")"
    End If
    
End Sub

Private Sub Timer2_Timer()
Dim sec As Integer
Dim min As Integer
Dim hor As Integer

globalT = globalT + 1
hor = Fix(globalT / 3600)
min = Fix(globalT / 60) Mod 60
sec = globalT Mod 60

Label3.Caption = Format(hor, "00") & ":" & Format(min, "00") & ":" & Format(sec, "00")

End Sub

'おまけ
Private Sub txtInput_GotFocus()
    If isInit = True Then txtInput.Text = ""
    isInit = False
End Sub

'チャットをminutes Tableへ登録処理
Private Sub put_MessageTB(user As Integer, c_type As Integer, minute As String)
    Dim cn   As ADODB.Connection
    Dim rs   As ADODB.Recordset
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    ' 接続文字列を設定
    cn.ConnectionString = _
        "Provider=MSDASQL.1;Persist Security Info=False;User ID=postgres;Data Source=PostgreSQL"
    cn.Open
    cn.Execute "INSERT INTO minutes VALUES(NEXTVAL('minutes_id_seq')," & user & "," & c_type & ",'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & minute & "','" & conf_id & "')"
    cn.Close
End Sub

'へぇ～をhexe Tableへ登録処理
Private Sub put_hexeTB(user As Integer, h_type As Integer, p_times As Integer)
    Dim cn   As ADODB.Connection
    Dim rs   As ADODB.Recordset
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    ' 接続文字列を設定
    cn.ConnectionString = _
        "Provider=MSDASQL.1;Persist Security Info=False;User ID=postgres;Data Source=PostgreSQL"
    cn.Open
    cn.Execute "INSERT INTO hexe VALUES(NEXTVAL('hexe_id_seq')," & user & "," & h_type & ",'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & p_times & "','" & conf_id & "')"
    cn.Close
End Sub

'へぇ～の押し判定
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
    num = num + 1
    
    '「鳴らさない」にチェックがついていたらへぇ～を鳴らさない
    If bHexesound.Value = False Then
    Call PlaySound("hexe.wav", 0, SND_FILENAME + SND_ASYNC)
    End If
    
    Label1.Caption = str(num) & " (" & hexe_sum & ")"
    End If
    
End Sub

'textInputにエンターのみ入力のときは議事録入力と判断＆送信
Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtInput = "" Then Exit Sub '入力が空白なら何もせずに終わる
    
    If KeyCode = vbKeyReturn Then
        'メッセージ送信
        sendMessage " 0:" & userName & " >" & txtInput.Text
        
        'DBへ登録
        If SEND_TO_SERVER Then
            put_MessageTB frmLogin.user_id, 0, txtInput.Text
        End If
        
        '入力ボックスの後処理
        With txtInput
            .Text = ""
            .SetFocus
        End With
    End If
End Sub

