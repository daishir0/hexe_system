VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "へぇ～システム"
   ClientHeight    =   4410
   ClientLeft      =   3165
   ClientTop       =   1230
   ClientWidth     =   8385
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8385
   Begin VB.CommandButton btnSend 
      Caption         =   "意見(&2)"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   3360
      Width           =   1300
   End
   Begin VB.TextBox minutes_text 
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直
      TabIndex        =   16
      Top             =   240
      Width           =   4095
   End
   Begin VB.CheckBox bZatsudan 
      Caption         =   "Check1"
      Height          =   180
      Left            =   6120
      TabIndex        =   13
      Top             =   4080
      Width           =   255
   End
   Begin VB.Timer hexetimer 
      Enabled         =   0   'False
      Left            =   5520
      Top             =   720
   End
   Begin VB.ComboBox conf_detail 
      Height          =   300
      ItemData        =   "frmMain.frx":08CA
      Left            =   1680
      List            =   "frmMain.frx":08E3
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   12
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "グラフを表示"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CheckBox bHexesound 
      Caption         =   "Check1"
      Height          =   255
      Left            =   6120
      TabIndex        =   9
      Top             =   3840
      Width           =   255
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
      Left            =   4080
      TabIndex        =   5
      Top             =   3360
      Width           =   1300
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "雑談(&3)"
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   4
      Top             =   3360
      Width           =   1300
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "議事録(&1)"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   1300
   End
   Begin VB.TextBox txtInput 
      Height          =   510
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直
      TabIndex        =   1
      Text            =   "frmMain.frx":093D
      Top             =   2760
      Width           =   6135
   End
   Begin VB.TextBox txtOutput 
      Height          =   2415
      Left            =   4320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label8 
      Caption         =   "<<議事録>>"
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "<<意見・雑談等>>"
      Height          =   255
      Left            =   5520
      TabIndex        =   17
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "システム起動時間"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "雑談を表示しない"
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7560
      Picture         =   "frmMain.frx":0968
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "へぇ～を鳴らさない"
      Height          =   255
      Left            =   6360
      TabIndex        =   10
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "へぇ"
      Height          =   255
      Left            =   7080
      TabIndex        =   7
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   3000
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
Dim username As String
Dim num As Integer
Dim hexe_sum As Integer
Dim SEND_TO_SERVER As Boolean
Dim conf_id As Integer
Dim now_conf_detail_id As Integer
Dim now_conf_stime As String
Dim cancel_ctrl As Boolean
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

Private Sub Command1_Click()
    frmGraph.Show
    Command1.Enabled = False
    frmMain.Show
End Sub

'会議詳細に変化があった場合
Private Sub conf_detail_Click()

    If now_conf_detail_id = -1 Then
    '初めて会議詳細を設定する場合
    '現在の会議詳細idと開始時間を初期化する
    now_conf_detail_id = conf_detail.ListIndex
    now_conf_stime = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    'メッセージ送信
    sendconfdetail conf_detail.Text, conf_detail.ListIndex
        
    Else
        '「議題の開始」以外で会議詳細に変化が無い場合何もしない
        If now_conf_detail_id = conf_detail.ListIndex And now_conf_detail_id > 0 Then
        'No Action

        Else
        '会議詳細に変化が有る場合，ＤＢに書き込む．会議詳細idは１～である
        put_cdTB frmLogin.user_id, now_conf_detail_id + 1, now_conf_stime
        '現在の会議詳細idと開始時間を初期化する，
        now_conf_detail_id = conf_detail.ListIndex
        now_conf_stime = Format(Now, "yyyy-mm-dd hh:mm:ss")
        
        sendconfdetail conf_detail.Text, conf_detail.ListIndex
        End If
    End If
        
End Sub

'フォーム読み込み
Private Sub Form_Load()


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
  username = frmLogin.username
'（暫定）conf_idをfrmLoginから持ってくる
  conf_id = frmLogin.conf_id
  now_conf_detail_id = -1
  cancel_ctrl = False
  Start_Date = Now
  Start_Time = Time
  
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
    
    'ユーザがmintuesだったら，議事のみの入力にする
    If username = "minutes" Then
    btnSend(0).Enabled = True
    btnSend(1).Enabled = False
    btnSend(2).Enabled = False
    btnSend(4).Enabled = False
    Else
    btnSend(0).Enabled = False
    btnSend(1).Enabled = True
    btnSend(2).Enabled = True
    btnSend(4).Enabled = True
    'txtInput.MultiLine = False
    End If
    
End Sub

'ボタンを押した時の処理
Private Sub btnSend_Click(index As Integer)
    If txtInput = "" Then Exit Sub
    
    Dim prefix As String
    'メッセージの頭にプレフィックスを付加
    Select Case index
        Case 0:
            prefix = "議"
        Case 1:
            prefix = "意"
        Case 2:
            prefix = "雑"
        Case 3:
            sendMessage "　（" & username & "さんからの議事切替です。【" & txtInput & "】）"
            
            'DBへ登録
            put_MessageTB frmLogin.user_id, index, txtInput

    
            '入力ボックスの後処理
            With txtInput
                .Text = ""
                .SetFocus
            End With
            Exit Sub
        Case 4:
            prefix = "結"
        Case 5:
            prefix = "来"
        Case Else:
            MsgBox "Error"
            End
    End Select
    
    'メッセージ送信
    '（）と時間をとりあえずはずす
    'sendMessage Label3.Caption & str(Index) & ":" & prefix & " " & userName & " ＞" & txtInput
    If index = 0 Then
    'sendMessage username & " >" & txtInput
    sendMessage txtInput
    Else
    sendMessage prefix & ":" & username & " >" & txtInput
    End If
    
    'DBへ登録
    put_MessageTB frmLogin.user_id, index, txtInput

    '入力ボックスの後処理
    With txtInput
        .Text = ""
        .SetFocus
    End With
End Sub

'閉じるボタンが押されたときの終了確認
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If MsgBox("議事録データが消去されます．本当に終了しますか？", vbOKCancel + vbDefaultButton1 + vbQuestion, "終了確認") = vbCancel Then
    Cancel = 1
    End If

End Sub

'システムが落ちるときの後処理
Private Sub Form_Unload(Cancel As Integer)

    '会議詳細を指定していないなら落ちる
    If now_conf_detail_id = -1 Then Exit Sub
    
    '最後の会議詳細を書き込む
    put_cdTB frmLogin.user_id, now_conf_detail_id + 1, now_conf_stime
    '
    MsgBox " 最後の会議詳細を書き込んだよ．終わっちゃうの？(((´･ω･`)ｶｯｸﾝ…"
End Sub

Private Sub hexetimer_Timer()
    If Image1.Visible = False Then
    Image1.Visible = True
    Else
    Image1.Visible = False
    End If

    If hexetimer.Interval <> 0 Then
    hexetimer.Interval = hexetimer.Interval - 25
    End If
    
End Sub

'ラベルをクリックしても「へぇ～を鳴らさない」のチェックが変わるように
Private Sub Label4_Click()
    If bHexesound.Value = 0 Then
    bHexesound.Value = 1
    Else
    bHexesound.Value = 0
    End If
End Sub

Private Sub Label5_Click()
    If bHexesound.Value = 0 Then
    bZatsudan.Value = 1
    Else
    bZatsudan.Value = 0
    End If
End Sub

'メニューで終了が押されたときの確認
Private Sub mnuQuit_Click()

    If MsgBox("議事録データが消去されます．本当に終了しますか？", vbOKCancel + vbDefaultButton1 + vbQuestion, "終了確認") = vbCancel Then
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
    Contents = "【以下，議事録】" & vbCrLf & minutes_text.Text & vbCrLf & "【以下，周りのメモ】" & vbCrLf & txtOutput.Text
    
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
    
    'minutesユーザ以外はマルチラインで入力されてしまったvbCrLfを消す
    If username <> "minutes" Then
    str = Replace(str, vbCrLf, "")
    End If
    
    '雑談を消すオプションが付いていたらsubを抜ける
    If Left(str, 1) = "雑" And bZatsudan.Value = 1 Then
    MsgBox "find zatsudan"
    Exit Sub
    End If
    
    '議事録だったら右ウインドウに出るようにする
    If Left(str, 1) = "結" Or Left(str, 1) = "雑" Or Left(str, 1) = "意" Then
    txtOutput.Text = txtOutput.Text & str & vbCrLf & vbCrLf
    txtOutput.SelStart = Len(txtOutput)
    Exit Sub
    End If

    minutes_text.Text = minutes_text.Text & str & vbCrLf
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
    put_hexeTB frmLogin.user_id, 0, num
    
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
    If SEND_TO_SERVER = False Then Exit Sub
    Dim cn   As ADODB.Connection
    Dim rs   As ADODB.Recordset
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    ' 接続文字列を設定
    cn.ConnectionString = _
        "Provider=MSDASQL.1;Persist Security Info=False;User ID=postgres;Data Source=PostgreSQL"
    cn.Open
    cn.Execute "INSERT INTO minutes VALUES(NEXTVAL('minutes_id_seq')," & user & "," & c_type & ",'" & "Now()" & "','" & minute & "','" & conf_id & "')"
    cn.Close
End Sub

'へぇ～をhexe Tableへ登録処理
Private Sub put_hexeTB(user As Integer, h_type As Integer, p_times As Integer)
    If SEND_TO_SERVER = False Then Exit Sub
    Dim cn   As ADODB.Connection
    Dim rs   As ADODB.Recordset
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    ' 接続文字列を設定
    cn.ConnectionString = _
        "Provider=MSDASQL.1;Persist Security Info=False;User ID=postgres;Data Source=PostgreSQL"
    cn.Open
    cn.Execute "INSERT INTO hexe VALUES(NEXTVAL('hexe_id_seq')," & user & "," & h_type & ",'" & "Now()" & "','" & p_times & "','" & conf_id & "')"
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
    
    'へぇ～を点滅させる
    hexetimer.Interval = 300
    hexetimer.Enabled = True
    End If

End Sub

'textInputにエンターのみ入力のときは議事録入力と判断＆送信
Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtInput = "" Then Exit Sub '入力が空白なら何もせずに終わる
    If username = "minutes" Then Exit Sub '議事録取っているひとならエンター入力させない
    
    'デフォルトではエンターのみ入力で意見にする
    If KeyCode = vbKeyReturn Then
        'メッセージ送信
        sendMessage "意" & username & " >" & txtInput.Text

        'DBへ登録
        put_MessageTB frmLogin.user_id, 1, txtInput.Text

        '入力ボックスの後処理
        With txtInput
            .Text = ""
            .SetFocus

        End With
    End If
End Sub

'議事詳細をconf_detail Tableへ登録処理 cd_type = confdetail_table
Private Sub put_cdTB(user As Integer, cd_type As Integer, s_time As String)
    If SEND_TO_SERVER = False Then Exit Sub
    Dim cn   As ADODB.Connection
    Dim rs   As ADODB.Recordset
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    ' 接続文字列を設定
    cn.ConnectionString = _
        "Provider=MSDASQL.1;Persist Security Info=False;User ID=postgres;Data Source=PostgreSQL"
    cn.Open
    cn.Execute "INSERT INTO conf_detail VALUES(NEXTVAL('minutes_id_seq')," & user & "," & cd_type & ",'" & s_time & "','" & "Now()" & "','" & conf_id & "')"
    cn.Close
End Sub

'会議詳細を切り替えた際のメッセージ送信
Private Sub sendconfdetail(msg As String, cd As Integer)
    Dim username As String
    
    If cd = 0 Then
    
    '議題の開始の場合は名前をゲットする
    username = InputBox("開始される議題の対象者を入力してください")
    If username = "" Then
    Exit Sub
    Else
    sendMessage "■■■■■■■" & username & "さんの" & msg & "です■■■■■■■"
    End If
    
    '入力ボックスの後処理
    txtInput.Text = ""
    Else
    sendMessage "【" & msg & "です】"
    End If
    
    txtInput.SetFocus
End Sub
