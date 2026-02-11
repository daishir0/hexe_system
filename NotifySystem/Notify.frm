VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "フォルダ監視"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton Command1 
      Caption         =   "実行"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "C:\test"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "タイムアウト："
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "10秒"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "監視ディレクトリ："
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "監視結果："
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub command1_Click()
    Dim strWatchPathName        As String
    Dim lngWatchSubTree         As Long
    Dim lngWatchFilter          As Long
    Dim lngChangeNotifyHandle   As Long
    Dim lngTimeOut              As Long
    Dim lngReturnEvent          As Long
    Dim lngWin32apiResultCode   As Long
    
    ' コマンドボタンを無効に設定
    Command1.Enabled = False
    ' ラベルを初期化
    Label3.Caption = ""
    
    DoEvents
    ' 監視するディレクトリを指定
    strWatchPathName = "C:\test"
    ' 監視するディレクトリに対するフラグを設定
    lngWatchSubTree = False
    ' 変更通知待ちを満たすフィルタ条件を指定
    lngWatchFilter = FILE_NOTIFY_CHANGE_FILE_NAME Or _
                     FILE_NOTIFY_CHANGE_DIR_NAME Or _
                     FILE_NOTIFY_CHANGE_ATTRIBUTES Or _
                     FILE_NOTIFY_CHANGE_SIZE Or _
                     FILE_NOTIFY_CHANGE_LAST_WRITE

    ' 変更通知ハンドルを作成
    lngChangeNotifyHandle = FindFirstChangeNotification(strWatchPathName, _
                                                        lngWatchSubTree, _
                                                        lngWatchFilter)
    ' 変更通知ハンドルを作成できた時は
    If lngChangeNotifyHandle <> INVALID_HANDLE_VALUE Then
        ' タイムアウト時間を指定
        lngTimeOut = 10000
        ' オブジェクトのシグナル状態を監視
        lngReturnEvent = WaitForSingleObject(lngChangeNotifyHandle, lngTimeOut)
        
        With Label3
            ' 制御を戻す原因となんったイベントを表示
            Select Case lngReturnEvent
                Case WAIT_ABANDONED
                    .Caption = "未開放ミューテックスオブジェクト"
                Case WAIT_FAILED
                    .Caption = "待機失敗"
                Case WAIT_OBJECT_0
                    .Caption = "変更通知を受信"
                Case WAIT_TIMEOUT
                    .Caption = "タイムアウト"
            End Select
        End With
        ' 変更通知ハンドルをクローズ
        lngWin32apiResultCode = FindCloseChangeNotification(lngChangeNotifyHandle)
    End If
    ' コマンドボタンを有効に設定
    Command1.Enabled = True
End Sub
