Attribute VB_Name = "Module2"
Option Explicit
' ファイルハンドルの値が向こうであることを示す定数の宣言
Public Const INVALID_HANDLE_VALUE = (-1)

' 変更通知待ちをフィルタ条件示す定数の宣言
Public Const FILE_NOTIFY_CHANGE_FILE_NAME = &H1&    ' ファイル名の変更
Public Const FILE_NOTIFY_CHANGE_DIR_NAME = &H2&     ' ディレクトリ名の変更
Public Const FILE_NOTIFY_CHANGE_ATTRIBUTES = &H4&   ' 属性の変更
Public Const FILE_NOTIFY_CHANGE_SIZE = &H8&         ' ファイルサイズの変更
Public Const FILE_NOTIFY_CHANGE_LAST_WRITE = &H10&         ' ファイルの最終書き込み時刻の変更
Public Const FILE_NOTIFY_CHANGE_SECURITY = &H100&   ' セキュリティ記述子の変更

' 変更通知ハンドルを作成しディレクトリを監視する関数の宣言
Declare Function FindFirstChangeNotification Lib "kernel32.dll" _
    Alias "FindFirstChangeNotificationA" _
   (ByVal lpPathName As String, _
    ByVal bWatchSubtree As Long, _
    ByVal dwNotifyFilter As Long) As Long

' ディレクトリ通知ハンドルの変更の漢詩を中止する関数の宣言
Declare Function FindCloseChangeNotification Lib "kernel32.dll" _
    (ByVal hChangeHandle As Long) As Long

' 関数が制御を戻した原因を示す定数の宣言
Public Const STATUS_ABANDONED_WAIT_0 = &H80&
Public Const STATUS_WAIT_0 = &H0&
Public Const STATUS_TIMEOUT = &H102&
Public Const WAIT_ABANDONED = ((STATUS_ABANDONED_WAIT_0) + 0)
Public Const WAIT_FAILED = &HFFFFFFFF
Public Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)
Public Const WAIT_TIMEOUT = STATUS_TIMEOUT

' オブジェクトがシグナル状態になった時または
' タイムアウト時間が経過した時に制御を戻す関数の宣言
Declare Function WaitForSingleObject Lib "kernel32.dll" _
    (ByVal hHandle As Long, _
     ByVal dwMilliseconds As Long) As Long


