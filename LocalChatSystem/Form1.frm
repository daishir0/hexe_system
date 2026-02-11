VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1560
      Width           =   3015
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   3000
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2280
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Winsock1.SendData Text2.Text
End Sub

Private Sub Form_Load()
    ' 送信側ソケット
    With Winsock1
      If .State = sckClosed Then
        'プロトコルをUDPにする。
        .Protocol = sckUDPProtocol
        'ローカルのポート番号を指定する。
        .Bind 8001
        '送信先の指定
        .RemoteHost = "150.37.222.255"
        'リモートポートの指定
        .RemotePort = 8002
      End If
    End With
    
    ' 受信側ソケット
    With Winsock2
        If .State = sckClosed Then
            .Protocol = sckUDPProtocol
            .Bind 8002
        End If
    End With
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
    Dim str As String
    
    Winsock2.GetData str
    Text1.Text = Text1.Text & vbCrLf & str & vbCrLf
End Sub

