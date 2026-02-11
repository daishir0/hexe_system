VERSION 5.00
Begin VB.Form へぇー 
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
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.Timer Timer 
      Interval        =   5000
      Left            =   3720
      Top             =   360
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "へぇー"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim num As Integer

'WAVを鳴らすAPI関数
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
(ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
'PlaySoundの定数
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000
Const SND_LOOP = &H8
Const SND_SYNC = &H10


Private Sub Command1_Click()
Call PlaySound("hexe.wav", 0, SND_FILENAME + SND_ASYNC)
End Sub

Private Sub Form_Load()
    'へぇーの数
    num = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Ctrlキーのコードが17らしい。Shiftって変数は同時に押してるか
    'どうかで変わると思うんだけど、よくわからん。
    If KeyCode = 17 Then
        If Shift = 2 Then
            num = num + 1
        'とりあえずShift+Ctrlで減らせるようにもしてみた。
        ElseIf Shift <> 0 And num > 0 Then
            num = num - 1
        End If
        Label1.Caption = num
    End If
    
    Label2.Caption = KeyCode & " " & Shift
End Sub

