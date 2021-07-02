VERSION 5.00
Begin VB.Form FrmLoading 
   BackColor       =   &H00000000&
   Caption         =   "Loading..."
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6795
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   ScaleHeight     =   4785
   ScaleWidth      =   6795
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   240
      Top             =   2400
   End
   Begin CSO.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   3360
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      Value           =   0
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Loading..."
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3840
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   1440
      Picture         =   "Form8.frx":628A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3780
   End
End
Attribute VB_Name = "FrmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
 이히 = 3


팁 = Int((3 * Rnd) + 1)

If 팁 = 1 Then
 Label1 = "◎Tip : 랭크는 Normal < Special < Rare < Unique < Elite < Legend입니다."
ElseIf 팁 = 2 Then
 Label1 = "◎Tip : 카드 합성을 할 경우, 능력치가 증가합니다."
ElseIf 팁 = 3 Then
 Label1 = "◎Tip : 카드 판매를 할 경우에는 원래 금액을 전부 돌려주지는 않습니다."
End If
End Sub

Private Sub Image1_Click()
If val(로딩) = 1 Then
 Unload Me
 로딩 = 0
 FrmGameInfo.Show
End If
End Sub


Private Sub Timer1_Timer()
If ProgressBar1.Value = 100 Then
 ProgressBar1.Text = "로딩 완료. 이미지를 클릭해주세요."
 로딩 = 1
Else
 If ProgressBar1.Value = 50 Then
  FrmPickSt.Show
  FrmPickSt.Visible = False
 ElseIf ProgressBar1.Value = 5 Then
  FrmPlayGame.Show
  FrmPlayGame.Visible = False
  FrmPlayGame.txtLoad.Text = "ㅠㅠ"
 End If
 ProgressBar1.Value = val(ProgressBar1.Value) + 1
 If val(이히) = 1 Then
  ProgressBar1.Text = "Loading."
  이히 = 2
 ElseIf val(이히) = 2 Then
  ProgressBar1.Text = "Loading.."
  이히 = 3
 ElseIf val(이히) = 3 Then
  ProgressBar1.Text = "Loading..."
  이히 = 1
 End If
End If
End Sub

