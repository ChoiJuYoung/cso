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
   StartUpPosition =   2  'ȭ�� ���
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
         Name            =   "����"
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
      Alignment       =   2  '��� ����
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
 ���� = 3


�� = Int((3 * Rnd) + 1)

If �� = 1 Then
 Label1 = "��Tip : ��ũ�� Normal < Special < Rare < Unique < Elite < Legend�Դϴ�."
ElseIf �� = 2 Then
 Label1 = "��Tip : ī�� �ռ��� �� ���, �ɷ�ġ�� �����մϴ�."
ElseIf �� = 3 Then
 Label1 = "��Tip : ī�� �ǸŸ� �� ��쿡�� ���� �ݾ��� ���� ���������� �ʽ��ϴ�."
End If
End Sub

Private Sub Image1_Click()
If val(�ε�) = 1 Then
 Unload Me
 �ε� = 0
 FrmGameInfo.Show
End If
End Sub


Private Sub Timer1_Timer()
If ProgressBar1.Value = 100 Then
 ProgressBar1.Text = "�ε� �Ϸ�. �̹����� Ŭ�����ּ���."
 �ε� = 1
Else
 If ProgressBar1.Value = 50 Then
  FrmPickSt.Show
  FrmPickSt.Visible = False
 ElseIf ProgressBar1.Value = 5 Then
  FrmPlayGame.Show
  FrmPlayGame.Visible = False
  FrmPlayGame.txtLoad.Text = "�Ф�"
 End If
 ProgressBar1.Value = val(ProgressBar1.Value) + 1
 If val(����) = 1 Then
  ProgressBar1.Text = "Loading."
  ���� = 2
 ElseIf val(����) = 2 Then
  ProgressBar1.Text = "Loading.."
  ���� = 3
 ElseIf val(����) = 3 Then
  ProgressBar1.Text = "Loading..."
  ���� = 1
 End If
End If
End Sub

