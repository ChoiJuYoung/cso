VERSION 5.00
Begin VB.Form FrmPickSt 
   BackColor       =   &H00808080&
   Caption         =   "�÷��� ��Ÿ�� ����"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5295
   Icon            =   "Form12.frx":0000
   LinkTopic       =   "Form12"
   ScaleHeight     =   3615
   ScaleWidth      =   5295
   StartUpPosition =   2  'ȭ�� ���
   Begin CSO.jcbutton CmdSt 
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   3240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "����� �÷���(-)"
      ForeColor       =   16776960
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin CSO.jcbutton CmdR 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   3240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "������ �÷���(-)"
      ForeColor       =   65535
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin CSO.jcbutton Cmd� 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "��� �÷���(-)"
      ForeColor       =   16711935
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin CSO.jcbutton CmdDe 
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   2880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "������ �÷���(-)"
      ForeColor       =   65280
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin CSO.jcbutton CmdAttack 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   2880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "������ �÷���(-)"
      ForeColor       =   255
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��� ����
      BackColor       =   &H00808080&
      Caption         =   "�÷��� ����â"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "- ������ �÷��� ��Ÿ���� �����Ͽ� �ֽʽÿ�. -"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   5295
   End
   Begin VB.Label lblName 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "���±�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   5295
   End
   Begin VB.Image ImgPla 
      Height          =   1500
      Left            =   1800
      Top             =   120
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  '�������� ����
      Height          =   2895
      Left            =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "FrmPickSt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd�_Click()
Style = "���"
FrmPlayGame.Visible = True
FrmPlayGame.Text1.Text = "������"
Unload Me
End Sub

Private Sub CmdAttack_Click()
Style = "������"
FrmPlayGame.Visible = True
FrmPlayGame.Text1.Text = "������"
Unload Me
End Sub

Private Sub CmdDe_Click()
Style = "������"
FrmPlayGame.Visible = True
FrmPlayGame.Text1.Text = "������"
Unload Me
End Sub

Private Sub CmdR_Click()
Style = "������"
FrmPlayGame.Visible = True
FrmPlayGame.Text1.Text = "������"
Unload Me
End Sub

Private Sub CmdSt_Click()
Style = "�����"
FrmPlayGame.Visible = True
FrmPlayGame.Text1.Text = "������"
Unload Me
End Sub

Private Sub Form_Load()
lblName = " ' " + MyName(����) + " ' "
If Len(Dir(App.Path & "\img\����\[" & Mid(MyYear(����), 2, 2) & "]" & MyName(����) & ".gif")) <> 0 Then
 ImgPla.Picture = LoadPicture(App.Path & "\img\����\[" & Mid(MyYear(����), 2, 2) & "]" & MyName(����) & ".gif")
Else
 ImgPla = LoadPicture(App.Path & "\img\����\" & MyName(����) & ".gif")
End If
Style = ""
End Sub

