VERSION 5.00
Begin VB.Form FrmBug 
   BackColor       =   &H00000000&
   Caption         =   "��������"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5970
   Icon            =   "Form24.frx":0000
   LinkTopic       =   "Form24"
   ScaleHeight     =   7545
   ScaleWidth      =   5970
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.TextBox txtPL�ؿ�� 
      Height          =   270
      Left            =   4680
      TabIndex        =   58
      Text            =   "Text2"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtPL��� 
      Height          =   270
      Left            =   4680
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtCo 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4560
      TabIndex        =   56
      Text            =   "Text8"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtSe 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4560
      TabIndex        =   55
      Text            =   "Text7"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtPa 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4560
      TabIndex        =   54
      Text            =   "Text6"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtDe 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4560
      TabIndex        =   53
      Text            =   "Text5"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtAm 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4560
      TabIndex        =   52
      Text            =   "Text4"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtSt 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4560
      TabIndex        =   51
      Text            =   "Text3"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtR 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4560
      TabIndex        =   50
      Text            =   "Text2"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtAt 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4560
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   480
      Width           =   1335
   End
   Begin CSO.jcbutton jcbutton1 
      Height          =   255
      Left            =   360
      TabIndex        =   40
      Top             =   7080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "����"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.TextBox txt��ũ 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4560
      TabIndex        =   39
      Text            =   "Text20"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtP���� 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   38
      Text            =   "Text19"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox TxtP�� 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   37
      Text            =   "Text18"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox txtZ���� 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   36
      Text            =   "Text17"
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox txtZ�� 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   35
      Text            =   "Text16"
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox txtT���� 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   34
      Text            =   "Text15"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox txtT�� 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   33
      Text            =   "Text14"
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox txtA���� 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   32
      Text            =   "Text13"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox TxtA�� 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   31
      Text            =   "Text12"
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox txtSevic 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   30
      Text            =   "Text11"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtVic 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   29
      Text            =   "Text10"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtSkill 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   28
      Text            =   "Text9"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtLev 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   27
      Text            =   "Text8"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtEXP 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   26
      Text            =   "Text7"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtPL 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   25
      Text            =   "Text6"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtPW 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   24
      Text            =   "Text5"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtZL 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   23
      Text            =   "Text4"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtZW 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   22
      Text            =   "Text3"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtTL 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtTW 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "PL�ؿ��"
      Height          =   255
      Left            =   3240
      TabIndex        =   60
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "PL���"
      Height          =   255
      Left            =   3240
      TabIndex        =   59
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "se"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   27
      Left            =   3120
      TabIndex        =   48
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "co"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   26
      Left            =   3120
      TabIndex        =   47
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "pa"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   25
      Left            =   3120
      TabIndex        =   46
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "de"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   24
      Left            =   3120
      TabIndex        =   45
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "am"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   23
      Left            =   3120
      TabIndex        =   44
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "st"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   22
      Left            =   3120
      TabIndex        =   43
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "r"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   21
      Left            =   3120
      TabIndex        =   42
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "at"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   3120
      TabIndex        =   41
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "��ũ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   3120
      TabIndex        =   19
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "P����"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   18
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "P��������"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   17
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Z����"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   16
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Z��������"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   15
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "T����"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   14
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "T��������"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "A����"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "A��������"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�ؿ��"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "���"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Skill"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "����"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "����ġ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "P��"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "P��"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Z��"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Z��"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "T��"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "T��"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "FrmBug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
txtTW = MyTW(����)
txtTL = MyTL(����)
txtZW = MyZW(����)
txtZL = MyZL(����)
txtPW = MyPW(����)
txtPL = MyPL(����)
txtEXP = MyExp(����)
txtLev = MyLev(����)
txtVic = MyVic(����)
txtSevic = MySeVic(����)
TxtA�� = MyA��(����)
txtA���� = MyA����(����)
txtT�� = MyT��(����)
txtT���� = MyT����(����)
txtZ�� = MyZ��(����)
txtZ���� = MyZ����(����)
TxtP�� = MyP��(����)
txtP���� = MyP����(����)
txt��ũ = MyRank(����)
txtAt = MyAt(����)
txtR = MyR(����)
txtSt = MySt(����)
txtAm = MyAm(����)
txtDe = MyDe(����)
txtPa = MyPa(����)
txtSe = MySe(����)
txtCo = MyCo(����)
txtSkill = MySkill(����)
txtPL��� = PL���
txtPL�ؿ�� = PL�ؿ��
End Sub

Private Sub jcbutton1_Click()
MyTW(����) = txtTW
MyTL(����) = txtTL
MyZW(����) = txtZW
MyZL(����) = txtZL
MyPW(����) = txtPW
MyPL(����) = txtPL
MyAW(����) = val(txtTW) + val(txtZW) + val(txtPW)
MyAL(����) = val(txtTL) + val(txtZL) + val(txtPL)
MyLev(����) = txtLev
MyExp(����) = txtEXP
MyVic(����) = txtVic
MySeVic(����) = txtSevic
MyA��(����) = TxtA��
MyA����(����) = txtA����
MyT��(����) = txtT��
MyT����(����) = txtT����
MyZ��(����) = txtZ��
MyZ����(����) = txtZ����
MyP��(����) = TxtP��
MyP����(����) = txtP����
MyRank(����) = txt��ũ
MyAt(����) = txtAt
MyR(����) = txtR
MySt(����) = txtSt
MyAm(����) = txtAm
MyDe(����) = txtDe
MyPa(����) = txtPa
MySe(����) = txtSe
MyCo(����) = txtCo
MySkill(����) = txtSkill
PL��� = txtPL���
PL�ؿ�� = txtPL�ؿ��
FrmMain.Timer12.Enabled = True
End Sub
