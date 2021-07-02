VERSION 5.00
Begin VB.Form FrmBug 
   BackColor       =   &H00000000&
   Caption         =   "직접진행"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5970
   Icon            =   "Form24.frx":0000
   LinkTopic       =   "Form24"
   ScaleHeight     =   7545
   ScaleWidth      =   5970
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtPL준우승 
      Height          =   270
      Left            =   4680
      TabIndex        =   58
      Text            =   "Text2"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtPL우승 
      Height          =   270
      Left            =   4680
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtCo 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
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
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
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
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
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
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
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
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
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
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
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
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
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
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
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
      Caption         =   "적용"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.TextBox txt랭크 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4560
      TabIndex        =   39
      Text            =   "Text20"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtP연승 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   38
      Text            =   "Text19"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox TxtP연 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   37
      Text            =   "Text18"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox txtZ연승 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   36
      Text            =   "Text17"
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox txtZ연 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   35
      Text            =   "Text16"
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox txtT연승 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   34
      Text            =   "Text15"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox txtT연 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   33
      Text            =   "Text14"
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox txtA연승 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   32
      Text            =   "Text13"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox TxtA연 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   31
      Text            =   "Text12"
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox txtSevic 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   30
      Text            =   "Text11"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtVic 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   29
      Text            =   "Text10"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtSkill 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   28
      Text            =   "Text9"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtLev 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   27
      Text            =   "Text8"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtEXP 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   26
      Text            =   "Text7"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtPL 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   25
      Text            =   "Text6"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtPW 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   24
      Text            =   "Text5"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtZL 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   23
      Text            =   "Text4"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtZW 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   22
      Text            =   "Text3"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtTL 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtTW 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "PL준우승"
      Height          =   255
      Left            =   3240
      TabIndex        =   60
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "PL우승"
      Height          =   255
      Left            =   3240
      TabIndex        =   59
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
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
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
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
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
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
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
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
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
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
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
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
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
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
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
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
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "랭크"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   3120
      TabIndex        =   19
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "P연승"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   18
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "P연승정보"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   17
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Z연승"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   16
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Z연승정보"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   15
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "T연승"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   14
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "T연승정보"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "A연승"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "A연승정보"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "준우승"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "우승"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
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
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "레벨"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "경험치"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "P패"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "P승"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Z패"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Z승"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "T패"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "T승"
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
txtTW = MyTW(돌려)
txtTL = MyTL(돌려)
txtZW = MyZW(돌려)
txtZL = MyZL(돌려)
txtPW = MyPW(돌려)
txtPL = MyPL(돌려)
txtEXP = MyExp(돌려)
txtLev = MyLev(돌려)
txtVic = MyVic(돌려)
txtSevic = MySeVic(돌려)
TxtA연 = MyA연(돌려)
txtA연승 = MyA연승(돌려)
txtT연 = MyT연(돌려)
txtT연승 = MyT연승(돌려)
txtZ연 = MyZ연(돌려)
txtZ연승 = MyZ연승(돌려)
TxtP연 = MyP연(돌려)
txtP연승 = MyP연승(돌려)
txt랭크 = MyRank(돌려)
txtAt = MyAt(돌려)
txtR = MyR(돌려)
txtSt = MySt(돌려)
txtAm = MyAm(돌려)
txtDe = MyDe(돌려)
txtPa = MyPa(돌려)
txtSe = MySe(돌려)
txtCo = MyCo(돌려)
txtSkill = MySkill(돌려)
txtPL우승 = PL우승
txtPL준우승 = PL준우승
End Sub

Private Sub jcbutton1_Click()
MyTW(돌려) = txtTW
MyTL(돌려) = txtTL
MyZW(돌려) = txtZW
MyZL(돌려) = txtZL
MyPW(돌려) = txtPW
MyPL(돌려) = txtPL
MyAW(돌려) = val(txtTW) + val(txtZW) + val(txtPW)
MyAL(돌려) = val(txtTL) + val(txtZL) + val(txtPL)
MyLev(돌려) = txtLev
MyExp(돌려) = txtEXP
MyVic(돌려) = txtVic
MySeVic(돌려) = txtSevic
MyA연(돌려) = TxtA연
MyA연승(돌려) = txtA연승
MyT연(돌려) = txtT연
MyT연승(돌려) = txtT연승
MyZ연(돌려) = txtZ연
MyZ연승(돌려) = txtZ연승
MyP연(돌려) = TxtP연
MyP연승(돌려) = txtP연승
MyRank(돌려) = txt랭크
MyAt(돌려) = txtAt
MyR(돌려) = txtR
MySt(돌려) = txtSt
MyAm(돌려) = txtAm
MyDe(돌려) = txtDe
MyPa(돌려) = txtPa
MySe(돌려) = txtSe
MyCo(돌려) = txtCo
MySkill(돌려) = txtSkill
PL우승 = txtPL우승
PL준우승 = txtPL준우승
FrmMain.Timer12.Enabled = True
End Sub
