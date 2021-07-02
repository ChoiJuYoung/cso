VERSION 5.00
Begin VB.Form FrmHighShop 
   BackColor       =   &H00FFFFFF&
   Caption         =   "고급상점"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Form19.frx":0000
   LinkTopic       =   "Form19"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  '화면 가운데
   Begin CSO.jcbutton jcbutton1 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "뽑기"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin CSO.jcbutton jcbutton2 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "뽑기"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Label Label7 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "구입 하시겠습니까?"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   4695
   End
   Begin VB.Label Label6 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[100000Cro]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   4695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   4  '대시-점
      X1              =   0
      X2              =   4680
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[Unique ~ Legend]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Width           =   4695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   4  '대시-점
      X1              =   0
      X2              =   4680
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "구입 하시겠습니까??"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[50000Cro]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[Normal ~ Legend]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "고급 상점에 오신것을 환영합니다. 두가지의 메뉴가 존재합니다. 무엇을 드시겠습니까?"
      BeginProperty Font 
         Name            =   "돋움"
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
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "FrmHighShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
If Mode = "Hell" Then
    Label2 = "[Normal ~ Champion]"
    Label5 = "[Unique ~ Champion]"
End If
End Sub

Private Sub jcbutton1_Click()
Randomize Oee
Oee = Int((801 * Rnd) + 0)
 If val(Money) >= 100000 Then
  If Mode = "Hard" Then
   Do Until (랭크(상점NPC) = "Unique") Or (랭크(상점NPC) = "Elite") Or (랭크(상점NPC) = "Legend")
    상점NPC = Int((800 * Rnd) + 1)
   Loop
  Else
   Do Until (랭크(상점NPC) = "Unique") Or (랭크(상점NPC) = "Elite") Or (랭크(상점NPC) = "Legend") Or (랭크(상점NPC) = "Secret") Or (랭크(상점NPC) = "Champion")
    상점NPC = Int((801 * Rnd) + 0)
   Loop
  End If
   구매가능 = "Yes"
   Money = val(Money) - 100000
 Else
  MsgBox "돈이 부족합니다. 10만원을 모아서 오십시오."
 End If
 
Dim 상점NPC보조 As Integer
상점NPC보조 = val(선수수) - 5
If 구매가능 = "Yes" Then
 SubName(상점NPC보조) = 이름(상점NPC)
 SubTeam(상점NPC보조) = Team(상점NPC)
 SubAt(상점NPC보조) = NPC공격력(상점NPC)
 SubR(상점NPC보조) = NPC견제(상점NPC)
 SubSt(상점NPC보조) = NPC전략(상점NPC)
 SubAm(상점NPC보조) = NPC물량(상점NPC)
 SubDe(상점NPC보조) = NPC수비력(상점NPC)
 SubPa(상점NPC보조) = NPC정찰(상점NPC)
 SubSe(상점NPC보조) = NPC센스(상점NPC)
 SubCo(상점NPC보조) = NPC컨트롤(상점NPC)
 SubRank(상점NPC보조) = 랭크(상점NPC)
 SubYear(상점NPC보조) = OYear(상점NPC)
 SubTribe(상점NPC보조) = 종족(상점NPC)
 SubLev(상점NPC보조) = 1
 SubExp(상점NPC보조) = 0
 SubMExp(상점NPC보조) = 50
 SubPoint(상점NPC보조) = 0
 SubNum(상점NPC보조) = val(상점NPC)
 SubAW(상점NPC보조) = 0
 SubAL(상점NPC보조) = 0
 SubTW(상점NPC보조) = 0
 SubTL(상점NPC보조) = 0
 SubZW(상점NPC보조) = 0
 SubZL(상점NPC보조) = 0
 SubPW(상점NPC보조) = 0
 SubPL(상점NPC보조) = 0
 SubVic(상점NPC보조) = 0
 SubSeVic(상점NPC보조) = 0
 SubCode(상점NPC보조) = "B"
 SubSkill(상점NPC보조) = Skill(상점NPC)
 선수수 = val(선수수) + 1
 구매가능 = "No"
 
If SubRank(상점NPC보조) = "Normal" Or SubRank(상점NPC보조) = "Special" Then
 SubNW(상점NPC보조) = "CB16"
ElseIf SubRank(상점NPC보조) = "Rare" Then
 SubNW(상점NPC보조) = "CA1"
ElseIf SubRank(상점NPC보조) = "Unique" Then
 SubNW(상점NPC보조) = "CA2"
ElseIf SubRank(상점NPC보조) = "Elite" Then
 SubNW(상점NPC보조) = "CA3"
Else
 SubNW(상점NPC보조) = "CS32"
End If

If 9.Text1 = "이히히" Then
    9.Text1 = "히히히"
Else
    9.Text1 = "이히히"
End If
 Unload FrmShop
 Unload Me
End If
End Sub

Private Sub jcbutton2_Click()
Randomize Oee
Oee = Int((801 * Rnd) + 0)
 If val(Money) >= 50000 Then
  If Mode = "Hard" Then
   Do Until (랭크(상점NPC) = "Normal") Or (랭크(상점NPC) = "Special") Or (랭크(상점NPC) = "Rare") Or (랭크(상점NPC) = "Unique") Or (랭크(상점NPC) = "Elite") Or (랭크(상점NPC) = "Legend")
    상점NPC = Int((800 * Rnd) + 1)
   Loop
  Else
   Do Until (랭크(상점NPC) = "Normal") Or (랭크(상점NPC) = "Special") Or (랭크(상점NPC) = "Rare") Or (랭크(상점NPC) = "Unique") Or (랭크(상점NPC) = "Elite") Or (랭크(상점NPC) = "Legend") Or (랭크(상점NPC) = "Secret") Or (랭크(상점NPC) = "Chapmion")
    상점NPC = Int((801 * Rnd) + 0)
   Loop
  End If
   구매가능 = "Yes"
   Money = val(Money) - 50000
 Else
  MsgBox "돈이 부족합니다. 5만원을 모아서 오십시오."
 End If
 
Dim 상점NPC보조 As Integer
상점NPC보조 = val(선수수) - 5
If 구매가능 = "Yes" Then
 SubName(상점NPC보조) = 이름(상점NPC)
 SubTeam(상점NPC보조) = Team(상점NPC)
 SubAt(상점NPC보조) = NPC공격력(상점NPC)
 SubR(상점NPC보조) = NPC견제(상점NPC)
 SubSt(상점NPC보조) = NPC전략(상점NPC)
 SubAm(상점NPC보조) = NPC물량(상점NPC)
 SubDe(상점NPC보조) = NPC수비력(상점NPC)
 SubPa(상점NPC보조) = NPC정찰(상점NPC)
 SubSe(상점NPC보조) = NPC센스(상점NPC)
 SubCo(상점NPC보조) = NPC컨트롤(상점NPC)
 SubRank(상점NPC보조) = 랭크(상점NPC)
 SubYear(상점NPC보조) = OYear(상점NPC)
 SubTribe(상점NPC보조) = 종족(상점NPC)
 SubLev(상점NPC보조) = 1
 SubExp(상점NPC보조) = 0
 SubMExp(상점NPC보조) = 50
 SubPoint(상점NPC보조) = 0
 SubNum(상점NPC보조) = val(상점NPC)
 SubAW(상점NPC보조) = 0
 SubAL(상점NPC보조) = 0
 SubTW(상점NPC보조) = 0
 SubTL(상점NPC보조) = 0
 SubZW(상점NPC보조) = 0
 SubZL(상점NPC보조) = 0
 SubPW(상점NPC보조) = 0
 SubPL(상점NPC보조) = 0
 SubVic(상점NPC보조) = 0
 SubSeVic(상점NPC보조) = 0
 SubCode(상점NPC보조) = "B"
 SubSkill(상점NPC보조) = Skill(상점NPC)
 선수수 = val(선수수) + 1
 구매가능 = "No"
If SubRank(상점NPC보조) = "Normal" Or SubRank(상점NPC보조) = "Special" Then
 SubNW(상점NPC보조) = "CB16"
ElseIf SubRank(상점NPC보조) = "Rare" Then
 SubNW(상점NPC보조) = "CA1"
ElseIf SubRank(상점NPC보조) = "Unique" Then
 SubNW(상점NPC보조) = "CA2"
ElseIf SubRank(상점NPC보조) = "Elite" Then
 SubNW(상점NPC보조) = "CA3"
Else
 SubNW(상점NPC보조) = "CS32"
End If
If 9.Text1 = "이히히" Then
    9.Text1 = "히히히"
Else
    9.Text1 = "이히히"
End If
 Unload FrmShop
 Unload Me
End If
End Sub
