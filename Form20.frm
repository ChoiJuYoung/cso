VERSION 5.00
Begin VB.Form FrmShopConf 
   Caption         =   "Confirm"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   Icon            =   "Form20.frx":0000
   LinkTopic       =   "Form20"
   ScaleHeight     =   6615
   ScaleWidth      =   6240
   StartUpPosition =   2  '화면 가운데
   Begin CSO.jcbutton jcbutton2 
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   6240
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   255
      Caption         =   "No"
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin CSO.jcbutton jcbutton1 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   6240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16711680
      Caption         =   "Yes"
      ForeColor       =   16777215
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin VB.Label Label8 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "<당신의 결정이 확실합니까?>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   5880
      Width           =   6255
   End
   Begin VB.Label lblAllPrice 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   5400
      Width           =   6255
   End
   Begin VB.Label Label6 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "총 지불할 금액,"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   4920
      Width           =   6255
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  '투명하지 않음
      Height          =   2175
      Left            =   0
      Top             =   4440
      Width           =   6255
   End
   Begin VB.Label lbl갯수 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3720
      Width           =   6255
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   6255
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   6255
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   6255
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '투명하지 않음
      Height          =   2895
      Left            =   0
      Top             =   1560
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "구입할 카드팩의 종류와 수량을 확인합니다."
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      Height          =   1575
      Left            =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "FrmShopConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
If val(구매) = 1 Then
 Label2 = "<Zerg Card Pack>"
 Label3 = "[Normal]"
 Label4 = "[3000Cro]"
 lblAllPrice = "[3000Cro]"
ElseIf val(구매) = 2 Then
 Label2 = "<Zerg Card Pack ver.A>"
 Label3 = "[Normal ~ Rare]"
 Label4 = "[7000Cro]"
 lblAllPrice = "[7000Cro]"
ElseIf val(구매) = 3 Then
 Label2 = "<Zerg Card Pack ver.S"
 Label3 = "[Normal ~ Elite]"
 Label4 = "[10000Cro]"
 lblAllPrice = "[10000Cro]"
ElseIf val(구매) = 4 Then
 Label2 = "<Terran Card Pack>"
 Label3 = "[Normal]"
 Label4 = "[3000Cro]"
 lblAllPrice = "[3000Cro]"
ElseIf val(구매) = 5 Then
 Label2 = "<Terran Card Pack ver.A>"
 Label3 = "[Normal ~ Rare]"
 Label4 = "[7000Cro]"
 lblAllPrice = "[7000Cro]"
ElseIf val(구매) = 6 Then
 Label2 = "<Terran Card Pack ver.S>"
 Label3 = "[Normal ~ Elite]"
 Label4 = "[10000Cro]"
 lblAllPrice = "[10000Cro]"
ElseIf val(구매) = 7 Then
 Label2 = "<Protoss Card Pack>"
 Label3 = "[Normal]"
 Label4 = "[3000Cro]"
 lblAllPrice = "[3000Cro]"
ElseIf val(구매) = 8 Then
 Label2 = "<Protoss Card Pack ver.A>"
 Label3 = "[Normal ~ Rare]"
 Label4 = "[7000Cro]"
 lblAllPrice = "[7000Cro]"
ElseIf val(구매) = 9 Then
 Label2 = "<Protoss Card Pack ver.S>"
 Label3 = "[Normal ~ Elite]"
 Label4 = "[10000Cro]"
 lblAllPrice = "[10000Cro]"
End If
lbl갯수 = "<수량 : 1>"
End Sub

Private Sub jcbutton1_Click()
상점NPC = Int((714 * Rnd) + 1)
If val(구매) = 1 Then
 If val(Money) >= 3000 Then
  Do Until (랭크(상점NPC) = "Normal") And 종족(상점NPC) = 2
    상점NPC = Int((800 * Rnd) + 1)
  Loop
   구매가능 = "Yes"
   Money = val(Money) - 3000
 Else
  MsgBox "돈이 부족합니다 ㅠㅠ"
 End If
ElseIf val(구매) = 2 Then
 If val(Money) >= 7000 Then
  Do Until (랭크(상점NPC) = "Normal" Or 랭크(상점NPC) = "Special" Or 랭크(상점NPC) = "Rare") And 종족(상점NPC) = 2
    상점NPC = Int((800 * Rnd) + 1)
  Loop
   구매가능 = "Yes"
   Money = val(Money) - 7000
 Else
  MsgBox "돈이 부족합니다 ㅠㅠ"
 End If
ElseIf val(구매) = 3 Then
 If val(Money) >= 10000 Then
  If Mode <> "Normal" Then
   Do Until (랭크(상점NPC) = "Normal" Or 랭크(상점NPC) = "Special" Or 랭크(상점NPC) = "Rare" Or 랭크(상점NPC) = "Unique" Or 랭크(상점NPC) = "Elite") And (종족(상점NPC) = 2)
    상점NPC = Int((723 * Rnd) + 1)
   Loop
  Else
   Do Until (랭크(상점NPC) = "Normal" Or 랭크(상점NPC) = "Special" Or 랭크(상점NPC) = "Rare" Or 랭크(상점NPC) = "Unique") And (종족(상점NPC) = 2)
    상점NPC = Int((723 * Rnd) + 1)
   Loop
  End If
   구매가능 = "Yes"
   Money = val(Money) - 10000
 Else
  MsgBox "돈이 부족합니다 ㅠㅠ"
 End If
ElseIf val(구매) = 4 Then
 If val(Money) >= 3000 Then
  Do Until (랭크(상점NPC) = "Normal") And 종족(상점NPC) = 1
    상점NPC = Int((800 * Rnd) + 1)
  Loop
   구매가능 = "Yes"
   Money = val(Money) - 3000
 Else
  MsgBox "돈이 부족합니다 ㅠㅠ"
 End If
ElseIf val(구매) = 5 Then
 If val(Money) >= 7000 Then
  Do Until (랭크(상점NPC) = "Normal" Or 랭크(상점NPC) = "Special" Or 랭크(상점NPC) = "Rare") And 종족(상점NPC) = 1
    상점NPC = Int((800 * Rnd) + 1)
  Loop
   구매가능 = "Yes"
   Money = val(Money) - 7000
 Else
  MsgBox "돈이 부족합니다 ㅠㅠ"
 End If
ElseIf val(구매) = 6 Then
 If val(Money) >= 10000 Then
  If Mode = "Normal" Then
   Do Until (랭크(상점NPC) = "Normal" Or 랭크(상점NPC) = "Special" Or 랭크(상점NPC) = "Rare" Or 랭크(상점NPC) = "Unique") And (종족(상점NPC) = 1)
    상점NPC = Int((723 * Rnd) + 1)
   Loop
  Else
   Do Until (랭크(상점NPC) = "Normal" Or 랭크(상점NPC) = "Special" Or 랭크(상점NPC) = "Rare" Or 랭크(상점NPC) = "Unique" Or 랭크(상점NPC) = "Elite") And (종족(상점NPC) = 1)
    상점NPC = Int((723 * Rnd) + 1)
   Loop
  End If
   구매가능 = "Yes"
   Money = val(Money) - 10000
 Else
  MsgBox "돈이 부족합니다 ㅠㅠ"
 End If
ElseIf val(구매) = 7 Then
 If val(Money) >= 3000 Then
  Do Until (랭크(상점NPC) = "Normal") And 종족(상점NPC) = 3
    상점NPC = Int((800 * Rnd) + 1)
  Loop
   구매가능 = "Yes"
   Money = val(Money) - 3000
 Else
  MsgBox "돈이 부족합니다 ㅠㅠ"
 End If
ElseIf val(구매) = 8 Then
 If val(Money) >= 7000 Then
  Do Until (랭크(상점NPC) = "Normal" Or 랭크(상점NPC) = "Special" Or 랭크(상점NPC) = "Rare") And 종족(상점NPC) = 3
    상점NPC = Int((800 * Rnd) + 1)
  Loop
   구매가능 = "Yes"
   Money = val(Money) - 7000
 Else
  MsgBox "돈이 부족합니다 ㅠㅠ"
 End If
ElseIf val(구매) = 9 Then
 If val(Money) >= 10000 Then
  If Mode = "Normal" Then
   Do Until (랭크(상점NPC) = "Normal" Or 랭크(상점NPC) = "Special" Or 랭크(상점NPC) = "Rare" Or 랭크(상점NPC) = "Unique") And (종족(상점NPC) = 3)
    상점NPC = Int((723 * Rnd) + 1)
   Loop
  Else
   Do Until (랭크(상점NPC) = "Normal" Or 랭크(상점NPC) = "Special" Or 랭크(상점NPC) = "Rare" Or 랭크(상점NPC) = "Unique" Or 랭크(상점NPC) = "Elite") And (종족(상점NPC) = 3)
    상점NPC = Int((723 * Rnd) + 1)
   Loop
  End If
   구매가능 = "Yes"
   Money = val(Money) - 10000
 Else
  MsgBox "돈이 부족합니다 ㅠㅠ"
 End If
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


 Call Save
 FrmMain.Timer12.Enabled = True
 Unload FrmShop
 Unload Me
End If
End Sub

Private Sub jcbutton2_Click()
Unload FrmShop
Unload Me
End Sub

