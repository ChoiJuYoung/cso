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
   StartUpPosition =   2  '鉢檎 亜錘汽
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
         Name            =   "閏顕"
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
         Name            =   "閏顕"
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
      Alignment       =   2  '亜錘汽 限茶
      BackStyle       =   0  '燈誤
      Caption         =   "<雁重税 衣舛戚 溌叔杯艦猿?>"
      BeginProperty Font 
         Name            =   "妓崇"
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
      Alignment       =   2  '亜錘汽 限茶
      BackStyle       =   0  '燈誤
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "妓崇"
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
      Alignment       =   2  '亜錘汽 限茶
      BackStyle       =   0  '燈誤
      Caption         =   "恥 走災拝 榎衝,"
      BeginProperty Font 
         Name            =   "妓崇"
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
      BackStyle       =   1  '燈誤馬走 省製
      Height          =   2175
      Left            =   0
      Top             =   4440
      Width           =   6255
   End
   Begin VB.Label lbl姐呪 
      Alignment       =   2  '亜錘汽 限茶
      BackStyle       =   0  '燈誤
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "妓崇"
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
      Alignment       =   2  '亜錘汽 限茶
      BackStyle       =   0  '燈誤
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "妓崇"
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
      Alignment       =   2  '亜錘汽 限茶
      BackStyle       =   0  '燈誤
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "妓崇"
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
      Alignment       =   2  '亜錘汽 限茶
      BackStyle       =   0  '燈誤
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "妓崇"
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
      BackStyle       =   1  '燈誤馬走 省製
      Height          =   2895
      Left            =   0
      Top             =   1560
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  '亜錘汽 限茶
      BackStyle       =   0  '燈誤
      Caption         =   "姥脊拝 朝球苫税 曽嫌人 呪勲聖 溌昔杯艦陥."
      BeginProperty Font 
         Name            =   "妓崇"
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
      BackStyle       =   1  '燈誤馬走 省製
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
If val(姥古) = 1 Then
 Label2 = "<Zerg Card Pack>"
 Label3 = "[Normal]"
 Label4 = "[3000Cro]"
 lblAllPrice = "[3000Cro]"
ElseIf val(姥古) = 2 Then
 Label2 = "<Zerg Card Pack ver.A>"
 Label3 = "[Normal ~ Rare]"
 Label4 = "[7000Cro]"
 lblAllPrice = "[7000Cro]"
ElseIf val(姥古) = 3 Then
 Label2 = "<Zerg Card Pack ver.S"
 Label3 = "[Normal ~ Elite]"
 Label4 = "[10000Cro]"
 lblAllPrice = "[10000Cro]"
ElseIf val(姥古) = 4 Then
 Label2 = "<Terran Card Pack>"
 Label3 = "[Normal]"
 Label4 = "[3000Cro]"
 lblAllPrice = "[3000Cro]"
ElseIf val(姥古) = 5 Then
 Label2 = "<Terran Card Pack ver.A>"
 Label3 = "[Normal ~ Rare]"
 Label4 = "[7000Cro]"
 lblAllPrice = "[7000Cro]"
ElseIf val(姥古) = 6 Then
 Label2 = "<Terran Card Pack ver.S>"
 Label3 = "[Normal ~ Elite]"
 Label4 = "[10000Cro]"
 lblAllPrice = "[10000Cro]"
ElseIf val(姥古) = 7 Then
 Label2 = "<Protoss Card Pack>"
 Label3 = "[Normal]"
 Label4 = "[3000Cro]"
 lblAllPrice = "[3000Cro]"
ElseIf val(姥古) = 8 Then
 Label2 = "<Protoss Card Pack ver.A>"
 Label3 = "[Normal ~ Rare]"
 Label4 = "[7000Cro]"
 lblAllPrice = "[7000Cro]"
ElseIf val(姥古) = 9 Then
 Label2 = "<Protoss Card Pack ver.S>"
 Label3 = "[Normal ~ Elite]"
 Label4 = "[10000Cro]"
 lblAllPrice = "[10000Cro]"
End If
lbl姐呪 = "<呪勲 : 1>"
End Sub

Private Sub jcbutton1_Click()
雌繊NPC = Int((714 * Rnd) + 1)
If val(姥古) = 1 Then
 If val(Money) >= 3000 Then
  Do Until (粂滴(雌繊NPC) = "Normal") And 曽膳(雌繊NPC) = 2
    雌繊NPC = Int((800 * Rnd) + 1)
  Loop
   姥古亜管 = "Yes"
   Money = val(Money) - 3000
 Else
  MsgBox "儀戚 採膳杯艦陥 ばば"
 End If
ElseIf val(姥古) = 2 Then
 If val(Money) >= 7000 Then
  Do Until (粂滴(雌繊NPC) = "Normal" Or 粂滴(雌繊NPC) = "Special" Or 粂滴(雌繊NPC) = "Rare") And 曽膳(雌繊NPC) = 2
    雌繊NPC = Int((800 * Rnd) + 1)
  Loop
   姥古亜管 = "Yes"
   Money = val(Money) - 7000
 Else
  MsgBox "儀戚 採膳杯艦陥 ばば"
 End If
ElseIf val(姥古) = 3 Then
 If val(Money) >= 10000 Then
  If Mode <> "Normal" Then
   Do Until (粂滴(雌繊NPC) = "Normal" Or 粂滴(雌繊NPC) = "Special" Or 粂滴(雌繊NPC) = "Rare" Or 粂滴(雌繊NPC) = "Unique" Or 粂滴(雌繊NPC) = "Elite") And (曽膳(雌繊NPC) = 2)
    雌繊NPC = Int((723 * Rnd) + 1)
   Loop
  Else
   Do Until (粂滴(雌繊NPC) = "Normal" Or 粂滴(雌繊NPC) = "Special" Or 粂滴(雌繊NPC) = "Rare" Or 粂滴(雌繊NPC) = "Unique") And (曽膳(雌繊NPC) = 2)
    雌繊NPC = Int((723 * Rnd) + 1)
   Loop
  End If
   姥古亜管 = "Yes"
   Money = val(Money) - 10000
 Else
  MsgBox "儀戚 採膳杯艦陥 ばば"
 End If
ElseIf val(姥古) = 4 Then
 If val(Money) >= 3000 Then
  Do Until (粂滴(雌繊NPC) = "Normal") And 曽膳(雌繊NPC) = 1
    雌繊NPC = Int((800 * Rnd) + 1)
  Loop
   姥古亜管 = "Yes"
   Money = val(Money) - 3000
 Else
  MsgBox "儀戚 採膳杯艦陥 ばば"
 End If
ElseIf val(姥古) = 5 Then
 If val(Money) >= 7000 Then
  Do Until (粂滴(雌繊NPC) = "Normal" Or 粂滴(雌繊NPC) = "Special" Or 粂滴(雌繊NPC) = "Rare") And 曽膳(雌繊NPC) = 1
    雌繊NPC = Int((800 * Rnd) + 1)
  Loop
   姥古亜管 = "Yes"
   Money = val(Money) - 7000
 Else
  MsgBox "儀戚 採膳杯艦陥 ばば"
 End If
ElseIf val(姥古) = 6 Then
 If val(Money) >= 10000 Then
  If Mode = "Normal" Then
   Do Until (粂滴(雌繊NPC) = "Normal" Or 粂滴(雌繊NPC) = "Special" Or 粂滴(雌繊NPC) = "Rare" Or 粂滴(雌繊NPC) = "Unique") And (曽膳(雌繊NPC) = 1)
    雌繊NPC = Int((723 * Rnd) + 1)
   Loop
  Else
   Do Until (粂滴(雌繊NPC) = "Normal" Or 粂滴(雌繊NPC) = "Special" Or 粂滴(雌繊NPC) = "Rare" Or 粂滴(雌繊NPC) = "Unique" Or 粂滴(雌繊NPC) = "Elite") And (曽膳(雌繊NPC) = 1)
    雌繊NPC = Int((723 * Rnd) + 1)
   Loop
  End If
   姥古亜管 = "Yes"
   Money = val(Money) - 10000
 Else
  MsgBox "儀戚 採膳杯艦陥 ばば"
 End If
ElseIf val(姥古) = 7 Then
 If val(Money) >= 3000 Then
  Do Until (粂滴(雌繊NPC) = "Normal") And 曽膳(雌繊NPC) = 3
    雌繊NPC = Int((800 * Rnd) + 1)
  Loop
   姥古亜管 = "Yes"
   Money = val(Money) - 3000
 Else
  MsgBox "儀戚 採膳杯艦陥 ばば"
 End If
ElseIf val(姥古) = 8 Then
 If val(Money) >= 7000 Then
  Do Until (粂滴(雌繊NPC) = "Normal" Or 粂滴(雌繊NPC) = "Special" Or 粂滴(雌繊NPC) = "Rare") And 曽膳(雌繊NPC) = 3
    雌繊NPC = Int((800 * Rnd) + 1)
  Loop
   姥古亜管 = "Yes"
   Money = val(Money) - 7000
 Else
  MsgBox "儀戚 採膳杯艦陥 ばば"
 End If
ElseIf val(姥古) = 9 Then
 If val(Money) >= 10000 Then
  If Mode = "Normal" Then
   Do Until (粂滴(雌繊NPC) = "Normal" Or 粂滴(雌繊NPC) = "Special" Or 粂滴(雌繊NPC) = "Rare" Or 粂滴(雌繊NPC) = "Unique") And (曽膳(雌繊NPC) = 3)
    雌繊NPC = Int((723 * Rnd) + 1)
   Loop
  Else
   Do Until (粂滴(雌繊NPC) = "Normal" Or 粂滴(雌繊NPC) = "Special" Or 粂滴(雌繊NPC) = "Rare" Or 粂滴(雌繊NPC) = "Unique" Or 粂滴(雌繊NPC) = "Elite") And (曽膳(雌繊NPC) = 3)
    雌繊NPC = Int((723 * Rnd) + 1)
   Loop
  End If
   姥古亜管 = "Yes"
   Money = val(Money) - 10000
 Else
  MsgBox "儀戚 採膳杯艦陥 ばば"
 End If
End If

Dim 雌繊NPC左繕 As Integer
雌繊NPC左繕 = val(識呪呪) - 5
If 姥古亜管 = "Yes" Then
 SubName(雌繊NPC左繕) = 戚硯(雌繊NPC)
 SubTeam(雌繊NPC左繕) = Team(雌繊NPC)
 SubAt(雌繊NPC左繕) = NPC因維径(雌繊NPC)
 SubR(雌繊NPC左繕) = NPC胃薦(雌繊NPC)
 SubSt(雌繊NPC左繕) = NPC穿繰(雌繊NPC)
 SubAm(雌繊NPC左繕) = NPC弘勲(雌繊NPC)
 SubDe(雌繊NPC左繕) = NPC呪搾径(雌繊NPC)
 SubPa(雌繊NPC左繕) = NPC舛茸(雌繊NPC)
 SubSe(雌繊NPC左繕) = NPC湿什(雌繊NPC)
 SubCo(雌繊NPC左繕) = NPC珍闘継(雌繊NPC)
 SubRank(雌繊NPC左繕) = 粂滴(雌繊NPC)
 SubYear(雌繊NPC左繕) = OYear(雌繊NPC)
 SubTribe(雌繊NPC左繕) = 曽膳(雌繊NPC)
 SubLev(雌繊NPC左繕) = 1
 SubExp(雌繊NPC左繕) = 0
 SubMExp(雌繊NPC左繕) = 50
 SubPoint(雌繊NPC左繕) = 0
 SubNum(雌繊NPC左繕) = val(雌繊NPC)
 SubAW(雌繊NPC左繕) = 0
 SubAL(雌繊NPC左繕) = 0
 SubTW(雌繊NPC左繕) = 0
 SubTL(雌繊NPC左繕) = 0
 SubZW(雌繊NPC左繕) = 0
 SubZL(雌繊NPC左繕) = 0
 SubPW(雌繊NPC左繕) = 0
 SubPL(雌繊NPC左繕) = 0
 SubVic(雌繊NPC左繕) = 0
 SubSeVic(雌繊NPC左繕) = 0
 SubCode(雌繊NPC左繕) = "B"
 SubSkill(雌繊NPC左繕) = Skill(雌繊NPC)
 識呪呪 = val(識呪呪) + 1
 姥古亜管 = "No"
 
If SubRank(雌繊NPC左繕) = "Normal" Or SubRank(雌繊NPC左繕) = "Special" Then
 SubNW(雌繊NPC左繕) = "CB16"
ElseIf SubRank(雌繊NPC左繕) = "Rare" Then
 SubNW(雌繊NPC左繕) = "CA1"
ElseIf SubRank(雌繊NPC左繕) = "Unique" Then
 SubNW(雌繊NPC左繕) = "CA2"
ElseIf SubRank(雌繊NPC左繕) = "Elite" Then
 SubNW(雌繊NPC左繕) = "CA3"
Else
 SubNW(雌繊NPC左繕) = "CS32"
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

