VERSION 5.00
Begin VB.Form FrmAbility 
   Caption         =   "Ability"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5880
   Icon            =   "Form13.frx":0000
   LinkTopic       =   "Form13"
   ScaleHeight     =   5400
   ScaleWidth      =   5880
   StartUpPosition =   2  '화면 가운데
   Begin CSO.jcbutton jcbutton1 
      Height          =   375
      Left            =   1320
      TabIndex        =   18
      Top             =   5040
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "스탯분배"
      CaptionEffects  =   0
   End
   Begin CSO.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2160
      TabIndex        =   19
      Top             =   5040
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      Value           =   0
      Theme           =   5
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
      Text            =   "경험치 %"
   End
   Begin VB.Label lblLeDe 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Label2"
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
      TabIndex        =   24
      Top             =   3480
      Width           =   5895
   End
   Begin VB.Label lbl연승 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Label2"
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
      TabIndex        =   23
      Top             =   2640
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Label1"
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
      TabIndex        =   22
      Top             =   3240
      Width           =   5895
   End
   Begin VB.Line Line7 
      X1              =   3840
      X2              =   3840
      Y1              =   3840
      Y2              =   5040
   End
   Begin VB.Line Line6 
      X1              =   1800
      X2              =   1800
      Y1              =   3840
      Y2              =   5040
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   4560
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   5880
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   5880
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5880
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      BackStyle       =   0  '투명
      Caption         =   "포인트 :"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label15 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BackStyle       =   0  '투명
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   20
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lblP 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "vsP : 0승 0패"
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
      Left            =   3840
      TabIndex        =   17
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label lblZ 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "vsZ : 0승 0패"
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
      Left            =   1800
      TabIndex        =   16
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label lblT 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "vsT : 0승 0패"
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
      TabIndex        =   15
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblA 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "vsA : 0승 0패"
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
      Left            =   3840
      TabIndex        =   14
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label lblVic 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "특이사항 : 없음"
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
      TabIndex        =   13
      Top             =   3000
      Width           =   5895
   End
   Begin VB.Label lblTri 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Tribe : Protoss"
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
      TabIndex        =   12
      Top             =   2400
      Width           =   5895
   End
   Begin VB.Label lblNa 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Name : 김가을"
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
      TabIndex        =   11
      Top             =   2160
      Width           =   5895
   End
   Begin VB.Image Img 
      Height          =   1500
      Left            =   2160
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label lblSe 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "센   스 : 1000"
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
      Left            =   1800
      TabIndex        =   10
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label lblAm 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "물   량 : 1000"
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
      TabIndex        =   9
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label lblTeam 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "소속팀 : 삼성전자"
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
      Left            =   3840
      TabIndex        =   8
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label lblR 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "견   제 : 1000"
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
      Left            =   1800
      TabIndex        =   7
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label lblPa 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "정   찰 : 1000"
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
      TabIndex        =   6
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lblExp 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "경험치 : 0 %"
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
      Left            =   3840
      TabIndex        =   5
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label lblCo 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "컨트롤 : 1000"
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
      Left            =   1800
      TabIndex        =   4
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label lblDe 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "수비력 : 1000"
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
      TabIndex        =   3
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lblLv 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Level : 1"
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
      Left            =   3840
      TabIndex        =   2
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lblSt 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "전   략 : 1000"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lblAt 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "공격력 : 1000"
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
      TabIndex        =   0
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00EAE2F1&
      BackStyle       =   1  '투명하지 않음
      Height          =   1575
      Left            =   0
      Top             =   3840
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00AE73E5&
      BackStyle       =   1  '투명하지 않음
      Height          =   3855
      Left            =   0
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "FrmAbility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
 If MyNW(선택) = "CB16" Then
  lblLeDe = "Code B 16강"
 ElseIf MyNW(선택) = "CB8" Then
  lblLeDe = "Code B 8강"
 ElseIf MyNW(선택) = "CB4" Then
  lblLeDe = "Code B 4강"
 ElseIf MyNW(선택) = "CBFin" Then
  lblLeDe = "Code B 결승전"
 ElseIf MyNW(선택) = "CA1" Then
  lblLeDe = "Code A 1Round"
 ElseIf MyNW(선택) = "CA2" Then
  lblLeDe = "Code A 2Round"
 ElseIf MyNW(선택) = "CA3" Then
  lblLeDe = "Code A 3Round"
 ElseIf MyNW(선택) = "CS32" Then
  lblLeDe = "Code S 32강"
 ElseIf MyNW(선택) = "CS16" Then
  lblLeDe = "Code S 16강"
 ElseIf MyNW(선택) = "CS8" Then
  lblLeDe = "Code S 8강"
 ElseIf MyNW(선택) = "CS4" Then
  lblLeDe = "Code S 4강"
 ElseIf MyNW(선택) = "CSFin" Then
  lblLeDe = "Code S 결승전"
 ElseIf MyNW(선택) = "UpADo" Then
  lblLeDe = "승격 강등전"
 End If
 
Dim 돌려 As Integer
For 돌려 = 1 To 6
 MyMExp(돌려) = val(MyLev(돌려)) * 50
Next
lbl연승 = MyA연승(선택)
If MyA연(선택) = "W" Then
 lbl연승 = "연승 : " & lbl연승 & "연승"
Else
 lbl연승 = "연패 : " & lbl연승 & "연패"
End If

If MySkill(선택) = 1 Then
    Label1 = "스킬 :皇制"
ElseIf MySkill(선택) = 2 Then
    Label1 = "스킬 :最終兵器"
ElseIf MySkill(선택) = 3 Then
    Label1 = "스킬 :暴風"
ElseIf MySkill(선택) = 4 Then
    Label1 = "스킬 :英雄"
ElseIf MySkill(선택) = 5 Then
    Label1 = "스킬 :天才"
ElseIf MySkill(선택) = 6 Then
    Label1 = "스킬 :Maestro"
ElseIf MySkill(선택) = 7 Then
    Label1 = "스킬 :怪物"
ElseIf MySkill(선택) = 8 Then
    Label1 = "스킬 :飛龍"
ElseIf MySkill(선택) = 9 Then
    Label1 = "스킬 :恐龍"
ElseIf MySkill(선택) = 10 Then
    Label1 = "스킬 :赤龍"
ElseIf MySkill(선택) = 11 Then
    Label1 = "스킬 :雲龍"
ElseIf MySkill(선택) = 12 Then
    Label1 = "스킬 :怪龍"
ElseIf MySkill(선택) = 13 Then
    Label1 = "스킬 :雷龍"
ElseIf MySkill(선택) = 14 Then
    Label1 = "스킬 :國本"
ElseIf MySkill(선택) = 15 Then
    Label1 = "스킬 :鬪神"
ElseIf MySkill(선택) = 16 Then
    Label1 = "스킬 :暴君"
ElseIf MySkill(선택) = 17 Then
    Label1 = "스킬 :大人"
ElseIf MySkill(선택) = 18 Then
    Label1 = "스킬 :死神"
ElseIf MySkill(선택) = 19 Then
    Label1 = "스킬 :牧童"
ElseIf MySkill(선택) = 20 Then
    Label1 = "스킬 :女制"
ElseIf MySkill(선택) = 21 Then
    Label1 = "스킬 :rEd sNipeR"
ElseIf MySkill(선택) = 22 Then
    Label1 = "스킬 :不死鳥"
ElseIf MySkill(선택) = 23 Then
    Label1 = "스킬 :Sun"
ElseIf MySkill(선택) = 24 Then
    Label1 = "스킬 :pErfecT tErraN"
ElseIf MySkill(선택) = 25 Then
    Label1 = "스킬 :Brain"
ElseIf MySkill(선택) = 26 Then
    Label1 = "스킬 :zErg sPeicaL kILLeR"
ElseIf MySkill(선택) = 27 Then
    Label1 = "스킬 :어린왕자"
ElseIf MySkill(선택) = 28 Then
    Label1 = "스킬 :鐵壁"
ElseIf MySkill(선택) = 29 Then
    Label1 = "스킬 :黑雲長"
ElseIf MySkill(선택) = 30 Then
    Label1 = "스킬 :夢想家"
ElseIf MySkill(선택) = 31 Then
    Label1 = "스킬 :공격력上"
ElseIf MySkill(선택) = 32 Then
    Label1 = "스킬 :견제上"
ElseIf MySkill(선택) = 33 Then
    Label1 = "스킬 :전략上"
ElseIf MySkill(선택) = 34 Then
    Label1 = "스킬 :물량上"
ElseIf MySkill(선택) = 35 Then
    Label1 = "스킬 :수비력上"
ElseIf MySkill(선택) = 36 Then
    Label1 = "스킬 :정찰上"
ElseIf MySkill(선택) = 37 Then
    Label1 = "스킬 :센스上"
ElseIf MySkill(선택) = 38 Then
    Label1 = "스킬 :컨트롤上"
Else
    Label1 = "스킬 :없음"
End If


Label15 = MyPoint(선택)
For 돌려 = 1 To 9
 SubMExp(돌려) = val(SubMExp(돌려)) * 50
Next
lblVic = "우승 : " & MyVic(선택) & "   " & "준우승 : " & MySeVic(선택)
lblAt = "공격력 : " & MyAt(선택)
lblSt = "전   략 : " & MySt(선택)
lblAm = "물   량 : " & MyAm(선택)
lblR = "견   제 : " & MyR(선택)
lblDe = "수비력 : " & MyDe(선택)
lblPa = "정   찰 : " & MyPa(선택)
lblCo = "컨트롤 : " & MyCo(선택)
lblSe = "센   스 : " & MySe(선택)
lblNa = "Name : " & MyYear(선택) & MyName(선택)
lblTeam = "소속팀 : " & MyTeam(선택)
lblLV = "Level : " & MyLev(선택)
lblExp = "경험치 : " & Int(val(MyExp(선택)) * 100 / val(MyMExp(선택))) & " %"
lblA = "vsA : " + MyAW(선택) + "승 " + MyAL(선택) + "패"
lblT = "vsT : " + MyTW(선택) + "승 " + MyTL(선택) + "패"
lblZ = "vsZ : " + MyZW(선택) + "승 " + MyZL(선택) + "패"
lblP = "vsP : " + MyPW(선택) + "승 " + MyPL(선택) + "패"
If MyTribe(선택) = 1 Then
lblTri = "Tribe : Terran"
ElseIf MyTribe(선택) = 2 Then
lblTri = "Tribe : Zerg"
ElseIf MyTribe(선택) = 3 Then
lblTri = "Tribe : Protoss"
End If
lblSpe = "특이사항 : 없음"
If Len(Dir(App.Path & "\img\선수\[" & Mid(MyYear(선택), 2, 2) & "]" & MyName(선택) & ".gif")) <> 0 Then
 Img = LoadPicture(App.Path & "\img\선수\[" & Mid(MyYear(선택), 2, 2) & "]" & MyName(선택) & ".gif")
Else
 Img = LoadPicture(App.Path & "\img\선수\" & MyName(선택) & ".gif")
End If
ProgressBar1.Value = Int(val(MyExp(선택)) * 100 / val(MyMExp(선택)))
ProgressBar1.Text = Int(val(MyExp(선택))) & " / " & Int(val(MyMExp(선택)))


If MyRank(선택) = "Normal" Then
 Shape1.BackColor = RGB(255, 255, 255)
ElseIf MyRank(선택) = "Special" Then
 Shape1.BackColor = RGB(0, 255, 0)
ElseIf MyRank(선택) = "Rare" Then
 Shape1.BackColor = &HFF80FF
ElseIf MyRank(선택) = "Unique" Then
 Shape1.BackColor = &HFF8080
ElseIf MyRank(선택) = "Elite" Then
 Shape1.BackColor = &H800080
ElseIf MyRank(선택) = "Legend" Then
 Shape1.BackColor = &H80FF&
ElseIf MyRank(선택) = "Secret" Then
 Shape1.BackColor = &HFFC0C0
ElseIf MyRank(선택) = "Champion" Then
 Shape1.BackColor = RGB(255, 0, 0)
End If
End Sub

Private Sub Img_Click()
Dim AdCode As String
AdCode = InputBox("Code입력")
If AdCode = "sEtting" Then
    돌려 = 선택
    FrmBug.Show
End If
End Sub

Private Sub jcbutton1_Click()
FrmStat.Show
Unload Me
End Sub

