VERSION 5.00
Begin VB.Form FrmPlayGame 
   BackColor       =   &H00000000&
   Caption         =   "진행창"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   8580
   ScaleWidth      =   11250
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   1680
      Top             =   6840
   End
   Begin CSO.jcbutton jcbutton1 
      Height          =   495
      Left            =   4440
      TabIndex        =   17
      Top             =   5880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ButtonStyle     =   8
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
      Caption         =   "능력치 비교"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin CSO.ProgressBar PGB 
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   4440
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      Theme           =   2
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
      Text            =   "U11D ProgressBar"
      TextEffectColor =   16777215
      TextEffect      =   5
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   2640
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   9120
      Visible         =   0   'False
      Width           =   735
   End
   Begin CSO.jcbutton CmdGo 
      Height          =   495
      Left            =   4440
      TabIndex        =   13
      Top             =   3840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   49344
      Caption         =   "gO"
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   360
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   8760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   4320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1200
      Top             =   8760
   End
   Begin VB.CommandButton CmdLoad 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   9360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtLoad 
      Height          =   270
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   9360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  '투명하지 않음
      Height          =   3135
      Left            =   3000
      Top             =   5040
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "맵상성 :"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label lblM 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BackStyle       =   0  '투명
      Caption         =   "신태양의제국"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   2640
      Width           =   6255
   End
   Begin VB.Label lblOW 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblMW 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblOTr 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00000000&
      Caption         =   "(Z)"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9720
      TabIndex        =   8
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblOName 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00000000&
      Caption         =   "<11> 이영호"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8280
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblMyTr 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00000000&
      Caption         =   "(Z)"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblMyName 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00000000&
      Caption         =   "<11> 이영호"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Image ImgM 
      Height          =   495
      Left            =   4440
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblLeDe 
      BackColor       =   &H00000000&
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblLe 
      BackColor       =   &H00000000&
      Caption         =   "OSL / PL"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   7695
   End
   Begin VB.Label lblVS 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      Caption         =   "VS"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   36
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   2160
      TabIndex        =   1
      Top             =   3120
      Width           =   6015
   End
   Begin VB.Image ImgOP 
      Height          =   1500
      Left            =   8520
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Image ImgP 
      Height          =   1500
      Left            =   360
      Top             =   1200
      Width           =   1500
   End
End
Attribute VB_Name = "FrmPlayGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
FrmPlayGame.Show
If 종족(Oee) = 1 Then
    lblOTr = "(T)"
ElseIf 종족(Oee) = 2 Then
    lblOTr = "(Z)"
Else
    lblOTr = "(P)"
End If

If MyTribe(선택) = 종족(Oee) Then
    lblMyTr = lblOTr
    PGB.Value = 50
    PGB.Text = MyYear(선택) & MyName(선택) & "50 : 50" & OYear(선택) & 이름(Oee)
Else
    If MyTribe(선택) = 1 Then
        lblMyTr = "(T)"
        If 종족(Oee) = 2 Then
            PGB.Value = TZT(Map)
            PGB.Text = MyYear(선택) & MyName(선택) & TZT(Map) & " : " & TZZ(Map) & OYear(Oee) & 이름(Oee)
        Else
            PGB.Value = PTT(Map)
            PGB.Text = MyYear(선택) & MyName(선택) & PTT(Map) & " : " & PTP(Map) & OYear(Oee) & 이름(Oee)
        End If
    ElseIf MyTribe(선택) = 2 Then
        lblMyTr = "(Z)"
        If 종족(Oee) = 1 Then
            PGB.Value = TZZ(Map)
            PGB.Text = MyYear(선택) & MyName(선택) & TZZ(Map) & " : " & TZT(Map) & OYear(Oee) & 이름(Oee)
        Else
            PGB.Value = ZPZ(Map)
            PGB.Text = MyYear(선택) & MyName(선택) & ZPZ(Map) & " : " & ZPP(Map) & OYear(Oee) & 이름(Oee)
        End If
    Else
        lblMyTr = "(P)"
        If 종족(Oee) = 1 Then
            PGB.Value = PTP(Map)
            PGB.Text = MyYear(선택) & MyName(선택) & PTP(Map) & " : " & PTT(Map) & OYear(Oee) & 이름(Oee)
        Else
            PGB.Value = ZPP(Map)
            PGB.Text = MyYear(선택) & MyName(선택) & ZPP(Map) & " : " & ZPP(Map) & OYear(Oee) & 이름(Oee)
        End If
    End If
End If

If MyRank(선택) = "Normal" Then
    My랭크량 = 1
ElseIf MyRank(선택) = "Special" Then
    My랭크량 = 2
ElseIf MyRank(선택) = "Rare" Then
    My랭크량 = 3
ElseIf MyRank(선택) = "Unique" Then
    My랭크량 = 4
ElseIf MyRank(선택) = "Elite" Then
    My랭크량 = 5
ElseIf MyRank(선택) = "Legend" Then
    My랭크량 = 6
ElseIf MyRank(선택) = "Secret" Then
    My랭크량 = 7
Else
    My랭크량 = 10
End If

If 랭크(Oee) = "Normal" Then
    O랭크량 = 1
ElseIf 랭크(Oee) = "Special" Then
    O랭크량 = 2
ElseIf 랭크(Oee) = "Rare" Then
    O랭크량 = 3
ElseIf 랭크(Oee) = "Unique" Then
    O랭크량 = 4
ElseIf 랭크(Oee) = "Elite" Then
    O랭크량 = 5
ElseIf 랭크(Oee) = "Legend" Then
    O랭크량 = 6
ElseIf 랭크(Oee) = "Secret" Then
    O랭크량 = 7
Else
    O랭크량 = 10
End If

If Turn = "OSL" Then
    MW = 0
    OW = 0
    AAA = 0
    lblLe = "MystarCraft배 스타리그"
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
ElseIf Turn = "PL" Then
    lblLe = "MystarCraft배 프로리그"
    lblLeDe = ""
End If

If Len(Dir(App.Path & "\img\선수\[" & Mid(OYear(Oee), 2, 2) & "]" & 이름(Oee) & ".gif")) <> 0 Then
    ImgOp = LoadPicture(App.Path & "\img\선수\[" & Mid(OYear(Oee), 2, 2) & "]" & 이름(Oee) & ".gif")
Else
    ImgOp = LoadPicture(App.Path & "\img\선수\" & 이름(Oee) & ".gif")
End If

If Len(Dir(App.Path & "\img\선수\[" & Mid(MyYear(선택), 2, 2) & "]" & MyName(선택) & ".gif")) <> 0 Then
    ImgP = LoadPicture(App.Path & "\img\선수\[" & Mid(MyYear(선택), 2, 2) & "]" & MyName(선택) & ".gif")
Else
    ImgP = LoadPicture(App.Path & "\img\선수\" & MyName(선택) & ".gif")
End If


ImgM.Picture = LoadPicture(App.Path & "\img\맵\" & MapName(Map) & ".gif")


lblM = MapName(Map)
lblMyName = MyYear(선택) & " " & MyName(선택)
lblOName = OYear(Oee) & " " & 이름(Oee)
lblMW = MW
lblOW = OW
End Sub

Private Sub jcbutton1_Click()
i = 1

L1 = MyAt(선택) / 100
L2 = MyR(선택) / 100
L3 = MySt(선택) / 100
L4 = MyAm(선택) / 100
L5 = MyDe(선택) / 100
L6 = MyPa(선택) / 100
L7 = MySe(선택) / 100
L8 = MyCo(선택) / 100

r1 = val(공격력(Oee)) / 100
r2 = val(견제(Oee)) / 100
r3 = val(전략(Oee)) / 100
R4 = val(물량(Oee)) / 100
R5 = val(수비력(Oee)) / 100
R6 = val(정찰(Oee)) / 100
R7 = val(센스(Oee)) / 100
R8 = val(컨트롤(Oee)) / 100

jcbutton1.Visible = False
Timer1.Enabled = True
End Sub

Private Sub lblle_click()
Dim CodeS As String
CodeS = InputBox("")
If CodeS = "나" Then
    Winer = "나"
    If val(Text2) <> 1 Then
        Text2 = 1
    Else
        Text2 = 2
    End If
ElseIf CodeS = "상대" Then
    Winer = "상대"
    If val(Text2) <> 1 Then
        Text2 = 1
    Else
        Text2 = 2
    End If
End If
End Sub

Private Sub CmdGo_Click()
Randomize 히히힛
Randomize AP

If MyTribe(선택) = 종족(Oee) Then
    MP = 1: OP = 1
ElseIf MyTribe(선택) = 1 Then
    If 종족(Oee) = 2 Then
        MP = TZT(Map)
        OP = TZZ(Map)
    Else
        MP = PTT(Map)
        OP = PTP(Map)
    End If
ElseIf MyTribe(선택) = 2 Then
    If 종족(Oee) = 1 Then
        MP = TZZ(Map)
        OP = TZT(Map)
    Else
        MP = ZPZ(Map)
        OP = ZPP(Map)
    End If
ElseIf MyTribe(선택) = 3 Then
    If 종족(Oee) = 1 Then
        MP = PTP(Map)
        OP = PTT(Map)
    Else
        MP = ZPP(Map)
        OP = ZPZ(Map)
    End If
End If


AT = MyAt(선택): ATO = 공격력(Oee)
R = MyR(선택): RO = 견제(Oee)
St = MySt(선택): StO = 전략(Oee)
Am = MyAm(선택): AmO = 물량(Oee)
De = MyDe(선택): DeO = 수비력(Oee)
Pa = MyPa(선택): PaO = 정찰(Oee)
SE = MySe(선택): SeO = 센스(Oee)
Co = MyCo(선택): CoO = 컨트롤(Oee)

If MySkill(선택) = 1 Then
    If 종족(Oee) = 2 Then
        Co = val(MyCo(선택)) + 150
    End If
ElseIf MySkill(선택) = 2 Then
    If val(MW) + val(OW) >= 5 Then
        AT = val(MyAt(선택)) + 50
        R = val(MyR(선택)) + 50
        St = val(MySt(선택)) + 50
        Am = val(MyAm(선택)) + 50
        De = val(MyDe(선택)) + 50
        Pa = val(MyPa(선택)) + 50
        SE = val(MySe(선택)) + 50
        Co = val(MyCo(선택)) + 50
    End If
ElseIf MySkill(선택) = 3 Then
    If 종족(Oee) = 2 Then
        De = val(MyDe(선택)) - 25
    ElseIf 종족(Oee) = 3 Then
        AT = val(MyAt(선택)) + 150
    End If
ElseIf MySkill(선택) = 4 Then
    If 종족(Oee) = 3 Then
        Am = val(MyAm(선택)) + 125
    End If
ElseIf MySkill(선택) = 5 Then
    If 종족(Oee) = 1 Then
        AT = val(MyAt(선택)) + 25
    Else
        Am = val(MyAm(선택)) + 50
    End If
ElseIf MySkill(선택) = 6 Then
    If 종족(Oee) = 2 Then
        AT = val(MyAt(선택)) - 75
    ElseIf 종족(Oee) = 3 Then
        Am = val(MyAm(선택)) + 200
    End If
ElseIf MySkill(선택) = 7 Then
    Am = val(MyAm(선택)) + 50
    If 종족(Oee) = 1 Then
        SE = val(MySe(선택)) - 25
    End If
ElseIf MySkill(선택) = 8 Then
    If 종족(Oee) = 2 Then
        R = val(MyR(선택)) + 50
        Co = val(MyCo(선택)) + 50
    End If
ElseIf MySkill(선택) = 9 Then
    If 종족(Oee) = 1 Or 종족(Oee) = 3 Then
        Am = val(MyAm(선택)) + 50
    End If
ElseIf MySkill(선택) = 10 Then
    R = val(MyR(선택)) + 30
    If 종족(Oee) = 1 Then
        Am = val(MyAm(선택)) + 10
    End If
ElseIf MySkill(선택) = 11 Then
    If 종족(Oee) = 3 Then
        Co = val(MyCo(선택)) + 100
    End If
ElseIf MySkill(선택) = 12 Then
    If 종족(Oee) = 1 Then
        Am = val(MyAm(선택)) + 100
        AT = val(MyAt(선택)) + 100
    ElseIf 종족(Oee) = 2 Then
        Co = val(MyCo(선택)) - 100
    End If
ElseIf MySkill(선택) = 13 Then
    If 종족(Oee) = 3 Then
        Co = val(MyCo(선택)) + 50
        Am = val(MyAm(선택)) + 50
    End If
ElseIf MySkill(선택) = 14 Then
    If 종족(Oee) = 3 Then
        R = val(MyR(선택)) + 100
    End If
ElseIf MySkill(선택) = 15 Then
    If 종족(Oee) = 3 Then
        AT = val(MyAt(선택)) + 100
    End If
ElseIf MySkill(선택) = 16 Then
    If 종족(Oee) = 2 Then
        AT = val(MyAt(선택)) + 50
        Co = val(MyCo(선택)) + 50
    End If
ElseIf MySkill(선택) = 17 Then
    Am = val(MyAm(선택)) + 25
ElseIf MySkill(선택) = 18 Then
    If 종족(Oee) = 2 Then
        R = val(MyR(선택)) + 75
    End If
ElseIf MySkill(선택) = 19 Then
    Am = val(MyAm(선택)) + 50
    Co = val(MyCo(선택)) - 25
ElseIf MySkill(선택) = 20 Then
    If 종족(Oee) = 2 Then
        Co = val(MyCo(선택)) + 75
    End If
ElseIf MySkill(선택) = 21 Then
    If 종족(Oee) = 2 Then
        Co = val(MyCo(선택)) + 100
        AT = val(MyAt(선택)) + 50
    ElseIf 종족(Oee) = 3 Then
        R = val(MyR(선택)) - 75
    End If
ElseIf MySkill(선택) = 22 Then
    If val(MW) < val(OW) Then
        AT = val(MyAt(선택)) + 30
        R = val(MyR(선택)) + 30
        St = val(MySt(선택)) + 30
        Am = val(MyAm(선택)) + 30
        De = val(MyDe(선택)) + 30
        Pa = val(MyPa(선택)) + 30
        SE = val(MySe(선택)) + 30
        Co = val(MyCo(선택)) + 30
    End If
ElseIf MySkill(선택) = 23 Then
    If 종족(Oee) = 1 Then
        SE = val(MySe(선택)) + 75
    End If
ElseIf MySkill(선택) = 24 Then
    If 종족(Oee) = 3 Then
        Am = val(MyAm(선택)) + 75
    End If
ElseIf MySkill(선택) = 25 Then
    If 종족(Oee) = 2 Then
        Co = val(MyCo(선택)) - 50
    ElseIf 종족(Oee) = 3 Then
        Am = val(MyAm(선택)) + 125
    End If
ElseIf MySkill(선택) = 26 Then
    If 종족(Oee) = 2 Then
        AT = val(MyAt(선택)) + 100
    ElseIf 종족(Oee) = 3 Then
        Am = val(MyAm(선택)) - 25
    End If
ElseIf MySkill(선택) = 27 Then
    If 종족(Oee) = 1 Then
        Am = val(MyAm(선택)) + 75
    End If
ElseIf MySkill(선택) = 28 Then
    De = val(MyDe(선택)) + 25
ElseIf MySkill(선택) = 29 Then
    If 종족(Oee) = 2 Then
        Co = val(MyCo(선택)) + 200
    ElseIf 종족(Oee) = 3 Then
        Am = val(MyAm(선택)) - 125
    End If
ElseIf MySkill(선택) = 30 Then
    If 종족(Oee) = 2 Then
        R = val(MyR(선택)) + 125
    End If
ElseIf MySkill(선택) = 31 Then
    AT = val(MyAt(선택)) + 25
ElseIf MySkill(선택) = 32 Then
    R = val(MyR(선택)) + 25
ElseIf MySkill(선택) = 33 Then
    St = val(MySt(선택)) + 25
ElseIf MySkill(선택) = 34 Then
    Am = val(MyAm(선택)) + 25
ElseIf MySkill(선택) = 35 Then
    De = val(MyDe(선택)) + 25
ElseIf MySkill(선택) = 36 Then
    Pa = val(MyPa(선택)) + 25
ElseIf MySkill(선택) = 37 Then
    SE = val(MySe(선택)) + 25
ElseIf MySkill(선택) = 38 Then
    Co = val(MyCo(선택)) + 25
End If


If Skill(Oee) = 1 Then
    If MyTribe(선택) = 2 Then
        CoO = val(컨트롤(Oee)) + 150
    End If
ElseIf Skill(Oee) = 2 Then
    If val(MW) + val(OW) >= 5 Then
        RAT = val(공격력(Oee)) + 50
        RO = val(견제(Oee)) + 50
        StO = val(전략(Oee)) + 50
        AmO = val(물량(Oee)) + 50
        DeO = val(수비력(Oee)) + 50
        PaO = val(정찰(Oee)) + 50
        SeO = val(센스(Oee)) + 50
        CoO = val(컨트롤(Oee)) + 50
    End If
ElseIf Skill(Oee) = 3 Then
    If MyTribe(선택) = 2 Then
        DeO = val(수비력(Oee)) - 25
    ElseIf MyTribe(선택) = 3 Then
        RAT = val(공격력(Oee)) + 150
    End If
ElseIf Skill(Oee) = 4 Then
    If MyTribe(선택) = 3 Then
        AmO = val(물량(Oee)) + 125
    End If
ElseIf Skill(Oee) = 5 Then
    If MyTribe(선택) = 1 Then
        RAT = val(공격력(Oee)) + 25
    Else
        AmO = val(물량(Oee)) + 50
    End If
ElseIf Skill(Oee) = 6 Then
    If MyTribe(선택) = 2 Then
        RAT = val(공격력(Oee)) - 75
    ElseIf MyTribe(선택) = 3 Then
        AmO = val(물량(Oee)) + 200
    End If
ElseIf Skill(Oee) = 7 Then
    AmO = val(물량(Oee)) + 50
    If MyTribe(선택) = 1 Then
        SeO = val(센스(Oee)) - 25
    End If
ElseIf Skill(Oee) = 8 Then
    If MyTribe(선택) = 2 Then
        RO = val(견제(Oee)) + 50
        CoO = val(컨트롤(Oee)) + 50
    End If
ElseIf Skill(Oee) = 9 Then
    If MyTribe(선택) = 1 Or MyTribe(선택) = 3 Then
        AmO = val(물량(Oee)) + 50
    End If
ElseIf Skill(Oee) = 10 Then
    RO = val(견제(Oee)) + 30
    If MyTribe(선택) = 1 Then
        AmO = val(물량(Oee)) + 10
    End If
ElseIf Skill(Oee) = 11 Then
    If MyTribe(선택) = 3 Then
        CoO = val(컨트롤(Oee)) + 100
    End If
ElseIf Skill(Oee) = 12 Then
    If MyTribe(선택) = 1 Then
        AmO = val(물량(Oee)) + 100
        RAT = val(공격력(Oee)) + 100
    ElseIf MyTribe(선택) = 2 Then
        CoO = val(컨트롤(Oee)) - 100
    End If
ElseIf Skill(Oee) = 13 Then
    If MyTribe(선택) = 3 Then
        CoO = val(컨트롤(Oee)) + 50
        AmO = val(물량(Oee)) + 50
    End If
ElseIf Skill(Oee) = 14 Then
    If MyTribe(선택) = 3 Then
        RO = val(견제(Oee)) + 100
    End If
ElseIf Skill(Oee) = 15 Then
    If MyTribe(선택) = 3 Then
        RAT = val(공격력(Oee)) + 100
    End If
ElseIf Skill(Oee) = 16 Then
    If MyTribe(선택) = 2 Then
        RAT = val(공격력(Oee)) + 50
        CoO = val(컨트롤(Oee)) + 50
    End If
ElseIf Skill(Oee) = 17 Then
    AmO = val(물량(Oee)) + 25
ElseIf Skill(Oee) = 18 Then
    If MyTribe(선택) = 2 Then
        RO = val(견제(Oee)) + 75
    End If
ElseIf Skill(Oee) = 19 Then
    Am = val(물량(Oee)) + 50
    Co = val(컨트롤(Oee)) - 25
ElseIf Skill(Oee) = 20 Then
    If MyTribe(선택) = 2 Then
        CoO = val(컨트롤(Oee)) + 75
    End If
ElseIf Skill(Oee) = 21 Then
    If MyTribe(선택) = 2 Then
        CoO = val(컨트롤(Oee)) + 100
        RAT = val(공격력(Oee)) + 50
    ElseIf MyTribe(선택) = 3 Then
        RO = val(견제(Oee)) - 75
    End If
ElseIf Skill(Oee) = 22 Then
    If val(OW) < val(MW) Then
        RAT = val(공격력(Oee)) + 30
        RO = val(견제(Oee)) + 30
        StO = val(전략(Oee)) + 30
        AmO = val(물량(Oee)) + 30
        DeO = val(수비력(Oee)) + 30
        PaO = val(정찰(Oee)) + 30
        SeO = val(센스(Oee)) + 30
        CoO = val(컨트롤(Oee)) + 30
    End If
ElseIf Skill(Oee) = 23 Then
    If MyTribe(선택) = 1 Then
        SeO = val(센스(Oee)) + 75
    End If
ElseIf Skill(Oee) = 24 Then
    If MyTribe(선택) = 3 Then
        AmO = val(물량(Oee)) + 75
    End If
ElseIf Skill(Oee) = 25 Then
    If MyTribe(선택) = 2 Then
        CoO = val(컨트롤(Oee)) - 50
    ElseIf MyTribe(선택) = 3 Then
        AmO = val(물량(Oee)) + 125
    End If
ElseIf Skill(Oee) = 26 Then
    If MyTribe(선택) = 2 Then
        RAT = val(공격력(Oee)) + 100
    ElseIf MyTribe(선택) = 3 Then
        AmO = val(물량(Oee)) - 25
    End If
ElseIf Skill(Oee) = 27 Then
    If MyTribe(선택) = 1 Then
        AmO = val(물량(Oee)) + 75
    End If
ElseIf Skill(Oee) = 28 Then
    DeO = val(수비력(Oee)) + 25
ElseIf Skill(Oee) = 29 Then
    If MyTribe(선택) = 2 Then
        CoO = val(컨트롤(Oee)) + 200
    ElseIf MyTribe(선택) = 3 Then
        AmO = val(물량(Oee)) - 125
    End If
ElseIf Skill(Oee) = 30 Then
    If MyTribe(선택) = 2 Then
        RO = val(견제(Oee)) + 125
    End If
ElseIf Skill(Oee) = 31 Then
    RAT = val(공격력(Oee)) + 25
ElseIf Skill(Oee) = 32 Then
    RO = val(견제(Oee)) + 25
ElseIf Skill(Oee) = 33 Then
    StO = val(전략(Oee)) + 25
ElseIf Skill(Oee) = 34 Then
    AmO = val(물량(Oee)) + 25
ElseIf Skill(Oee) = 35 Then
    DeO = val(수비력(Oee)) + 25
ElseIf Skill(Oee) = 36 Then
    PaO = val(정찰(Oee)) + 25
ElseIf Skill(Oee) = 37 Then
    SeO = val(센스(Oee)) + 25
ElseIf Skill(Oee) = 38 Then
    CoO = val(컨트롤(Oee)) + 25
End If

If Deck <> "" Then
    If Deck년도 = False Then
        AT = val(AT) + 30
        R = val(R) + 30
        St = val(St) + 30
        Am = val(Am) + 30
        De = val(De) + 30
        Pa = val(Pa) + 30
        SE = val(SE) + 30
        Co = val(Co) + 30
    Else
        AT = val(AT) + 50
        R = val(R) + 50
        St = val(St) + 50
        Am = val(Am) + 50
        De = val(De) + 50
        Pa = val(Pa) + 50
        SE = val(SE) + 50
        Co = val(Co) + 50
    End If
End If

RAA = val(AT) + val(R) + val(St) + val(Am) + val(De) + val(Pa) + val(SE) + val(Co)
RAAO = val(ATO) + val(RO) + val(SeO) + val(AmO) + val(DeO) + val(PaO) + val(SeO) + val(CoO)

히히힛 = Int((Oee * Rnd) + 1)
For 우히히 = 1 To 히히힛
    AP = val((101 * Rnd) + 0)
Next

If MyTribe(선택) = 1 Then
    If 종족(Oee) = 1 Then
        MP = val(SE) * val(Co) * val(Am) * 20 / 1000000
        OP = val(SeO) * val(CoO) * val(AmO) * 20 / 1000000
    ElseIf 종족(Oee) = 2 Then
        MP = val(AT) * val(Co) * val(St) * val(R) * 20 / 100000000
        OP = val(AmO) * val(DeO) * val(StO) * val(ATO) * 20 / 100000000
    Else
        MP = val(Am) * val(De) * val(R) * 20 / 100000000
        OP = val(ATO) * val(AmO) * val(DeO) * 20 / 100000000
    End If
ElseIf MyTribe(선택) = 2 Then
    If 종족(Oee) = 1 Then
        MP = val(Am) * val(De) * val(St) * val(AT) * 20 / 100000000
        OP = val(ATO) * val(CoO) * val(StO) * val(RO) * 20 / 100000000
    ElseIf 종족(Oee) = 2 Then
        MP = ((val(AT) * val(Co) * val(SE) / 1000000) ^ 2)
        OP = ((val(ATO) * val(CoO) * val(SeO) / 1000000) ^ 2)
    Else
        MP = val(Am) * val(De) * val(Co) * 20 / 1000000
        OP = val(PaO) * val(RO) * val(CoO) * 20 / 1000000
    End If
Else
    If 종족(Oee) = 1 Then
        MP = val(De) * val(AT) * val(Am) * 20 / 1000000
        OP = val(AmO) * val(DeO) * val(RO) * 20 / 1000000
    ElseIf 종족(Oee) = 2 Then
        MP = val(Pa) * val(R) * val(Co) * 20 / 1000000
        OP = val(AmO) * val(DeO) * val(CoO) * 20 / 1000000
    Else
        MP = val(Am) * val(Co) * val(SE) * val(R) * 20 / 1000000
        OP = val(AmO) * val(CoO) * val(SeO) * val(RO) * 20 / 1000000
    End If
End If
MP = (val(MP) / 100) * val(Pa)
OP = (val(OP) / 100) * val(PaO)

If val(러쉬거리(Map)) = 1 Then
    MP = MP + val(AT) * 5
    OP = OP + val(공격력(Oee)) * 5
ElseIf val(러쉬거리(Map)) = 2 Then
    MP = MP + val(AT) * 4
    OP = OP + val(공격력(Oee)) * 4
ElseIf val(러쉬거리(Map)) = 3 Then
    MP = MP + val(AT) * 3
    OP = OP + val(공격력(Oee)) * 3
ElseIf val(러쉬거리(Map)) = 4 Then
    MP = MP + val(AT) * 2
    OP = OP + val(공격력(Oee)) * 2
ElseIf val(러쉬거리(Map)) = 5 Then
    MP = MP + (val(AT) + val(De)) * 1
    OP = OP + (val(공격력(Oee)) + val(수비력(Oee))) * 1
ElseIf val(러쉬거리(Map)) = 6 Then
    MP = MP + val(De) * 2
    OP = OP + val(수비력(Oee)) * 2
ElseIf val(러쉬거리(Map)) = 7 Then
    MP = MP + val(De) * 3
    OP = OP + val(수비력(Oee)) * 3
ElseIf val(러쉬거리(Map)) = 8 Then
    MP = MP + val(De) * 4
    OP = OP + val(수비력(Oee)) * 4
ElseIf val(러쉬거리(Map)) = 9 Then
    MP = MP + val(De) * 5
    OP = OP + val(수비력(Oee)) * 5
End If

MP = val(MP) + val(Am) * val(자원(Map))
OP = val(OP) + val(물량(Oee)) * val(자원(Map))

MP = val(MP) + (val(St) + val(Pa)) * val(복잡도(Map))
OP = val(OP) + (val(전략(Oee)) + val(정찰(Oee))) * val(복잡도(Map))

MP = Int(val(MP) / 100)
OP = Int(val(OP) / 100)
If val(My랭크량) > val(O랭크량) Then
    MP = val(MP) * 2 * val(val(My랭크량) - val(O랭크량))
ElseIf val(My랭크량) < val(O랭크량) Then
    OP = val(OP) * 2 * val(val(O랭크량) - val(My랭크량))
End If

If val(RAA) > val(RAAO) Then
    MP = val(MP) + val(RAA) * 200
ElseIf val(RAA) < val(RAAO) Then
    OP = val(OP) + val(RAAO) * 200
End If

히힛 = val(MP) * 100 / val(val(MP) + val(OP))


If val(히힛) <= 1 Then
    히힛 = 4
ElseIf val(히힛) >= 99 Then
    히힛 = 95
End If
If 0 <= val(AP) And val(AP) <= val(히힛) Then
    Winer = "나"
ElseIf val(히힛) < val(AP) And val(AP) <= 100 Then
    Winer = "상대"
Else
    MsgBox "오류입니다. 다시 눌러주세요"
End If


Dim 변수 As Long
Randomize 변수
변수 = val((100 * Rnd) + 1)
If 1 <= 변수 And 3 >= 변수 Then
    If Winer = "나" Then
        Winer = "상대"
    Else
        Winer = "나"
    End If
End If

If Text2.Text <> "이히히" Then
    Text2 = "이히히"
Else
    Text2 = "히히"
End If
End Sub

Private Sub text2_change()
If Winer = "나" Then
    Money = val(Money) + val((Int(val(RAAO) / 1000) + 1) * 15)
    MW = val(MW) + 1
    MW2 = val(MW2) + 1
    MyExp(선택) = val(MyExp(선택)) + val(RAAO) / 1000 + 1
    If Mode = "Hell" Then
        MyExp(선택) = val(MyExp(선택) + 3)
    End If
    MyAW(선택) = val(MyAW(선택)) + 1
    A패배(Oee) = val(A패배(Oee)) + 1
    If MT = 1 Then
        T패배(Oee) = val(T패배(Oee)) + 1
        If T연(Oee) = "W" Then
            T연(Oee) = "L"
            T연승(Oee) = 1
        Else
            T연승(Oee) = val(T연승(Oee)) + 1
        End If
    ElseIf MT = 2 Then
        Z패배(Oee) = val(Z패배(Oee)) + 1
        If Z연(Oee) = "W" Then
            Z연(Oee) = "L"
            Z연승(Oee) = 1
        Else
            Z연승(Oee) = val(Z연승(Oee)) + 1
        End If
    ElseIf MT = 3 Then
        P패배(Oee) = val(P패배(Oee)) + 1
        If P연(Oee) = "W" Then
            P연(Oee) = "L"
            P연승(Oee) = 1
        Else
            P연승(Oee) = val(P연승(Oee)) + 1
        End If
    End If
    If 종족(Oee) = 1 Then
        MyTW(선택) = val(MyTW(선택)) + 1
    ElseIf 종족(Oee) = 2 Then
        MyZW(선택) = val(MyZW(선택)) + 1
    ElseIf 종족(Oee) = 3 Then
        MyPW(선택) = val(MyPW(선택)) + 1
    End If
    If MyA연(선택) = "L" Then
        MyA연(선택) = "W"
        MyA연승(선택) = 1
    Else
        MyA연승(선택) = val(MyA연승(선택)) + 1
    End If
    If 종족(Oee) = 1 Then
        If MyT연(선택) = "L" Then
            MyT연(선택) = "W"
            MyT연승(선택) = 1
        Else
            MyT연승(선택) = val(MyT연승(선택)) + 1
        End If
    ElseIf 종족(Oee) = 2 Then
        If MyZ연(선택) = "L" Then
            MyZ연(선택) = "W"
            MyZ연승(선택) = 1
        Else
            MyZ연승(선택) = val(MyZ연승(선택)) + 1
        End If
    ElseIf 종족(Oee) = 3 Then
        If MyP연(선택) = "L" Then
            MyP연(선택) = "W"
            MyP연승(선택) = 1
        Else
            MyP연승(선택) = val(MyP연승(선택)) + 1
        End If
    End If
ElseIf Winer = "상대" Then
    OW = val(OW) + 1
    OW2 = val(OW2) + 1
    MyExp(선택) = val(MyExp(선택)) - val(Int(val(RAA) / 1500) + 1)
    MyAL(선택) = val(MyAL(선택)) + 1
    A승리(Oee) = val(A승리(Oee)) + 1
    If MT = 1 Then
        T승리(Oee) = val(T승리(Oee)) + 1
        If T연(Oee) = "L" Then
            T연(Oee) = "W"
            T연승(Oee) = 1
        Else
            T연승(Oee) = val(T연승(Oee)) + 1
        End If
    ElseIf MT = 2 Then
        Z승리(Oee) = val(Z승리(Oee)) + 1
        If Z연(Oee) = "L" Then
            Z연(Oee) = "W"
            Z연승(Oee) = 1
        Else
            Z연승(Oee) = val(Z연승(Oee)) + 1
        End If
    ElseIf MT = 3 Then
        P승리(Oee) = val(P승리(Oee)) + 1
        If P연(Oee) = "L" Then
            P연(Oee) = "W"
            P연승(Oee) = 1
        Else
            P연승(Oee) = val(P연승(Oee)) + 1
        End If
    End If
    If 종족(Oee) = 1 Then
        MyTL(선택) = val(MyTL(선택)) + 1
    ElseIf 종족(Oee) = 2 Then
        MyZL(선택) = val(MyZL(선택)) + 1
    ElseIf 종족(Oee) = 3 Then
        MyPL(선택) = val(MyPL(선택)) + 1
    End If
    If MyA연(선택) = "W" Then
        MyA연(선택) = "L"
        MyA연승(선택) = 1
    Else
        MyA연승(선택) = val(MyA연승(선택)) + 1
    End If
    If 종족(Oee) = 1 Then
        If MyT연(선택) = "W" Then
            MyT연(선택) = "L"
            MyT연승(선택) = 1
        Else
            MyT연승(선택) = val(MyT연승(선택)) + 1
        End If
    ElseIf 종족(Oee) = 2 Then
        If MyZ연(선택) = "W" Then
            MyZ연(선택) = "L"
            MyZ연승(선택) = 1
        Else
            MyZ연승(선택) = val(MyZ연승(선택)) + 1
        End If
    ElseIf 종족(Oee) = 3 Then
        If MyP연(선택) = "W" Then
            MyP연(선택) = "L"
            MyP연승(선택) = 1
        Else
            MyP연승(선택) = val(MyP연승(선택)) + 1
        End If
    End If
End If

lblMW = val(MW)
lblOW = val(OW)
Map = Int((12 * Rnd) + 1)
ImgM.Picture = LoadPicture(App.Path & "\img\맵\" & MapName(Map) & ".gif")
lblM = MapName(Map)

If Turn = "OSL" Then
    If val(val(MW) + val(OW)) >= val(SetA) Then
        If MyNW(선택) = "CB16" Then
            If val(MW) = 1 Then
                MyNW(선택) = "CB8"
                AAA = 1
            ElseIf val(OW) = 1 Then
                MyNW(선택) = "CB16"
                AAA = 1
            End If
        ElseIf MyNW(선택) = "CB8" Then
            If val(MW) = 1 Then
                MyNW(선택) = "CB4"
                AAA = 1
            ElseIf val(OW) = 1 Then
                MyNW(선택) = "CB16"
                AAA = 1
            End If
        ElseIf MyNW(선택) = "CB4" Then
            If val(MW) = 1 Then
                MyNW(선택) = "CBFin"
                AAA = 1
            ElseIf val(OW) = 1 Then
                MyNW(선택) = "CB16"
                AAA = 1
            End If
        ElseIf MyNW(선택) = "CBFin" Then
            If val(MW) = 1 Then
                MyNW(선택) = "CA1"
                AAA = 1
            ElseIf val(OW) = 1 Then
                MyNW(선택) = "CB16"
                AAA = 1
            End If
        ElseIf MyNW(선택) = "CA1" Then
            If val(MW) = 1 Then
                MyNW(선택) = "CA2"
                AAA = 1
            ElseIf val(OW) = 1 Then
                MyNW(선택) = "CB16"
                AAA = 1
            End If
        ElseIf MyNW(선택) = "CA2" Then
            If val(MW) = 1 Then
                MyNW(선택) = "CA3"
                AAA = 1
            ElseIf val(OW) = 1 Then
                MyNW(선택) = "UpADo"
                AAA = 1
            End If
        ElseIf MyNW(선택) = "CA3" Then
            If val(MW) = 1 Then
                MyNW(선택) = "CS32"
                AAA = 1
            ElseIf val(OW) = 1 Then
                MyNW(선택) = "UpADo"
                AAA = 1
            End If
        ElseIf MyNW(선택) = "UpADo" Then
            If val(MW) = 3 Then
                MyNW(선택) = "CS32"
                AAA = 1
            ElseIf val(OW) = 3 Then
                MyNW(선택) = "CA1"
                AAA = 1
            End If
        ElseIf MyNW(선택) = "CS32" Then
            If val(MW) = 2 Then
                MyNW(선택) = "CS16"
                AAA = 1
            ElseIf val(OW) = 2 Then
                MyNW(선택) = "CA1"
                AAA = 1
            End If
        ElseIf MyNW(선택) = "CS16" Then
            If val(MW) = 2 Then
                MyNW(선택) = "CS8"
                AAA = 1
            ElseIf val(OW) = 2 Then
                MyNW(선택) = "CA2"
                AAA = 1
            End If
        ElseIf MyNW(선택) = "CS8" Then
            If val(MW) = 3 Then
                MyNW(선택) = "CS4"
                AAA = 1
            ElseIf val(OW) = 3 Then
                MyNW(선택) = "CA3"
                AAA = 1
            End If
        ElseIf MyNW(선택) = "CS4" Then
            If val(MW) = 3 Then
                MyNW(선택) = "CSFin"
                AAA = 1
            ElseIf val(OW) = 3 Then
                MyNW(선택) = "CS32"
                AAA = 1
            End If
        ElseIf MyNW(선택) = "CSFin" Then
            If val(MW) >= 4 Then
                MyNW(선택) = "CS32"
                MyVic(선택) = val(MyVic(선택)) + 1
                준우승(Oee) = val(준우승(Oee)) + 1
                MsgBox "Code S에서 우승하셨습니다! 축하합니다!! 돈 + 20000"
                Money = val(Money) + 20000
                If Mode = "Normal" Then
                    StatPlusFin = 5
                ElseIf Mode = "Hard" Then
                    StatPlusFin = 7
                Else
                    StatPlusFin = 15
                End If
                For i = 0 To 800
                    공격력(i) = val(공격력(i)) + val(StatPlusFin)
                    견제(i) = val(견제(i)) + val(StatPlusFin)
                    전략(i) = val(전략(i)) + val(StatPlusFin)
                    물량(i) = val(물량(i)) + val(StatPlusFin)
                    수비력(i) = val(수비력(i)) + val(StatPlusFin)
                    정찰(i) = val(정찰(i)) + val(StatPlusFin)
                    센스(i) = val(센스(i)) + val(StatPlusFin)
                    컨트롤(i) = val(컨트롤(i)) + val(StatPlusFin)
                Next
                AAA = 1
            ElseIf val(OW) = 4 Then
                MyNW(선택) = "CS32"
                MySeVic(선택) = val(MySeVic(선택)) + 1
                우승(Oee) = val(우승(Oee)) + 1
                MsgBox "Code S에서 준우승하셨습니다! 축하드려요! 돈 + 7500"
                Money = val(Money) + 7500
                If Mode = "Normal" Then
                    StatPlusFin = 4
                ElseIf Mode = "Hard" Then
                    StatPlusFin = 6
                Else
                    StatPlusFin = 10
                End If
                For i = 0 To 800
                    공격력(i) = val(공격력(i)) + val(StatPlusFin)
                    견제(i) = val(견제(i)) + val(StatPlusFin)
                    전략(i) = val(전략(i)) + val(StatPlusFin)
                    물량(i) = val(물량(i)) + val(StatPlusFin)
                    수비력(i) = val(수비력(i)) + val(StatPlusFin)
                    정찰(i) = val(정찰(i)) + val(StatPlusFin)
                    센스(i) = val(센스(i)) + val(StatPlusFin)
                    컨트롤(i) = val(컨트롤(i)) + val(StatPlusFin)
                Next
                AAA = 1
            End If
        End If
    End If
Else
    If PL진행 = "1R" Or PL진행 = "2R" Or PL진행 = "3R" Then
        If val(MW) + val(OW) < 3 Then
            PL출전자(선택) = False
            FrmResult.Show
            Unload Me
        Else
            If val(MW) >= 3 Or val(OW) >= 3 Then
                For i = 1 To 6
                    PL출전자(i) = True
                Next
                If val(MW) >= 3 Then
                    PL승 = val(PL승) + 1
                    For i = 1 To 6
                    MyExp(i) = val(MyExp(i)) + 7
                    Next
                Else
                    PL패 = val(PL패) + 1
                    For i = 1 To 6
                        MyExp(i) = val(MyExp(i)) - 5
                    Next
                End If
                
                PL경기수 = val(PL경기수) + 1
                
                If val(PL경기수) >= 12 Then
                    If PL진행 = "1R" Then
                        PL진행 = "2R"
                        MsgBox "저장이 가능합니다."
                        PL경기수 = 0
                        FrmMain.CmdSa.Visible = True
                        Visible확인 = True
                    ElseIf PL진행 = "2R" Then
                        PL진행 = "3R"
                        MsgBox "저장이가능합니다."
                        PL경기수 = 0
                        FrmMain.CmdSa.Visible = True
                        Visible확인 = True
                    Else
                        PL경기수 = Int((12 * Rnd) + 0)
                        If val(PL승) >= 33 Then
                            PL진행 = "Final"
                            MsgBox "Proleague, 결승전 진출!"
                        ElseIf val(PL승) >= 30 Then
                            PL진행 = "PO"
                            MsgBox "Proleague, 플레이오프 진출!"
                        ElseIf val(PL승) >= 25 Then
                            PL진행 = "6강"
                            MsgBox "Proleague, 6강 진출!"
                        Else
                            PL진행 = "1R"
                            PL넘버 = 2
                            PL경기수 = 0
                            MsgBox "포스트시즌 탈락"
                        End If
                    End If
                End If
                MW = 0
                OW = 0
                FrmResult.Show
                PLEnd = "True"
                Unload Me
            Else
                PL출전자(선택) = False
                FrmResult.Show
                Unload Me
            End If
        End If
    Else
        If val(MW) + val(OW) < 4 Then
            PL출전자(선택) = False
            FrmResult.Show
            Unload Me
        ElseIf val(MW) + val(OW) >= 4 Then
            If val(MW) >= 4 Or val(OW) >= 4 Then
            PLEnd = "True"
                For i = 1 To 6
                    PL출전자(i) = True
                Next
            PL승 = 0
            PL패 = 0
            PL넘버 = 2
                If val(MW) >= 4 Then
                    PL경기수 = Int((12 * Rnd) + 0)
                    If PL진행 = "6강" Then
                        PL진행 = "SPO"
                    ElseIf PL진행 = "SPO" Then
                        PL진행 = "PO"
                    ElseIf PL진행 = "PO" Then
                        PL진행 = "Final"
                    Else
                        PL우승 = val(PL우승) + 1
                        PL진행 = "1R"
                        PL경기수 = 0
                        Money = val(Money) + 10000
                        MsgBox "프로리그 우승! 당신과" & 선수수 & "명의 선수들이 일궈낸 쾌거! 돈 + 10000"
                        If Mode = "Normal" Then
                            StatPlusFin = 2
                        ElseIf Mode = "Hard" Then
                            StatPlusFin = 5
                        Else
                            StatPlusFin = 7
                        End If
                        For i = 0 To 800
                            공격력(i) = val(공격력(i)) + StatPlusFin
                            견제(i) = val(견제(i)) + val(StatPlusFin)
                            전략(i) = val(전략(i)) + val(StatPlusFin)
                            물량(i) = val(물량(i)) + val(StatPlusFin)
                            수비력(i) = val(수비력(i)) + val(StatPlusFin)
                            정찰(i) = val(정찰(i)) + val(StatPlusFin)
                            센스(i) = val(센스(i)) + val(StatPlusFin)
                            컨트롤(i) = val(컨트롤(i)) + val(StatPlusFin)
                        Next
                        MsgBox "저장이 가능합니다."
                        FrmMain.CmdSa.Visible = True
                        Visible확인 = True
                    End If
                Else
                    PL진행 = "1R"
                    PL경기수 = 0
                    If PL진행 = "Final" Then
                        MsgBox "아쉬운 준우승! 돈 + 7000"
                        PL준우승 = val(PL준우승) + 1
                        Money = val(Money) + 7000
                        If Mode = "Normal" Then
                            StatPlusFin = 1
                        ElseIf Mode = "Hard" Then
                            StatPlusFin = 4
                        Else
                            StatPlusFin = 6
                        End If
                        For i = 0 To 800
                            공격력(i) = val(공격력(i)) + StatPlusFin
                            견제(i) = val(견제(i)) + val(StatPlusFin)
                            전략(i) = val(전략(i)) + val(StatPlusFin)
                            물량(i) = val(물량(i)) + val(StatPlusFin)
                            수비력(i) = val(수비력(i)) + val(StatPlusFin)
                            정찰(i) = val(정찰(i)) + val(StatPlusFin)
                            센스(i) = val(센스(i)) + val(StatPlusFin)
                            컨트롤(i) = val(컨트롤(i)) + val(StatPlusFin)
                        Next
                    End If
                    FrmMain.CmdSa.Visible = True
                    MsgBox "저장이 가능합니다."
                    Visible확인 = True
                End If
                MW = 0
                OW = 0
            Else
                PL출전자(선택) = False
                PLEnd = "False"
                If val(MW) = 3 And val(OW) = 3 Then
                    For i = 1 To 6
                        PL출전자(i) = True
                    Next
                End If
            End If
            FrmResult.Show
            Unload Me
        End If
    End If
End If

If val(AAA) = 1 Then
    FrmResult.Show
    Unload Me
End If
End Sub

Private Sub Timer1_Timer()
Dim X As Long, Y As Long

X = 5160
Y = 6240

If i < 100 Then
    If i Mod 4 = 0 Then
        Shape1.BackColor = RGB(0, 0, 1)
        Shape1.BackColor = RGB(0, 0, 0)
    End If
    W1 = val(i)
    W2 = val(i)
    W3 = val(i)
    W4 = val(i)
    W5 = val(i)
    W6 = val(i)
    W7 = val(i)
    W8 = val(i)
    Line (X + 11 * W1, Y)-(X + (55 / 10) * Sqr(2) * W1, Y + (55 / 10) * Sqr(2) * W1), RGB(255, 255, 255)
    Line (X + (55 / 10) * Sqr(2) * W2, Y + (55 / 10) * Sqr(2) * W2)-(X, Y + 11 * W3), RGB(255, 255, 255)
    Line (X, Y + 11 * W3)-(X - (55 / 10) * Sqr(2) * W4, Y + (55 / 10) * Sqr(2) * W4), RGB(255, 255, 255)
    Line (X - (55 / 10) * Sqr(2) * W4, Y + (55 / 10) * Sqr(2) * W4)-(X - 11 * W5, Y), RGB(255, 255, 255)
    Line (X - 11 * W5, Y)-(X - (55 / 10) * Sqr(2) * W6, Y - (55 / 10) * Sqr(2) * W6), RGB(255, 255, 255)
    Line (X - (55 / 10) * Sqr(2) * W6, Y - (55 / 10) * Sqr(2) * W6)-(X, Y - 11 * W7), RGB(255, 255, 255)
    Line (X, Y - 11 * W7)-(X + (55 / 10) * Sqr(2) * W8, Y - (55 / 10) * Sqr(2) * W8), RGB(255, 255, 255)
    Line (X + (55 / 10) * Sqr(2) * W8, Y - (55 / 10) * Sqr(2) * W8)-(X + 11 * W1, Y), RGB(255, 255, 255)
    
    Line (X + L1 * W1, Y)-(X + (L2 / 2) * Sqr(2) * W1, Y + (L2 / 2) * Sqr(2) * W1), RGB(255, 0, 0)
    Line (X + (L2 / 2) * Sqr(2) * W2, Y + (L2 / 2) * Sqr(2) * W2)-(X, Y + L3 * W3), RGB(255, 0, 0)
    Line (X, Y + L3 * W3)-(X - (L4 / 2) * Sqr(2) * W4, Y + (L4 / 2) * Sqr(2) * W4), RGB(255, 0, 0)
    Line (X - (L4 / 2) * Sqr(2) * W4, Y + (L4 / 2) * Sqr(2) * W4)-(X - L5 * W5, Y), RGB(255, 0, 0)
    Line (X - L5 * W5, Y)-(X - (L6 / 2) * Sqr(2) * W6, Y - (L6 / 2) * Sqr(2) * W6), RGB(255, 0, 0)
    Line (X - (L6 / 2) * Sqr(2) * W6, Y - (L6 / 2) * Sqr(2) * W6)-(X, Y - L7 * W7), RGB(255, 0, 0)
    Line (X, Y - L7 * W7)-(X + (L8 / 2) * Sqr(2) * W8, Y - (L8 / 2) * Sqr(2) * W8), RGB(255, 0, 0)
    Line (X + (L8 / 2) * Sqr(2) * W8, Y - (L8 / 2) * Sqr(2) * W8)-(X + L1 * W1, Y), RGB(255, 0, 0)
    
    Line (X + r1 * W1, Y)-(X + (r2 / 2) * Sqr(2) * W2, Y + (r2 / 2) * Sqr(2) * W2), RGB(0, 255, 255)
    Line (X + (r2 / 2) * Sqr(2) * W2, Y + (r2 / 2) * Sqr(2) * W2)-(X, Y + r3 * W3), RGB(0, 255, 255)
    Line (X, Y + r3 * W3)-(X - (R4 / 2) * Sqr(2) * W4, Y + (R4 / 2) * Sqr(2) * W4), RGB(0, 255, 255)
    Line (X - (R4 / 2) * Sqr(2) * W4, Y + (R4 / 2) * Sqr(2) * W4)-(X - R5 * W5, Y), RGB(0, 255, 255)
    Line (X - R5 * W5, Y)-(X - (R6 / 2) * Sqr(2) * W6, Y - (R6 / 2) * Sqr(2) * W6), RGB(0, 255, 255)
    Line (X - (R6 / 2) * Sqr(2) * W6, Y - (R6 / 2) * Sqr(2) * W6)-(X, Y - R7 * W7), RGB(0, 255, 255)
    Line (X, Y - R7 * W7)-(X + (R8 / 2) * Sqr(2) * W8, Y - (R8 / 2) * Sqr(2) * W8), RGB(0, 255, 255)
    Line (X + (R8 / 2) * Sqr(2) * W8, Y - (R8 / 2) * Sqr(2) * W8)-(X + r1 * W1, Y), RGB(0, 255, 255)
    
    i = i + 1
Else
    Shape1.BackColor = RGB(0, 0, 1)
    Shape1.BackColor = RGB(0, 0, 0)
    W1 = 100
    W2 = 100
    W3 = 100
    W4 = 100
    W5 = 100
    W6 = 100
    W7 = 100
    W8 = 100
    Line (X + 11 * W1, Y)-(X + (55 / 10) * Sqr(2) * W1, Y + (55 / 10) * Sqr(2) * W1), RGB(255, 255, 255)
    Line (X + (55 / 10) * Sqr(2) * W2, Y + (55 / 10) * Sqr(2) * W2)-(X, Y + 11 * W3), RGB(255, 255, 255)
    Line (X, Y + 11 * W3)-(X - (55 / 10) * Sqr(2) * W4, Y + (55 / 10) * Sqr(2) * W4), RGB(255, 255, 255)
    Line (X - (55 / 10) * Sqr(2) * W4, Y + (55 / 10) * Sqr(2) * W4)-(X - 11 * W5, Y), RGB(255, 255, 255)
    Line (X - 11 * W5, Y)-(X - (55 / 10) * Sqr(2) * W6, Y - (55 / 10) * Sqr(2) * W6), RGB(255, 255, 255)
    Line (X - (55 / 10) * Sqr(2) * W6, Y - (55 / 10) * Sqr(2) * W6)-(X, Y - 11 * W7), RGB(255, 255, 255)
    Line (X, Y - 11 * W7)-(X + (55 / 10) * Sqr(2) * W8, Y - (55 / 10) * Sqr(2) * W8), RGB(255, 255, 255)
    Line (X + (55 / 10) * Sqr(2) * W8, Y - (55 / 10) * Sqr(2) * W8)-(X + 11 * W1, Y), RGB(255, 255, 255)
    
    Line (X + L1 * W1, Y)-(X + (L2 / 2) * Sqr(2) * W1, Y + (L2 / 2) * Sqr(2) * W1), RGB(255, 0, 0)
    Line (X + (L2 / 2) * Sqr(2) * W2, Y + (L2 / 2) * Sqr(2) * W2)-(X, Y + L3 * W3), RGB(255, 0, 0)
    Line (X, Y + L3 * W3)-(X - (L4 / 2) * Sqr(2) * W4, Y + (L4 / 2) * Sqr(2) * W4), RGB(255, 0, 0)
    Line (X - (L4 / 2) * Sqr(2) * W4, Y + (L4 / 2) * Sqr(2) * W4)-(X - L5 * W5, Y), RGB(255, 0, 0)
    Line (X - L5 * W5, Y)-(X - (L6 / 2) * Sqr(2) * W6, Y - (L6 / 2) * Sqr(2) * W6), RGB(255, 0, 0)
    Line (X - (L6 / 2) * Sqr(2) * W6, Y - (L6 / 2) * Sqr(2) * W6)-(X, Y - L7 * W7), RGB(255, 0, 0)
    Line (X, Y - L7 * W7)-(X + (L8 / 2) * Sqr(2) * W8, Y - (L8 / 2) * Sqr(2) * W8), RGB(255, 0, 0)
    Line (X + (L8 / 2) * Sqr(2) * W8, Y - (L8 / 2) * Sqr(2) * W8)-(X + L1 * W1, Y), RGB(255, 0, 0)
    
    Line (X + r1 * W1, Y)-(X + (r2 / 2) * Sqr(2) * W2, Y + (r2 / 2) * Sqr(2) * W2), RGB(0, 255, 255)
    Line (X + (r2 / 2) * Sqr(2) * W2, Y + (r2 / 2) * Sqr(2) * W2)-(X, Y + r3 * W3), RGB(0, 255, 255)
    Line (X, Y + r3 * W3)-(X - (R4 / 2) * Sqr(2) * W4, Y + (R4 / 2) * Sqr(2) * W4), RGB(0, 255, 255)
    Line (X - (R4 / 2) * Sqr(2) * W4, Y + (R4 / 2) * Sqr(2) * W4)-(X - R5 * W5, Y), RGB(0, 255, 255)
    Line (X - R5 * W5, Y)-(X - (R6 / 2) * Sqr(2) * W6, Y - (R6 / 2) * Sqr(2) * W6), RGB(0, 255, 255)
    Line (X - (R6 / 2) * Sqr(2) * W6, Y - (R6 / 2) * Sqr(2) * W6)-(X, Y - R7 * W7), RGB(0, 255, 255)
    Line (X, Y - R7 * W7)-(X + (R8 / 2) * Sqr(2) * W8, Y - (R8 / 2) * Sqr(2) * W8), RGB(0, 255, 255)
    Line (X + (R8 / 2) * Sqr(2) * W8, Y - (R8 / 2) * Sqr(2) * W8)-(X + r1 * W1, Y), RGB(0, 255, 255)
    Timer1.Enabled = False
    Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
Dim X As Long, Y As Long

X = 5160
Y = 6240
Line (X + 11 * W1, Y)-(X + (55 / 10) * Sqr(2) * W1, Y + (55 / 10) * Sqr(2) * W1), RGB(255, 255, 255)
Line (X + (55 / 10) * Sqr(2) * W2, Y + (55 / 10) * Sqr(2) * W2)-(X, Y + 11 * W3), RGB(255, 255, 255)
Line (X, Y + 11 * W3)-(X - (55 / 10) * Sqr(2) * W4, Y + (55 / 10) * Sqr(2) * W4), RGB(255, 255, 255)
Line (X - (55 / 10) * Sqr(2) * W4, Y + (55 / 10) * Sqr(2) * W4)-(X - 11 * W5, Y), RGB(255, 255, 255)
Line (X - 11 * W5, Y)-(X - (55 / 10) * Sqr(2) * W6, Y - (55 / 10) * Sqr(2) * W6), RGB(255, 255, 255)
Line (X - (55 / 10) * Sqr(2) * W6, Y - (55 / 10) * Sqr(2) * W6)-(X, Y - 11 * W7), RGB(255, 255, 255)
Line (X, Y - 11 * W7)-(X + (55 / 10) * Sqr(2) * W8, Y - (55 / 10) * Sqr(2) * W8), RGB(255, 255, 255)
Line (X + (55 / 10) * Sqr(2) * W8, Y - (55 / 10) * Sqr(2) * W8)-(X + 11 * W1, Y), RGB(255, 255, 255)

Line (X + L1 * W1, Y)-(X + (L2 / 2) * Sqr(2) * W1, Y + (L2 / 2) * Sqr(2) * W1), RGB(255, 0, 0)
Line (X + (L2 / 2) * Sqr(2) * W2, Y + (L2 / 2) * Sqr(2) * W2)-(X, Y + L3 * W3), RGB(255, 0, 0)
Line (X, Y + L3 * W3)-(X - (L4 / 2) * Sqr(2) * W4, Y + (L4 / 2) * Sqr(2) * W4), RGB(255, 0, 0)
Line (X - (L4 / 2) * Sqr(2) * W4, Y + (L4 / 2) * Sqr(2) * W4)-(X - L5 * W5, Y), RGB(255, 0, 0)
Line (X - L5 * W5, Y)-(X - (L6 / 2) * Sqr(2) * W6, Y - (L6 / 2) * Sqr(2) * W6), RGB(255, 0, 0)
Line (X - (L6 / 2) * Sqr(2) * W6, Y - (L6 / 2) * Sqr(2) * W6)-(X, Y - L7 * W7), RGB(255, 0, 0)
Line (X, Y - L7 * W7)-(X + (L8 / 2) * Sqr(2) * W8, Y - (L8 / 2) * Sqr(2) * W8), RGB(255, 0, 0)
Line (X + (L8 / 2) * Sqr(2) * W8, Y - (L8 / 2) * Sqr(2) * W8)-(X + L1 * W1, Y), RGB(255, 0, 0)

Line (X + r1 * W1, Y)-(X + (r2 / 2) * Sqr(2) * W2, Y + (r2 / 2) * Sqr(2) * W2), RGB(0, 255, 255)
Line (X + (r2 / 2) * Sqr(2) * W2, Y + (r2 / 2) * Sqr(2) * W2)-(X, Y + r3 * W3), RGB(0, 255, 255)
Line (X, Y + r3 * W3)-(X - (R4 / 2) * Sqr(2) * W4, Y + (R4 / 2) * Sqr(2) * W4), RGB(0, 255, 255)
Line (X - (R4 / 2) * Sqr(2) * W4, Y + (R4 / 2) * Sqr(2) * W4)-(X - R5 * W5, Y), RGB(0, 255, 255)
Line (X - R5 * W5, Y)-(X - (R6 / 2) * Sqr(2) * W6, Y - (R6 / 2) * Sqr(2) * W6), RGB(0, 255, 255)
Line (X - (R6 / 2) * Sqr(2) * W6, Y - (R6 / 2) * Sqr(2) * W6)-(X, Y - R7 * W7), RGB(0, 255, 255)
Line (X, Y - R7 * W7)-(X + (R8 / 2) * Sqr(2) * W8, Y - (R8 / 2) * Sqr(2) * W8), RGB(0, 255, 255)
Line (X + (R8 / 2) * Sqr(2) * W8, Y - (R8 / 2) * Sqr(2) * W8)-(X + r1 * W1, Y), RGB(0, 255, 255)

End Sub

Private Sub Timer3_Timer()
AP = Int((101 * Rnd) + 0)
End Sub
