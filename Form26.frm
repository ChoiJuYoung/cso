VERSION 5.00
Begin VB.Form FrmCoupon 
   BackColor       =   &H00FFFFFF&
   Caption         =   "상점 메뉴"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4815
   Icon            =   "Form26.frx":0000
   LinkTopic       =   "Form26"
   ScaleHeight     =   1575
   ScaleWidth      =   4815
   StartUpPosition =   2  '화면 가운데
   Begin CSO.xFrame xFrame2 
      Height          =   1575
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2778
      BackColor       =   8421504
      Caption         =   "쿠폰 사용"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      HeaderGradientBottom=   12611136
      Begin CSO.jcbutton jcbutton9 
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   49152
         Caption         =   "뒤로 가기"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin CSO.jcbutton jcbutton8 
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   49152
         Caption         =   "팀별 선수구매"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin CSO.jcbutton jcbutton7 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   49152
         Caption         =   "Training"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin CSO.jcbutton jcbutton6 
         Height          =   375
         Left            =   3240
         TabIndex        =   6
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   49152
         Caption         =   "Skill Change"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin CSO.jcbutton jcbutton5 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   49152
         Caption         =   "선수 구매"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin CSO.jcbutton jcbutton4 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   49152
         Caption         =   "Lotto"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
   End
   Begin CSO.xFrame xFrame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2778
      BackColor       =   16777215
      Caption         =   "환영합니다."
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      HeaderGradientBottom=   12611136
      Begin CSO.jcbutton jcbutton3 
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16576
         Caption         =   "쿠폰 사용"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin CSO.jcbutton jcbutton1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16576
         Caption         =   "카드 상점"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
   End
End
Attribute VB_Name = "FrmCoupon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub jcbutton1_Click()
 FrmShop.Show
 Unload Me
End Sub


Private Sub jcbutton2_Click()
If val(Money) >= 50000 Then
 MsgBox 로또 & "Cro 당첨"
 Money = val(Money) - 50000
 Money = val(Money) + val(로또)
 MsgBox 로또 & "원 당첨."
Else
 MsgBox "50000Cro가 필요해용"
End If
End Sub

Private Sub jcbutton3_Click()
If val(쿠폰) >= 1 Then
 xFrame1.Visible = False
 xFrame2.Visible = True
Else
 MsgBox "쿠폰이 없습니다 ㅠㅠ...."
End If
End Sub

Private Sub jcbutton4_Click()
If val(쿠폰) >= 1 Then
    쿠폰 = val(쿠폰) - 1
    Money = val(Money) + val(로또)
    MsgBox 로또 & "원 당첨"
Else
    MsgBox "쿠폰 없자나요; 왜그래요 아마추어같이;"
End If
End Sub

Private Sub jcbutton5_Click()
If val(쿠폰) >= 1 Then
    If val(선수수) <= 14 Then
        구매가능 = "Yes"
        쿠폰 = val(쿠폰) - 1
        Do Until 랭크(상점NPC) = "Elite" Or 랭크(상점NPC) = "Normal" Or 랭크(상점NPC) = "Legend" Or 랭크(상점NPC) = "Secret"
            상점NPC = Int((800 * Rnd) + 1)
        Loop
        Dim 쿠폰NPC보조 As Integer
        쿠폰NPC보조 = val(선수수) - 5
        If 구매가능 = "Yes" Then
            SubName(쿠폰NPC보조) = 이름(상점NPC)
            SubTeam(쿠폰NPC보조) = Team(상점NPC)
            SubAt(쿠폰NPC보조) = NPC공격력(상점NPC)
            SubR(쿠폰NPC보조) = NPC견제(상점NPC)
            SubSt(쿠폰NPC보조) = NPC전략(상점NPC)
            SubAm(쿠폰NPC보조) = NPC물량(상점NPC)
            SubDe(쿠폰NPC보조) = NPC수비력(상점NPC)
            SubPa(쿠폰NPC보조) = NPC정찰(상점NPC)
            SubSe(쿠폰NPC보조) = NPC센스(상점NPC)
            SubCo(쿠폰NPC보조) = NPC컨트롤(상점NPC)
            SubRank(쿠폰NPC보조) = 랭크(상점NPC)
            SubYear(쿠폰NPC보조) = OYear(상점NPC)
            SubTribe(쿠폰NPC보조) = 종족(상점NPC)
            SubLev(쿠폰NPC보조) = 1
            SubExp(쿠폰NPC보조) = 0
            SubMExp(쿠폰NPC보조) = 50
            SubPoint(쿠폰NPC보조) = 0
            SubNum(쿠폰NPC보조) = val(상점NPC)
            SubAW(쿠폰NPC보조) = 0
            SubAL(쿠폰NPC보조) = 0
            SubTW(쿠폰NPC보조) = 0
            SubTL(쿠폰NPC보조) = 0
            SubZW(쿠폰NPC보조) = 0
            SubZL(쿠폰NPC보조) = 0
            SubPW(쿠폰NPC보조) = 0
            SubPL(쿠폰NPC보조) = 0
            SubVic(쿠폰NPC보조) = 0
            SubSeVic(쿠폰NPC보조) = 0
            SubCode(쿠폰NPC보조) = "B"
            SubSkill(쿠폰NPC보조) = Skill(상점NPC)
            선수수 = val(선수수) + 1
            구매가능 = "No"
            If SubRank(쿠폰NPC보조) = "Normal" Or SubRank(쿠폰NPC보조) = "Special" Then
                SubNW(쿠폰NPC보조) = "CB16"
            ElseIf SubRank(쿠폰NPC보조) = "Rare" Then
                SubNW(쿠폰NPC보조) = "CA1"
            ElseIf SubRank(쿠폰NPC보조) = "Unique" Then
                SubNW(쿠폰NPC보조) = "CA2"
            ElseIf SubRank(쿠폰NPC보조) = "Elite" Then
                SubNW(쿠폰NPC보조) = "CA3"
            Else
                SubNW(쿠폰NPC보조) = "CS32"
            End If
                If 9.Text1 = "이히히" Then
                    9.Text1 = "히히히"
                Else
                    9.Text1 = "이히히"
                End If
                Unload FrmCoupon
                Unload Me
        End If
    Else
        MsgBox "선수수가 최대입니다."
    End If
Else
    MsgBox "쿠폰 없자나요; 왜그래요 아마추어같이;"
End If
End Sub

Private Sub jcbutton6_Click()
Dim SKillChange As String
SKillChange = 0
If val(쿠폰) >= 1 Then
    Do Until val(SKillChange) >= 1 And val(SKillChange) <= 6
        SKillChange = InputBox("선수번호를 입력하세요. 1 = " & MyYear(1) & MyName(1) & " 2 = " & MyYear(2) & MyName(2) & " 3 = " & MyYear(3) & MyName(3) & " 4 = " & MyYear(4) & MyName(4) & " 5 = " & MyYear(5) & MyName(5) & " 6 = " & MyYear(6) & MyName(6))
    Loop
    
    MySkill(val(SKillChange)) = Int((38 * Rnd) + 1)
    쿠폰 = val(쿠폰) - 1
Else
    MsgBox "쿠폰 없자나요; 왜그래요 아마추어같이;"
End If

End Sub

Private Sub jcbutton7_Click()
If val(쿠폰) >= 1 Then
    쿠폰 = val(쿠폰) - 1
    Dim L As Long
    For i = 1 To 6
        L = Int((50 * Rnd) + 1)
        MyAt(i) = val(MyAt(i)) + L
        L = Int((50 * Rnd) + 1)
        MyR(i) = val(MyR(i)) + L
        L = Int((50 * Rnd) + 1)
        MySt(i) = val(MySt(i)) + L
        L = Int((50 * Rnd) + 1)
        MyAm(i) = val(MyAm(i)) + L
        L = Int((50 * Rnd) + 1)
        MyDe(i) = val(MyDe(i)) + L
        L = Int((50 * Rnd) + 1)
        MyPa(i) = val(MyPa(i)) + L
        L = Int((50 * Rnd) + 1)
        MySe(i) = val(MySe(i)) + L
        L = Int((50 * Rnd) + 1)
        MyCo(i) = val(MyCo(i)) + L
        L = Int((50 * Rnd) + 1)
    Next
    
    For i = 1 To val(선수수 - 5)
        L = Int((50 * Rnd) + 1)
        SubAt(i) = val(SubAt(i)) + L
        L = Int((50 * Rnd) + 1)
        SubR(i) = val(SubR(i)) + L
        L = Int((50 * Rnd) + 1)
        SubSt(i) = val(SubSt(i)) + L
        L = Int((50 * Rnd) + 1)
        SubAm(i) = val(SubAm(i)) + L
        L = Int((50 * Rnd) + 1)
        SubDe(i) = val(SubDe(i)) + L
        L = Int((50 * Rnd) + 1)
        SubPa(i) = val(SubPa(i)) + L
        L = Int((50 * Rnd) + 1)
        SubSe(i) = val(SubSe(i)) + L
        L = Int((50 * Rnd) + 1)
        SubCo(i) = val(SubCo(i)) + L
        L = Int((50 * Rnd) + 1)
    Next
Else
    MsgBox "쿠폰 없자나요; 왜그래요 아마추어같이;"
End If
End Sub

Private Sub jcbutton8_Click()
If val(쿠폰) >= 1 Then
    If val(선수수) <= 14 Then
        Dim 쿠폰팀 As String
        쿠폰팀 = InputBox("원하는 팀 이름을 적어주세요. 삼성전자 eSTRO 공군 MBC POS CJ GO 온게임넷 하이트 STX 르까프 화승 PLUS Mystar 8th 웅진 한빛 SK Orion IS 4U 폭스 Toona Pantech Curitel중에 골라야 합니다. 대소문자 구분은 정확히 해주세요.")
        구매가능 = "Yes"
        쿠폰 = val(쿠폰) - 1
        Do Until (Team(상점NPC) = 쿠폰팀) And (랭크(상점NPC) <> "Champion") And (랭크(상점NPC) <> "Normal")
            상점NPC = Int((800 * Rnd) + 1)
        Loop
        Dim 쿠폰NPC보조 As Integer
        쿠폰NPC보조 = val(선수수) - 5
        If 구매가능 = "Yes" Then
            SubName(쿠폰NPC보조) = 이름(상점NPC)
            SubTeam(쿠폰NPC보조) = Team(상점NPC)
            SubAt(쿠폰NPC보조) = NPC공격력(상점NPC)
            SubR(쿠폰NPC보조) = NPC견제(상점NPC)
            SubSt(쿠폰NPC보조) = NPC전략(상점NPC)
            SubAm(쿠폰NPC보조) = NPC물량(상점NPC)
            SubDe(쿠폰NPC보조) = NPC수비력(상점NPC)
            SubPa(쿠폰NPC보조) = NPC정찰(상점NPC)
            SubSe(쿠폰NPC보조) = NPC센스(상점NPC)
            SubCo(쿠폰NPC보조) = NPC컨트롤(상점NPC)
            SubRank(쿠폰NPC보조) = 랭크(상점NPC)
            SubYear(쿠폰NPC보조) = OYear(상점NPC)
            SubTribe(쿠폰NPC보조) = 종족(상점NPC)
            SubLev(쿠폰NPC보조) = 1
            SubExp(쿠폰NPC보조) = 0
            SubMExp(쿠폰NPC보조) = 50
            SubPoint(쿠폰NPC보조) = 0
            SubNum(쿠폰NPC보조) = val(상점NPC)
            SubAW(쿠폰NPC보조) = 0
            SubAL(쿠폰NPC보조) = 0
            SubTW(쿠폰NPC보조) = 0
            SubTL(쿠폰NPC보조) = 0
            SubZW(쿠폰NPC보조) = 0
            SubZL(쿠폰NPC보조) = 0
            SubPW(쿠폰NPC보조) = 0
            SubPL(쿠폰NPC보조) = 0
            SubVic(쿠폰NPC보조) = 0
            SubSeVic(쿠폰NPC보조) = 0
            SubCode(쿠폰NPC보조) = "B"
            SubSkill(쿠폰NPC보조) = Skill(상점NPC)
            선수수 = val(선수수) + 1
            구매가능 = "No"
            If SubRank(쿠폰NPC보조) = "Normal" Or SubRank(쿠폰NPC보조) = "Special" Then
                SubNW(쿠폰NPC보조) = "CB16"
            ElseIf SubRank(쿠폰NPC보조) = "Rare" Then
                SubNW(쿠폰NPC보조) = "CA1"
            ElseIf SubRank(쿠폰NPC보조) = "Unique" Then
                SubNW(쿠폰NPC보조) = "CA2"
            ElseIf SubRank(쿠폰NPC보조) = "Elite" Then
                SubNW(쿠폰NPC보조) = "CA3"
            Else
                SubNW(쿠폰NPC보조) = "CS32"
            End If
                If 9.Text1 = "이히히" Then
                    9.Text1 = "히히히"
                Else
                    9.Text1 = "이히히"
                End If
                Unload FrmCoupon
                Unload Me
        End If
    Else
        MsgBox "선수수가 최대입니다."
    End If
End If
End Sub

Private Sub jcbutton9_Click()
xFrame1.Visible = True
xFrame2.Visible = False
End Sub
