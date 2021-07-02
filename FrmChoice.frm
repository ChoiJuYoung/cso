VERSION 5.00
Begin VB.Form FrmChoice 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Choice Cards"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11415
   Icon            =   "FrmChoice.frx":0000
   LinkTopic       =   "Form32"
   ScaleHeight     =   6765
   ScaleWidth      =   11415
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer TimOee 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6960
      Top             =   4920
   End
   Begin VB.Timer TimRan 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   4920
   End
   Begin CSO.jcbutton CmdGO 
      Height          =   255
      Left            =   6480
      TabIndex        =   2
      Top             =   6240
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
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
      BackColor       =   14935011
      Caption         =   "Start!"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin VB.TextBox TxtName 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6480
      TabIndex        =   1
      Text            =   "Player NickName"
      Top             =   5880
      Width           =   2415
   End
   Begin CSO.jcbutton CmdSel 
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   6240
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
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
      Caption         =   "Select"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Label lblName 
      BackStyle       =   0  '투명
      Caption         =   "NAme"
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
      Index           =   5
      Left            =   8160
      TabIndex        =   13
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  '투명
      Caption         =   "NAme"
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
      Index           =   4
      Left            =   6480
      TabIndex        =   12
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  '투명
      Caption         =   "NAme"
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
      Index           =   3
      Left            =   9840
      TabIndex        =   11
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  '투명
      Caption         =   "NAme"
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
      Index           =   2
      Left            =   8160
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  '투명
      Caption         =   "NAme"
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
      Index           =   1
      Left            =   6480
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  '투명
      Caption         =   "NAme"
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
      Index           =   6
      Left            =   9840
      TabIndex        =   8
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblSum 
      BackStyle       =   0  '투명
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
      Left            =   9000
      TabIndex        =   7
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label lblPTribe 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
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
      Left            =   4560
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblPSum 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
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
      Left            =   4560
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblPRank 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
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
      Left            =   4560
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblPName 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
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
      Left            =   4560
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.Image ImgCon 
      Height          =   1500
      Index           =   5
      Left            =   8160
      Top             =   3360
      Width           =   1500
   End
   Begin VB.Image ImgCon 
      Height          =   1500
      Index           =   4
      Left            =   6480
      Top             =   3360
      Width           =   1500
   End
   Begin VB.Image ImgCon 
      Height          =   1500
      Index           =   3
      Left            =   9840
      Top             =   360
      Width           =   1500
   End
   Begin VB.Image ImgCon 
      Height          =   1500
      Index           =   2
      Left            =   8160
      Top             =   360
      Width           =   1500
   End
   Begin VB.Image ImgCon 
      Height          =   1500
      Index           =   1
      Left            =   6480
      Top             =   360
      Width           =   1500
   End
   Begin VB.Image ImgCon 
      Height          =   1500
      Index           =   6
      Left            =   9840
      Top             =   3360
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   2880
      Top             =   360
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  '투명하지 않음
      Height          =   6735
      Left            =   2160
      Top             =   0
      Width           =   4215
   End
   Begin VB.Image ImgChoice 
      Height          =   1500
      Index           =   2
      Left            =   120
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Image ImgChoice 
      Height          =   1500
      Index           =   1
      Left            =   120
      Top             =   120
      Width           =   1500
   End
   Begin VB.Image ImgChoice 
      Height          =   1500
      Index           =   3
      Left            =   120
      Top             =   5040
      Width           =   1500
   End
End
Attribute VB_Name = "FrmChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private 클릭 As Integer

Private Sub CmdGo_Click()
For 갠리그 = 1 To 6
    If MyRank(갠리그) = "Normal" Or MyRank(갠리그) = "Special" Then
        MyNW(갠리그) = "CB16"
    ElseIf MyRank(갠리그) = "Rare" Then
        MyNW(갠리그) = "CA1"
    ElseIf MyRank(갠리그) = "Unique" Then
        MyNW(갠리그) = "CA2"
    ElseIf MyRank(갠리그) = "Elite" Then
        MyNW(갠리그) = "CA3"
    Else
        MyNW(갠리그) = "CS32"
    End If
Next
하향 = 0
하향횟수 = 0

Money = 5000
TeamName = TxtName

FrmMain.Show
Unload Me
End Sub

Private Sub CmdSel_Click()
MyName(선수수) = 이름(명단(클릭))
MyTribe(선수수) = 종족(명단(클릭))
MyAt(선수수) = 공격력(명단(클릭))
MyR(선수수) = 견제(명단(클릭))
MySt(선수수) = 전략(명단(클릭))
MyAm(선수수) = 물량(명단(클릭))
MyDe(선수수) = 수비력(명단(클릭))
MyPa(선수수) = 정찰(명단(클릭))
MySe(선수수) = 센스(명단(클릭))
MyCo(선수수) = 컨트롤(명단(클릭))
MyYear(선수수) = OYear(명단(클릭))
MyRank(선수수) = 랭크(명단(클릭))
MyTeam(선수수) = Team(명단(클릭))
MySkill(선수수) = Skill(명단(클릭))
PlayNumber(선수수) = Oee
Randomize Oee
Call LoadImage(ImgCon(선수수), MyName(선수수), MyYear(선수수))
Number = Number + val(공격력(명단(클릭))) + val(견제(명단(클릭))) + val(전략(명단(클릭))) + val(물량(명단(클릭))) + val(수비력(명단(클릭))) + val(정찰(명단(클릭))) + val(센스(명단(클릭))) + val(컨트롤(명단(클릭)))
Call lblNameAlter(lblName(선수수), 1, val(선수수))
DoEvents
If lblName(선수수).ForeColor = RGB(255, 255, 255) Then
    lblName(선수수).ForeColor = RGB(0, 0, 0)
End If


선수수 = val(선수수) + 1
If val(선수수) >= 7 Then
    CmdGo.Visible = True
    CmdSel.Visible = False
    lblSum = "능력치 평균 : " & Int(Number / 6)
    선수수 = 6
Else
    TimOee.Enabled = False
    TimRan.Enabled = True
End If
End Sub

Private Sub Form_Load()
Number = 0
Oee = Int((800 * Rnd) + 1)
선수수 = 1
If 추첨경우 = 1 Then
    For i = 1 To 3
        Do Until (랭크(Oee) = "Unique") And (종족(Oee) = 1)
            Oee = Int((800 * Rnd) + 1)
        Loop
        명단(i) = Oee
        Oee = Int((800 * Rnd) + 1)
    Next
ElseIf 추첨경우 = 2 Then
    For i = 1 To 3
        Do Until (랭크(Oee) = "Unique") And (종족(Oee) = 1)
            Oee = Int((800 * Rnd) + 1)
        Loop
        명단(i) = Oee
        Oee = Int((800 * Rnd) + 1)
    Next
ElseIf 추첨경우 = 3 Then
    For i = 1 To 3
        Do Until 랭크(Oee) = "Special" And 종족(Oee) = 1
            Oee = Int((800 * Rnd) + 1)
        Loop
        명단(i) = Oee
        Oee = Int((800 * Rnd) + 1)
    Next
ElseIf 추첨경우 = 4 Then
    For i = 1 To 3
        Do Until 랭크(Oee) = "Rare" And 종족(Oee) = 1
            Oee = Int((800 * Rnd) + 1)
        Loop
        명단(i) = Oee
        Oee = Int((800 * Rnd) + 1)
    Next
ElseIf 추첨경우 = 5 Then
    For i = 1 To 3
        Do Until 랭크(Oee) = "Rare" And 종족(Oee) = 1
            Oee = Int((800 * Rnd) + 1)
        Loop
        명단(i) = Oee
        Oee = Int((800 * Rnd) + 1)
    Next
Else
    For i = 1 To 3
        Do Until 랭크(Oee) = "Special" And 종족(Oee) = 1
            Oee = Int((800 * Rnd) + 1)
        Loop
        명단(i) = Oee
        Oee = Int((800 * Rnd) + 1)
    Next
End If

For i = 1 To 3
    Call LoadImage(ImgChoice(i), 이름(명단(i)), OYear(명단(i)))
Next

TimOee.Enabled = True
End Sub

Private Sub ImgChoice_Click(Index As Integer)
Shape1.BackColor = &HC0C0C1
Shape1.BackColor = &HC0C0C0
클릭 = Index
DoEvents
Call MakeLineCom(Me, 명단(Index), 4320, 3840)
Call LoadImage(Image1, 이름(명단(Index)), OYear(명단(Index)))
lblPName = OYear(명단(Index)) & 이름(명단(Index))
lblPRank = 랭크(명단(Index))
Call lblTribeAlter(lblPTribe, val(종족(명단(Index))))
lblPTribe = "종족 : " & Left(Right(lblPTribe, 2), 1)
lblPSum = "능력치 합계 : " & val(공격력(명단(Index))) + val(견제(명단(Index))) + val(전략(명단(Index))) + val(물량(명단(Index))) + val(수비력(명단(Index))) + val(정찰(명단(Index))) + val(센스(명단(Index))) + val(컨트롤(명단(Index)))
End Sub

Private Sub TimOee_Timer()
Oee = Int((800 * Rnd) + 1)
End Sub

Private Sub TimRan_Timer()
TimRan.Enabled = False

If 선수수 = 2 Then
    If 추첨경우 = 1 Then
        For i = 1 To 3
            Do Until (랭크(Oee) = "Rare") And (종족(Oee) = 2)
                Oee = Int((800 * Rnd) + 1)
            Loop
            명단(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    ElseIf 추첨경우 = 2 Then
        For i = 1 To 3
            Do Until (랭크(Oee) = "Special") And (종족(Oee) = 2)
                Oee = Int((800 * Rnd) + 1)
            Loop
            명단(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    ElseIf 추첨경우 = 3 Then
        For i = 1 To 3
            Do Until 랭크(Oee) = "Unique" And 종족(Oee) = 2
                Oee = Int((800 * Rnd) + 1)
            Loop
            명단(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    ElseIf 추첨경우 = 4 Then
        For i = 1 To 3
            Do Until 랭크(Oee) = "Unique" And 종족(Oee) = 2
                Oee = Int((800 * Rnd) + 1)
            Loop
            명단(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    ElseIf 추첨경우 = 5 Then
        For i = 1 To 3
            Do Until 랭크(Oee) = "Special" And 종족(Oee) = 2
                Oee = Int((800 * Rnd) + 1)
            Loop
            명단(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    Else
        For i = 1 To 3
            Do Until 랭크(Oee) = "Rare" And 종족(Oee) = 2
                Oee = Int((800 * Rnd) + 1)
            Loop
            명단(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    End If
ElseIf 선수수 = 3 Then
    If 추첨경우 = 1 Then
        For i = 1 To 3
            Do Until (랭크(Oee) = "Special") And (종족(Oee) = 3)
                Oee = Int((800 * Rnd) + 1)
            Loop
            명단(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    ElseIf 추첨경우 = 2 Then
        For i = 1 To 3
            Do Until (랭크(Oee) = "Rare") And (종족(Oee) = 3)
                Oee = Int((800 * Rnd) + 1)
            Loop
            명단(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    ElseIf 추첨경우 = 3 Then
        For i = 1 To 3
            Do Until 랭크(Oee) = "Rare" And 종족(Oee) = 3
                Oee = Int((800 * Rnd) + 1)
            Loop
            명단(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    ElseIf 추첨경우 = 4 Then
        For i = 1 To 3
            Do Until 랭크(Oee) = "Special" And 종족(Oee) = 3
                Oee = Int((800 * Rnd) + 1)
            Loop
            명단(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    ElseIf 추첨경우 = 5 Then
        For i = 1 To 3
            Do Until 랭크(Oee) = "Unique" And 종족(Oee) = 3
                Oee = Int((800 * Rnd) + 1)
            Loop
            명단(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    Else
        For i = 1 To 3
            Do Until 랭크(Oee) = "Unique" And 종족(Oee) = 3
                Oee = Int((800 * Rnd) + 1)
            Loop
            명단(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    End If
ElseIf 선수수 >= 4 Then
    For i = 1 To 3
        Do Until (랭크(Oee) = "Normal") And (종족(Oee) = val(선수수) - 3)
            Oee = Int((800 * Rnd) + 1)
        Loop
        명단(i) = Oee
        Oee = Int((800 * Rnd) + 1)
    Next
End If

For i = 1 To 3
    Call LoadImage(ImgChoice(i), 이름(명단(i)), OYear(명단(i)))
Next

End Sub
