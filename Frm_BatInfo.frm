VERSION 5.00
Begin VB.Form Frm_BatInfo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   Caption         =   "Battle Information"
   ClientHeight    =   10200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "Frm_BatInfo.frx":0000
   LinkTopic       =   "Form32"
   ScaleHeight     =   10200
   ScaleWidth      =   15240
   StartUpPosition =   2  '화면 가운데
   Begin CSO.jcbutton CmdGo 
      Height          =   255
      Left            =   5520
      TabIndex        =   30
      Top             =   7800
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
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
      Caption         =   "1Set"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin CSO.jcbutton CmdCom 
      Height          =   255
      Left            =   5760
      TabIndex        =   29
      Top             =   8160
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14800597
      Caption         =   "배정 완료"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin CSO.jcbutton CmdSet 
      Height          =   255
      Left            =   5760
      TabIndex        =   28
      Top             =   7800
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14800597
      Caption         =   "Set"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin VB.ComboBox CmbSet 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      Style           =   2  '드롭다운 목록
      TabIndex        =   6
      Top             =   3120
      Width           =   4095
   End
   Begin VB.Label lblSet 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   5760
      TabIndex        =   27
      Top             =   7080
      Width           =   3855
   End
   Begin VB.Label lblSet 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5760
      TabIndex        =   26
      Top             =   6720
      Width           =   3855
   End
   Begin VB.Label lblSet 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   5760
      TabIndex        =   25
      Top             =   6360
      Width           =   3855
   End
   Begin VB.Label lblSet 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5760
      TabIndex        =   24
      Top             =   6000
      Width           =   3855
   End
   Begin VB.Label lblSet 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   23
      Top             =   5640
      Width           =   3855
   End
   Begin VB.Label lblSet 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   22
      Top             =   5280
      Width           =   3855
   End
   Begin VB.Label lblSet 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   5760
      TabIndex        =   21
      Top             =   7440
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   72
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   7320
      TabIndex        =   20
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblOP 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   72
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   9720
      TabIndex        =   19
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblMP 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   72
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3720
      TabIndex        =   18
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblPLInfo 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "프로리그 - 포스트 시즌 준 플레이 오프"
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
      Left            =   3720
      TabIndex        =   17
      Top             =   360
      Width           =   7695
   End
   Begin VB.Label lblMName 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
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
      Left            =   7320
      TabIndex        =   16
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label lblMPT 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
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
      Left            =   7320
      TabIndex        =   15
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label lblMZP 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
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
      Left            =   7320
      TabIndex        =   14
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label lblMTZ 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
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
      Left            =   7320
      TabIndex        =   13
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Image ImgMap 
      Height          =   1500
      Left            =   5760
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Label lblLS 
      BackStyle       =   0  '투명
      Caption         =   "준우승 : "
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
      Left            =   2400
      TabIndex        =   12
      Top             =   9480
      Width           =   2895
   End
   Begin VB.Label lblLV 
      BackStyle       =   0  '투명
      Caption         =   "우승 : "
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
      Left            =   2400
      TabIndex        =   11
      Top             =   9120
      Width           =   2895
   End
   Begin VB.Label lblLA 
      BackStyle       =   0  '투명
      Caption         =   "vs All : "
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
      Left            =   2400
      TabIndex        =   10
      Top             =   8760
      Width           =   2895
   End
   Begin VB.Label lblLP 
      BackStyle       =   0  '투명
      Caption         =   "vs P : "
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
      Left            =   2400
      TabIndex        =   9
      Top             =   8400
      Width           =   2895
   End
   Begin VB.Label lblLZ 
      BackStyle       =   0  '투명
      Caption         =   "vs Z : "
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
      Left            =   2400
      TabIndex        =   8
      Top             =   8040
      Width           =   2895
   End
   Begin VB.Label lblLT 
      BackStyle       =   0  '투명
      Caption         =   "vs T : "
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
      Left            =   2400
      TabIndex        =   7
      Top             =   7680
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      Height          =   2295
      Left            =   2280
      Top             =   7560
      Width           =   3135
   End
   Begin VB.Label lblTriR 
      BackStyle       =   0  '투명
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
      Left            =   11280
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblNameR 
      BackStyle       =   0  '투명
      Caption         =   "<10>이영호[Ex]"
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
      Left            =   9840
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblTriL 
      BackStyle       =   0  '투명
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
      Left            =   5280
      TabIndex        =   3
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblNameL 
      BackStyle       =   0  '투명
      Caption         =   "<10>이영호[Ex]"
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
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Image ImgSelR 
      Height          =   1500
      Left            =   9840
      Top             =   2880
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image ImgSelL 
      Height          =   1500
      Left            =   3840
      Top             =   2880
      Width           =   1500
   End
   Begin VB.Image Img_RPlay 
      Height          =   1500
      Index           =   5
      Left            =   13560
      Top             =   6840
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Img_RPlay 
      Height          =   1500
      Index           =   4
      Left            =   13560
      Top             =   5280
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Img_RPlay 
      Height          =   1500
      Index           =   3
      Left            =   13560
      Top             =   3720
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Img_RPlay 
      Height          =   1500
      Index           =   2
      Left            =   13560
      Top             =   2160
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Img_RPlay 
      Height          =   1500
      Index           =   1
      Left            =   13560
      Top             =   600
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Img_RPlay 
      Height          =   1500
      Index           =   6
      Left            =   13560
      Top             =   8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblTeamR 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   13080
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image Img_LPlay 
      Height          =   1500
      Index           =   5
      Left            =   120
      Top             =   6840
      Width           =   1500
   End
   Begin VB.Image Img_LPlay 
      Height          =   1500
      Index           =   4
      Left            =   120
      Top             =   5280
      Width           =   1500
   End
   Begin VB.Image Img_LPlay 
      Height          =   1500
      Index           =   3
      Left            =   120
      Top             =   3720
      Width           =   1500
   End
   Begin VB.Image Img_LPlay 
      Height          =   1500
      Index           =   2
      Left            =   120
      Top             =   2160
      Width           =   1500
   End
   Begin VB.Image Img_LPlay 
      Height          =   1500
      Index           =   1
      Left            =   120
      Top             =   600
      Width           =   1500
   End
   Begin VB.Image Img_LPlay 
      Height          =   1500
      Index           =   6
      Left            =   120
      Top             =   8400
      Width           =   1500
   End
   Begin VB.Label lblTeamL 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Frm_BatInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmbSet_Click()
SetNum = CmbSet.ListIndex + 1
lblMName = MapName(MapL(SetNum))
lblMTZ = "T vs Z = " & TZT(MapL(SetNum)) & " vs " & TZZ(MapL(SetNum))
lblMZP = "Z vs P = " & ZPZ(MapL(SetNum)) & " vs " & ZPP(MapL(SetNum))
lblMPT = "P vs T = " & PTP(MapL(SetNum)) & " vs " & PTT(MapL(SetNum))
Call LoadMapImg(ImgMap, MapName(MapL(SetNum)))

If CmdSet.Visible = False Then
    Call MakeLine(Me, SetL(SetNum), 4380, 6120)
    Call LoadImage(ImgSelL, MyName(SetL(SetNum)), MyYear(SetL(SetNum)))
    Call lblNameAlter(lblNameL, 1, SetL(SetNum))
    Call lblTribeAlter(lblTriL, val(MyTribe(SetL(SetNum))))

    Call MakeLineCom(Me, SetR(SetNum), 10740, 6120)
    Call LoadImage(ImgSelR, 이름(SetR(SetNum)), OYear(SetR(SetNum)))
    Call lblNameAlter(lblNameR, 2, SetR(SetNum))
    Call lblTribeAlter(lblTriR, val(종족(SetR(SetNum))))
End If
End Sub

Private Sub CmdCom_Click()
CmdSet.Visible = False
CmdCom.Visible = False
CmdGO.Visible = True
CmdGO.Caption = 진행Set & "Set : " & MyYear(SetL(진행Set)) & MyName(SetL(진행Set)) & " vs " & OYear(SetR(진행Set)) & 이름(SetR(진행Set)) & "[" & MapName(MapL(진행Set)) & "]"
Call MakeLineCom(Me, SetR(SetNum), 10740, 6120)
Call LoadImage(ImgSelR, 이름(SetR(SetNum)), OYear(SetR(SetNum)))
Call lblNameAlter(lblNameR, 2, SetR(SetNum))
Call lblTribeAlter(lblTriR, val(종족(SetR(SetNum))))

If PL진행 = "1R" Or PL진행 = "2R" Or PL진행 = "3R" Then
    For i = 1 To 5
        Call LoadImage(Img_LPlay(i), MyName(SetL(i)), MyYear(SetL(i)))
    Next
    Img_LPlay(6).Visible = False
Else
    For i = 1 To 6
        Call LoadImage(Img_LPlay(i), MyName(SetL(i)), MyYear(SetL(i)))
    Next
End If
For i = 1 To 6
    Img_RPlay(i).Visible = True
Next
ImgSelR.Visible = True
lblNameR.Visible = True
lblTriR.Visible = True
End Sub

Private Sub CmdGo_Click()
선택 = SetL(진행Set)
Oee = SetR(진행Set)

MN = MyName(선택)
Year = MyYear(선택)
AT = MyAt(선택)
R = MyR(선택)
St = MySt(선택)
Am = MyAm(선택)
De = MyDe(선택)
Pa = MyPa(선택)
SE = MySe(선택)
Co = MyCo(선택)
MT = MyTribe(선택)

FrmLoading.Show
Me.Visible = False
End Sub

Private Sub CmdSet_Click()
On Error GoTo Some:
If SetSel(SelLNum) = True Then
    If SetL(SetNum) <> 0 Then
        SetSel(SetL(SetNum)) = True
    End If
    SetSel(SelLNum) = False
    SetL(SetNum) = SelLNum
    lblSet(SetNum) = SetNum & "Set : " & MyYear(SetL(SetNum)) & MyName(SetL(SetNum)) & "[" & MapName(MapL(SetNum)) & "]"
Else
    MsgBox MyYear(SelLNum) & MyName(SelLNum) & "선수는 이미 배정된 선수입니다."
End If

If PL진행 = "1R" Or PL진행 = "2R" Or PL진행 = "3R" Then
    For i = 1 To 5
        If lblSet(i) = "" Then
            완료여부 = False
            Exit For
        End If
        완료여부 = True
    Next
    If 완료여부 = True Then
        CmdCom.Visible = True
    End If
Else
    For i = 1 To 6
        If lblSet(i) = "" Then
            완료여부 = False
            Exit For
        End If
        완료여부 = True
    Next
    If 완료여부 = True Then
        CmdCom.Visible = True
    End If
End If
Exit Sub

Some:
MsgBox "세트를 선택해 주세요."
End Sub

Private Sub Form_Load()
'초기 세팅들
진행Set = 1
lblMP = MW
lblOP = OW

For i = 1 To 7
    lblSet(i) = ""
    SetR(i) = 1
Next
For i = 1 To 6
    SetSel(i) = True
    SetL(i) = 0
Next


Me.Hide

CmbSet.AddItem ("1Set")
CmbSet.AddItem ("2Set")
CmbSet.AddItem ("3Set")
CmbSet.AddItem ("4Set")
CmbSet.AddItem ("5Set")

'1Set 맵
MapL(1) = Int((12 * Rnd) + 1)

'2Set 맵
Do Until MapL(2) <> MapL(1)
    MapL(2) = Int((12 * Rnd) + 1)
Loop

'3Set 맵
Do Until MapL(3) <> MapL(2) And MapL(3) <> MapL(1)
    MapL(3) = Int((12 * Rnd) + 1)
Loop

'4Set 맵
Do Until MapL(4) <> MapL(3) And MapL(4) <> MapL(2) And MapL(4) <> MapL(1)
    MapL(4) = Int((12 * Rnd) + 1)
Loop

'5Set 맵
Do Until MapL(5) <> MapL(4) And MapL(5) <> MapL(3) And MapL(5) <> MapL(2) And MapL(5) <> MapL(1)
    MapL(5) = Int((12 * Rnd) + 1)
Loop
    

If PL진행 = "1R" Or PL진행 = "2R" Or PL진행 = "3R" Then
Else
    CmbSet.AddItem ("6Set")
    Do Until MapL(6) <> MapL(5) And MapL(6) <> MapL(4) And MapL(6) <> MapL(3) And MapL(6) <> MapL(2) And MapL(6) <> MapL(1)
        MapL(6) = Int((12 * Rnd) + 1)
    Loop
    
    Do Until MapL(7) <> MapL(6) And MapL(7) <> MapL(5) And MapL(7) <> MapL(4) And MapL(7) <> MapL(3) And MapL(7) <> MapL(2) And MapL(7) <> MapL(1)
        MapL(7) = Int((12 * Rnd) + 1)
    Loop
End If

For i = 1 To 6
    Call LoadImage(Img_LPlay(i), MyName(i), MyYear(i))
Next
Call MakeLine(Me, 1, 4380, 6120)
Call LoadImage(ImgSelL, MyName(1), MyYear(1))
Call lblNameAlter(lblNameL, 1, 1)
Call lblTribeAlter(lblTriL, val(MyTribe(1)))
SelLNum = 1
CmdSet.Caption = "1Set : " & MyYear(1) & MyName(1)

'팀이름 입력부
lblTeamL = TeamName
If PL경기수 = 0 Then
    lblTeamR = "Vs 삼성"
ElseIf PL경기수 = 1 Then
    lblTeamR = "Vs eSTRO & 공군"
ElseIf PL경기수 = 2 Then
    lblTeamR = "Vs MBC"
ElseIf PL경기수 = 3 Then
    lblTeamR = "Vs CJ"
ElseIf PL경기수 = 4 Then
    lblTeamR = "Vs Hite"
ElseIf PL경기수 = 5 Then
    lblTeamR = "Vs STX"
ElseIf PL경기수 = 6 Then
    lblTeamR = "Vs Oz"
ElseIf PL경기수 = 7 Then
    lblTeamR = "Vs Mystar & 8th"
ElseIf PL경기수 = 8 Then
    lblTeamR = "Vs 웅진"
ElseIf PL경기수 = 9 Then
    lblTeamR = "Vs SK"
ElseIf PL경기수 = 10 Then
    lblTeamR = "Vs KT"
ElseIf PL경기수 = 11 Then
    FrmMain.CmdSa.Visible = True
    lblTeamR = "Vs 폭스"
End If

Call RandomOee(SetR(1), val(PL경기수))
Call RandomOee(SetR(2), val(PL경기수))
Call RandomOee(SetR(3), val(PL경기수))
Call RandomOee(SetR(4), val(PL경기수))
Call RandomOee(SetR(5), val(PL경기수))
Call RandomOee(SetR(6), val(PL경기수))

''''///컴퓨터 선수 배정
Call RandomOee(SetR(1), val(PL경기수))

'2Set 선수
Do Until (SetR(2) <> SetR(1))
    Call RandomOee(SetR(2), val(PL경기수))
    DoEvents
Loop

'3Set 선수
Do Until SetR(3) <> SetR(2) And SetR(3) <> SetR(1)
    Call RandomOee(SetR(3), val(PL경기수))
    DoEvents
Loop

'4Set 선수
Do Until SetR(4) <> SetR(3) And SetR(4) <> SetR(2) And SetR(4) <> SetR(1)
    Call RandomOee(SetR(4), val(PL경기수))
    DoEvents
Loop

'5Set 선수
Do Until SetR(5) <> SetR(4) And SetR(5) <> SetR(3) And SetR(5) <> SetR(2) And SetR(5) <> SetR(1)
    Call RandomOee(SetR(5), val(PL경기수))
    DoEvents
Loop

'6Set 선수 & 5Set 선수 수정
If PL진행 = "1R" Or PL진행 = "2R" Or PL진행 = "3R" Then
    If PL경기수 = 0 Then
        SetR(5) = 136
    ElseIf PL경기수 = 1 Then
        SetR(5) = 600
    ElseIf PL경기수 = 2 Then
        SetR(5) = 569 Or SetR(5) = 102 Or SetR(5) = 104
    ElseIf PL경기수 = 3 Then
        SetR(5) = 209 Or SetR(5) = 552 Or SetR(5) = 568
    ElseIf PL경기수 = 4 Then
        SetR(5) = 370
    ElseIf PL경기수 = 5 Then
        SetR(5) = 437
    ElseIf PL경기수 = 6 Then
        SetR(5) = 495 Or SetR(5) = 638
    ElseIf PL경기수 = 7 Then
        SetR(5) = 722 Or SetR(5) = 723 Or 800
    ElseIf PL경기수 = 8 Then
        SetR(5) = 560
    ElseIf PL경기수 = 9 Then
        SetR(5) = 540 Or SetR(5) = 544 Or SetR(5) = 547 Or SetR(5) = 553
    ElseIf PL경기수 = 10 Then
        SetR(5) = 649 Or SetR(5) = 109
    ElseIf PL경기수 = 11 Then
        SetR(5) = 549 Or SetR(5) = 585
    ElseIf PL경기수 = 12 Then
        SetR(5) = 719 Or SetR(5) = 713
    Else
        SetR(5) = Int((801 * Rnd) + 0)
    End If
Else
    Do Until SetR(6) <> SetR(5) And SetR(6) <> SetR(4) And SetR(6) <> SetR(3) And SetR(6) <> SetR(2) And SetR(6) <> SetR(1)
        Call RandomOee(SetR(6), val(PL경기수))
    DoEvents
    Loop
End If

'7Set 선수
If PL경기수 = 0 Then
    SetR(7) = 136
ElseIf PL경기수 = 1 Then
    SetR(7) = 600
ElseIf PL경기수 = 2 Then
    SetR(7) = 569 Or SetR(7) = 102 Or SetR(7) = 104
ElseIf PL경기수 = 3 Then
    SetR(7) = 209 Or SetR(7) = 552 Or SetR(7) = 568
ElseIf PL경기수 = 4 Then
    SetR(7) = 370
ElseIf PL경기수 = 5 Then
    SetR(7) = 437
ElseIf PL경기수 = 6 Then
    SetR(7) = 495 Or SetR(7) = 638
ElseIf PL경기수 = 7 Then
    SetR(7) = 722 Or SetR(7) = 723 Or 800
ElseIf PL경기수 = 8 Then
    SetR(7) = 560
ElseIf PL경기수 = 9 Then
    SetR(7) = 540 Or SetR(7) = 544 Or SetR(7) = 547 Or SetR(7) = 553
ElseIf PL경기수 = 10 Then
    SetR(7) = 649 Or SetR(7) = 109
ElseIf PL경기수 = 11 Then
    SetR(7) = 549 Or SetR(7) = 585
ElseIf PL경기수 = 12 Then
    SetR(7) = 719 Or SetR(7) = 713
Else
    SetR(7) = Int((801 * Rnd) + 0)
End If

Dim n As Integer
n = 6
If PL진행 = "1R" Or PL진행 = "2R" Or PL진행 = "3R" Then
    n = 5
End If

For i = 1 To n
    Call LoadImage(Img_RPlay(i), 이름(SetR(i)), OYear(SetR(i)))
Next

Me.Show



'스탯 입력
lblLT = "vs T : " & MyTW(1) & "승 " & MyTL(1) & "패"
lblLZ = "vs Z : " & MyZW(1) & "승 " & MyZL(1) & "패"
lblLP = "vs P : " & MyPW(1) & "승 " & MyPL(1) & "패"
lblLA = "vs A : " & MyAW(1) & "승 " & MyAL(1) & "패"
lblLV = "우승 : " & MyVic(1)
lblLS = "준우승 : " & MySeVic(1)
End Sub

Private Sub Img_LPlay_Click(Index As Integer)
If CmdSet.Visible = True Then
    SelLNum = Index
    CmdSet.Caption = SetNum & "Set : " & MyYear(Index) & MyName(Index)
    '기본 정보 입력
    Call MakeLine(Me, Index, 4380, 6120)
    Call LoadImage(ImgSelL, MyName(Index), MyYear(Index))
    Call lblNameAlter(lblNameL, 1, Index)
    Call lblTribeAlter(lblTriL, val(MyTribe(Index)))
    
    '스탯 입력
    lblLT = "vs T : " & MyTW(Index) & "승 " & MyTL(Index) & "패"
    lblLZ = "vs Z : " & MyZW(Index) & "승 " & MyZL(Index) & "패"
    lblLP = "vs P : " & MyPW(Index) & "승 " & MyPL(Index) & "패"
    lblLA = "vs A : " & MyAW(Index) & "승 " & MyAL(Index) & "패"
    lblLV = "우승 : " & MyVic(Index)
    lblLS = "준우승 : " & MySeVic(Index)
End If
End Sub
