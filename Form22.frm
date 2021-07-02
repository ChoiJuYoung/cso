VERSION 5.00
Begin VB.Form FrmFire 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Card Delete"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6405
   Icon            =   "Form22.frx":0000
   LinkTopic       =   "Form22"
   ScaleHeight     =   2430
   ScaleWidth      =   6405
   StartUpPosition =   2  '화면 가운데
   Begin CSO.jcbutton Command10 
      Height          =   255
      Left            =   1320
      TabIndex        =   18
      Top             =   2040
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "선택"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin CSO.jcbutton Command8 
      Height          =   255
      Left            =   1320
      TabIndex        =   17
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "선택"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin CSO.jcbutton Command7 
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "선택"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin CSO.jcbutton Command6 
      Height          =   255
      Left            =   1320
      TabIndex        =   15
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "선택"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin CSO.jcbutton Command5 
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "선택"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin CSO.jcbutton Command4 
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "선택"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin CSO.jcbutton Command3 
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "선택"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin CSO.jcbutton command2 
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "선택"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin CSO.jcbutton Command1 
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "선택"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin CSO.jcbutton jcbutton1 
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   33023
      Caption         =   "Fire!!!"
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   3120
      Picture         =   "Form22.frx":628A
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label lblS 
      BackStyle       =   0  '투명
      Caption         =   "<11>이영호"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblS 
      BackStyle       =   0  '투명
      Caption         =   "<11>이영호"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblS 
      BackStyle       =   0  '투명
      Caption         =   "<11>이영호"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblS 
      BackStyle       =   0  '투명
      Caption         =   "<11>이영호"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblS 
      BackStyle       =   0  '투명
      Caption         =   "<11>이영호"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblS 
      BackStyle       =   0  '투명
      Caption         =   "<11>이영호"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblS 
      BackStyle       =   0  '투명
      Caption         =   "<11>이영호"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblS 
      BackStyle       =   0  '투명
      Caption         =   "<11>이영호"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblS 
      BackStyle       =   0  '투명
      Caption         =   "<11>이영호"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "FrmFire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
선택 = 1
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(선택) & ".gif")
End If
End Sub

Private Sub Command2_Click()
선택 = 2
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(선택) & ".gif")
End If
End Sub

Private Sub Command3_Click()
선택 = 3
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(선택) & ".gif")
End If
End Sub

Private Sub Command4_Click()
선택 = 4
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(선택) & ".gif")
End If
End Sub

Private Sub Command5_Click()
선택 = 5
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(선택) & ".gif")
End If
End Sub

Private Sub Command6_Click()
선택 = 6
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(선택) & ".gif")
End If
End Sub

Private Sub Command7_Click()
선택 = 7
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(선택) & ".gif")
End If
End Sub

Private Sub Command8_Click()
선택 = 8
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(선택) & ".gif")
End If
End Sub

Private Sub Command9_Click()
선택 = 9
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(선택), 2, 2) & "]" & SubName(선택) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(선택) & ".gif")
End If
End Sub

Private Sub Form_Load()
For 세팅 = 0 To 8
 lblS(세팅) = SubYear(세팅 + 1) & SubName(세팅 + 1)
Next
Command1_Click
End Sub

Private Sub jcbutton1_Click()
If SubRank(선택) = "Normal" Then
 Money = val(Money) + 500
ElseIf SubRank(선택) = "Special" Then
 Money = val(Money) + 1000
ElseIf SubRank(선택) = "Rare" Then
 Money = val(Money) + 2000
ElseIf SubRank(선택) = "Unique" Then
 Money = val(Money) + 4000
ElseIf SubRank(선택) = "Elite" Then
 Money = val(Money) + 8000
ElseIf SubRank(선택) = "Legend" Then
 Money = val(Money) + 16000
ElseIf SubRank(선택) = "Secret" Then
 Money = val(Money) + 32000
ElseIf SubRank(선택) = "Champion" Then
 Money = val(Money) + 64000
End If

선수수 = val(선수수) - 1
SubName(선택) = ""
SubTribe(선택) = ""
SubAt(선택) = ""
SubR(선택) = ""
SubSt(선택) = ""
SubAm(선택) = ""
SubDe(선택) = ""
SubPa(선택) = ""
SubSe(선택) = ""
SubCo(선택) = ""
SubYear(선택) = ""
SubRank(선택) = ""
SubAW(선택) = ""
SubAL(선택) = ""
SubTW(선택) = ""
SubTL(선택) = ""
SubZW(선택) = ""
SubZL(선택) = ""
SubPW(선택) = ""
SubPL(선택) = ""
SubTeam(선택) = ""
SubLev(선택) = ""
SubExp(선택) = ""
SubMExp(선택) = ""
SubPoint(선택) = ""
SubVic(선택) = ""
SubSeVic(선택) = ""
SubT연승(선택) = 0
SubZ연승(선택) = 0
SubP연승(선택) = 0
SubA연승(선택) = 0
SubT연(선택) = "W"
SubZ연(선택) = "W"
SubP연(선택) = "W"
SubA연(선택) = "W"
SubNum(선택) = 0
SubNW(선택) = ""
SubSkill(선택) = 0
Do Until 선택 = val(선수수) - 5
 SubName(선택) = SubName(선택 + 1)
 SubTribe(선택) = SubTribe(선택 + 1)
 SubAt(선택) = SubAt(선택 + 1)
 SubR(선택) = SubR(선택 + 1)
 SubSt(선택) = SubSt(선택 + 1)
 SubAm(선택) = SubAm(선택 + 1)
 SubDe(선택) = SubDe(선택 + 1)
 SubPa(선택) = SubPa(선택 + 1)
 SubSe(선택) = SubSe(선택 + 1)
 SubCo(선택) = SubCo(선택 + 1)
 SubYear(선택) = SubYear(선택 + 1)
 SubRank(선택) = SubRank(선택 + 1)
 SubAW(선택) = SubAW(선택 + 1)
 SubAL(선택) = SubAL(선택 + 1)
 SubTW(선택) = SubTW(선택 + 1)
 SubTL(선택) = SubTL(선택 + 1)
 SubZW(선택) = SubZW(선택 + 1)
 SubZL(선택) = SubZL(선택 + 1)
 SubPW(선택) = SubPW(선택 + 1)
 SubPL(선택) = SubPL(선택 + 1)
 SubTeam(선택) = SubTeam(선택 + 1)
 SubLev(선택) = SubLev(선택 + 1)
 SubExp(선택) = SubExp(선택 + 1)
 SubMExp(선택) = SubMExp(선택 + 1)
 SubPoint(선택) = SubPoint(선택 + 1)
 SubVic(선택) = SubVic(선택 + 1)
 SubSeVic(선택) = SubSeVic(선택 + 1)
 SubNum(선택) = SubNum(선택 + 1)
 SubNW(선택) = SubNW(선택 + 1)
 SubT연승(선택) = SubT연승(선택 + 1)
 SubZ연승(선택) = SubZ연승(선택 + 1)
 SubP연승(선택) = SubP연승(선택 + 1)
 SubA연승(선택) = SubA연승(선택 + 1)
 SubT연(선택) = SubT연(선택 + 1)
 SubZ연(선택) = SubZ연(선택 + 1)
 SubP연(선택) = SubP연(선택 + 1)
 SubA연(선택) = SubA연(선택 + 1)
 SubSkill(선택) = SubSkill(선택 + 1)
 선택 = val(선택) + 1
Loop

SubName(선택) = ""
SubTribe(선택) = ""
SubAt(선택) = ""
SubR(선택) = ""
SubSt(선택) = ""
SubAm(선택) = ""
SubDe(선택) = ""
SubPa(선택) = ""
SubSe(선택) = ""
SubCo(선택) = ""
SubYear(선택) = ""
SubRank(선택) = ""
SubAW(선택) = ""
SubAL(선택) = ""
SubTW(선택) = ""
SubTL(선택) = ""
SubZW(선택) = ""
SubZL(선택) = ""
SubPW(선택) = ""
SubPL(선택) = ""
SubTeam(선택) = ""
SubLev(선택) = ""
SubExp(선택) = ""
SubMExp(선택) = ""
SubPoint(선택) = ""
SubVic(선택) = ""
SubSeVic(선택) = ""
SubT연승(선택) = 0
SubZ연승(선택) = 0
SubP연승(선택) = 0
SubA연승(선택) = 0
SubT연(선택) = "W"
SubZ연(선택) = "W"
SubP연(선택) = "W"
SubA연(선택) = "W"
SubNum(선택) = 0
SubNW(선택) = ""
SubSkill(선택) = 0
For 세팅 = 0 To 8
 lblS(세팅) = SubYear(세팅 + 1) & SubName(세팅 + 1)
Next
FrmMain.Timer14.Enabled = True
FrmMain.Timer12.Enabled = True
If val(선수수) >= 7 Then
Command1_Click
Else
Unload Me
End If
End Sub

