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
   StartUpPosition =   2  'ȭ�� ���
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
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "����"
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
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "����"
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
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "����"
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
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "����"
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
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "����"
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
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "����"
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
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "����"
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
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "����"
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
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "����"
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
         Name            =   "����"
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
      BackStyle       =   0  '����
      Caption         =   "<11>�̿�ȣ"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblS 
      BackStyle       =   0  '����
      Caption         =   "<11>�̿�ȣ"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblS 
      BackStyle       =   0  '����
      Caption         =   "<11>�̿�ȣ"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblS 
      BackStyle       =   0  '����
      Caption         =   "<11>�̿�ȣ"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblS 
      BackStyle       =   0  '����
      Caption         =   "<11>�̿�ȣ"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblS 
      BackStyle       =   0  '����
      Caption         =   "<11>�̿�ȣ"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblS 
      BackStyle       =   0  '����
      Caption         =   "<11>�̿�ȣ"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblS 
      BackStyle       =   0  '����
      Caption         =   "<11>�̿�ȣ"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblS 
      BackStyle       =   0  '����
      Caption         =   "<11>�̿�ȣ"
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
���� = 1
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(����) & ".gif")
End If
End Sub

Private Sub Command2_Click()
���� = 2
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(����) & ".gif")
End If
End Sub

Private Sub Command3_Click()
���� = 3
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(����) & ".gif")
End If
End Sub

Private Sub Command4_Click()
���� = 4
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(����) & ".gif")
End If
End Sub

Private Sub Command5_Click()
���� = 5
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(����) & ".gif")
End If
End Sub

Private Sub Command6_Click()
���� = 6
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(����) & ".gif")
End If
End Sub

Private Sub Command7_Click()
���� = 7
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(����) & ".gif")
End If
End Sub

Private Sub Command8_Click()
���� = 8
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(����) & ".gif")
End If
End Sub

Private Sub Command9_Click()
���� = 9
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(����), 2, 2) & "]" & SubName(����) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(����) & ".gif")
End If
End Sub

Private Sub Form_Load()
For ���� = 0 To 8
 lblS(����) = SubYear(���� + 1) & SubName(���� + 1)
Next
Command1_Click
End Sub

Private Sub jcbutton1_Click()
If SubRank(����) = "Normal" Then
 Money = val(Money) + 500
ElseIf SubRank(����) = "Special" Then
 Money = val(Money) + 1000
ElseIf SubRank(����) = "Rare" Then
 Money = val(Money) + 2000
ElseIf SubRank(����) = "Unique" Then
 Money = val(Money) + 4000
ElseIf SubRank(����) = "Elite" Then
 Money = val(Money) + 8000
ElseIf SubRank(����) = "Legend" Then
 Money = val(Money) + 16000
ElseIf SubRank(����) = "Secret" Then
 Money = val(Money) + 32000
ElseIf SubRank(����) = "Champion" Then
 Money = val(Money) + 64000
End If

������ = val(������) - 1
SubName(����) = ""
SubTribe(����) = ""
SubAt(����) = ""
SubR(����) = ""
SubSt(����) = ""
SubAm(����) = ""
SubDe(����) = ""
SubPa(����) = ""
SubSe(����) = ""
SubCo(����) = ""
SubYear(����) = ""
SubRank(����) = ""
SubAW(����) = ""
SubAL(����) = ""
SubTW(����) = ""
SubTL(����) = ""
SubZW(����) = ""
SubZL(����) = ""
SubPW(����) = ""
SubPL(����) = ""
SubTeam(����) = ""
SubLev(����) = ""
SubExp(����) = ""
SubMExp(����) = ""
SubPoint(����) = ""
SubVic(����) = ""
SubSeVic(����) = ""
SubT����(����) = 0
SubZ����(����) = 0
SubP����(����) = 0
SubA����(����) = 0
SubT��(����) = "W"
SubZ��(����) = "W"
SubP��(����) = "W"
SubA��(����) = "W"
SubNum(����) = 0
SubNW(����) = ""
SubSkill(����) = 0
Do Until ���� = val(������) - 5
 SubName(����) = SubName(���� + 1)
 SubTribe(����) = SubTribe(���� + 1)
 SubAt(����) = SubAt(���� + 1)
 SubR(����) = SubR(���� + 1)
 SubSt(����) = SubSt(���� + 1)
 SubAm(����) = SubAm(���� + 1)
 SubDe(����) = SubDe(���� + 1)
 SubPa(����) = SubPa(���� + 1)
 SubSe(����) = SubSe(���� + 1)
 SubCo(����) = SubCo(���� + 1)
 SubYear(����) = SubYear(���� + 1)
 SubRank(����) = SubRank(���� + 1)
 SubAW(����) = SubAW(���� + 1)
 SubAL(����) = SubAL(���� + 1)
 SubTW(����) = SubTW(���� + 1)
 SubTL(����) = SubTL(���� + 1)
 SubZW(����) = SubZW(���� + 1)
 SubZL(����) = SubZL(���� + 1)
 SubPW(����) = SubPW(���� + 1)
 SubPL(����) = SubPL(���� + 1)
 SubTeam(����) = SubTeam(���� + 1)
 SubLev(����) = SubLev(���� + 1)
 SubExp(����) = SubExp(���� + 1)
 SubMExp(����) = SubMExp(���� + 1)
 SubPoint(����) = SubPoint(���� + 1)
 SubVic(����) = SubVic(���� + 1)
 SubSeVic(����) = SubSeVic(���� + 1)
 SubNum(����) = SubNum(���� + 1)
 SubNW(����) = SubNW(���� + 1)
 SubT����(����) = SubT����(���� + 1)
 SubZ����(����) = SubZ����(���� + 1)
 SubP����(����) = SubP����(���� + 1)
 SubA����(����) = SubA����(���� + 1)
 SubT��(����) = SubT��(���� + 1)
 SubZ��(����) = SubZ��(���� + 1)
 SubP��(����) = SubP��(���� + 1)
 SubA��(����) = SubA��(���� + 1)
 SubSkill(����) = SubSkill(���� + 1)
 ���� = val(����) + 1
Loop

SubName(����) = ""
SubTribe(����) = ""
SubAt(����) = ""
SubR(����) = ""
SubSt(����) = ""
SubAm(����) = ""
SubDe(����) = ""
SubPa(����) = ""
SubSe(����) = ""
SubCo(����) = ""
SubYear(����) = ""
SubRank(����) = ""
SubAW(����) = ""
SubAL(����) = ""
SubTW(����) = ""
SubTL(����) = ""
SubZW(����) = ""
SubZL(����) = ""
SubPW(����) = ""
SubPL(����) = ""
SubTeam(����) = ""
SubLev(����) = ""
SubExp(����) = ""
SubMExp(����) = ""
SubPoint(����) = ""
SubVic(����) = ""
SubSeVic(����) = ""
SubT����(����) = 0
SubZ����(����) = 0
SubP����(����) = 0
SubA����(����) = 0
SubT��(����) = "W"
SubZ��(����) = "W"
SubP��(����) = "W"
SubA��(����) = "W"
SubNum(����) = 0
SubNW(����) = ""
SubSkill(����) = 0
For ���� = 0 To 8
 lblS(����) = SubYear(���� + 1) & SubName(���� + 1)
Next
FrmMain.Timer14.Enabled = True
FrmMain.Timer12.Enabled = True
If val(������) >= 7 Then
Command1_Click
Else
Unload Me
End If
End Sub

