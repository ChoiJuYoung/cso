VERSION 5.00
Begin VB.Form FrmNewGame 
   Caption         =   "Card Choice"
   ClientHeight    =   6735
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10815
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   6735
   ScaleWidth      =   10815
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Tim12 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3720
      Top             =   3000
   End
   Begin CSO.jcbutton CmdHel 
      Height          =   615
      Left            =   7200
      TabIndex        =   17
      Top             =   5160
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   255
      Caption         =   "helL"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin CSO.jcbutton CmdHar 
      Height          =   615
      Left            =   3600
      TabIndex        =   16
      Top             =   5160
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16711680
      Caption         =   "Hard"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin CSO.jcbutton CmdNor 
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   5160
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "NormaL"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin CSO.jcbutton jcbutton2 
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   7560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ButtonStyle     =   13
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Tester, Click Here."
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Timer TimAdd 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3240
      Top             =   3000
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2760
      Top             =   3000
   End
   Begin VB.Timer TimElse 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2280
      Top             =   3000
   End
   Begin VB.Timer Tim11 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4200
      Top             =   3480
   End
   Begin VB.Timer Tim10 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3720
      Top             =   3480
   End
   Begin VB.Timer Tim09 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3240
      Top             =   3480
   End
   Begin VB.Timer Tim08 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2760
      Top             =   3480
   End
   Begin VB.Timer Tim07 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2280
      Top             =   3480
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   600
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   7800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   8040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   8040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   3360
   End
   Begin CSO.jcbutton jcbutton1 
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   6120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   661
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   11432933
      Caption         =   "선택 완료"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   120
      Top             =   3360
   End
   Begin CSO.jcbutton CmdB3 
      Height          =   855
      Left            =   7200
      TabIndex        =   5
      Top             =   2280
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1508
      ButtonStyle     =   8
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12583104
      Caption         =   "Zerg - Type C"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin CSO.jcbutton CmdB2 
      Height          =   855
      Left            =   3600
      TabIndex        =   4
      Top             =   2280
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1508
      ButtonStyle     =   8
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16744576
      Caption         =   "Zerg - Type B"
      ForeColor       =   16711935
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin CSO.jcbutton CmdB1 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   2280
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1508
      ButtonStyle     =   8
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421631
      Caption         =   "Zerg - Type A"
      ForeColor       =   16711680
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin CSO.jcbutton CmdA3 
      Height          =   855
      Left            =   7200
      TabIndex        =   2
      Top             =   720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1508
      ButtonStyle     =   8
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12583104
      Caption         =   "Terran - Type C"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin CSO.jcbutton CmdA2 
      Height          =   855
      Left            =   3600
      TabIndex        =   1
      Top             =   720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1508
      ButtonStyle     =   8
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16744576
      Caption         =   "Terran - Type B"
      ForeColor       =   16711935
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin CSO.jcbutton CmdA1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1508
      ButtonStyle     =   8
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421631
      Caption         =   "Terran - Type A"
      ForeColor       =   16711680
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   1800
      Top             =   8160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1200
      Top             =   8160
   End
   Begin CSO.jcbutton CmdC1 
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   3840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1508
      ButtonStyle     =   8
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421631
      Caption         =   "Protoss - Type A"
      ForeColor       =   16711680
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin CSO.jcbutton CmdC2 
      Height          =   855
      Left            =   3600
      TabIndex        =   7
      Top             =   3840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1508
      ButtonStyle     =   8
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16744576
      Caption         =   "Protoss - Type B"
      ForeColor       =   16711935
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin CSO.jcbutton CmdC3 
      Height          =   855
      Left            =   7200
      TabIndex        =   8
      Top             =   3840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1508
      ButtonStyle     =   8
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12583104
      Caption         =   "Protoss - Type C"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  '투명하지 않음
      Height          =   855
      Left            =   0
      Top             =   5040
      Width           =   10815
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "잠시만 기다려 주십시오."
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
      Top             =   120
      Width           =   10815
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '투명하지 않음
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   10815
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  '투명하지 않음
      Height          =   855
      Left            =   0
      Top             =   5880
      Width           =   10815
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  '투명하지 않음
      Height          =   1575
      Index           =   6
      Left            =   7200
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  '투명하지 않음
      Height          =   1575
      Index           =   5
      Left            =   3600
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  '투명하지 않음
      Height          =   1575
      Index           =   4
      Left            =   0
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  '투명하지 않음
      Height          =   1575
      Index           =   3
      Left            =   7200
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  '투명하지 않음
      Height          =   1575
      Index           =   2
      Left            =   3600
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  '투명하지 않음
      Height          =   1575
      Index           =   1
      Left            =   0
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  '투명하지 않음
      Height          =   1575
      Index           =   0
      Left            =   7200
      Top             =   360
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '투명하지 않음
      Height          =   1575
      Left            =   3600
      Top             =   360
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      Height          =   1575
      Left            =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "FrmNewGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdA1_Click()
If CmdA2.Enabled = False Or CmdA3.Enabled = False Then
CmdA1.Enabled = False
CmdA2.Enabled = True
CmdA3.Enabled = True
Else
CmdA1.Enabled = False
선택량 = val(선택량) + 1
End If
Randomize Oee
Randomize CR
Oee = Int((800 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
End Sub

Private Sub CmdA2_Click()
If CmdA1.Enabled = False Or CmdA3.Enabled = False Then
CmdA2.Enabled = False
CmdA1.Enabled = True
CmdA3.Enabled = True
Else
CmdA2.Enabled = False
선택량 = val(선택량) + 1
End If
Randomize Oee
Randomize CR
Oee = Int((800 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
End Sub

Private Sub CmdA3_Click()
If CmdA2.Enabled = False Or CmdA1.Enabled = False Then
CmdA3.Enabled = False
CmdA2.Enabled = True
CmdA1.Enabled = True
Else
CmdA3.Enabled = False
선택량 = val(선택량) + 1
End If
Randomize Oee
Randomize CR
Oee = Int((800 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
End Sub

Private Sub CmdB1_Click()
If CmdB2.Enabled = False Or CmdB3.Enabled = False Then
CmdB1.Enabled = False
CmdB2.Enabled = True
CmdB3.Enabled = True
Else
CmdB1.Enabled = False
선택량 = val(선택량) + 1
End If
Randomize Oee
Randomize CR
Oee = Int((800 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
End Sub

Private Sub CmdB2_Click()
If CmdB1.Enabled = False Or CmdB3.Enabled = False Then
CmdB2.Enabled = False
CmdB1.Enabled = True
CmdB3.Enabled = True
Else
CmdB2.Enabled = False
선택량 = val(선택량) + 1
End If
Randomize Oee
Randomize CR
Oee = Int((800 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
End Sub

Private Sub CmdB3_Click()
If CmdB2.Enabled = False Or CmdB1.Enabled = False Then
CmdB3.Enabled = False
CmdB2.Enabled = True
CmdB1.Enabled = True
Else
CmdB3.Enabled = False
선택량 = val(선택량) + 1
End If
Randomize Oee
Randomize CR
Oee = Int((800 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
End Sub

Private Sub CmdC1_Click()
If CmdC2.Enabled = False Or CmdC3.Enabled = False Then
CmdC1.Enabled = False
CmdC2.Enabled = True
CmdC3.Enabled = True
Else
CmdC1.Enabled = False
선택량 = val(선택량) + 1
End If
Randomize Oee
Randomize CR
Oee = Int((800 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
End Sub

Private Sub CmdC2_Click()
If CmdC1.Enabled = False Or CmdC3.Enabled = False Then
CmdC2.Enabled = False
CmdC1.Enabled = True
CmdC3.Enabled = True
Else
CmdC2.Enabled = False
선택량 = val(선택량) + 1
End If
Randomize Oee
Randomize CR
Oee = Int((800 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
End Sub

Private Sub CmdC3_Click()
If CmdC2.Enabled = False Or CmdC1.Enabled = False Then
CmdC3.Enabled = False
CmdC2.Enabled = True
CmdC1.Enabled = True
Else
CmdC3.Enabled = False
선택량 = val(선택량) + 1
End If
Randomize Oee
Randomize CR
Oee = Int((800 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
End Sub

Private Sub CmdHar_Click()
CmdNor.Enabled = True
CmdHel.Enabled = True
CmdHar.Enabled = False
Mode = "Hard"
End Sub

Private Sub CmdHel_Click()
CmdNor.Enabled = True
CmdHar.Enabled = True
CmdHel.Enabled = False
Mode = "Hell"
End Sub

Private Sub CmdNor_Click()
CmdHar.Enabled = True
CmdHel.Enabled = True
CmdNor.Enabled = False
Mode = "Normal"
End Sub

Private Sub Command1_Click()
Dim CodeName As String
CodeName = InputBox("코드입력")
If CodeName = "최주영꺼" Then
선수수 = 6
Oee = 540
MyName(1) = 이름(Oee)
MyTribe(1) = 1
MyAt(1) = 공격력(Oee)
MyR(1) = 견제(Oee)
MySt(1) = 전략(Oee)
MyAm(1) = 물량(Oee)
MyDe(1) = 수비력(Oee)
MyPa(1) = 정찰(Oee)
MySe(1) = 센스(Oee)
MyCo(1) = 컨트롤(Oee)
MyYear(1) = OYear(Oee)
MyRank(1) = 랭크(Oee)
MyTeam(1) = Team(Oee)
MySkill(1) = Skill(Oee)
PlayNumber(1) = Oee

Oee = 93
MyName(2) = 이름(Oee)
MyTribe(2) = 2
MyAt(2) = 공격력(Oee)
MyR(2) = 견제(Oee)
MySt(2) = 전략(Oee)
MyAm(2) = 물량(Oee)
MyDe(2) = 수비력(Oee)
MyPa(2) = 정찰(Oee)
MySe(2) = 센스(Oee)
MyCo(2) = 컨트롤(Oee)
MyYear(2) = OYear(Oee)
MyRank(2) = 랭크(Oee)
MyTeam(2) = Team(Oee)
MySkill(2) = Skill(Oee)
PlayNumber(2) = Oee

Oee = 113
MyName(3) = 이름(Oee)
MyTribe(3) = 3
MyAt(3) = 공격력(Oee)
MyR(3) = 견제(Oee)
MySt(3) = 전략(Oee)
MyAm(3) = 물량(Oee)
MyDe(3) = 수비력(Oee)
MyPa(3) = 정찰(Oee)
MySe(3) = 센스(Oee)
MyCo(3) = 컨트롤(Oee)
MyYear(3) = OYear(Oee)
MyRank(3) = 랭크(Oee)
MyTeam(3) = Team(Oee)
MySkill(3) = Skill(Oee)
PlayNumber(3) = Oee

Oee = 114
MyName(4) = 이름(Oee)
MyTribe(4) = 1
MyAt(4) = 공격력(Oee)
MyR(4) = 견제(Oee)
MySt(4) = 전략(Oee)
MyAm(4) = 물량(Oee)
MyDe(4) = 수비력(Oee)
MyPa(4) = 정찰(Oee)
MySe(4) = 센스(Oee)
MyCo(4) = 컨트롤(Oee)
MyYear(4) = OYear(Oee)
MyRank(4) = 랭크(Oee)
MyTeam(4) = Team(Oee)
MySkill(4) = Skill(Oee)
PlayNumber(4) = Oee

Oee = 175
MyName(5) = 이름(Oee)
MyTribe(5) = 2
MyAt(5) = 공격력(Oee)
MyR(5) = 견제(Oee)
MySt(5) = 전략(Oee)
MyAm(5) = 물량(Oee)
MyDe(5) = 수비력(Oee)
MyPa(5) = 정찰(Oee)
MySe(5) = 센스(Oee)
MyCo(5) = 컨트롤(Oee)
MyYear(5) = OYear(Oee)
MyRank(5) = 랭크(Oee)
MyTeam(5) = Team(Oee)
MySkill(5) = Skill(Oee)
PlayNumber(5) = Oee

Oee = 320
MyName(6) = 이름(Oee)
MyTribe(6) = 3
MyAt(6) = 공격력(Oee)
MyR(6) = 견제(Oee)
MySt(6) = 전략(Oee)
MyAm(6) = 물량(Oee)
MyDe(6) = 수비력(Oee)
MyPa(6) = 정찰(Oee)
MySe(6) = 센스(Oee)
MyCo(6) = 컨트롤(Oee)
MyYear(6) = OYear(Oee)
MyRank(6) = 랭크(Oee)
MyTeam(6) = Team(Oee)
MySkill(6) = Skill(Oee)
PlayNumber(6) = Oee

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
Next 갠리그

TeamName = InputBox("닉네임을 입력하세요.")
If Mode = "Normal" Then
    For i = 1 To 6
        MyAt(i) = val(MyAt(i)) + 50
        MyR(i) = val(MyR(i)) + 50
        MySt(i) = val(MySt(i)) + 50
        MyAm(i) = val(MyAm(i)) + 50
        MyDe(i) = val(MyDe(i)) + 50
        MyPa(i) = val(MyPa(i)) + 50
        MySe(i) = val(MySe(i)) + 50
        MyCo(i) = val(MyCo(i)) + 50
    Next
    For Oee = 0 To 800
        NPC공격력(Oee) = val(NPC공격력(Oee)) + 50
        NPC견제(Oee) = val(NPC견제(Oee)) + 50
        NPC전략(Oee) = val(NPC전략(Oee)) + 50
        NPC물량(Oee) = val(NPC물량(Oee)) + 50
        NPC수비력(Oee) = val(NPC수비력(Oee)) + 50
        NPC정찰(Oee) = val(NPC정찰(Oee)) + 50
        NPC센스(Oee) = val(NPC센스(Oee)) + 50
        NPC컨트롤(Oee) = val(NPC컨트롤(Oee)) + 50
    Next
ElseIf Mode = "Hell" Then
    For Oee = 0 To 800
        공격력(Oee) = val(공격력(Oee)) + 50
        견제(Oee) = val(견제(Oee)) + 50
        전략(Oee) = val(전략(Oee)) + 50
        물량(Oee) = val(물량(Oee)) + 50
        수비력(Oee) = val(수비력(Oee)) + 50
        정찰(Oee) = val(정찰(Oee)) + 50
        센스(Oee) = val(센스(Oee)) + 50
        컨트롤(Oee) = val(컨트롤(Oee)) + 50
    Next
End If
하향 = 0
하향횟수 = 0
FrmMain.Show
Money = 5000
Unload Me
ElseIf CodeName = "SoulDeck" Then
선수수 = 6
Oee = 114
MyName(1) = 이름(Oee)
MyTribe(1) = 1
MyAt(1) = 공격력(Oee)
MyR(1) = 견제(Oee)
MySt(1) = 전략(Oee)
MyAm(1) = 물량(Oee)
MyDe(1) = 수비력(Oee)
MyPa(1) = 정찰(Oee)
MySe(1) = 센스(Oee)
MyCo(1) = 컨트롤(Oee)
MyYear(1) = OYear(Oee)
MyRank(1) = 랭크(Oee)
MyTeam(1) = Team(Oee)
MySkill(1) = Skill(Oee)
PlayNumber(1) = Oee

Oee = 544
MyName(2) = 이름(Oee)
MyTribe(2) = 2
MyAt(2) = 공격력(Oee)
MyR(2) = 견제(Oee)
MySt(2) = 전략(Oee)
MyAm(2) = 물량(Oee)
MyDe(2) = 수비력(Oee)
MyPa(2) = 정찰(Oee)
MySe(2) = 센스(Oee)
MyCo(2) = 컨트롤(Oee)
MyYear(2) = OYear(Oee)
MyRank(2) = 랭크(Oee)
MyTeam(2) = Team(Oee)
MySkill(2) = Skill(Oee)
PlayNumber(2) = Oee

Oee = 136
MyName(3) = 이름(Oee)
MyTribe(3) = 3
MyAt(3) = 공격력(Oee)
MyR(3) = 견제(Oee)
MySt(3) = 전략(Oee)
MyAm(3) = 물량(Oee)
MyDe(3) = 수비력(Oee)
MyPa(3) = 정찰(Oee)
MySe(3) = 센스(Oee)
MyCo(3) = 컨트롤(Oee)
MyYear(3) = OYear(Oee)
MyRank(3) = 랭크(Oee)
MyTeam(3) = Team(Oee)
MySkill(3) = Skill(Oee)
PlayNumber(3) = Oee

Oee = 288
MyName(4) = 이름(Oee)
MyTribe(4) = 3
MyAt(4) = 공격력(Oee)
MyR(4) = 견제(Oee)
MySt(4) = 전략(Oee)
MyAm(4) = 물량(Oee)
MyDe(4) = 수비력(Oee)
MyPa(4) = 정찰(Oee)
MySe(4) = 센스(Oee)
MyCo(4) = 컨트롤(Oee)
MyYear(4) = OYear(Oee)
MyRank(4) = 랭크(Oee)
MyTeam(4) = Team(Oee)
MySkill(4) = Skill(Oee)
PlayNumber(4) = Oee

Oee = 112
MyName(5) = 이름(Oee)
MyTribe(5) = 2
MyAt(5) = 공격력(Oee)
MyR(5) = 견제(Oee)
MySt(5) = 전략(Oee)
MyAm(5) = 물량(Oee)
MyDe(5) = 수비력(Oee)
MyPa(5) = 정찰(Oee)
MySe(5) = 센스(Oee)
MyCo(5) = 컨트롤(Oee)
MyYear(5) = OYear(Oee)
MyRank(5) = 랭크(Oee)
MyTeam(5) = Team(Oee)
MySkill(5) = Skill(Oee)
PlayNumber(5) = Oee

Oee = 320
MyName(6) = 이름(Oee)
MyTribe(6) = 3
MyAt(6) = 공격력(Oee)
MyR(6) = 견제(Oee)
MySt(6) = 전략(Oee)
MyAm(6) = 물량(Oee)
MyDe(6) = 수비력(Oee)
MyPa(6) = 정찰(Oee)
MySe(6) = 센스(Oee)
MyCo(6) = 컨트롤(Oee)
MyYear(6) = OYear(Oee)
MyRank(6) = 랭크(Oee)
MyTeam(6) = Team(Oee)
MySkill(6) = Skill(Oee)
PlayNumber(6) = Oee

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
Next 갠리그

TeamName = InputBox("닉네임을 입력하세요.")
If Mode = "Normal" Then
    For i = 1 To 6
        MyAt(i) = val(MyAt(i)) + 50
        MyR(i) = val(MyR(i)) + 50
        MySt(i) = val(MySt(i)) + 50
        MyAm(i) = val(MyAm(i)) + 50
        MyDe(i) = val(MyDe(i)) + 50
        MyPa(i) = val(MyPa(i)) + 50
        MySe(i) = val(MySe(i)) + 50
        MyCo(i) = val(MyCo(i)) + 50
    Next
    For Oee = 0 To 800
        NPC공격력(Oee) = val(NPC공격력(Oee)) + 50
        NPC견제(Oee) = val(NPC견제(Oee)) + 50
        NPC전략(Oee) = val(NPC전략(Oee)) + 50
        NPC물량(Oee) = val(NPC물량(Oee)) + 50
        NPC수비력(Oee) = val(NPC수비력(Oee)) + 50
        NPC정찰(Oee) = val(NPC정찰(Oee)) + 50
        NPC센스(Oee) = val(NPC센스(Oee)) + 50
        NPC컨트롤(Oee) = val(NPC컨트롤(Oee)) + 50
    Next
ElseIf Mode = "Hell" Then
    For Oee = 0 To 800
        공격력(Oee) = val(공격력(Oee)) + 50
        견제(Oee) = val(견제(Oee)) + 50
        전략(Oee) = val(전략(Oee)) + 50
        물량(Oee) = val(물량(Oee)) + 50
        수비력(Oee) = val(수비력(Oee)) + 50
        정찰(Oee) = val(정찰(Oee)) + 50
        센스(Oee) = val(센스(Oee)) + 50
        컨트롤(Oee) = val(컨트롤(Oee)) + 50
    Next
End If
하향 = 0
하향횟수 = 0
FrmMain.Show
Money = 5000
Unload Me
End If

End Sub

Private Sub Command2_Click()
Dim CodeName As String
CodeName = InputBox("코드입력")
If CodeName = "Crow" Then
Money = 10000
선수수 = 6
Oee = 714
MyName(1) = 이름(Oee)
MyTribe(1) = 1
MyAt(1) = 공격력(Oee)
MyR(1) = 견제(Oee)
MySt(1) = 전략(Oee)
MyAm(1) = 물량(Oee)
MyDe(1) = 수비력(Oee)
MyPa(1) = 정찰(Oee)
MySe(1) = 센스(Oee)
MyCo(1) = 컨트롤(Oee)
MyYear(1) = OYear(Oee)
MyRank(1) = 랭크(Oee)
MyTeam(1) = Team(Oee)
MySkill(1) = Skill(Oee)
PlayNumber(1) = Oee

Oee = 80
MyName(2) = 이름(Oee)
MyTribe(2) = 2
MyAt(2) = 공격력(Oee)
MyR(2) = 견제(Oee)
MySt(2) = 전략(Oee)
MyAm(2) = 물량(Oee)
MyDe(2) = 수비력(Oee)
MyPa(2) = 정찰(Oee)
MySe(2) = 센스(Oee)
MyCo(2) = 컨트롤(Oee)
MyYear(2) = OYear(Oee)
MyRank(2) = 랭크(Oee)
MyTeam(2) = Team(Oee)
MySkill(2) = Skill(Oee)
PlayNumber(2) = Oee

Oee = 136
MyName(3) = 이름(Oee)
MyTribe(3) = 3
MyAt(3) = 공격력(Oee)
MyR(3) = 견제(Oee)
MySt(3) = 전략(Oee)
MyAm(3) = 물량(Oee)
MyDe(3) = 수비력(Oee)
MyPa(3) = 정찰(Oee)
MySe(3) = 센스(Oee)
MyCo(3) = 컨트롤(Oee)
MyYear(3) = OYear(Oee)
MyRank(3) = 랭크(Oee)
MyTeam(3) = Team(Oee)
MySkill(3) = Skill(Oee)
PlayNumber(3) = Oee

Oee = 659
MyName(4) = 이름(Oee)
MyTribe(4) = 1
MyAt(4) = 공격력(Oee)
MyR(4) = 견제(Oee)
MySt(4) = 전략(Oee)
MyAm(4) = 물량(Oee)
MyDe(4) = 수비력(Oee)
MyPa(4) = 정찰(Oee)
MySe(4) = 센스(Oee)
MyCo(4) = 컨트롤(Oee)
MyYear(4) = OYear(Oee)
MyRank(4) = 랭크(Oee)
MyTeam(4) = Team(Oee)
MySkill(4) = Skill(Oee)
PlayNumber(4) = Oee

Oee = 660
MyName(5) = 이름(Oee)
MyTribe(5) = 2
MyAt(5) = 공격력(Oee)
MyR(5) = 견제(Oee)
MySt(5) = 전략(Oee)
MyAm(5) = 물량(Oee)
MyDe(5) = 수비력(Oee)
MyPa(5) = 정찰(Oee)
MySe(5) = 센스(Oee)
MyCo(5) = 컨트롤(Oee)
MyYear(5) = OYear(Oee)
MyRank(5) = 랭크(Oee)
MyTeam(5) = Team(Oee)
MySkill(5) = Skill(Oee)
PlayNumber(5) = Oee

Oee = 288
MyName(6) = 이름(Oee)
MyTribe(6) = 3
MyAt(6) = 공격력(Oee)
MyR(6) = 견제(Oee)
MySt(6) = 전략(Oee)
MyAm(6) = 물량(Oee)
MyDe(6) = 수비력(Oee)
MyPa(6) = 정찰(Oee)
MySe(6) = 센스(Oee)
MyCo(6) = 컨트롤(Oee)
MyYear(6) = OYear(Oee)
MyRank(6) = 랭크(Oee)
MyTeam(6) = Team(Oee)
MySkill(6) = Skill(Oee)
PlayNumber(6) = Oee

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
Next 갠리그

TeamName = InputBox("닉네임을 입력하세요.")
If Mode = "Normal" Then
    For i = 1 To 6
        MyAt(i) = val(MyAt(i)) + 50
        MyR(i) = val(MyR(i)) + 50
        MySt(i) = val(MySt(i)) + 50
        MyAm(i) = val(MyAm(i)) + 50
        MyDe(i) = val(MyDe(i)) + 50
        MyPa(i) = val(MyPa(i)) + 50
        MySe(i) = val(MySe(i)) + 50
        MyCo(i) = val(MyCo(i)) + 50
    Next
    For Oee = 0 To 800
        NPC공격력(Oee) = val(NPC공격력(Oee)) + 50
        NPC견제(Oee) = val(NPC견제(Oee)) + 50
        NPC전략(Oee) = val(NPC전략(Oee)) + 50
        NPC물량(Oee) = val(NPC물량(Oee)) + 50
        NPC수비력(Oee) = val(NPC수비력(Oee)) + 50
        NPC정찰(Oee) = val(NPC정찰(Oee)) + 50
        NPC센스(Oee) = val(NPC센스(Oee)) + 50
        NPC컨트롤(Oee) = val(NPC컨트롤(Oee)) + 50
    Next
ElseIf Mode = "Hell" Then
    For Oee = 0 To 800
        공격력(Oee) = val(공격력(Oee)) + 50
        견제(Oee) = val(견제(Oee)) + 50
        전략(Oee) = val(전략(Oee)) + 50
        물량(Oee) = val(물량(Oee)) + 50
        수비력(Oee) = val(수비력(Oee)) + 50
        정찰(Oee) = val(정찰(Oee)) + 50
        센스(Oee) = val(센스(Oee)) + 50
        컨트롤(Oee) = val(컨트롤(Oee)) + 50
    Next
End If
하향 = 0
하향횟수 = 0
FrmMain.Show
Unload Me
ElseIf CodeName = "SecretDeck" Then
Money = 299792458
선수수 = 6
Oee = 649
MyName(1) = 이름(Oee)
MyTribe(1) = 1
MyAt(1) = 공격력(Oee)
MyR(1) = 견제(Oee)
MySt(1) = 전략(Oee)
MyAm(1) = 물량(Oee)
MyDe(1) = 수비력(Oee)
MyPa(1) = 정찰(Oee)
MySe(1) = 센스(Oee)
MyCo(1) = 컨트롤(Oee)
MyYear(1) = OYear(Oee)
MyRank(1) = 랭크(Oee)
MyTeam(1) = Team(Oee)
MySkill(1) = Skill(Oee)
PlayNumber(1) = Oee

Oee = 544
MyName(2) = 이름(Oee)
MyTribe(2) = 2
MyAt(2) = 공격력(Oee)
MyR(2) = 견제(Oee)
MySt(2) = 전략(Oee)
MyAm(2) = 물량(Oee)
MyDe(2) = 수비력(Oee)
MyPa(2) = 정찰(Oee)
MySe(2) = 센스(Oee)
MyCo(2) = 컨트롤(Oee)
MyYear(2) = OYear(Oee)
MyRank(2) = 랭크(Oee)
MyTeam(2) = Team(Oee)
MySkill(2) = Skill(Oee)
PlayNumber(2) = Oee

Oee = 560
MyName(3) = 이름(Oee)
MyTribe(3) = 3
MyAt(3) = 공격력(Oee)
MyR(3) = 견제(Oee)
MySt(3) = 전략(Oee)
MyAm(3) = 물량(Oee)
MyDe(3) = 수비력(Oee)
MyPa(3) = 정찰(Oee)
MySe(3) = 센스(Oee)
MyCo(3) = 컨트롤(Oee)
MyYear(3) = OYear(Oee)
MyRank(3) = 랭크(Oee)
MyTeam(3) = Team(Oee)
MySkill(3) = Skill(Oee)
PlayNumber(3) = Oee

Oee = 540
MyName(4) = 이름(Oee)
MyTribe(4) = 1
MyAt(4) = 공격력(Oee)
MyR(4) = 견제(Oee)
MySt(4) = 전략(Oee)
MyAm(4) = 물량(Oee)
MyDe(4) = 수비력(Oee)
MyPa(4) = 정찰(Oee)
MySe(4) = 센스(Oee)
MyCo(4) = 컨트롤(Oee)
MyYear(4) = OYear(Oee)
MyRank(4) = 랭크(Oee)
MyTeam(4) = Team(Oee)
MySkill(4) = Skill(Oee)
PlayNumber(4) = Oee

Oee = 547
MyName(5) = 이름(Oee)
MyTribe(5) = 1
MyAt(5) = 공격력(Oee)
MyR(5) = 견제(Oee)
MySt(5) = 전략(Oee)
MyAm(5) = 물량(Oee)
MyDe(5) = 수비력(Oee)
MyPa(5) = 정찰(Oee)
MySe(5) = 센스(Oee)
MyCo(5) = 컨트롤(Oee)
MyYear(5) = OYear(Oee)
MyRank(5) = 랭크(Oee)
MyTeam(5) = Team(Oee)
MySkill(5) = Skill(Oee)
PlayNumber(5) = Oee

Oee = 553
MyName(6) = 이름(Oee)
MyTribe(6) = 1
MyAt(6) = 공격력(Oee)
MyR(6) = 견제(Oee)
MySt(6) = 전략(Oee)
MyAm(6) = 물량(Oee)
MyDe(6) = 수비력(Oee)
MyPa(6) = 정찰(Oee)
MySe(6) = 센스(Oee)
MyCo(6) = 컨트롤(Oee)
MyYear(6) = OYear(Oee)
MyRank(6) = 랭크(Oee)
MyTeam(6) = Team(Oee)
MySkill(6) = Skill(Oee)
PlayNumber(6) = Oee

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
Next 갠리그

TeamName = InputBox("닉네임을 입력하세요.")
If Mode = "Normal" Then
    For i = 1 To 6
        MyAt(i) = val(MyAt(i)) + 50
        MyR(i) = val(MyR(i)) + 50
        MySt(i) = val(MySt(i)) + 50
        MyAm(i) = val(MyAm(i)) + 50
        MyDe(i) = val(MyDe(i)) + 50
        MyPa(i) = val(MyPa(i)) + 50
        MySe(i) = val(MySe(i)) + 50
        MyCo(i) = val(MyCo(i)) + 50
    Next
    For Oee = 0 To 800
        NPC공격력(Oee) = val(NPC공격력(Oee)) + 50
        NPC견제(Oee) = val(NPC견제(Oee)) + 50
        NPC전략(Oee) = val(NPC전략(Oee)) + 50
        NPC물량(Oee) = val(NPC물량(Oee)) + 50
        NPC수비력(Oee) = val(NPC수비력(Oee)) + 50
        NPC정찰(Oee) = val(NPC정찰(Oee)) + 50
        NPC센스(Oee) = val(NPC센스(Oee)) + 50
        NPC컨트롤(Oee) = val(NPC컨트롤(Oee)) + 50
    Next
ElseIf Mode = "Hell" Then
    For Oee = 0 To 800
        공격력(Oee) = val(공격력(Oee)) + 50
        견제(Oee) = val(견제(Oee)) + 50
        전략(Oee) = val(전략(Oee)) + 50
        물량(Oee) = val(물량(Oee)) + 50
        수비력(Oee) = val(수비력(Oee)) + 50
        정찰(Oee) = val(정찰(Oee)) + 50
        센스(Oee) = val(센스(Oee)) + 50
        컨트롤(Oee) = val(컨트롤(Oee)) + 50
    Next
End If
하향 = 0
하향횟수 = 0
FrmMain.Show
Unload Me

ElseIf CodeName = "Mystar" Then
Money = 10000
선수수 = 6
Oee = 713
MyName(1) = 이름(Oee)
MyTribe(1) = 1
MyAt(1) = 공격력(Oee)
MyR(1) = 견제(Oee)
MySt(1) = 전략(Oee)
MyAm(1) = 물량(Oee)
MyDe(1) = 수비력(Oee)
MyPa(1) = 정찰(Oee)
MySe(1) = 센스(Oee)
MyCo(1) = 컨트롤(Oee)
MyYear(1) = OYear(Oee)
MyRank(1) = 랭크(Oee)
MyTeam(1) = Team(Oee)
MySkill(1) = Skill(Oee)
PlayNumber(1) = Oee

Oee = 709
MyName(2) = 이름(Oee)
MyTribe(2) = 2
MyAt(2) = 공격력(Oee)
MyR(2) = 견제(Oee)
MySt(2) = 전략(Oee)
MyAm(2) = 물량(Oee)
MyDe(2) = 수비력(Oee)
MyPa(2) = 정찰(Oee)
MySe(2) = 센스(Oee)
MyCo(2) = 컨트롤(Oee)
MyYear(2) = OYear(Oee)
MyRank(2) = 랭크(Oee)
MyTeam(2) = Team(Oee)
MySkill(2) = Skill(Oee)
PlayNumber(2) = Oee

Oee = 711
MyName(3) = 이름(Oee)
MyTribe(3) = 3
MyAt(3) = 공격력(Oee)
MyR(3) = 견제(Oee)
MySt(3) = 전략(Oee)
MyAm(3) = 물량(Oee)
MyDe(3) = 수비력(Oee)
MyPa(3) = 정찰(Oee)
MySe(3) = 센스(Oee)
MyCo(3) = 컨트롤(Oee)
MyYear(3) = OYear(Oee)
MyRank(3) = 랭크(Oee)
MyTeam(3) = Team(Oee)
MySkill(3) = Skill(Oee)
PlayNumber(3) = Oee

Oee = 710
MyName(4) = 이름(Oee)
MyTribe(4) = 1
MyAt(4) = 공격력(Oee)
MyR(4) = 견제(Oee)
MySt(4) = 전략(Oee)
MyAm(4) = 물량(Oee)
MyDe(4) = 수비력(Oee)
MyPa(4) = 정찰(Oee)
MySe(4) = 센스(Oee)
MyCo(4) = 컨트롤(Oee)
MyYear(4) = OYear(Oee)
MyRank(4) = 랭크(Oee)
MyTeam(4) = Team(Oee)
MySkill(4) = Skill(Oee)
PlayNumber(4) = Oee

Oee = 118
MyName(5) = 이름(Oee)
MyTribe(5) = 1
MyAt(5) = 공격력(Oee)
MyR(5) = 견제(Oee)
MySt(5) = 전략(Oee)
MyAm(5) = 물량(Oee)
MyDe(5) = 수비력(Oee)
MyPa(5) = 정찰(Oee)
MySe(5) = 센스(Oee)
MyCo(5) = 컨트롤(Oee)
MyYear(5) = OYear(Oee)
MyRank(5) = 랭크(Oee)
MyTeam(5) = Team(Oee)
MySkill(5) = Skill(Oee)
PlayNumber(5) = Oee

Oee = 712
MyName(6) = 이름(Oee)
MyTribe(6) = 1
MyAt(6) = 공격력(Oee)
MyR(6) = 견제(Oee)
MySt(6) = 전략(Oee)
MyAm(6) = 물량(Oee)
MyDe(6) = 수비력(Oee)
MyPa(6) = 정찰(Oee)
MySe(6) = 센스(Oee)
MyCo(6) = 컨트롤(Oee)
MyYear(6) = OYear(Oee)
MyRank(6) = 랭크(Oee)
MyTeam(6) = Team(Oee)
MySkill(6) = Skill(Oee)
PlayNumber(6) = Oee

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
Next 갠리그

TeamName = InputBox("닉네임을 입력하세요.")
If Mode = "Normal" Then
    For i = 1 To 6
        MyAt(i) = val(MyAt(i)) + 50
        MyR(i) = val(MyR(i)) + 50
        MySt(i) = val(MySt(i)) + 50
        MyAm(i) = val(MyAm(i)) + 50
        MyDe(i) = val(MyDe(i)) + 50
        MyPa(i) = val(MyPa(i)) + 50
        MySe(i) = val(MySe(i)) + 50
        MyCo(i) = val(MyCo(i)) + 50
    Next
    For Oee = 0 To 800
        NPC공격력(Oee) = val(NPC공격력(Oee)) + 50
        NPC견제(Oee) = val(NPC견제(Oee)) + 50
        NPC전략(Oee) = val(NPC전략(Oee)) + 50
        NPC물량(Oee) = val(NPC물량(Oee)) + 50
        NPC수비력(Oee) = val(NPC수비력(Oee)) + 50
        NPC정찰(Oee) = val(NPC정찰(Oee)) + 50
        NPC센스(Oee) = val(NPC센스(Oee)) + 50
        NPC컨트롤(Oee) = val(NPC컨트롤(Oee)) + 50
    Next
ElseIf Mode = "Hell" Then
    For Oee = 0 To 800
        공격력(Oee) = val(공격력(Oee)) + 50
        견제(Oee) = val(견제(Oee)) + 50
        전략(Oee) = val(전략(Oee)) + 50
        물량(Oee) = val(물량(Oee)) + 50
        수비력(Oee) = val(수비력(Oee)) + 50
        정찰(Oee) = val(정찰(Oee)) + 50
        센스(Oee) = val(센스(Oee)) + 50
        컨트롤(Oee) = val(컨트롤(Oee)) + 50
    Next
End If
하향 = 0
하향횟수 = 0
FrmMain.Show
Unload Me


End If
End Sub


Private Sub Form_Load()
CmdHar_Click
For Map = 1 To 12
 If val(Map) = 1 Then
  MapName(Map) = "글레디에이터"
  러쉬거리(Map) = 5
  자원(Map) = 5
  복잡도(Map) = 5
  TZT(Map) = 65
  TZZ(Map) = 35
  ZPZ(Map) = 50
  ZPP(Map) = 50
  PTP(Map) = 60
  PTT(Map) = 40
 ElseIf val(Map) = 2 Then
  MapName(Map) = "네오벨트웨이"
  러쉬거리(Map) = 6
  자원(Map) = 6
  복잡도(Map) = 4
  TZT(Map) = 55
  TZZ(Map) = 45
  ZPZ(Map) = 55
  ZPP(Map) = 45
  PTP(Map) = 45
  PTT(Map) = 55
 ElseIf val(Map) = 3 Then
  MapName(Map) = "네오아즈텍"
  러쉬거리(Map) = 8
  자원(Map) = 6
  복잡도(Map) = 8
  TZT(Map) = 45
  TZZ(Map) = 55
  ZPZ(Map) = 50
  ZPP(Map) = 50
  PTP(Map) = 60
  PTT(Map) = 40
 ElseIf val(Map) = 4 Then
  MapName(Map) = "라만차"
  러쉬거리(Map) = 5
  자원(Map) = 5
  복잡도(Map) = 5
  TZT(Map) = 60
  TZZ(Map) = 40
  ZPZ(Map) = 60
  ZPP(Map) = 40
  PTP(Map) = 60
  PTT(Map) = 40
 ElseIf val(Map) = 5 Then
  MapName(Map) = "루나"
  러쉬거리(Map) = 5
  자원(Map) = 5
  복잡도(Map) = 1
  TZT(Map) = 45
  TZZ(Map) = 55
  ZPZ(Map) = 55
  ZPP(Map) = 45
  PTP(Map) = 55
  PTT(Map) = 45
 ElseIf val(Map) = 6 Then
  MapName(Map) = "신태양의제국"
  러쉬거리(Map) = 8
  자원(Map) = 5
  복잡도(Map) = 5
  TZT(Map) = 65
  TZZ(Map) = 35
  ZPZ(Map) = 50
  ZPP(Map) = 50
  PTP(Map) = 45
  PTT(Map) = 55
 ElseIf val(Map) = 7 Then
  MapName(Map) = "신피의능선"
  러쉬거리(Map) = 8
  자원(Map) = 6
  복잡도(Map) = 9
  TZT(Map) = 40
  TZZ(Map) = 60
  ZPZ(Map) = 55
  ZPP(Map) = 45
  PTP(Map) = 60
  PTT(Map) = 40
 ElseIf val(Map) = 8 Then
  MapName(Map) = "써킷브레이커"
  러쉬거리(Map) = 5
  자원(Map) = 5
  복잡도(Map) = 5
  TZT(Map) = 50
  TZZ(Map) = 50
  ZPZ(Map) = 50
  ZPP(Map) = 50
  PTP(Map) = 50
  PTT(Map) = 50
 ElseIf val(Map) = 9 Then
  MapName(Map) = "얼터너티브"
  러쉬거리(Map) = 5
  자원(Map) = 5
  복잡도(Map) = 9
  TZT(Map) = 45
  TZZ(Map) = 55
  ZPZ(Map) = 50
  ZPP(Map) = 50
  PTP(Map) = 50
  PTT(Map) = 50
 ElseIf val(Map) = 10 Then
  MapName(Map) = "투혼"
  러쉬거리(Map) = 6
  자원(Map) = 8
  복잡도(Map) = 6
  TZT(Map) = 55
  TZZ(Map) = 45
  ZPZ(Map) = 45
  ZPP(Map) = 55
  PTP(Map) = 50
  PTT(Map) = 50
 ElseIf val(Map) = 11 Then
  MapName(Map) = "파이썬"
  러쉬거리(Map) = 3
  자원(Map) = 3
  복잡도(Map) = 1
  TZT(Map) = 55
  TZZ(Map) = 45
  ZPZ(Map) = 60
  ZPP(Map) = 40
  PTP(Map) = 45
  PTT(Map) = 55
 ElseIf val(Map) = 12 Then
  MapName(Map) = "패스파인더"
  러쉬거리(Map) = 4
  자원(Map) = 5
  복잡도(Map) = 5
  TZT(Map) = 65
  TZZ(Map) = 35
  ZPZ(Map) = 65
  ZPP(Map) = 35
  PTP(Map) = 50
  PTT(Map) = 50
 End If
Next

크로우생산 = "No"
PL우승 = 0
PL준우승 = 0
불러옴 = False
jcbutton1.Caption = "잠시만 기다려주세요."
Dim 돌려 As Integer
For 돌려 = 1 To 9
 SubName(돌려) = ""
 SubTeam(돌려) = ""
 SubAt(돌려) = ""
 SubR(돌려) = ""
 SubSt(돌려) = ""
 SubAm(돌려) = ""
 SubDe(돌려) = ""
 SubPa(돌려) = ""
 SubSe(돌려) = ""
 SubCo(돌려) = ""
 SubLev(돌려) = 1
 SubExp(돌려) = 0
 SubMExp(돌려) = 50
 SubAW(돌려) = 0
 SubAL(돌려) = 0
 SubTW(돌려) = 0
 SubTL(돌려) = 0
 SubZW(돌려) = 0
 SubZL(돌려) = 0
 SubPW(돌려) = 0
 SubPL(돌려) = 0
 SubT연승(돌려) = 0
 SubZ연승(돌려) = 0
 SubP연승(돌려) = 0
 SubA연승(돌려) = 0
 SubT연(돌려) = "W"
 SubZ연(돌려) = "W"
 SubP연(돌려) = "W"
 SubA연(돌려) = "W"
 SubSkill(돌려) = 0
Next 돌려

For 돌려 = 1 To 6
PL출전자(돌려) = True
Next 돌려

PLEnd = "False"
PL넘버 = 1
Money = 5000
  '****선수 선택
 'F


'Secret크로우
Oee = 0
 이름(Oee) = "크로우"
 OYear(Oee) = "<10>"
 랭크(Oee) = "Champion"
 Team(Oee) = "MyStar"
 종족(Oee) = 1
 공격력(Oee) = 950
 견제(Oee) = 950
 전략(Oee) = 900
 물량(Oee) = 900
 수비력(Oee) = 900
 정찰(Oee) = 900
 센스(Oee) = 900
 컨트롤(Oee) = 1000
 우승(Oee) = 0
 준우승(Oee) = 0
 컨디션(Oee) = 100
 A승리(Oee) = 0
 A패배(Oee) = 0
 P승리(Oee) = 0
 P패배(Oee) = 0
 T승리(Oee) = 0
 T패배(Oee) = 0
 Z승리(Oee) = 0
 Z패배(Oee) = 0
 T연승(Oee) = 0
 Z연승(Oee) = 0
 P연승(Oee) = 0
 A연승(Oee) = 0
 T연(Oee) = "W"
 Z연(Oee) = "W"
 P연(Oee) = "W"
 A연(Oee) = "W"
Close #1
PL경기수 = 0
PL승 = 0
PL패 = 0
PL진행 = "1R"
Text1.Text = "히리리"
End Sub


Private Sub jcbutton1_Click()
추첨경우 = Int((6 * Rnd) + 1)
FrmChoice.Show
Unload Me
End Sub

Private Sub jcbutton2_Click()
Dim 테스터코드 As String
테스터코드 = InputBox("Code")
If 테스터코드 = "tOdaY" Then
선수수 = 6
Randomize Oee

CR = Int((120 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
End If

Do Until (종족(Oee) = 1)
CR = Int((120 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
CR = 51
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
End If

Loop
MyName(1) = 이름(Oee)
MyTribe(1) = 1
MyAt(1) = 공격력(Oee)
MyR(1) = 견제(Oee)
MySt(1) = 전략(Oee)
MyAm(1) = 물량(Oee)
MyDe(1) = 수비력(Oee)
MyPa(1) = 정찰(Oee)
MySe(1) = 센스(Oee)
MyCo(1) = 컨트롤(Oee)
MyYear(1) = OYear(Oee)
MyRank(1) = 랭크(Oee)
MyTeam(1) = Team(Oee)
MySkill(1) = Skill(Oee)
PlayNumber(1) = Oee
Randomize Oee

CR = Int((120 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
End If

Do Until (종족(Oee) = 2)
CR = Int((120 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
CR = 51
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
End If

Loop
MyName(2) = 이름(Oee)
MyTribe(2) = 2
MyAt(2) = 공격력(Oee)
MyR(2) = 견제(Oee)
MySt(2) = 전략(Oee)
MyAm(2) = 물량(Oee)
MyDe(2) = 수비력(Oee)
MyPa(2) = 정찰(Oee)
MySe(2) = 센스(Oee)
MyCo(2) = 컨트롤(Oee)
MyYear(2) = OYear(Oee)
MyRank(2) = 랭크(Oee)
MyTeam(2) = Team(Oee)
MySkill(2) = Skill(Oee)
PlayNumber(2) = Oee
Randomize Oee


Oee = 710
MyName(3) = 이름(Oee)
MyTribe(3) = 3
MyAt(3) = 공격력(Oee)
MyR(3) = 견제(Oee)
MySt(3) = 전략(Oee)
MyAm(3) = 물량(Oee)
MyDe(3) = 수비력(Oee)
MyPa(3) = 정찰(Oee)
MySe(3) = 센스(Oee)
MyCo(3) = 컨트롤(Oee)
MyYear(3) = OYear(Oee)
MyRank(3) = 랭크(Oee)
MyTeam(3) = Team(Oee)
MySkill(3) = Skill(Oee)
PlayNumber(3) = Oee
Randomize Oee

CR = Int((120 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Do Until (종족(Oee) = 1)
CR = Int((120 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
CR = 51
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Loop
MyName(4) = 이름(Oee)
MyTribe(4) = 1
MyAt(4) = 공격력(Oee)
MyR(4) = 견제(Oee)
MySt(4) = 전략(Oee)
MyAm(4) = 물량(Oee)
MyDe(4) = 수비력(Oee)
MyPa(4) = 정찰(Oee)
MySe(4) = 센스(Oee)
MyCo(4) = 컨트롤(Oee)
MyYear(4) = OYear(Oee)
MyRank(4) = 랭크(Oee)
MyTeam(4) = Team(Oee)
MySkill(4) = Skill(Oee)
PlayNumber(4) = Oee
Randomize Oee

CR = Int((120 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Do Until (종족(Oee) = 2)
CR = Int((120 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
CR = 51
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Loop
MyName(5) = 이름(Oee)
MyTribe(5) = 2
MyAt(5) = 공격력(Oee)
MyR(5) = 견제(Oee)
MySt(5) = 전략(Oee)
MyAm(5) = 물량(Oee)
MyDe(5) = 수비력(Oee)
MyPa(5) = 정찰(Oee)
MySe(5) = 센스(Oee)
MyCo(5) = 컨트롤(Oee)
MyYear(5) = OYear(Oee)
MyRank(5) = 랭크(Oee)
MyTeam(5) = Team(Oee)
MySkill(5) = Skill(Oee)
PlayNumber(5) = Oee
Randomize Oee

Oee = 710
MyName(6) = 이름(Oee)
MyTribe(6) = 3
MyAt(6) = 공격력(Oee)
MyR(6) = 견제(Oee)
MySt(6) = 전략(Oee)
MyAm(6) = 물량(Oee)
MyDe(6) = 수비력(Oee)
MyPa(6) = 정찰(Oee)
MySe(6) = 센스(Oee)
MyCo(6) = 컨트롤(Oee)
MyYear(6) = OYear(Oee)
MyRank(6) = 랭크(Oee)
MyTeam(6) = Team(Oee)
MySkill(6) = Skill(Oee)
PlayNumber(6) = Oee
Randomize Oee


돈량 = 0
For 돌려 = 1 To 6
 If MyRank(돌려) = "Normal" Then
  돈량 = val(돈량) + 1
 ElseIf MyRank(돌려) = "Special" Then
  돈량 = val(돈량) + 2
 ElseIf MyRank(돌려) = "Rare" Then
  돈량 = val(돈량) + 3
 ElseIf MyRank(돌려) = "Unique" Then
  돈량 = val(돈량) + 4
 ElseIf MyRank(돌려) = "Elite" Then
  돈량 = val(돈량) + 5
 ElseIf MyRank(돌려) = "Legend" Then
  돈량 = val(돈량) + 6
 ElseIf MyRank(돌려) = "Secret" Then
  돈량 = val(돈량) + 7
 End If
Next

If val(돈량) = 6 Then
 Money = 25000
ElseIf (val(돈량) >= 7) And (val(돈량) <= 12) Then
 Money = 20000
ElseIf (val(돈량) >= 13) And (val(돈량) <= 18) Then
 Money = 15000
ElseIf (val(돈량) >= 19) And (val(돈량) <= 24) Then
 Money = 10000
ElseIf (val(돈량) >= 25) And (val(돈량) <= 30) Then
 Money = 5000
ElseIf (val(돈량) >= 31) And (val(돈량) <= 36) Then
 Money = 2000
ElseIf (val(돈량) >= 37) And (val(돈량) <= 42) Then
 Money = 1000
End If
확인용1 = val(Money) / 2

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
Next 갠리그

TeamName = InputBox("닉네임을 입력하세요.")
If Mode = "Normal" Then
    For i = 1 To 6
        MyAt(i) = val(MyAt(i)) + 50
        MyR(i) = val(MyR(i)) + 50
        MySt(i) = val(MySt(i)) + 50
        MyAm(i) = val(MyAm(i)) + 50
        MyDe(i) = val(MyDe(i)) + 50
        MyPa(i) = val(MyPa(i)) + 50
        MySe(i) = val(MySe(i)) + 50
        MyCo(i) = val(MyCo(i)) + 50
    Next
    For Oee = 0 To 800
        NPC공격력(Oee) = val(NPC공격력(Oee)) + 50
        NPC견제(Oee) = val(NPC견제(Oee)) + 50
        NPC전략(Oee) = val(NPC전략(Oee)) + 50
        NPC물량(Oee) = val(NPC물량(Oee)) + 50
        NPC수비력(Oee) = val(NPC수비력(Oee)) + 50
        NPC정찰(Oee) = val(NPC정찰(Oee)) + 50
        NPC센스(Oee) = val(NPC센스(Oee)) + 50
        NPC컨트롤(Oee) = val(NPC컨트롤(Oee)) + 50
    Next
ElseIf Mode = "Hell" Then
    For Oee = 0 To 800
        공격력(Oee) = val(공격력(Oee)) + 50
        견제(Oee) = val(견제(Oee)) + 50
        전략(Oee) = val(전략(Oee)) + 50
        물량(Oee) = val(물량(Oee)) + 50
        수비력(Oee) = val(수비력(Oee)) + 50
        정찰(Oee) = val(정찰(Oee)) + 50
        센스(Oee) = val(센스(Oee)) + 50
        컨트롤(Oee) = val(컨트롤(Oee)) + 50
    Next
End If
하향 = 0
하향횟수 = 0
FrmMain.Show
Unload Me
ElseIf 테스터코드 = "은하랑" Then
선수수 = 6
Randomize Oee
Oee = 697
MyName(1) = 이름(Oee)
MyTribe(1) = 1
MyAt(1) = 공격력(Oee)
MyR(1) = 견제(Oee)
MySt(1) = 전략(Oee)
MyAm(1) = 물량(Oee)
MyDe(1) = 수비력(Oee)
MyPa(1) = 정찰(Oee)
MySe(1) = 센스(Oee)
MyCo(1) = 컨트롤(Oee)
MyYear(1) = OYear(Oee)
MyRank(1) = 랭크(Oee)
MyTeam(1) = Team(Oee)
MySkill(1) = Skill(Oee)
PlayNumber(1) = Oee
Randomize Oee
Oee = 799
MyName(2) = 이름(Oee)
MyTribe(2) = 2
MyAt(2) = 공격력(Oee)
MyR(2) = 견제(Oee)
MySt(2) = 전략(Oee)
MyAm(2) = 물량(Oee)
MyDe(2) = 수비력(Oee)
MyPa(2) = 정찰(Oee)
MySe(2) = 센스(Oee)
MyCo(2) = 컨트롤(Oee)
MyYear(2) = OYear(Oee)
MyRank(2) = 랭크(Oee)
MyTeam(2) = Team(Oee)
MySkill(2) = Skill(Oee)
PlayNumber(2) = Oee
Randomize Oee


Oee = 716
MyName(3) = 이름(Oee)
MyTribe(3) = 3
MyAt(3) = 공격력(Oee)
MyR(3) = 견제(Oee)
MySt(3) = 전략(Oee)
MyAm(3) = 물량(Oee)
MyDe(3) = 수비력(Oee)
MyPa(3) = 정찰(Oee)
MySe(3) = 센스(Oee)
MyCo(3) = 컨트롤(Oee)
MyYear(3) = OYear(Oee)
MyRank(3) = 랭크(Oee)
MyTeam(3) = Team(Oee)
MySkill(3) = Skill(Oee)
PlayNumber(3) = Oee
Randomize Oee
Oee = 20
MyName(4) = 이름(Oee)
MyTribe(4) = 1
MyAt(4) = 공격력(Oee)
MyR(4) = 견제(Oee)
MySt(4) = 전략(Oee)
MyAm(4) = 물량(Oee)
MyDe(4) = 수비력(Oee)
MyPa(4) = 정찰(Oee)
MySe(4) = 센스(Oee)
MyCo(4) = 컨트롤(Oee)
MyYear(4) = OYear(Oee)
MyRank(4) = 랭크(Oee)
MyTeam(4) = Team(Oee)
MySkill(4) = Skill(Oee)
PlayNumber(4) = Oee
Randomize Oee
Oee = 93
MyName(5) = 이름(Oee)
MyTribe(5) = 2
MyAt(5) = 공격력(Oee)
MyR(5) = 견제(Oee)
MySt(5) = 전략(Oee)
MyAm(5) = 물량(Oee)
MyDe(5) = 수비력(Oee)
MyPa(5) = 정찰(Oee)
MySe(5) = 센스(Oee)
MyCo(5) = 컨트롤(Oee)
MyYear(5) = OYear(Oee)
MyRank(5) = 랭크(Oee)
MyTeam(5) = Team(Oee)
MySkill(5) = Skill(Oee)
PlayNumber(5) = Oee
Randomize Oee

If Mode = "Normal" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hard" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hell" Then
    CR = Int((120 * Rnd) + 1)
End If
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Do Until (종족(Oee) = 3)
If Mode = "Normal" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hard" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hell" Then
    CR = Int((120 * Rnd) + 1)
End If
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Loop
MyName(6) = 이름(Oee)
MyTribe(6) = 3
MyAt(6) = 공격력(Oee)
MyR(6) = 견제(Oee)
MySt(6) = 전략(Oee)
MyAm(6) = 물량(Oee)
MyDe(6) = 수비력(Oee)
MyPa(6) = 정찰(Oee)
MySe(6) = 센스(Oee)
MyCo(6) = 컨트롤(Oee)
MyYear(6) = OYear(Oee)
MyRank(6) = 랭크(Oee)
MyTeam(6) = Team(Oee)
MySkill(6) = Skill(Oee)
PlayNumber(6) = Oee
Randomize Oee


돈량 = 0
For 돌려 = 1 To 6
 If MyRank(돌려) = "Normal" Then
  돈량 = val(돈량) + 1
 ElseIf MyRank(돌려) = "Special" Then
  돈량 = val(돈량) + 2
 ElseIf MyRank(돌려) = "Rare" Then
  돈량 = val(돈량) + 3
 ElseIf MyRank(돌려) = "Unique" Then
  돈량 = val(돈량) + 4
 ElseIf MyRank(돌려) = "Elite" Then
  돈량 = val(돈량) + 5
 ElseIf MyRank(돌려) = "Legend" Then
  돈량 = val(돈량) + 6
 ElseIf MyRank(돌려) = "Secret" Then
  돈량 = val(돈량) + 7
 End If
Next

If val(돈량) = 6 Then
 Money = 25000
ElseIf (val(돈량) >= 7) And (val(돈량) <= 12) Then
 Money = 20000
ElseIf (val(돈량) >= 13) And (val(돈량) <= 18) Then
 Money = 15000
ElseIf (val(돈량) >= 19) And (val(돈량) <= 24) Then
 Money = 10000
ElseIf (val(돈량) >= 25) And (val(돈량) <= 30) Then
 Money = 5000
ElseIf (val(돈량) >= 31) And (val(돈량) <= 36) Then
 Money = 2000
ElseIf (val(돈량) >= 37) And (val(돈량) <= 42) Then
 Money = 1000
End If
확인용1 = val(Money) / 2

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
Next 갠리그

TeamName = InputBox("닉네임을 입력하세요.")
If Mode = "Normal" Then
    For i = 1 To 6
        MyAt(i) = val(MyAt(i)) + 50
        MyR(i) = val(MyR(i)) + 50
        MySt(i) = val(MySt(i)) + 50
        MyAm(i) = val(MyAm(i)) + 50
        MyDe(i) = val(MyDe(i)) + 50
        MyPa(i) = val(MyPa(i)) + 50
        MySe(i) = val(MySe(i)) + 50
        MyCo(i) = val(MyCo(i)) + 50
    Next
    For Oee = 0 To 800
        NPC공격력(Oee) = val(NPC공격력(Oee)) + 50
        NPC견제(Oee) = val(NPC견제(Oee)) + 50
        NPC전략(Oee) = val(NPC전략(Oee)) + 50
        NPC물량(Oee) = val(NPC물량(Oee)) + 50
        NPC수비력(Oee) = val(NPC수비력(Oee)) + 50
        NPC정찰(Oee) = val(NPC정찰(Oee)) + 50
        NPC센스(Oee) = val(NPC센스(Oee)) + 50
        NPC컨트롤(Oee) = val(NPC컨트롤(Oee)) + 50
    Next
ElseIf Mode = "Hell" Then
    For Oee = 0 To 800
        공격력(Oee) = val(공격력(Oee)) + 50
        견제(Oee) = val(견제(Oee)) + 50
        전략(Oee) = val(전략(Oee)) + 50
        물량(Oee) = val(물량(Oee)) + 50
        수비력(Oee) = val(수비력(Oee)) + 50
        정찰(Oee) = val(정찰(Oee)) + 50
        센스(Oee) = val(센스(Oee)) + 50
        컨트롤(Oee) = val(컨트롤(Oee)) + 50
    Next
End If
하향 = 0
하향횟수 = 0
FrmMain.Show
Unload Me
ElseIf 테스터코드 = "moonlight" Then
선수수 = 6
Randomize Oee
Oee = 153
MyName(1) = 이름(Oee)
MyTribe(1) = 1
MyAt(1) = 공격력(Oee)
MyR(1) = 견제(Oee)
MySt(1) = 전략(Oee)
MyAm(1) = 물량(Oee)
MyDe(1) = 수비력(Oee)
MyPa(1) = 정찰(Oee)
MySe(1) = 센스(Oee)
MyCo(1) = 컨트롤(Oee)
MyYear(1) = OYear(Oee)
MyRank(1) = 랭크(Oee)
MyTeam(1) = Team(Oee)
MySkill(1) = Skill(Oee)
PlayNumber(1) = Oee
Randomize Oee

If Mode = "Normal" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hard" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hell" Then
    CR = Int((120 * Rnd) + 1)
End If
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
End If

Do Until (종족(Oee) = 2)
If Mode = "Normal" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hard" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hell" Then
    CR = Int((120 * Rnd) + 1)
End If
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
End If

Loop
MyName(2) = 이름(Oee)
MyTribe(2) = 2
MyAt(2) = 공격력(Oee)
MyR(2) = 견제(Oee)
MySt(2) = 전략(Oee)
MyAm(2) = 물량(Oee)
MyDe(2) = 수비력(Oee)
MyPa(2) = 정찰(Oee)
MySe(2) = 센스(Oee)
MyCo(2) = 컨트롤(Oee)
MyYear(2) = OYear(Oee)
MyRank(2) = 랭크(Oee)
MyTeam(2) = Team(Oee)
MySkill(2) = Skill(Oee)
PlayNumber(2) = Oee
Randomize Oee

Oee = 311
MyName(3) = 이름(Oee)
MyTribe(3) = 3
MyAt(3) = 공격력(Oee)
MyR(3) = 견제(Oee)
MySt(3) = 전략(Oee)
MyAm(3) = 물량(Oee)
MyDe(3) = 수비력(Oee)
MyPa(3) = 정찰(Oee)
MySe(3) = 센스(Oee)
MyCo(3) = 컨트롤(Oee)
MyYear(3) = OYear(Oee)
MyRank(3) = 랭크(Oee)
MyTeam(3) = Team(Oee)
MySkill(3) = Skill(Oee)
PlayNumber(3) = Oee
Randomize Oee

If Mode = "Normal" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hard" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hell" Then
    CR = Int((120 * Rnd) + 1)
End If
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Do Until (종족(Oee) = 1)
If Mode = "Normal" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hard" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hell" Then
    CR = Int((120 * Rnd) + 1)
End If
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Loop
MyName(4) = 이름(Oee)
MyTribe(4) = 1
MyAt(4) = 공격력(Oee)
MyR(4) = 견제(Oee)
MySt(4) = 전략(Oee)
MyAm(4) = 물량(Oee)
MyDe(4) = 수비력(Oee)
MyPa(4) = 정찰(Oee)
MySe(4) = 센스(Oee)
MyCo(4) = 컨트롤(Oee)
MyYear(4) = OYear(Oee)
MyRank(4) = 랭크(Oee)
MyTeam(4) = Team(Oee)
MySkill(4) = Skill(Oee)
PlayNumber(4) = Oee
Randomize Oee

If Mode = "Normal" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hard" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hell" Then
    CR = Int((120 * Rnd) + 1)
End If
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Do Until (종족(Oee) = 2)
If Mode = "Normal" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hard" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hell" Then
    CR = Int((120 * Rnd) + 1)
End If
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Loop
MyName(5) = 이름(Oee)
MyTribe(5) = 2
MyAt(5) = 공격력(Oee)
MyR(5) = 견제(Oee)
MySt(5) = 전략(Oee)
MyAm(5) = 물량(Oee)
MyDe(5) = 수비력(Oee)
MyPa(5) = 정찰(Oee)
MySe(5) = 센스(Oee)
MyCo(5) = 컨트롤(Oee)
MyYear(5) = OYear(Oee)
MyRank(5) = 랭크(Oee)
MyTeam(5) = Team(Oee)
MySkill(5) = Skill(Oee)
PlayNumber(5) = Oee
Randomize Oee

If Mode = "Normal" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hard" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hell" Then
    CR = Int((120 * Rnd) + 1)
End If
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Do Until (종족(Oee) = 3)
If Mode = "Normal" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hard" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hell" Then
    CR = Int((120 * Rnd) + 1)
End If
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until 랭크(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until 랭크(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until 랭크(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until 랭크(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until 랭크(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until 랭크(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Loop
MyName(6) = 이름(Oee)
MyTribe(6) = 3
MyAt(6) = 공격력(Oee)
MyR(6) = 견제(Oee)
MySt(6) = 전략(Oee)
MyAm(6) = 물량(Oee)
MyDe(6) = 수비력(Oee)
MyPa(6) = 정찰(Oee)
MySe(6) = 센스(Oee)
MyCo(6) = 컨트롤(Oee)
MyYear(6) = OYear(Oee)
MyRank(6) = 랭크(Oee)
MyTeam(6) = Team(Oee)
MySkill(6) = Skill(Oee)
PlayNumber(6) = Oee
Randomize Oee


돈량 = 0
For 돌려 = 1 To 6
 If MyRank(돌려) = "Normal" Then
  돈량 = val(돈량) + 1
 ElseIf MyRank(돌려) = "Special" Then
  돈량 = val(돈량) + 2
 ElseIf MyRank(돌려) = "Rare" Then
  돈량 = val(돈량) + 3
 ElseIf MyRank(돌려) = "Unique" Then
  돈량 = val(돈량) + 4
 ElseIf MyRank(돌려) = "Elite" Then
  돈량 = val(돈량) + 5
 ElseIf MyRank(돌려) = "Legend" Then
  돈량 = val(돈량) + 6
 ElseIf MyRank(돌려) = "Secret" Then
  돈량 = val(돈량) + 7
 End If
Next

If val(돈량) = 6 Then
 Money = 25000
ElseIf (val(돈량) >= 7) And (val(돈량) <= 12) Then
 Money = 20000
ElseIf (val(돈량) >= 13) And (val(돈량) <= 18) Then
 Money = 15000
ElseIf (val(돈량) >= 19) And (val(돈량) <= 24) Then
 Money = 10000
ElseIf (val(돈량) >= 25) And (val(돈량) <= 30) Then
 Money = 5000
ElseIf (val(돈량) >= 31) And (val(돈량) <= 36) Then
 Money = 2000
ElseIf (val(돈량) >= 37) And (val(돈량) <= 42) Then
 Money = 1000
End If
확인용1 = val(Money) / 2

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
Next 갠리그

TeamName = InputBox("닉네임을 입력하세요.")
If Mode = "Normal" Then
    For i = 1 To 6
        MyAt(i) = val(MyAt(i)) + 50
        MyR(i) = val(MyR(i)) + 50
        MySt(i) = val(MySt(i)) + 50
        MyAm(i) = val(MyAm(i)) + 50
        MyDe(i) = val(MyDe(i)) + 50
        MyPa(i) = val(MyPa(i)) + 50
        MySe(i) = val(MySe(i)) + 50
        MyCo(i) = val(MyCo(i)) + 50
    Next
    For Oee = 0 To 800
        NPC공격력(Oee) = val(NPC공격력(Oee)) + 50
        NPC견제(Oee) = val(NPC견제(Oee)) + 50
        NPC전략(Oee) = val(NPC전략(Oee)) + 50
        NPC물량(Oee) = val(NPC물량(Oee)) + 50
        NPC수비력(Oee) = val(NPC수비력(Oee)) + 50
        NPC정찰(Oee) = val(NPC정찰(Oee)) + 50
        NPC센스(Oee) = val(NPC센스(Oee)) + 50
        NPC컨트롤(Oee) = val(NPC컨트롤(Oee)) + 50
    Next
ElseIf Mode = "Hell" Then
    For Oee = 0 To 800
        공격력(Oee) = val(공격력(Oee)) + 50
        견제(Oee) = val(견제(Oee)) + 50
        전략(Oee) = val(전략(Oee)) + 50
        물량(Oee) = val(물량(Oee)) + 50
        수비력(Oee) = val(수비력(Oee)) + 50
        정찰(Oee) = val(정찰(Oee)) + 50
        센스(Oee) = val(센스(Oee)) + 50
        컨트롤(Oee) = val(컨트롤(Oee)) + 50
    Next
End If
하향 = 0
하향횟수 = 0
FrmMain.Show
Unload Me
Else

MsgBox "코드오류"
End If
End Sub

Private Sub Label1_Click()
Dim 전체코드 As String
전체코드 = InputBox("코드입력")
If 전체코드 = "55886248" Then
 Command1.Enabled = True
 Command2.Enabled = True
 Command1.Visible = True
 Command2.Visible = True
ElseIf 전체코드 = "Data 수정용" Then
    Open App.Path & "\Data\Data수정용.txt" For Output As #1
        For i = 1 To 800
            Print #1, "작은" & i & "번째 선수"
            Print #1, "이름(" & i & ") = 큰따옴표" & 이름(i) & "큰따옴표"
            Print #1, "랭크(" & i & ") = 큰따옴표" & 랭크(i) & "큰따옴표"
            Print #1, "OYear(" & i & ") = 큰따옴표" & OYear(i) & "큰따옴표"
            Print #1, "Team(" & i & ") = 큰따옴표" & Team(i) & "큰따옴표"
            Print #1, "종족(" & i & ") = 큰따옴표" & 종족(i) & "큰따옴표"
            Print #1, "공격력(" & i & ") = 큰따옴표" & 공격력(i) & "큰따옴표"
            Print #1, "견제(" & i & ") = 큰따옴표" & 견제(i) & "큰따옴표"
            Print #1, "전략(" & i & ") = 큰따옴표" & 전략(i) & "큰따옴표"
            Print #1, "물량(" & i & ") = 큰따옴표" & 물량(i) & "큰따옴표"
            Print #1, "수비력(" & i & ") = 큰따옴표" & 수비력(i) & "큰따옴표"
            Print #1, "정찰(" & i & ") = 큰따옴표" & 정찰(i) & "큰따옴표"
            Print #1, "센스(" & i & ") = 큰따옴표" & 센스(i) & "큰따옴표"
            Print #1, "컨트롤(" & i & ") = 큰따옴표" & 컨트롤(i) & "큰따옴표"
            Print #1, ""
        Next
    Close #1
    MsgBox "완료"
ElseIf 전체코드 = "전체Setting" Then
    Open App.Path & "\전체 능력치.txt" For Output As #1
        For i = 1 To 800
            Print #1, "선수번호 : " & i
            Print #1, OYear(i) & 이름(i)
            Print #1, 랭크(i)
            Print #1, 공격력(i)
            Print #1, 견제(i)
            Print #1, 전략(i)
            Print #1, 물량(i)
            Print #1, 수비력(i)
            Print #1, 정찰(i)
            Print #1, 센스(i)
            Print #1, 컨트롤(i)
            Print #1, "----------------"
            Print #1, "----------------"
        Next
    Close #1
ElseIf 전체코드 = "Setting" Then
    For i = 1 To 800
        통계(1, i) = 공격력(i)
        통계(2, i) = 견제(i)
        통계(3, i) = 전략(i)
        통계(4, i) = 물량(i)
        통계(5, i) = 수비력(i)
        통계(6, i) = 정찰(i)
        통계(7, i) = 센스(i)
        통계(8, i) = 컨트롤(i)
        이름통계(i) = 이름(i)
        랭크통계(i) = 랭크(i)
        년도통계(i) = OYear(i)
    Next
    
    For M우세 = 1 To 800
        For O우세 = 1 To 8
            총합(M우세) = 총합(M우세) + 통계(O우세, M우세)
        Next
    Next
    
    For M우세 = 1 To 1000
        For i = 1 To 799
            If 총합(i) < 총합(i + 1) Then
                통계도움(4) = 총합(i)
                총합(i) = 총합(i + 1)
                총합(i + 1) = val(통계도움(4))
                
                통계도움(1) = 이름통계(i)
                이름통계(i) = 이름통계(i + 1)
                이름통계(i + 1) = 통계도움(1)
                
                통계도움(2) = 랭크통계(i)
                랭크통계(i) = 랭크통계(i + 1)
                랭크통계(i + 1) = 통계도움(2)
                
                통계도움(3) = 년도통계(i)
                년도통계(i) = 년도통계(i + 1)
                년도통계(i + 1) = 통계도움(3)
            End If
        Next
    Next

    Open App.Path & "\능력치 통계.txt" For Output As #1
        For i = 1 To 800
                Print #1, i & ". " & 년도통계(i) & 이름통계(i) & " ㅡ " & 랭크통계(i) & ", 총합 : " & 총합(i)
        Next
    Close #1

    MsgBox "완료"
ElseIf 전체코드 = "Say" Then
    Open App.Path & "\능력치잘못된애들.txt" For Output As #1
    '능력치 확인
    ''Normal
    Print #1, "ㅡNormal"
    For Oee = 1 To 800
        If 랭크(Oee) = "Normal" Then
            Print #1, OYear(Oee) & 이름(Oee) & " ㅡ " & 랭크(Oee)
        End If
    Next
    
    ''Special
    Print #1, "ㅡSpecial"
    For Oee = 1 To 800
        If 랭크(Oee) = "Special" Then
            Print #1, OYear(Oee) & 이름(Oee) & " ㅡ " & 랭크(Oee)
        End If
    Next
    
    ''Rare
    Print #1, "ㅡRare"
    For Oee = 1 To 800
        If 랭크(Oee) = "Rare" Then
            Print #1, OYear(Oee) & 이름(Oee) & " ㅡ " & 랭크(Oee)
        End If
    Next
    
    ''Unique
    Print #1, "ㅡUnique"
    For Oee = 1 To 800
        If 랭크(Oee) = "Unique" Then
            Print #1, OYear(Oee) & 이름(Oee) & " ㅡ " & 랭크(Oee)
        End If
    Next
    
    ''Elite
    Print #1, "ㅡElite"
    For Oee = 1 To 800
        If 랭크(Oee) = "Elite" Then
            Print #1, OYear(Oee) & 이름(Oee) & " ㅡ " & 랭크(Oee)
        End If
    Next
    
    ''Legend
    Print #1, "ㅡLegend"
    For Oee = 1 To 800
        If 랭크(Oee) = "Legend" Then
            Print #1, OYear(Oee) & 이름(Oee) & " ㅡ " & 랭크(Oee)
        End If
    Next
    
    ''Secret
    Print #1, "ㅡSecret"
    For Oee = 1 To 800
        If 랭크(Oee) = "Secret" Then
            Print #1, OYear(Oee) & 이름(Oee) & " ㅡ " & 랭크(Oee)
        End If
    Next
    
    ''Champion
    Print #1, "ㅡChampion"
    For Oee = 1 To 800
        If 랭크(Oee) = "Champion" Then
            Print #1, OYear(Oee) & 이름(Oee) & " ㅡ " & 랭크(Oee)
        End If
    Next
    DoEvents
    MsgBox "완료"
ElseIf 전체코드 = "목록" Then
    '랭크별 인원 체크 & 랭크별 선수 목록 체크
    Dim k As Integer, sum As Integer
    sum = 0
    
    Open App.Path & "\목록.txt" For Output As #1
    For i = 1 To 8
        Select Case i
        Case 1
            For k = 0 To 800
                If 랭크(k) = "Normal" Then
                    sum = sum + 1
                End If
            Next
            Print #1, "ㅡNormal : " & sum & "명ㅡ"
            For k = 0 To 800
                If 랭크(k) = "Normal" Then
                    Print #1, "<" & OYear(k) & ">  " & 이름(k)
                End If
            Next
        Case 2
            For k = 0 To 800
                If 랭크(k) = "Special" Then
                    sum = sum + 1
                End If
            Next
            Print #1, "ㅡSpecial : " & sum & "명ㅡ"
            For k = 0 To 800
                If 랭크(k) = "Special" Then
                    Print #1, "<" & OYear(k) & ">  " & 이름(k)
                End If
            Next
        Case 3
            For k = 0 To 800
                If 랭크(k) = "Rare" Then
                    sum = sum + 1
                End If
            Next
            Print #1, "ㅡRare : " & sum & "명ㅡ"
            For k = 0 To 800
                If 랭크(k) = "Rare" Then
                    Print #1, "<" & OYear(k) & ">  " & 이름(k)
                End If
            Next
        Case 4
            For k = 0 To 800
                If 랭크(k) = "Unique" Then
                    sum = sum + 1
                End If
            Next
            Print #1, "ㅡUnique : " & sum & "명ㅡ"
            For k = 0 To 800
                If 랭크(k) = "Unique" Then
                    Print #1, "<" & OYear(k) & ">  " & 이름(k)
                End If
            Next
        Case 5
            For k = 0 To 800
                If 랭크(k) = "Elite" Then
                    sum = sum + 1
                End If
            Next
            Print #1, "ㅡElite : " & sum & "명ㅡ"
            For k = 0 To 800
                If 랭크(k) = "Elite" Then
                    Print #1, "<" & OYear(k) & ">  " & 이름(k)
                End If
            Next
        Case 6
            For k = 0 To 800
                If 랭크(k) = "Legend" Then
                    sum = sum + 1
                End If
            Next
            Print #1, "ㅡLegend : " & sum & "명ㅡ"
            For k = 0 To 800
                If 랭크(k) = "Legend" Then
                    Print #1, "<" & OYear(k) & ">  " & 이름(k)
                End If
            Next
        Case 7
            For k = 0 To 800
                If 랭크(k) = "Secret" Then
                    sum = sum + 1
                End If
            Next
            Print #1, "ㅡSecret : " & sum & "명ㅡ"
            For k = 0 To 800
                If 랭크(k) = "Secret" Then
                    Print #1, "<" & OYear(k) & ">  " & 이름(k)
                End If
            Next
        Case 8
            For k = 0 To 800
                If 랭크(k) = "Champion" Then
                    sum = sum + 1
                End If
            Next
            Print #1, "ㅡChampion : " & sum & "명ㅡ"
            For k = 0 To 800
                If 랭크(k) = "Champion" Then
                    Print #1, "<" & OYear(k) & ">  " & 이름(k)
                End If
            Next
        End Select
    sum = 0
    Next
    MsgBox "완료"
    Close #1
Else
 MsgBox "코드 오류입니다."
End If
End Sub

Private Sub Text1_Change()
Timer5.Enabled = True
End Sub


Private Sub Tim07_Timer()
For Oee = 119 To 260

If Oee = 119 Then
이름(Oee) = "강민"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 700
전략(Oee) = 750
물량(Oee) = 600
수비력(Oee) = 750
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 120 Then
이름(Oee) = "김동수"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 600
수비력(Oee) = 400
정찰(Oee) = 450
센스(Oee) = 450
컨트롤(Oee) = 500



ElseIf Oee = 121 Then
이름(Oee) = "김윤환"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 122 Then
이름(Oee) = "박정석"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 500
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 850



ElseIf Oee = 123 Then
이름(Oee) = "배병우"
랭크(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
종족(Oee) = 2
공격력(Oee) = 900
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 850
수비력(Oee) = 850
정찰(Oee) = 700
센스(Oee) = 700
컨트롤(Oee) = 800



ElseIf Oee = 124 Then
이름(Oee) = "변길섭"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
종족(Oee) = 1
공격력(Oee) = 850
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 550
수비력(Oee) = 500
정찰(Oee) = 400
센스(Oee) = 400
컨트롤(Oee) = 650



ElseIf Oee = 125 Then
이름(Oee) = "이병민"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
종족(Oee) = 1
공격력(Oee) = 550
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 550



ElseIf Oee = 126 Then
이름(Oee) = "이영호"
랭크(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
종족(Oee) = 1
공격력(Oee) = 900
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 850
수비력(Oee) = 750
정찰(Oee) = 600
센스(Oee) = 750
컨트롤(Oee) = 750



ElseIf Oee = 127 Then
이름(Oee) = "이영호1"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 700
전략(Oee) = 750
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 700



ElseIf Oee = 128 Then
이름(Oee) = "임재덕"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 600



ElseIf Oee = 129 Then
이름(Oee) = "정명호"
랭크(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 600
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 750
컨트롤(Oee) = 750



ElseIf Oee = 130 Then
이름(Oee) = "조용호"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 500
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 500
컨트롤(Oee) = 600



ElseIf Oee = 131 Then
이름(Oee) = "홍진호"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
종족(Oee) = 2
공격력(Oee) = 850
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 500
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 700
컨트롤(Oee) = 600



ElseIf Oee = 132 Then
이름(Oee) = "김동건"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "삼성전자"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 133 Then
이름(Oee) = "박성준1"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 800
견제(Oee) = 600
전략(Oee) = 500
물량(Oee) = 700
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 500
컨트롤(Oee) = 600



ElseIf Oee = 134 Then
이름(Oee) = "박성훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "삼성전자"
종족(Oee) = 3
공격력(Oee) = 550
견제(Oee) = 650
전략(Oee) = 800
물량(Oee) = 500
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 600


ElseIf Oee = 135 Then
이름(Oee) = "변은종"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 800
견제(Oee) = 600
전략(Oee) = 450
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 500
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 136 Then
이름(Oee) = "송병구"
랭크(Oee) = "Legend"
OYear(Oee) = "<07>"
Team(Oee) = "삼성전자"
종족(Oee) = 3
공격력(Oee) = 900
견제(Oee) = 750
전략(Oee) = 800
물량(Oee) = 950
수비력(Oee) = 700
정찰(Oee) = 800
센스(Oee) = 950
컨트롤(Oee) = 950



ElseIf Oee = 137 Then
이름(Oee) = "이성은"
랭크(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "삼성전자"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 800
전략(Oee) = 850
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 650
센스(Oee) = 850
컨트롤(Oee) = 700



ElseIf Oee = 138 Then
이름(Oee) = "이재황"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 139 Then
이름(Oee) = "이창훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 500
전략(Oee) = 650
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 140 Then
이름(Oee) = "임채성"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "삼성전자"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 600
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 700



ElseIf Oee = 141 Then
이름(Oee) = "장용석"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "삼성전자"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 750



ElseIf Oee = 142 Then
이름(Oee) = "주영달"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 800
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 800



ElseIf Oee = 142 Then
이름(Oee) = "허영무"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "삼성전자"
종족(Oee) = 3
공격력(Oee) = 850
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 850
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 143 Then
이름(Oee) = "김구현"
랭크(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
종족(Oee) = 3
공격력(Oee) = 800
견제(Oee) = 900
전략(Oee) = 600
물량(Oee) = 900
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 800



ElseIf Oee = 144 Then
이름(Oee) = "김민제"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 600



ElseIf Oee = 145 Then
이름(Oee) = "김윤중"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 500
정찰(Oee) = 600
센스(Oee) = 450
컨트롤(Oee) = 550



ElseIf Oee = 146 Then
이름(Oee) = "김윤환1"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 650



ElseIf Oee = 147 Then
이름(Oee) = "박성진"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 550
수비력(Oee) = 700
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 500



ElseIf Oee = 148 Then
이름(Oee) = "박정욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 500
물량(Oee) = 800
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 149 Then
이름(Oee) = "박종수"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 700
전략(Oee) = 650
물량(Oee) = 600
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 150 Then
이름(Oee) = "서지수"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 400
물량(Oee) = 350
수비력(Oee) = 350
정찰(Oee) = 350
센스(Oee) = 400
컨트롤(Oee) = 500



ElseIf Oee = 151 Then
이름(Oee) = "이철민"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 600
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 152 Then
이름(Oee) = "조일장"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 153 Then
이름(Oee) = "진영수"
랭크(Oee) = "Unique"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 850
견제(Oee) = 950
전략(Oee) = 700
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 900
센스(Oee) = 850
컨트롤(Oee) = 900



ElseIf Oee = 154 Then
이름(Oee) = "최연식"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 700
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 450
컨트롤(Oee) = 550



ElseIf Oee = 155 Then
이름(Oee) = "김남기"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "한빛"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 700
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 156 Then
이름(Oee) = "김동주"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "한빛"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 157 Then
이름(Oee) = "김명운"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "한빛"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 500
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 800



ElseIf Oee = 158 Then
이름(Oee) = "김병욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "한빛"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 550



ElseIf Oee = 159 Then
이름(Oee) = "김인기"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "한빛"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 650



ElseIf Oee = 160 Then
이름(Oee) = "김승현"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "한빛"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 550
전략(Oee) = 650
물량(Oee) = 750
수비력(Oee) = 550
정찰(Oee) = 500
센스(Oee) = 550
컨트롤(Oee) = 700



ElseIf Oee = 161 Then
이름(Oee) = "김준영"
랭크(Oee) = "Unique"
OYear(Oee) = "<07>"
Team(Oee) = "한빛"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 950
수비력(Oee) = 950
정찰(Oee) = 900
센스(Oee) = 750
컨트롤(Oee) = 750



ElseIf Oee = 162 Then
이름(Oee) = "문지훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "한빛"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 600
물량(Oee) = 600
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 550



ElseIf Oee = 163 Then
이름(Oee) = "박경락"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "한빛"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 850
전략(Oee) = 700
물량(Oee) = 500
수비력(Oee) = 550
정찰(Oee) = 500
센스(Oee) = 550
컨트롤(Oee) = 550



ElseIf Oee = 164 Then
이름(Oee) = "설현호"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "한빛"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 550
전략(Oee) = 700
물량(Oee) = 650
수비력(Oee) = 450
정찰(Oee) = 500
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 165 Then
이름(Oee) = "신정민"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "한빛"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 450
물량(Oee) = 800
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 550



ElseIf Oee = 166 Then
이름(Oee) = "윤용태"
랭크(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "한빛"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 750
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 750



ElseIf Oee = 167 Then
이름(Oee) = "임진묵"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "한빛"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 500
센스(Oee) = 550
컨트롤(Oee) = 700



ElseIf Oee = 168 Then
이름(Oee) = "고인규"
랭크(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 850
수비력(Oee) = 750
정찰(Oee) = 700
센스(Oee) = 650
컨트롤(Oee) = 750



ElseIf Oee = 169 Then
이름(Oee) = "권오혁"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 500
컨트롤(Oee) = 550



ElseIf Oee = 170 Then
이름(Oee) = "김성제"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
종족(Oee) = 3
공격력(Oee) = 400
견제(Oee) = 800
전략(Oee) = 600
물량(Oee) = 400
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 550
컨트롤(Oee) = 800



ElseIf Oee = 171 Then
이름(Oee) = "도재욱"
랭크(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
종족(Oee) = 3
공격력(Oee) = 950
견제(Oee) = 750
전략(Oee) = 500
물량(Oee) = 950
수비력(Oee) = 750
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 650



ElseIf Oee = 172 Then
이름(Oee) = "박대경"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
종족(Oee) = 3
공격력(Oee) = 850
견제(Oee) = 750
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 500
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 173 Then
이름(Oee) = "박성준"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
종족(Oee) = 2
공격력(Oee) = 800
견제(Oee) = 900
전략(Oee) = 600
물량(Oee) = 800
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 800



ElseIf Oee = 174 Then
이름(Oee) = "박용욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 550
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 750



ElseIf Oee = 175 Then
이름(Oee) = "박태민"
랭크(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 900
수비력(Oee) = 900
정찰(Oee) = 850
센스(Oee) = 650
컨트롤(Oee) = 750



ElseIf Oee = 176 Then
이름(Oee) = "손승재"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 450
컨트롤(Oee) = 550



ElseIf Oee = 177 Then
이름(Oee) = "오충훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 800
수비력(Oee) = 800
정찰(Oee) = 650
센스(Oee) = 500
컨트롤(Oee) = 700



ElseIf Oee = 178 Then
이름(Oee) = "윤종민"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 450
물량(Oee) = 750
수비력(Oee) = 750
정찰(Oee) = 700
센스(Oee) = 450
컨트롤(Oee) = 600



ElseIf Oee = 179 Then
이름(Oee) = "이건준"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 700



ElseIf Oee = 180 Then
이름(Oee) = "전상욱"
랭크(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 700
물량(Oee) = 850
수비력(Oee) = 900
정찰(Oee) = 850
센스(Oee) = 600
컨트롤(Oee) = 700



ElseIf Oee = 181 Then
이름(Oee) = "최연성"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
종족(Oee) = 1
공격력(Oee) = 550
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 750
정찰(Oee) = 650
센스(Oee) = 750
컨트롤(Oee) = 750



ElseIf Oee = 182 Then
이름(Oee) = "강구열"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 183 Then
이름(Oee) = "고석현"
랭크(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 850
견제(Oee) = 600
전략(Oee) = 800
물량(Oee) = 800
수비력(Oee) = 850
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 900



ElseIf Oee = 184 Then
이름(Oee) = "김동현"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 500
컨트롤(Oee) = 600



ElseIf Oee = 185 Then
이름(Oee) = "김재훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 800
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 186 Then
이름(Oee) = "김태훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 550
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 450
컨트롤(Oee) = 600



ElseIf Oee = 187 Then
이름(Oee) = "김택용"
랭크(Oee) = "Elite"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
종족(Oee) = 3
공격력(Oee) = 800
견제(Oee) = 950
전략(Oee) = 700
물량(Oee) = 850
수비력(Oee) = 700
정찰(Oee) = 950
센스(Oee) = 950
컨트롤(Oee) = 800



ElseIf Oee = 188 Then
이름(Oee) = "민찬기"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 550
수비력(Oee) = 550
정찰(Oee) = 500
센스(Oee) = 550
컨트롤(Oee) = 750



ElseIf Oee = 189 Then
이름(Oee) = "박지호"
랭크(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
종족(Oee) = 3
공격력(Oee) = 900
견제(Oee) = 700
전략(Oee) = 650
물량(Oee) = 900
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 600



ElseIf Oee = 190 Then
이름(Oee) = "서경종"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 500
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 600
컨트롤(Oee) = 800



ElseIf Oee = 191 Then
이름(Oee) = "이재호"
랭크(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 800
전략(Oee) = 600
물량(Oee) = 800
수비력(Oee) = 850
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 750



ElseIf Oee = 192 Then
이름(Oee) = "정영철"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 500
컨트롤(Oee) = 550



ElseIf Oee = 193 Then
이름(Oee) = "염보성"
랭크(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 700
전략(Oee) = 750
물량(Oee) = 900
수비력(Oee) = 850
정찰(Oee) = 700
센스(Oee) = 700
컨트롤(Oee) = 650



ElseIf Oee = 194 Then
이름(Oee) = "구성훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "르까프"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 195 Then
이름(Oee) = "김경모"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "르까프"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 500
컨트롤(Oee) = 600



ElseIf Oee = 196 Then
이름(Oee) = "김성곤"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "르까프"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 500
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 550



ElseIf Oee = 197 Then
이름(Oee) = "김정환"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "르까프"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 500
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 700



ElseIf Oee = 198 Then
이름(Oee) = "박지수"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "르까프"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 500
물량(Oee) = 900
수비력(Oee) = 900
정찰(Oee) = 800
센스(Oee) = 700
컨트롤(Oee) = 600



ElseIf Oee = 199 Then
이름(Oee) = "오영종"
랭크(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "르까프"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 850
전략(Oee) = 850
물량(Oee) = 900
수비력(Oee) = 550
정찰(Oee) = 750
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 200 Then
이름(Oee) = "이유석"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "르까프"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 500
컨트롤(Oee) = 500



ElseIf Oee = 201 Then
이름(Oee) = "이제동"
랭크(Oee) = "Unique"
OYear(Oee) = "<07>"
Team(Oee) = "르까프"
종족(Oee) = 2
공격력(Oee) = 900
견제(Oee) = 950
전략(Oee) = 700
물량(Oee) = 850
수비력(Oee) = 650
정찰(Oee) = 750
센스(Oee) = 800
컨트롤(Oee) = 950



ElseIf Oee = 202 Then
이름(Oee) = "이학주"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "르까프"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 500
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 550
컨트롤(Oee) = 500



ElseIf Oee = 203 Then
이름(Oee) = "손주흥"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "르까프"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 700
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 600



ElseIf Oee = 204 Then
이름(Oee) = "손찬웅"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "르까프"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 800
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 700
센스(Oee) = 600
컨트롤(Oee) = 700



ElseIf Oee = 205 Then
이름(Oee) = "최가람"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "르까프"
종족(Oee) = 2
공격력(Oee) = 900
견제(Oee) = 700
전략(Oee) = 600
물량(Oee) = 400
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 750



ElseIf Oee = 206 Then
이름(Oee) = "권수현"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 207 Then
이름(Oee) = "김민호"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 450
컨트롤(Oee) = 600



ElseIf Oee = 208 Then
이름(Oee) = "김성기"
랭크(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
종족(Oee) = 1
공격력(Oee) = 900
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 900
수비력(Oee) = 800
정찰(Oee) = 700
센스(Oee) = 600
컨트롤(Oee) = 550



ElseIf Oee = 209 Then
이름(Oee) = "마재윤"
랭크(Oee) = "Secret"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 850
견제(Oee) = 850
전략(Oee) = 850
물량(Oee) = 950
수비력(Oee) = 950
정찰(Oee) = 900
센스(Oee) = 850
컨트롤(Oee) = 800
Skill(Oee) = 6


ElseIf Oee = 210 Then
이름(Oee) = "박영민"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 600
전략(Oee) = 800
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 700
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 211 Then
이름(Oee) = "변형태"
랭크(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
종족(Oee) = 1
공격력(Oee) = 950
견제(Oee) = 800
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 800
컨트롤(Oee) = 800



ElseIf Oee = 212 Then
이름(Oee) = "서지훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 700
센스(Oee) = 750
컨트롤(Oee) = 800



ElseIf Oee = 213 Then
이름(Oee) = "손재범"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 550



ElseIf Oee = 214 Then
이름(Oee) = "장육"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 700
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 700



ElseIf Oee = 215 Then
이름(Oee) = "조병세"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 500
센스(Oee) = 450
컨트롤(Oee) = 600



ElseIf Oee = 216 Then
이름(Oee) = "주현준"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 217 Then
이름(Oee) = "한상봉"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 900
견제(Oee) = 750
전략(Oee) = 700
물량(Oee) = 600
수비력(Oee) = 500
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 750



ElseIf Oee = 218 Then
이름(Oee) = "김광섭"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "온게임넷"
종족(Oee) = 2
공격력(Oee) = 550
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 219 Then
이름(Oee) = "김상욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "온게임넷"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 700
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 700
센스(Oee) = 550
컨트롤(Oee) = 700



ElseIf Oee = 220 Then
이름(Oee) = "신상문"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "온게임넷"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 221 Then
이름(Oee) = "안상원"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "온게임넷"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 550
컨트롤(Oee) = 550



ElseIf Oee = 222 Then
이름(Oee) = "이승훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "온게임넷"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 800
전략(Oee) = 700
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 800
컨트롤(Oee) = 700



ElseIf Oee = 223 Then
이름(Oee) = "이종미"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "온게임넷"
종족(Oee) = 2
공격력(Oee) = 350
견제(Oee) = 300
전략(Oee) = 350
물량(Oee) = 400
수비력(Oee) = 400
정찰(Oee) = 400
센스(Oee) = 350
컨트롤(Oee) = 350



ElseIf Oee = 224 Then
이름(Oee) = "임원기"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "온게임넷"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 700



ElseIf Oee = 225 Then
이름(Oee) = "전태규"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "온게임넷"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 500
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 750
센스(Oee) = 650
컨트롤(Oee) = 500



ElseIf Oee = 226 Then
이름(Oee) = "차재욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "온게임넷"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 550
물량(Oee) = 550
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 227 Then
이름(Oee) = "곽동훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
종족(Oee) = 2
공격력(Oee) = 550
견제(Oee) = 500
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 550



ElseIf Oee = 228 Then
이름(Oee) = "김덕인"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 500
컨트롤(Oee) = 500



ElseIf Oee = 229 Then
이름(Oee) = "김민구"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 550
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 550
컨트롤(Oee) = 700



ElseIf Oee = 230 Then
이름(Oee) = "김원기"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 750



ElseIf Oee = 231 Then
이름(Oee) = "남승현"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 550
전략(Oee) = 550
물량(Oee) = 500
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 600



ElseIf Oee = 232 Then
이름(Oee) = "박문기"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 750
전략(Oee) = 500
물량(Oee) = 800
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 750



ElseIf Oee = 233 Then
이름(Oee) = "서기수"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
종족(Oee) = 3
공격력(Oee) = 850
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 850
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 650



ElseIf Oee = 234 Then
이름(Oee) = "신상호"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 800
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 235 Then
이름(Oee) = "조용성"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 600



ElseIf Oee = 236 Then
이름(Oee) = "최욱명"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 650
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 500
컨트롤(Oee) = 550



ElseIf Oee = 237 Then
이름(Oee) = "강도경"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "공군"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 500
물량(Oee) = 550
수비력(Oee) = 450
정찰(Oee) = 450
센스(Oee) = 500
컨트롤(Oee) = 500



ElseIf Oee = 238 Then
이름(Oee) = "김선기"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "공군"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 450
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 500



ElseIf Oee = 239 Then
이름(Oee) = "김환중"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "공군"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 550



ElseIf Oee = 240 Then
이름(Oee) = "박대만"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "공군"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 500
전략(Oee) = 450
물량(Oee) = 750
수비력(Oee) = 500
정찰(Oee) = 600
센스(Oee) = 500
컨트롤(Oee) = 550



ElseIf Oee = 241 Then
이름(Oee) = "성학승"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "공군"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 600



ElseIf Oee = 242 Then
이름(Oee) = "이재훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "공군"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 450
전략(Oee) = 450
물량(Oee) = 700
수비력(Oee) = 500
정찰(Oee) = 600
센스(Oee) = 450
컨트롤(Oee) = 600



ElseIf Oee = 243 Then
이름(Oee) = "이주영"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "공군"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 600
물량(Oee) = 900
수비력(Oee) = 900
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 550



ElseIf Oee = 244 Then
이름(Oee) = "임요환"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "공군"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 800
전략(Oee) = 750
물량(Oee) = 500
수비력(Oee) = 550
정찰(Oee) = 700
센스(Oee) = 900
컨트롤(Oee) = 650



ElseIf Oee = 245 Then
이름(Oee) = "조형근"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "공군"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 700
물량(Oee) = 500
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 550



ElseIf Oee = 246 Then
이름(Oee) = "최인규"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "공군"
종족(Oee) = 1
공격력(Oee) = 500
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 500
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 550



ElseIf Oee = 247 Then
이름(Oee) = "김상우"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "위메이드"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 550
전략(Oee) = 550
물량(Oee) = 500
수비력(Oee) = 450
정찰(Oee) = 500
센스(Oee) = 450
컨트롤(Oee) = 650



ElseIf Oee = 248 Then
이름(Oee) = "김재춘"
랭크(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "위메이드"
종족(Oee) = 2
공격력(Oee) = 800
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 850
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 850



ElseIf Oee = 249 Then
이름(Oee) = "나도현"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "위메이드"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 400
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 600
컨트롤(Oee) = 750



ElseIf Oee = 250 Then
이름(Oee) = "박영훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "위메이드"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 450
수비력(Oee) = 450
정찰(Oee) = 500
센스(Oee) = 450
컨트롤(Oee) = 650



ElseIf Oee = 251 Then
이름(Oee) = "박세정"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "위메이드"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 800
전략(Oee) = 650
물량(Oee) = 550
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 252 Then
이름(Oee) = "박성균"
랭크(Oee) = "Unique"
OYear(Oee) = "<07>"
Team(Oee) = "위메이드"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 900
수비력(Oee) = 950
정찰(Oee) = 850
센스(Oee) = 800
컨트롤(Oee) = 800



ElseIf Oee = 253 Then
이름(Oee) = "손영훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "위메이드"
종족(Oee) = 3
공격력(Oee) = 550
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 550
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 550
컨트롤(Oee) = 550



ElseIf Oee = 254 Then
이름(Oee) = "심소명"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "위메이드"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 550
전략(Oee) = 800
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 550



ElseIf Oee = 255 Then
이름(Oee) = "안기효"
랭크(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "위메이드"
종족(Oee) = 3
공격력(Oee) = 850
견제(Oee) = 600
전략(Oee) = 700
물량(Oee) = 950
수비력(Oee) = 800
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 900



ElseIf Oee = 256 Then
이름(Oee) = "임동혁"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "위메이드"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 700
물량(Oee) = 500
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 450
컨트롤(Oee) = 550



ElseIf Oee = 257 Then
이름(Oee) = "이윤열"
랭크(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "위메이드"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 750
물량(Oee) = 850
수비력(Oee) = 800
정찰(Oee) = 700
센스(Oee) = 800
컨트롤(Oee) = 800



ElseIf Oee = 258 Then
이름(Oee) = "전태양"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "위메이드"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 700
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 450
컨트롤(Oee) = 550



ElseIf Oee = 259 Then
이름(Oee) = "한동욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "위메이드"
종족(Oee) = 1
공격력(Oee) = 850
견제(Oee) = 850
전략(Oee) = 600
물량(Oee) = 450
수비력(Oee) = 450
정찰(Oee) = 500
센스(Oee) = 600
컨트롤(Oee) = 950



ElseIf Oee = 260 Then
이름(Oee) = "한동훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "위메이드"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 700
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 700
End If
 우승(Oee) = 0
 준우승(Oee) = 0
 컨디션(Oee) = 100
 A승리(Oee) = 0
 A패배(Oee) = 0
 P승리(Oee) = 0
 P패배(Oee) = 0
 T승리(Oee) = 0
 T패배(Oee) = 0
 Z승리(Oee) = 0
 Z패배(Oee) = 0
 T연승(Oee) = 0
 Z연승(Oee) = 0
 P연승(Oee) = 0
 A연승(Oee) = 0
 T연(Oee) = "W"
 Z연(Oee) = "W"
 P연(Oee) = "W"
 A연(Oee) = "W"
Next Oee

Tim08.Enabled = True
Tim07.Enabled = False
End Sub

Private Sub Tim08_Timer()
For Oee = 261 To 407
If Oee = 261 Then
이름(Oee) = "고강민"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
종족(Oee) = 2
공격력(Oee) = 550
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 550



ElseIf Oee = 262 Then
이름(Oee) = "김대엽"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 700
컨트롤(Oee) = 550



ElseIf Oee = 263 Then
이름(Oee) = "김영진"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 264 Then
이름(Oee) = "김윤환"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 650



ElseIf Oee = 265 Then
이름(Oee) = "김재춘"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 700



ElseIf Oee = 266 Then
이름(Oee) = "박재영"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 900
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 650
컨트롤(Oee) = 550



ElseIf Oee = 267 Then
이름(Oee) = "박준우"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 700
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 500
정찰(Oee) = 450
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 268 Then
이름(Oee) = "배병우"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 269 Then
이름(Oee) = "우정호"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 270 Then
이름(Oee) = "이영호"
랭크(Oee) = "Unique"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 750
전략(Oee) = 850
물량(Oee) = 900
수비력(Oee) = 900
정찰(Oee) = 750
센스(Oee) = 900
컨트롤(Oee) = 650



ElseIf Oee = 271 Then
이름(Oee) = "이영호1"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 272 Then
이름(Oee) = "임재덕"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 600



ElseIf Oee = 273 Then
이름(Oee) = "장주현"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 450
정찰(Oee) = 650
센스(Oee) = 500
컨트롤(Oee) = 500



ElseIf Oee = 274 Then
이름(Oee) = "정명호"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 275 Then
이름(Oee) = "김동건"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "삼성전자"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 600



ElseIf Oee = 276 Then
이름(Oee) = "박동수"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "삼성전자"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 500
물량(Oee) = 500
수비력(Oee) = 550
정찰(Oee) = 500
센스(Oee) = 450
컨트롤(Oee) = 700



ElseIf Oee = 277 Then
이름(Oee) = "박성훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 550
견제(Oee) = 600
전략(Oee) = 750
물량(Oee) = 600
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 278 Then
이름(Oee) = "손석희"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "삼성전자"
종족(Oee) = 3
공격력(Oee) = 550
견제(Oee) = 500
전략(Oee) = 600
물량(Oee) = 500
수비력(Oee) = 650
정찰(Oee) = 500
센스(Oee) = 600
컨트롤(Oee) = 700



ElseIf Oee = 279 Then
이름(Oee) = "송병구"
랭크(Oee) = "Elite"
OYear(Oee) = "<08>"
Team(Oee) = "삼성전자"
종족(Oee) = 3
공격력(Oee) = 900
견제(Oee) = 850
전략(Oee) = 850
물량(Oee) = 850
수비력(Oee) = 850
정찰(Oee) = 850
센스(Oee) = 700
컨트롤(Oee) = 850



ElseIf Oee = 280 Then
이름(Oee) = "유준희"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 281 Then
이름(Oee) = "이성은"
랭크(Oee) = "Rare"
OYear(Oee) = "<08>"
Team(Oee) = "삼성전자"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 800
전략(Oee) = 750
물량(Oee) = 800
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 900
컨트롤(Oee) = 800



ElseIf Oee = 282 Then
이름(Oee) = "이재황"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 283 Then
이름(Oee) = "임채성"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "삼성전자"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 550
물량(Oee) = 550
수비력(Oee) = 450
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 700



ElseIf Oee = 284 Then
이름(Oee) = "임태규"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "삼성전자"
종족(Oee) = 3
공격력(Oee) = 550
견제(Oee) = 600
전략(Oee) = 500
물량(Oee) = 550
수비력(Oee) = 450
정찰(Oee) = 450
센스(Oee) = 450
컨트롤(Oee) = 550



ElseIf Oee = 285 Then
이름(Oee) = "주영달"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 700



ElseIf Oee = 286 Then
이름(Oee) = "차명환"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 700
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 287 Then
이름(Oee) = "최윤선"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "삼성전자"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 500
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 650



ElseIf Oee = 288 Then
이름(Oee) = "허영무"
랭크(Oee) = "Elite"
OYear(Oee) = "<08>"
Team(Oee) = "삼성전자"
종족(Oee) = 3
공격력(Oee) = 850
견제(Oee) = 850
전략(Oee) = 750
물량(Oee) = 900
수비력(Oee) = 700
정찰(Oee) = 750
센스(Oee) = 800
컨트롤(Oee) = 950



ElseIf Oee = 289 Then
이름(Oee) = "김경효"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 290 Then
이름(Oee) = "김구현"
랭크(Oee) = "Rare"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
종족(Oee) = 3
공격력(Oee) = 800
견제(Oee) = 900
전략(Oee) = 850
물량(Oee) = 800
수비력(Oee) = 650
정찰(Oee) = 750
센스(Oee) = 750
컨트롤(Oee) = 800



ElseIf Oee = 291 Then
이름(Oee) = "김민제"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 650
컨트롤(Oee) = 600



ElseIf Oee = 292 Then
이름(Oee) = "김윤중"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 500
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 293 Then
이름(Oee) = "김윤환1"
랭크(Oee) = "Rare"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 850
견제(Oee) = 750
전략(Oee) = 700
물량(Oee) = 900
수비력(Oee) = 750
정찰(Oee) = 650
센스(Oee) = 750
컨트롤(Oee) = 650



ElseIf Oee = 294 Then
이름(Oee) = "김현우"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 550
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 650
컨트롤(Oee) = 700



ElseIf Oee = 295 Then
이름(Oee) = "박성준"
랭크(Oee) = "Unique"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 900
견제(Oee) = 800
전략(Oee) = 600
물량(Oee) = 900
수비력(Oee) = 800
정찰(Oee) = 650
센스(Oee) = 850
컨트롤(Oee) = 900



ElseIf Oee = 296 Then
이름(Oee) = "박종수"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 297 Then
이름(Oee) = "서지수"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 450
수비력(Oee) = 450
정찰(Oee) = 500
센스(Oee) = 450
컨트롤(Oee) = 650



ElseIf Oee = 298 Then
이름(Oee) = "이신형"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 550



ElseIf Oee = 299 Then
이름(Oee) = "조성호"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 500



ElseIf Oee = 300 Then
이름(Oee) = "조일장"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 600



ElseIf Oee = 301 Then
이름(Oee) = "진영수"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 800
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 700
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 302 Then
이름(Oee) = "강민구"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "웅진"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 750
센스(Oee) = 750
컨트롤(Oee) = 800



ElseIf Oee = 303 Then
이름(Oee) = "김남기"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "웅진"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 304 Then
이름(Oee) = "김동주"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "웅진"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 305 Then
이름(Oee) = "김명운"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "웅진"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 700
전략(Oee) = 650
물량(Oee) = 750
수비력(Oee) = 750
정찰(Oee) = 700
센스(Oee) = 700
컨트롤(Oee) = 750



ElseIf Oee = 306 Then
이름(Oee) = "김승현"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "웅진"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 700

ElseIf Oee = 307 Then
이름(Oee) = "김인기"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "웅진"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 650



ElseIf Oee = 308 Then
이름(Oee) = "김준영"
랭크(Oee) = "Special"
OYear(Oee) = "<08>"
Team(Oee) = "웅진"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 850
수비력(Oee) = 800
정찰(Oee) = 800
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 309 Then
이름(Oee) = "문지훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "웅진"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 310 Then
이름(Oee) = "신정민"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "웅진"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 450
물량(Oee) = 800
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 550



ElseIf Oee = 311 Then
이름(Oee) = "윤용태"
랭크(Oee) = "Special"
OYear(Oee) = "<08>"
Team(Oee) = "웅진"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 750
정찰(Oee) = 800
센스(Oee) = 650
컨트롤(Oee) = 750



ElseIf Oee = 312 Then
이름(Oee) = "이동준"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "웅진"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 600
수비력(Oee) = 700
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 550



ElseIf Oee = 313 Then
이름(Oee) = "이형연"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "웅진"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 500



ElseIf Oee = 314 Then
이름(Oee) = "임진묵"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "웅진"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 650



ElseIf Oee = 315 Then
이름(Oee) = "정종현"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "웅진"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 316 Then
이름(Oee) = "고인규"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 700
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 700
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 317 Then
이름(Oee) = "권오혁"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 318 Then
이름(Oee) = "김성제"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
종족(Oee) = 3
공격력(Oee) = 500
견제(Oee) = 800
전략(Oee) = 600
물량(Oee) = 500
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 750



ElseIf Oee = 319 Then
이름(Oee) = "김택용"
랭크(Oee) = "Unique"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 900
전략(Oee) = 750
물량(Oee) = 850
수비력(Oee) = 750
정찰(Oee) = 850
센스(Oee) = 900
컨트롤(Oee) = 850



ElseIf Oee = 320 Then
이름(Oee) = "도재욱"
랭크(Oee) = "Unique"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
종족(Oee) = 3
공격력(Oee) = 950
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 950
수비력(Oee) = 700
정찰(Oee) = 750
센스(Oee) = 900
컨트롤(Oee) = 750



ElseIf Oee = 321 Then
이름(Oee) = "박대경"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 322 Then
이름(Oee) = "박재혁"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 323 Then
이름(Oee) = "박태민"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
종족(Oee) = 2
공격력(Oee) = 550
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 700
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 324 Then
이름(Oee) = "오충훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
종족(Oee) = 1
공격력(Oee) = 550
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 750
수비력(Oee) = 750
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 550



ElseIf Oee = 325 Then
이름(Oee) = "윤종민"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 326 Then
이름(Oee) = "이승석"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
종족(Oee) = 2
공격력(Oee) = 550
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 327 Then
이름(Oee) = "임요환"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 800
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 328 Then
이름(Oee) = "전상욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 329 Then
이름(Oee) = "정영철"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
종족(Oee) = 2
공격력(Oee) = 550
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 700



ElseIf Oee = 330 Then
이름(Oee) = "최연성"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
종족(Oee) = 1
공격력(Oee) = 500
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 331 Then
이름(Oee) = "강구열"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 600
전략(Oee) = 750
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 450
센스(Oee) = 700
컨트롤(Oee) = 600



ElseIf Oee = 332 Then
이름(Oee) = "고석현"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 800
견제(Oee) = 750
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 333 Then
이름(Oee) = "김동현"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 650



ElseIf Oee = 334 Then
이름(Oee) = "김재훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 800
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 600



ElseIf Oee = 335 Then
이름(Oee) = "김태훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 500
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 450
컨트롤(Oee) = 650



ElseIf Oee = 336 Then
이름(Oee) = "민찬기"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 750
전략(Oee) = 700
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 800



ElseIf Oee = 337 Then
이름(Oee) = "박수범"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 700
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 600



ElseIf Oee = 338 Then
이름(Oee) = "박지호"
랭크(Oee) = "Special"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
종족(Oee) = 3
공격력(Oee) = 900
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 800
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 600



ElseIf Oee = 339 Then
이름(Oee) = "서경종"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 750
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 700



ElseIf Oee = 340 Then
이름(Oee) = "염보성"
랭크(Oee) = "Special"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 750
전략(Oee) = 650
물량(Oee) = 800
수비력(Oee) = 800
정찰(Oee) = 750
센스(Oee) = 700
컨트롤(Oee) = 750



ElseIf Oee = 341 Then
이름(Oee) = "이재호"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 700
전략(Oee) = 750
물량(Oee) = 600
수비력(Oee) = 650
정찰(Oee) = 700
센스(Oee) = 700
컨트롤(Oee) = 650



ElseIf Oee = 342 Then
이름(Oee) = "전흥식"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 600



ElseIf Oee = 343 Then
이름(Oee) = "구성훈"
랭크(Oee) = "Special"
OYear(Oee) = "<08>"
Team(Oee) = "르까프"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 750
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 344 Then
이름(Oee) = "김경모"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "르까프"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 345 Then
이름(Oee) = "김민혁"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "르까프"
종족(Oee) = 1
공격력(Oee) = 500
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 500
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 500



ElseIf Oee = 346 Then
이름(Oee) = "김정환"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "르까프"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 650
물량(Oee) = 550
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 347 Then
이름(Oee) = "김태균"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "르까프"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 500



ElseIf Oee = 348 Then
이름(Oee) = "노영훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "르까프"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 550
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 500
센스(Oee) = 600
컨트롤(Oee) = 600


ElseIf Oee = 349 Then
이름(Oee) = "박지수"
랭크(Oee) = "Unique"
OYear(Oee) = "<08>"
Team(Oee) = "르까프"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 800
전략(Oee) = 800
물량(Oee) = 850
수비력(Oee) = 900
정찰(Oee) = 700
센스(Oee) = 800
컨트롤(Oee) = 750

 

ElseIf Oee = 350 Then
이름(Oee) = "손주흥"
랭크(Oee) = "Special"
OYear(Oee) = "<08>"
Team(Oee) = "르까프"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 800



ElseIf Oee = 351 Then
이름(Oee) = "손찬웅"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "르까프"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 800
수비력(Oee) = 500
정찰(Oee) = 700
센스(Oee) = 800
컨트롤(Oee) = 650



ElseIf Oee = 352 Then
이름(Oee) = "이제동"
랭크(Oee) = "Unique"
OYear(Oee) = "<08>"
Team(Oee) = "르까프"
종족(Oee) = 2
공격력(Oee) = 900
견제(Oee) = 850
전략(Oee) = 750
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 900
컨트롤(Oee) = 900



ElseIf Oee = 353 Then
이름(Oee) = "이학주"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "르까프"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 354 Then
이름(Oee) = "황보건우"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "르까프"
종족(Oee) = 2
공격력(Oee) = 500
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 500
컨트롤(Oee) = 500



ElseIf Oee = 355 Then
이름(Oee) = "권수현"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 356 Then
이름(Oee) = "김국군"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 550
물량(Oee) = 550
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 700



ElseIf Oee = 357 Then
이름(Oee) = "김민호"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 500
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 450
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 358 Then
이름(Oee) = "김정우"
랭크(Oee) = "Rare"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 750
전략(Oee) = 700
물량(Oee) = 800
수비력(Oee) = 800
정찰(Oee) = 700
센스(Oee) = 750
컨트롤(Oee) = 750



ElseIf Oee = 359 Then
이름(Oee) = "박영민"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 750
물량(Oee) = 750
수비력(Oee) = 650
정찰(Oee) = 700
센스(Oee) = 700
컨트롤(Oee) = 750



ElseIf Oee = 360 Then
이름(Oee) = "변형태"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 800
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 361 Then
이름(Oee) = "서지훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 362 Then
이름(Oee) = "손재범"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 650
컨트롤(Oee) = 500



ElseIf Oee = 363 Then
이름(Oee) = "조병세"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 364 Then
이름(Oee) = "주현준"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 700
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 365 Then
이름(Oee) = "진영화"
랭크(Oee) = "Rare"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
종족(Oee) = 3
공격력(Oee) = 800
견제(Oee) = 850
전략(Oee) = 650
물량(Oee) = 800
수비력(Oee) = 850
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 366 Then
이름(Oee) = "한상봉"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 900
견제(Oee) = 700
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 600
컨트롤(Oee) = 800



ElseIf Oee = 367 Then
이름(Oee) = "김광섭"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "온게임넷"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 368 Then
이름(Oee) = "김상욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "온게임넷"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 650
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 369 Then
이름(Oee) = "김학수"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "온게임넷"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 550
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 550



ElseIf Oee = 370 Then
이름(Oee) = "신상문"
랭크(Oee) = "Legend"
OYear(Oee) = "<08>"
Team(Oee) = "온게임넷"
종족(Oee) = 1
공격력(Oee) = 950
견제(Oee) = 800
전략(Oee) = 750
물량(Oee) = 900
수비력(Oee) = 750
정찰(Oee) = 750
센스(Oee) = 950
컨트롤(Oee) = 950



ElseIf Oee = 371 Then
이름(Oee) = "정명훈"
랭크(Oee) = "Rare"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 800
전략(Oee) = 800
물량(Oee) = 850
수비력(Oee) = 800
정찰(Oee) = 650
센스(Oee) = 850
컨트롤(Oee) = 650



ElseIf Oee = 372 Then
이름(Oee) = "안상원"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "온게임넷"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 750
센스(Oee) = 700
컨트롤(Oee) = 650



ElseIf Oee = 373 Then
이름(Oee) = "이경민"
랭크(Oee) = "Special"
OYear(Oee) = "<08>"
Team(Oee) = "온게임넷"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 800
전략(Oee) = 750
물량(Oee) = 850
수비력(Oee) = 750
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 374 Then
이름(Oee) = "이승훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "온게임넷"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 750
전략(Oee) = 700
물량(Oee) = 700
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 375 Then
이름(Oee) = "임원기"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "온게임넷"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 700
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 376 Then
이름(Oee) = "조재걸"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "온게임넷"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 650
컨트롤(Oee) = 700



ElseIf Oee = 377 Then
이름(Oee) = "김지성"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
종족(Oee) = 2
공격력(Oee) = 500
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 500
컨트롤(Oee) = 550



ElseIf Oee = 378 Then
이름(Oee) = "남승현"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 600



ElseIf Oee = 379 Then
이름(Oee) = "박문기"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 750
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 380 Then
이름(Oee) = "박상우"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 750
정찰(Oee) = 700
센스(Oee) = 750
컨트롤(Oee) = 650



ElseIf Oee = 381 Then
이름(Oee) = "서기수"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 382 Then
이름(Oee) = "신대근"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 383 Then
이름(Oee) = "신상호"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
종족(Oee) = 3
공격력(Oee) = 850
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 650



ElseIf Oee = 384 Then
이름(Oee) = "안수형"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 500



ElseIf Oee = 385 Then
이름(Oee) = "이호준"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
종족(Oee) = 1
공격력(Oee) = 550
견제(Oee) = 500
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 750
정찰(Oee) = 700
센스(Oee) = 600
컨트롤(Oee) = 500



ElseIf Oee = 386 Then
이름(Oee) = "최지성"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 650
전략(Oee) = 500
물량(Oee) = 500
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 650



ElseIf Oee = 387 Then
이름(Oee) = "김선기"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "공군"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 388 Then
이름(Oee) = "김환중"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "공군"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 389 Then
이름(Oee) = "박대만"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "공군"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 700



ElseIf Oee = 389 Then
이름(Oee) = "박정석"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "공군"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 750



ElseIf Oee = 390 Then
이름(Oee) = "성학승"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "공군"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 750
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 600


ElseIf Oee = 391 Then
이름(Oee) = "오영종"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "공군"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 700
전략(Oee) = 650
물량(Oee) = 800
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 700



ElseIf Oee = 392 Then
이름(Oee) = "이재훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "공군"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 700



ElseIf Oee = 393 Then
이름(Oee) = "이주영"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "공군"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 750



ElseIf Oee = 394 Then
이름(Oee) = "차재욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "공군"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 550
물량(Oee) = 550
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 550



ElseIf Oee = 395 Then
이름(Oee) = "한동욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "공군"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 750
전략(Oee) = 650
물량(Oee) = 600
수비력(Oee) = 450
정찰(Oee) = 500
센스(Oee) = 650
컨트롤(Oee) = 800



ElseIf Oee = 396 Then
이름(Oee) = "홍진호"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "공군"
종족(Oee) = 2
공격력(Oee) = 800
견제(Oee) = 750
전략(Oee) = 650
물량(Oee) = 500
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 650
컨트롤(Oee) = 600



ElseIf Oee = 397 Then
이름(Oee) = "김성진"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "위메이드"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 500
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 398 Then
이름(Oee) = "박성균"
랭크(Oee) = "Special"
OYear(Oee) = "<08>"
Team(Oee) = "위메이드"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 750
전략(Oee) = 700
물량(Oee) = 800
수비력(Oee) = 800
정찰(Oee) = 750
센스(Oee) = 800
컨트롤(Oee) = 700



ElseIf Oee = 399 Then
이름(Oee) = "박세정"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "위메이드"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 650



ElseIf Oee = 400 Then
이름(Oee) = "손영훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "위메이드"
종족(Oee) = 3
공격력(Oee) = 550
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 550
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 550
컨트롤(Oee) = 550



ElseIf Oee = 401 Then
이름(Oee) = "신노열"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "위메이드"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 402 Then
이름(Oee) = "안기효"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "위메이드"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 403 Then
이름(Oee) = "이영한"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "위메이드"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 700
전략(Oee) = 650
물량(Oee) = 600
수비력(Oee) = 500
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 700



ElseIf Oee = 404 Then
이름(Oee) = "이윤열"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "위메이드"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 700
센스(Oee) = 600
컨트롤(Oee) = 700



ElseIf Oee = 405 Then
이름(Oee) = "임동혁"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "위메이드"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 700



ElseIf Oee = 406 Then
이름(Oee) = "전태양"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "위메이드"
종족(Oee) = 1
공격력(Oee) = 550
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 750
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 600



ElseIf Oee = 407 Then
이름(Oee) = "한동훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "위메이드"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 650
End If
 우승(Oee) = 0
 준우승(Oee) = 0
 컨디션(Oee) = 100
 A승리(Oee) = 0
 A패배(Oee) = 0
 P승리(Oee) = 0
 P패배(Oee) = 0
 T승리(Oee) = 0
 T패배(Oee) = 0
 Z승리(Oee) = 0
 Z패배(Oee) = 0
 T연승(Oee) = 0
 Z연승(Oee) = 0
 P연승(Oee) = 0
 A연승(Oee) = 0
 T연(Oee) = "W"
 Z연(Oee) = "W"
 P연(Oee) = "W"
 A연(Oee) = "W"

Next Oee
Tim09.Enabled = True
Tim08.Enabled = False
End Sub

Private Sub Tim09_Timer()
For Oee = 408 To 539
If Oee = 408 Then
이름(Oee) = "고강민"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 700
수비력(Oee) = 650
정찰(Oee) = 700
센스(Oee) = 600
컨트롤(Oee) = 550



ElseIf Oee = 409 Then
이름(Oee) = "김대엽"
랭크(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
종족(Oee) = 3
공격력(Oee) = 850
견제(Oee) = 700
전략(Oee) = 600
물량(Oee) = 850
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 410 Then
이름(Oee) = "김재춘"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 700



ElseIf Oee = 411 Then
이름(Oee) = "남승현"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 600



ElseIf Oee = 412 Then
이름(Oee) = "박재영"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
종족(Oee) = 3
공격력(Oee) = 550
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 800
수비력(Oee) = 700
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 413 Then
이름(Oee) = "박지수"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 800
정찰(Oee) = 700
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 414 Then
이름(Oee) = "박찬수"
랭크(Oee) = "Unique"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
종족(Oee) = 2
공격력(Oee) = 900
견제(Oee) = 800
전략(Oee) = 650
물량(Oee) = 750
수비력(Oee) = 800
정찰(Oee) = 750
센스(Oee) = 800
컨트롤(Oee) = 950



ElseIf Oee = 415 Then
이름(Oee) = "배병우"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 750
수비력(Oee) = 550
정찰(Oee) = 700
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 416 Then
이름(Oee) = "우정호"
랭크(Oee) = "Elite"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
종족(Oee) = 3
공격력(Oee) = 900
견제(Oee) = 800
전략(Oee) = 750
물량(Oee) = 850
수비력(Oee) = 850
정찰(Oee) = 850
센스(Oee) = 750
컨트롤(Oee) = 850



ElseIf Oee = 417 Then
이름(Oee) = "이영호"
랭크(Oee) = "Unique"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 700
전략(Oee) = 750
물량(Oee) = 950
수비력(Oee) = 900
정찰(Oee) = 800
센스(Oee) = 950
컨트롤(Oee) = 700



ElseIf Oee = 418 Then
이름(Oee) = "황병영"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 750
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 550



ElseIf Oee = 419 Then
이름(Oee) = "박대호"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "삼성전자"
종족(Oee) = 1
공격력(Oee) = 550
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 600
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 420 Then
이름(Oee) = "박동수"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "삼성전자"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 500
센스(Oee) = 550
컨트롤(Oee) = 700



ElseIf Oee = 421 Then
이름(Oee) = "손석희"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "삼성전자"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 750
정찰(Oee) = 700
센스(Oee) = 600
컨트롤(Oee) = 750



ElseIf Oee = 422 Then
이름(Oee) = "유준희"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 423 Then
이름(Oee) = "이성은"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "삼성전자"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 750
전략(Oee) = 800
물량(Oee) = 550
수비력(Oee) = 500
정찰(Oee) = 650
센스(Oee) = 800
컨트롤(Oee) = 800



ElseIf Oee = 424 Then
이름(Oee) = "이재황"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 500
컨트롤(Oee) = 550



ElseIf Oee = 425 Then
이름(Oee) = "이정현"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 500
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 550



ElseIf Oee = 426 Then
이름(Oee) = "임채성"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "삼성전자"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 550
수비력(Oee) = 450
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 700



ElseIf Oee = 427 Then
이름(Oee) = "임태규"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "삼성전자"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 450
컨트롤(Oee) = 550



ElseIf Oee = 428 Then
이름(Oee) = "주영달"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 700



ElseIf Oee = 429 Then
이름(Oee) = "차명환"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 750
물량(Oee) = 700
수비력(Oee) = 750
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 430 Then
이름(Oee) = "최윤선"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "삼성전자"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 550
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 700



ElseIf Oee = 431 Then
이름(Oee) = "허영무"
랭크(Oee) = "Rare"
OYear(Oee) = "<09>"
Team(Oee) = "삼성전자"
종족(Oee) = 3
공격력(Oee) = 800
견제(Oee) = 800
전략(Oee) = 650
물량(Oee) = 800
수비력(Oee) = 800
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 900



ElseIf Oee = 432 Then
이름(Oee) = "김경효"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 433 Then
이름(Oee) = "김구현"
랭크(Oee) = "Rare"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
종족(Oee) = 3
공격력(Oee) = 850
견제(Oee) = 800
전략(Oee) = 700
물량(Oee) = 850
수비력(Oee) = 650
정찰(Oee) = 700
센스(Oee) = 800
컨트롤(Oee) = 750



ElseIf Oee = 434 Then
이름(Oee) = "김동건"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 600



ElseIf Oee = 435 Then
이름(Oee) = "김성현"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 550
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 436 Then
이름(Oee) = "김윤중"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 650
물량(Oee) = 800
수비력(Oee) = 550
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 437 Then
이름(Oee) = "김윤환1"
랭크(Oee) = "Elite"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 850
견제(Oee) = 750
전략(Oee) = 850
물량(Oee) = 800
수비력(Oee) = 850
정찰(Oee) = 800
센스(Oee) = 900
컨트롤(Oee) = 900



ElseIf Oee = 438 Then
이름(Oee) = "김현우"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 650
컨트롤(Oee) = 750



ElseIf Oee = 439 Then
이름(Oee) = "박성준"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 900
견제(Oee) = 700
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 750



ElseIf Oee = 440 Then
이름(Oee) = "박종수"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 441 Then
이름(Oee) = "서지수"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 450
수비력(Oee) = 450
정찰(Oee) = 500
센스(Oee) = 450
컨트롤(Oee) = 650



ElseIf Oee = 442 Then
이름(Oee) = "이신형"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 500
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 600



ElseIf Oee = 443 Then
이름(Oee) = "조성호"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 500



ElseIf Oee = 444 Then
이름(Oee) = "조일장"
랭크(Oee) = "Rare"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 800
견제(Oee) = 750
전략(Oee) = 500
물량(Oee) = 900
수비력(Oee) = 900
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 800



ElseIf Oee = 445 Then
이름(Oee) = "진영수"
랭크(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 850
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 750
센스(Oee) = 750
컨트롤(Oee) = 800



ElseIf Oee = 446 Then
이름(Oee) = "강민구"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "웅진"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 447 Then
이름(Oee) = "김동주"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "웅진"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 700
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 448 Then
이름(Oee) = "김명운"
랭크(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "웅진"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 750
물량(Oee) = 800
수비력(Oee) = 800
정찰(Oee) = 800
센스(Oee) = 700
컨트롤(Oee) = 800



ElseIf Oee = 449 Then
이름(Oee) = "김승현"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "웅진"
종족(Oee) = 3
공격력(Oee) = 800
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 800
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 700



ElseIf Oee = 450 Then
이름(Oee) = "김영진"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "웅진"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 451 Then
이름(Oee) = "노준규"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "웅진"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 452 Then
이름(Oee) = "박대만"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "웅진"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 500
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 700



ElseIf Oee = 453 Then
이름(Oee) = "박정훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "웅진"
종족(Oee) = 3
공격력(Oee) = 550
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 700
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 650



ElseIf Oee = 454 Then
이름(Oee) = "윤용태"
랭크(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "웅진"
종족(Oee) = 3
공격력(Oee) = 800
견제(Oee) = 750
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 900



ElseIf Oee = 455 Then
이름(Oee) = "이동준"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "웅진"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 600
수비력(Oee) = 700
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 550



ElseIf Oee = 456 Then
이름(Oee) = "임정현"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "웅진"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 650
수비력(Oee) = 450
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 700



ElseIf Oee = 457 Then
이름(Oee) = "임진묵"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "웅진"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 700
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 750
센스(Oee) = 600
컨트롤(Oee) = 850



ElseIf Oee = 458 Then
이름(Oee) = "정종현"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "웅진"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 459 Then
이름(Oee) = "한상봉"
랭크(Oee) = "Rare"
OYear(Oee) = "<09>"
Team(Oee) = "웅진"
종족(Oee) = 2
공격력(Oee) = 900
견제(Oee) = 850
전략(Oee) = 700
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 700
센스(Oee) = 750
컨트롤(Oee) = 950



ElseIf Oee = 460 Then
이름(Oee) = "고인규"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 750
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 700
센스(Oee) = 700
컨트롤(Oee) = 750



ElseIf Oee = 461 Then
이름(Oee) = "권오혁"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
종족(Oee) = 3
공격력(Oee) = 550
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 462 Then
이름(Oee) = "김택용"
랭크(Oee) = "Unique"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 950
전략(Oee) = 750
물량(Oee) = 800
수비력(Oee) = 750
정찰(Oee) = 900
센스(Oee) = 850
컨트롤(Oee) = 900



ElseIf Oee = 463 Then
이름(Oee) = "도재욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
종족(Oee) = 3
공격력(Oee) = 800
견제(Oee) = 700
전략(Oee) = 650
물량(Oee) = 900
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 650



ElseIf Oee = 464 Then
이름(Oee) = "박재혁"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 700
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 750
컨트롤(Oee) = 800



ElseIf Oee = 465 Then
이름(Oee) = "어윤수"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 500
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 550



ElseIf Oee = 466 Then
이름(Oee) = "이승석"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 467 Then
이름(Oee) = "임요환"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 800
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 468 Then
이름(Oee) = "정명훈"
랭크(Oee) = "Rare"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 800
전략(Oee) = 800
물량(Oee) = 800
수비력(Oee) = 900
정찰(Oee) = 750
센스(Oee) = 850
컨트롤(Oee) = 600



ElseIf Oee = 469 Then
이름(Oee) = "정영철"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 750
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 470 Then
이름(Oee) = "최연성"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
종족(Oee) = 1
공격력(Oee) = 550
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 471 Then
이름(Oee) = "최호선"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 500
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 500
센스(Oee) = 650
컨트롤(Oee) = 600



ElseIf Oee = 472 Then
이름(Oee) = "고석현"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 800
견제(Oee) = 750
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 700



ElseIf Oee = 473 Then
이름(Oee) = "김동현"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 650



ElseIf Oee = 474 Then
이름(Oee) = "김재훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 475 Then
이름(Oee) = "김태훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 500
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 476 Then
이름(Oee) = "박수범"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 600



ElseIf Oee = 477 Then
이름(Oee) = "박지호"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
종족(Oee) = 3
공격력(Oee) = 850
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 800
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 478 Then
이름(Oee) = "서경종"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 750
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 700



ElseIf Oee = 479 Then
이름(Oee) = "염보성"
랭크(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 850
수비력(Oee) = 800
정찰(Oee) = 800
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 480 Then
이름(Oee) = "이재호"
랭크(Oee) = "Rare"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 800
전략(Oee) = 750
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 800
센스(Oee) = 750
컨트롤(Oee) = 800



ElseIf Oee = 481 Then
이름(Oee) = "임성진"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 700
전략(Oee) = 600
물량(Oee) = 500
수비력(Oee) = 500
정찰(Oee) = 650
센스(Oee) = 500
컨트롤(Oee) = 500



ElseIf Oee = 482 Then
이름(Oee) = "장민철"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 650



ElseIf Oee = 483 Then
이름(Oee) = "전흥식"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 600



ElseIf Oee = 484 Then
이름(Oee) = "정우서"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 485 Then
이름(Oee) = "강동현"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "화승"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 450
물량(Oee) = 550
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 450
컨트롤(Oee) = 650



ElseIf Oee = 486 Then
이름(Oee) = "구성훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "화승"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 750
전략(Oee) = 750
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 600



ElseIf Oee = 487 Then
이름(Oee) = "김경모"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "화승"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 488 Then
이름(Oee) = "김민혁"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "화승"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 550
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 489 Then
이름(Oee) = "김태균"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "화승"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 490 Then
이름(Oee) = "노영훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "화승"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 450
컨트롤(Oee) = 500



ElseIf Oee = 491 Then
이름(Oee) = "박준오"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "화승"
종족(Oee) = 2
공격력(Oee) = 800
견제(Oee) = 700
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 700
센스(Oee) = 550
컨트롤(Oee) = 750



ElseIf Oee = 492 Then
이름(Oee) = "방태수"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "화승"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 500
수비력(Oee) = 450
정찰(Oee) = 450
센스(Oee) = 450
컨트롤(Oee) = 550



ElseIf Oee = 493 Then
이름(Oee) = "손주흥"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "화승"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 700
수비력(Oee) = 750
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 750



ElseIf Oee = 494 Then
이름(Oee) = "손찬웅"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "화승"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 900
전략(Oee) = 700
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 750



ElseIf Oee = 495 Then
이름(Oee) = "이제동"
랭크(Oee) = "Legend"
OYear(Oee) = "<09>"
Team(Oee) = "화승"
종족(Oee) = 2
공격력(Oee) = 950
견제(Oee) = 950
전략(Oee) = 750
물량(Oee) = 800
수비력(Oee) = 800
정찰(Oee) = 750
센스(Oee) = 950
컨트롤(Oee) = 950



ElseIf Oee = 496 Then
이름(Oee) = "이학주"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "화승"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 497 Then
이름(Oee) = "임원기"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "화승"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 700
센스(Oee) = 550
컨트롤(Oee) = 650



ElseIf Oee = 498 Then
이름(Oee) = "권수현"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 499 Then
이름(Oee) = "김민호"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 500
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 500 Then
이름(Oee) = "김정우"
랭크(Oee) = "Elite"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 900
견제(Oee) = 800
전략(Oee) = 650
물량(Oee) = 900
수비력(Oee) = 900
정찰(Oee) = 850
센스(Oee) = 800
컨트롤(Oee) = 800


ElseIf Oee = 501 Then
이름(Oee) = "변형태"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
종족(Oee) = 1
공격력(Oee) = 850
견제(Oee) = 750
전략(Oee) = 650
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 650



ElseIf Oee = 502 Then
이름(Oee) = "손재범"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 650
컨트롤(Oee) = 500



ElseIf Oee = 503 Then
이름(Oee) = "신동원"
랭크(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 800
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 900
수비력(Oee) = 750
정찰(Oee) = 650
센스(Oee) = 550
컨트롤(Oee) = 750



ElseIf Oee = 504 Then
이름(Oee) = "이재훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 700



ElseIf Oee = 505 Then
이름(Oee) = "이주영"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 750
컨트롤(Oee) = 750



ElseIf Oee = 506 Then
이름(Oee) = "장윤철"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 550



ElseIf Oee = 507 Then
이름(Oee) = "조병세"
랭크(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
종족(Oee) = 1
공격력(Oee) = 900
견제(Oee) = 750
전략(Oee) = 650
물량(Oee) = 750
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 750
컨트롤(Oee) = 700



ElseIf Oee = 508 Then
이름(Oee) = "진영화"
랭크(Oee) = "Rare"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 750
전략(Oee) = 750
물량(Oee) = 800
수비력(Oee) = 850
정찰(Oee) = 650
센스(Oee) = 750
컨트롤(Oee) = 700



ElseIf Oee = 509 Then
이름(Oee) = "김상욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "하이트"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 700
전략(Oee) = 600
물량(Oee) = 800
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 600



ElseIf Oee = 510 Then
이름(Oee) = "김학수"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "하이트"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 550
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 550



ElseIf Oee = 511 Then
이름(Oee) = "신상문"
랭크(Oee) = "Rare"
OYear(Oee) = "<09>"
Team(Oee) = "하이트"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 800
전략(Oee) = 750
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 700
센스(Oee) = 850
컨트롤(Oee) = 800



ElseIf Oee = 512 Then
이름(Oee) = "안준영"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "하이트"
종족(Oee) = 2
공격력(Oee) = 550
견제(Oee) = 550
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 650



ElseIf Oee = 513 Then
이름(Oee) = "이경민"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "하이트"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 850
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 514 Then
이름(Oee) = "김도우"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 650
정찰(Oee) = 500
센스(Oee) = 450
컨트롤(Oee) = 550



ElseIf Oee = 515 Then
이름(Oee) = "김성대"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 500
물량(Oee) = 700
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 516 Then
이름(Oee) = "박상우"
랭크(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 850
수비력(Oee) = 800
정찰(Oee) = 750
센스(Oee) = 750
컨트롤(Oee) = 650



ElseIf Oee = 517 Then
이름(Oee) = "서기수"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 600



ElseIf Oee = 518 Then
이름(Oee) = "신대근"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 750
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 700
센스(Oee) = 700
컨트롤(Oee) = 700



ElseIf Oee = 519 Then
이름(Oee) = "신상호"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
종족(Oee) = 3
공격력(Oee) = 850
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 650



ElseIf Oee = 520 Then
이름(Oee) = "신재욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 750
수비력(Oee) = 800
정찰(Oee) = 700
센스(Oee) = 600
컨트롤(Oee) = 700



ElseIf Oee = 521 Then
이름(Oee) = "안수형"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 500



ElseIf Oee = 522 Then
이름(Oee) = "정명호"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 523 Then
이름(Oee) = "최지성"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 650



ElseIf Oee = 524 Then
이름(Oee) = "민찬기"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "공군"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 750
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 750



ElseIf Oee = 525 Then
이름(Oee) = "박영민"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "공군"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 750
물량(Oee) = 800
수비력(Oee) = 650
정찰(Oee) = 700
센스(Oee) = 600
컨트롤(Oee) = 700



ElseIf Oee = 526 Then
이름(Oee) = "박정석"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "공군"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 500
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 750



ElseIf Oee = 527 Then
이름(Oee) = "박태민"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "공군"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 700
센스(Oee) = 650
컨트롤(Oee) = 700



ElseIf Oee = 527 Then
이름(Oee) = "서지훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "공군"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 528 Then
이름(Oee) = "오영종"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "공군"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 750
전략(Oee) = 650
물량(Oee) = 800
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 700



ElseIf Oee = 529 Then
이름(Oee) = "차재욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "공군"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 550
물량(Oee) = 550
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 550



ElseIf Oee = 530 Then
이름(Oee) = "한동욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "공군"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 750
전략(Oee) = 650
물량(Oee) = 600
수비력(Oee) = 450
정찰(Oee) = 500
센스(Oee) = 600
컨트롤(Oee) = 800



ElseIf Oee = 531 Then
이름(Oee) = "홍진호"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "공군"
종족(Oee) = 2
공격력(Oee) = 850
견제(Oee) = 750
전략(Oee) = 750
물량(Oee) = 500
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 531 Then
이름(Oee) = "강정우"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "위메이드"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 600



ElseIf Oee = 532 Then
이름(Oee) = "박성균"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "위메이드"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 700
전략(Oee) = 650
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 750
센스(Oee) = 700
컨트롤(Oee) = 650



ElseIf Oee = 533 Then
이름(Oee) = "박세정"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "위메이드"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 800
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 534 Then
이름(Oee) = "신노열"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "위메이드"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 750
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 700
정찰(Oee) = 700
센스(Oee) = 600
컨트롤(Oee) = 750



ElseIf Oee = 535 Then
이름(Oee) = "안기효"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "위메이드"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 650



ElseIf Oee = 536 Then
이름(Oee) = "이영한"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "위메이드"
종족(Oee) = 2
공격력(Oee) = 850
견제(Oee) = 750
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 750



ElseIf Oee = 537 Then
이름(Oee) = "이영호1"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "위메이드"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 550
센스(Oee) = 650
컨트롤(Oee) = 650



ElseIf Oee = 538 Then
이름(Oee) = "전상욱"
랭크(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "위메이드"
종족(Oee) = 1
공격력(Oee) = 850
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 850
수비력(Oee) = 800
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 600



ElseIf Oee = 539 Then
이름(Oee) = "전태양"
랭크(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "위메이드"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 800
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 650
End If
 우승(Oee) = 0
 준우승(Oee) = 0
 컨디션(Oee) = 100
 A승리(Oee) = 0
 A패배(Oee) = 0
 P승리(Oee) = 0
 P패배(Oee) = 0
 T승리(Oee) = 0
 T패배(Oee) = 0
 Z승리(Oee) = 0
 Z패배(Oee) = 0
 T연승(Oee) = 0
 Z연승(Oee) = 0
 P연승(Oee) = 0
 A연승(Oee) = 0
 T연(Oee) = "W"
 Z연(Oee) = "W"
 P연(Oee) = "W"
 A연(Oee) = "W"

Next Oee
Tim10.Enabled = True
Tim09.Enabled = False
End Sub

Private Sub Tim10_Timer()
For Oee = 576 To 715
If Oee = 576 Then
 이름(Oee) = "강정우"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "폭스"
 종족(Oee) = 1
 공격력(Oee) = 700
 견제(Oee) = 550
 전략(Oee) = 600
 물량(Oee) = 700
 수비력(Oee) = 550
 정찰(Oee) = 550
 센스(Oee) = 550
 컨트롤(Oee) = 600

ElseIf Oee = 577 Then
 이름(Oee) = "박성균"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "폭스"
 종족(Oee) = 1
 공격력(Oee) = 700
 견제(Oee) = 700
 전략(Oee) = 550
 물량(Oee) = 700
 수비력(Oee) = 800
 정찰(Oee) = 700
 센스(Oee) = 700
 컨트롤(Oee) = 550

ElseIf Oee = 578 Then
 이름(Oee) = "박세정"
 랭크(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "폭스"
 종족(Oee) = 3
 공격력(Oee) = 700
 견제(Oee) = 700
 전략(Oee) = 700
 물량(Oee) = 800
 수비력(Oee) = 700
 정찰(Oee) = 800
 센스(Oee) = 800
 컨트롤(Oee) = 850

ElseIf Oee = 579 Then
 이름(Oee) = "신노열"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "폭스"
 종족(Oee) = 2
 공격력(Oee) = 650
 견제(Oee) = 750
 전략(Oee) = 550
 물량(Oee) = 750
 수비력(Oee) = 650
 정찰(Oee) = 700
 센스(Oee) = 800
 컨트롤(Oee) = 750

ElseIf Oee = 580 Then
 이름(Oee) = "이영호1"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "폭스"
 종족(Oee) = 3
 공격력(Oee) = 700
 견제(Oee) = 600
 전략(Oee) = 650
 물량(Oee) = 650
 수비력(Oee) = 550
 정찰(Oee) = 550
 센스(Oee) = 650
 컨트롤(Oee) = 650

ElseIf Oee = 581 Then
 이름(Oee) = "이영한"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "폭스"
 종족(Oee) = 2
 공격력(Oee) = 850
 견제(Oee) = 750
 전략(Oee) = 650
 물량(Oee) = 700
 수비력(Oee) = 550
 정찰(Oee) = 600
 센스(Oee) = 650
 컨트롤(Oee) = 750

ElseIf Oee = 582 Then
 이름(Oee) = "이예훈"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "폭스"
 종족(Oee) = 2
 공격력(Oee) = 700
 견제(Oee) = 600
 전략(Oee) = 500
 물량(Oee) = 650
 수비력(Oee) = 600
 정찰(Oee) = 550
 센스(Oee) = 550
 컨트롤(Oee) = 550

ElseIf Oee = 583 Then
 이름(Oee) = "이윤열"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "폭스"
 종족(Oee) = 1
 공격력(Oee) = 650
 견제(Oee) = 600
 전략(Oee) = 700
 물량(Oee) = 700
 수비력(Oee) = 750
 정찰(Oee) = 700
 센스(Oee) = 650
 컨트롤(Oee) = 700

ElseIf Oee = 584 Then
 이름(Oee) = "전상욱"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "폭스"
 종족(Oee) = 1
 공격력(Oee) = 750
 견제(Oee) = 700
 전략(Oee) = 650
 물량(Oee) = 700
 수비력(Oee) = 750
 정찰(Oee) = 700
 센스(Oee) = 700
 컨트롤(Oee) = 650

ElseIf Oee = 585 Then
 이름(Oee) = "전태양"
 랭크(Oee) = "Elite"
 OYear(Oee) = "<10>"
 Team(Oee) = "폭스"
 종족(Oee) = 1
 공격력(Oee) = 850
 견제(Oee) = 850
 전략(Oee) = 750
 물량(Oee) = 850
 수비력(Oee) = 900
 정찰(Oee) = 750
 센스(Oee) = 800
 컨트롤(Oee) = 850

ElseIf Oee = 586 Then
 이름(Oee) = "김경모"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "공군"
 종족(Oee) = 2
 공격력(Oee) = 650
 견제(Oee) = 650
 전략(Oee) = 500
 물량(Oee) = 600
 수비력(Oee) = 600
 정찰(Oee) = 650
 센스(Oee) = 550
 컨트롤(Oee) = 650

ElseIf Oee = 587 Then
 이름(Oee) = "민찬기"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "공군"
 종족(Oee) = 1
 공격력(Oee) = 750
 견제(Oee) = 750
 전략(Oee) = 650
 물량(Oee) = 650
 수비력(Oee) = 650
 정찰(Oee) = 600
 센스(Oee) = 600
 컨트롤(Oee) = 800

ElseIf Oee = 588 Then
 이름(Oee) = "박영민"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "공군"
 종족(Oee) = 3
 공격력(Oee) = 700
 견제(Oee) = 600
 전략(Oee) = 700
 물량(Oee) = 800
 수비력(Oee) = 600
 정찰(Oee) = 700
 센스(Oee) = 600
 컨트롤(Oee) = 700

ElseIf Oee = 589 Then
 이름(Oee) = "박정석"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "공군"
 종족(Oee) = 3
 공격력(Oee) = 750
 견제(Oee) = 650
 전략(Oee) = 650
 물량(Oee) = 700
 수비력(Oee) = 500
 정찰(Oee) = 600
 센스(Oee) = 550
 컨트롤(Oee) = 750

ElseIf Oee = 590 Then
 이름(Oee) = "박태민"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "공군"
 종족(Oee) = 2
 공격력(Oee) = 600
 견제(Oee) = 600
 전략(Oee) = 600
 물량(Oee) = 750
 수비력(Oee) = 700
 정찰(Oee) = 700
 센스(Oee) = 650
 컨트롤(Oee) = 700

ElseIf Oee = 591 Then
 이름(Oee) = "서지훈"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "공군"
 종족(Oee) = 1
 공격력(Oee) = 650
 견제(Oee) = 600
 전략(Oee) = 600
 물량(Oee) = 650
 수비력(Oee) = 600
 정찰(Oee) = 650
 센스(Oee) = 650
 컨트롤(Oee) = 650

ElseIf Oee = 592 Then
 이름(Oee) = "손석희"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "공군"
 종족(Oee) = 3
 공격력(Oee) = 650
 견제(Oee) = 500
 전략(Oee) = 600
 물량(Oee) = 600
 수비력(Oee) = 550
 정찰(Oee) = 500
 센스(Oee) = 550
 컨트롤(Oee) = 650

ElseIf Oee = 593 Then
 이름(Oee) = "안기효"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "공군"
 종족(Oee) = 3
 공격력(Oee) = 650
 견제(Oee) = 650
 전략(Oee) = 650
 물량(Oee) = 700
 수비력(Oee) = 600
 정찰(Oee) = 600
 센스(Oee) = 600
 컨트롤(Oee) = 650

ElseIf Oee = 594 Then
 이름(Oee) = "오영종"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "공군"
 종족(Oee) = 3
 공격력(Oee) = 700
 견제(Oee) = 700
 전략(Oee) = 650
 물량(Oee) = 800
 수비력(Oee) = 650
 정찰(Oee) = 600
 센스(Oee) = 650
 컨트롤(Oee) = 700

ElseIf Oee = 595 Then
 이름(Oee) = "차재욱"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "공군"
 종족(Oee) = 1
 공격력(Oee) = 600
 견제(Oee) = 500
 전략(Oee) = 550
 물량(Oee) = 550
 수비력(Oee) = 650
 정찰(Oee) = 600
 센스(Oee) = 600
 컨트롤(Oee) = 550

ElseIf Oee = 596 Then
 이름(Oee) = "한동욱"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "공군"
 종족(Oee) = 1
 공격력(Oee) = 800
 견제(Oee) = 700
 전략(Oee) = 600
 물량(Oee) = 600
 수비력(Oee) = 450
 정찰(Oee) = 500
 센스(Oee) = 600
 컨트롤(Oee) = 800

ElseIf Oee = 597 Then
 이름(Oee) = "홍진호"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "공군"
 종족(Oee) = 2
 공격력(Oee) = 850
 견제(Oee) = 750
 전략(Oee) = 750
 물량(Oee) = 500
 수비력(Oee) = 500
 정찰(Oee) = 500
 센스(Oee) = 600
 컨트롤(Oee) = 650

ElseIf Oee = 598 Then
 이름(Oee) = "김도우"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 종족(Oee) = 1
 공격력(Oee) = 700
 견제(Oee) = 600
 전략(Oee) = 600
 물량(Oee) = 650
 수비력(Oee) = 750
 정찰(Oee) = 600
 센스(Oee) = 650
 컨트롤(Oee) = 700

ElseIf Oee = 599 Then
 이름(Oee) = "김성대"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 종족(Oee) = 2
 공격력(Oee) = 650
 견제(Oee) = 700
 전략(Oee) = 550
 물량(Oee) = 800
 수비력(Oee) = 800
 정찰(Oee) = 700
 센스(Oee) = 650
 컨트롤(Oee) = 700

ElseIf Oee = 600 Then
 이름(Oee) = "박상우"
 랭크(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 종족(Oee) = 1
 공격력(Oee) = 800
 견제(Oee) = 600
 전략(Oee) = 650
 물량(Oee) = 950
 수비력(Oee) = 850
 정찰(Oee) = 750
 센스(Oee) = 800
 컨트롤(Oee) = 600

ElseIf Oee = 601 Then
 이름(Oee) = "신대근"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 종족(Oee) = 2
 공격력(Oee) = 750
 견제(Oee) = 650
 전략(Oee) = 650
 물량(Oee) = 700
 수비력(Oee) = 700
 정찰(Oee) = 700
 센스(Oee) = 650
 컨트롤(Oee) = 550

ElseIf Oee = 602 Then
 이름(Oee) = "신상호"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 종족(Oee) = 3
 공격력(Oee) = 800
 견제(Oee) = 650
 전략(Oee) = 650
 물량(Oee) = 700
 수비력(Oee) = 600
 정찰(Oee) = 550
 센스(Oee) = 550
 컨트롤(Oee) = 550

ElseIf Oee = 603 Then
 이름(Oee) = "신재욱"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 종족(Oee) = 3
 공격력(Oee) = 650
 견제(Oee) = 650
 전략(Oee) = 750
 물량(Oee) = 650
 수비력(Oee) = 800
 정찰(Oee) = 700
 센스(Oee) = 650
 컨트롤(Oee) = 750

ElseIf Oee = 604 Then
 이름(Oee) = "안수형"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 종족(Oee) = 3
 공격력(Oee) = 650
 견제(Oee) = 500
 전략(Oee) = 500
 물량(Oee) = 650
 수비력(Oee) = 500
 정찰(Oee) = 500
 센스(Oee) = 500
 컨트롤(Oee) = 500

ElseIf Oee = 605 Then
 이름(Oee) = "유병준"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 종족(Oee) = 3
 공격력(Oee) = 550
 견제(Oee) = 550
 전략(Oee) = 500
 물량(Oee) = 600
 수비력(Oee) = 600
 정찰(Oee) = 500
 센스(Oee) = 500
 컨트롤(Oee) = 550

ElseIf Oee = 606 Then
 이름(Oee) = "정명호"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 종족(Oee) = 2
 공격력(Oee) = 700
 견제(Oee) = 650
 전략(Oee) = 550
 물량(Oee) = 650
 수비력(Oee) = 600
 정찰(Oee) = 550
 센스(Oee) = 600
 컨트롤(Oee) = 650

ElseIf Oee = 607 Then
 이름(Oee) = "최지성"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 종족(Oee) = 1
 공격력(Oee) = 700
 견제(Oee) = 600
 전략(Oee) = 600
 물량(Oee) = 600
 수비력(Oee) = 550
 정찰(Oee) = 500
 센스(Oee) = 500
 컨트롤(Oee) = 650

ElseIf Oee = 608 Then
 이름(Oee) = "강석"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "하이트"
 종족(Oee) = 2
 공격력(Oee) = 650
 견제(Oee) = 550
 전략(Oee) = 500
 물량(Oee) = 600
 수비력(Oee) = 600
 정찰(Oee) = 500
 센스(Oee) = 550
 컨트롤(Oee) = 600

ElseIf Oee = 609 Then
 이름(Oee) = "김봉준"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "하이트"
 종족(Oee) = 3
 공격력(Oee) = 550
 견제(Oee) = 600
 전략(Oee) = 600
 물량(Oee) = 750
 수비력(Oee) = 500
 정찰(Oee) = 500
 센스(Oee) = 550
 컨트롤(Oee) = 600

ElseIf Oee = 610 Then
 이름(Oee) = "김상욱"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "하이트"
 종족(Oee) = 2
 공격력(Oee) = 700
 견제(Oee) = 650
 전략(Oee) = 600
 물량(Oee) = 800
 수비력(Oee) = 700
 정찰(Oee) = 650
 센스(Oee) = 600
 컨트롤(Oee) = 650

ElseIf Oee = 611 Then
 이름(Oee) = "김학수"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "하이트"
 종족(Oee) = 3
 공격력(Oee) = 700
 견제(Oee) = 550
 전략(Oee) = 650
 물량(Oee) = 650
 수비력(Oee) = 550
 정찰(Oee) = 550
 센스(Oee) = 550
 컨트롤(Oee) = 550

ElseIf Oee = 612 Then
 이름(Oee) = "신상문"
 랭크(Oee) = "Special"
 OYear(Oee) = "<10>"
 Team(Oee) = "하이트"
 종족(Oee) = 1
 공격력(Oee) = 700
 견제(Oee) = 700
 전략(Oee) = 750
 물량(Oee) = 750
 수비력(Oee) = 700
 정찰(Oee) = 700
 센스(Oee) = 700
 컨트롤(Oee) = 700

ElseIf Oee = 613 Then
 이름(Oee) = "안준영"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "하이트"
 종족(Oee) = 2
 공격력(Oee) = 550
 견제(Oee) = 550
 전략(Oee) = 550
 물량(Oee) = 650
 수비력(Oee) = 550
 정찰(Oee) = 550
 센스(Oee) = 500
 컨트롤(Oee) = 650

ElseIf Oee = 614 Then
 이름(Oee) = "이경민"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "하이트"
 종족(Oee) = 3
 공격력(Oee) = 700
 견제(Oee) = 650
 전략(Oee) = 750
 물량(Oee) = 850
 수비력(Oee) = 600
 정찰(Oee) = 600
 센스(Oee) = 700
 컨트롤(Oee) = 750

ElseIf Oee = 615 Then
 이름(Oee) = "이호준"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "하이트"
 종족(Oee) = 1
 공격력(Oee) = 600
 견제(Oee) = 550
 전략(Oee) = 600
 물량(Oee) = 750
 수비력(Oee) = 750
 정찰(Oee) = 750
 센스(Oee) = 600
 컨트롤(Oee) = 550

ElseIf Oee = 616 Then
 이름(Oee) = "조재걸"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "하이트"
 종족(Oee) = 3
 공격력(Oee) = 700
 견제(Oee) = 650
 전략(Oee) = 550
 물량(Oee) = 650
 수비력(Oee) = 550
 정찰(Oee) = 550
 센스(Oee) = 650
 컨트롤(Oee) = 650

ElseIf Oee = 617 Then
 이름(Oee) = "최홍희"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "하이트"
 종족(Oee) = 3
 공격력(Oee) = 550
 견제(Oee) = 550
 전략(Oee) = 500
 물량(Oee) = 600
 수비력(Oee) = 550
 정찰(Oee) = 500
 센스(Oee) = 500
 컨트롤(Oee) = 500

ElseIf Oee = 618 Then
 이름(Oee) = "하태준"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "하이트"
 종족(Oee) = 3
 공격력(Oee) = 650
 견제(Oee) = 550
 전략(Oee) = 600
 물량(Oee) = 650
 수비력(Oee) = 600
 정찰(Oee) = 600
 센스(Oee) = 550
 컨트롤(Oee) = 550

ElseIf Oee = 619 Then
 이름(Oee) = "권수현"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 종족(Oee) = 2
 공격력(Oee) = 600
 견제(Oee) = 600
 전략(Oee) = 550
 물량(Oee) = 600
 수비력(Oee) = 700
 정찰(Oee) = 600
 센스(Oee) = 550
 컨트롤(Oee) = 600

ElseIf Oee = 620 Then
 이름(Oee) = "김민호"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 종족(Oee) = 2
 공격력(Oee) = 700
 견제(Oee) = 650
 전략(Oee) = 500
 물량(Oee) = 650
 수비력(Oee) = 500
 정찰(Oee) = 500
 센스(Oee) = 600
 컨트롤(Oee) = 650

ElseIf Oee = 621 Then
 이름(Oee) = "김정우"
 랭크(Oee) = "Unique"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 종족(Oee) = 2
 공격력(Oee) = 850
 견제(Oee) = 800
 전략(Oee) = 750
 물량(Oee) = 850
 수비력(Oee) = 800
 정찰(Oee) = 800
 센스(Oee) = 750
 컨트롤(Oee) = 850

ElseIf Oee = 622 Then
 이름(Oee) = "변형태"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 종족(Oee) = 1
 공격력(Oee) = 850
 견제(Oee) = 700
 전략(Oee) = 650
 물량(Oee) = 700
 수비력(Oee) = 600
 정찰(Oee) = 650
 센스(Oee) = 700
 컨트롤(Oee) = 650

ElseIf Oee = 623 Then
 이름(Oee) = "손재범"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 종족(Oee) = 3
 공격력(Oee) = 650
 견제(Oee) = 650
 전략(Oee) = 550
 물량(Oee) = 700
 수비력(Oee) = 650
 정찰(Oee) = 550
 센스(Oee) = 600
 컨트롤(Oee) = 500

ElseIf Oee = 624 Then
 이름(Oee) = "신동원"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 종족(Oee) = 2
 공격력(Oee) = 750
 견제(Oee) = 600
 전략(Oee) = 600
 물량(Oee) = 700
 수비력(Oee) = 700
 정찰(Oee) = 650
 센스(Oee) = 650
 컨트롤(Oee) = 750

ElseIf Oee = 625 Then
 이름(Oee) = "조병세"
 랭크(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 종족(Oee) = 1
 공격력(Oee) = 950
 견제(Oee) = 750
 전략(Oee) = 700
 물량(Oee) = 750
 수비력(Oee) = 650
 정찰(Oee) = 600
 센스(Oee) = 950
 컨트롤(Oee) = 800

ElseIf Oee = 626 Then
 이름(Oee) = "장윤철"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 종족(Oee) = 3
 공격력(Oee) = 750
 견제(Oee) = 600
 전략(Oee) = 650
 물량(Oee) = 800
 수비력(Oee) = 650
 정찰(Oee) = 600
 센스(Oee) = 700
 컨트롤(Oee) = 850

ElseIf Oee = 627 Then
 이름(Oee) = "정우용"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 종족(Oee) = 1
 공격력(Oee) = 650
 견제(Oee) = 550
 전략(Oee) = 550
 물량(Oee) = 650
 수비력(Oee) = 500
 정찰(Oee) = 600
 센스(Oee) = 650
 컨트롤(Oee) = 600

ElseIf Oee = 628 Then
 이름(Oee) = "진영화"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 종족(Oee) = 3
 공격력(Oee) = 600
 견제(Oee) = 700
 전략(Oee) = 700
 물량(Oee) = 750
 수비력(Oee) = 750
 정찰(Oee) = 750
 센스(Oee) = 650
 컨트롤(Oee) = 700

ElseIf Oee = 629 Then
 이름(Oee) = "한두열"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 종족(Oee) = 2
 공격력(Oee) = 550
 견제(Oee) = 600
 전략(Oee) = 550
 물량(Oee) = 600
 수비력(Oee) = 600
 정찰(Oee) = 600
 센스(Oee) = 550
 컨트롤(Oee) = 600

ElseIf Oee = 630 Then
 이름(Oee) = "강동현"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "화승"
 종족(Oee) = 2
 공격력(Oee) = 650
 견제(Oee) = 600
 전략(Oee) = 450
 물량(Oee) = 550
 수비력(Oee) = 500
 정찰(Oee) = 550
 센스(Oee) = 450
 컨트롤(Oee) = 650

ElseIf Oee = 631 Then
 이름(Oee) = "구성훈"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "화승"
 종족(Oee) = 1
 공격력(Oee) = 650
 견제(Oee) = 600
 전략(Oee) = 650
 물량(Oee) = 650
 수비력(Oee) = 750
 정찰(Oee) = 700
 센스(Oee) = 750
 컨트롤(Oee) = 650

ElseIf Oee = 632 Then
 이름(Oee) = "김민혁"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "화승"
 종족(Oee) = 1
 공격력(Oee) = 600
 견제(Oee) = 500
 전략(Oee) = 500
 물량(Oee) = 550
 수비력(Oee) = 500
 정찰(Oee) = 550
 센스(Oee) = 600
 컨트롤(Oee) = 600

ElseIf Oee = 633 Then
 이름(Oee) = "김태균"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "화승"
 종족(Oee) = 3
 공격력(Oee) = 750
 견제(Oee) = 650
 전략(Oee) = 550
 물량(Oee) = 700
 수비력(Oee) = 550
 정찰(Oee) = 600
 센스(Oee) = 600
 컨트롤(Oee) = 600

ElseIf Oee = 634 Then
 이름(Oee) = "박준오"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "화승"
 종족(Oee) = 2
 공격력(Oee) = 750
 견제(Oee) = 700
 전략(Oee) = 550
 물량(Oee) = 650
 수비력(Oee) = 600
 정찰(Oee) = 700
 센스(Oee) = 600
 컨트롤(Oee) = 750

ElseIf Oee = 635 Then
 이름(Oee) = "방태수"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "화승"
 종족(Oee) = 2
 공격력(Oee) = 600
 견제(Oee) = 500
 전략(Oee) = 500
 물량(Oee) = 500
 수비력(Oee) = 450
 정찰(Oee) = 450
 센스(Oee) = 450
 컨트롤(Oee) = 550

ElseIf Oee = 636 Then
 이름(Oee) = "손주흥"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "화승"
 종족(Oee) = 1
 공격력(Oee) = 700
 견제(Oee) = 650
 전략(Oee) = 700
 물량(Oee) = 700
 수비력(Oee) = 650
 정찰(Oee) = 700
 센스(Oee) = 600
 컨트롤(Oee) = 750

ElseIf Oee = 637 Then
 이름(Oee) = "손찬웅"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "화승"
 종족(Oee) = 3
 공격력(Oee) = 700
 견제(Oee) = 900
 전략(Oee) = 700
 물량(Oee) = 600
 수비력(Oee) = 550
 정찰(Oee) = 600
 센스(Oee) = 600
 컨트롤(Oee) = 750

ElseIf Oee = 638 Then
 이름(Oee) = "이제동"
 랭크(Oee) = "Elite"
 OYear(Oee) = "<10>"
 Team(Oee) = "화승"
 종족(Oee) = 2
 공격력(Oee) = 950
 견제(Oee) = 950
 전략(Oee) = 800
 물량(Oee) = 800
 수비력(Oee) = 750
 정찰(Oee) = 750
 센스(Oee) = 800
 컨트롤(Oee) = 800

ElseIf Oee = 639 Then
 이름(Oee) = "이학주"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "화승"
 종족(Oee) = 1
 공격력(Oee) = 700
 견제(Oee) = 650
 전략(Oee) = 550
 물량(Oee) = 650
 수비력(Oee) = 700
 정찰(Oee) = 650
 센스(Oee) = 600
 컨트롤(Oee) = 650

ElseIf Oee = 640 Then
 이름(Oee) = "임원기"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "화승"
 종족(Oee) = 3
 공격력(Oee) = 650
 견제(Oee) = 650
 전략(Oee) = 700
 물량(Oee) = 650
 수비력(Oee) = 500
 정찰(Oee) = 700
 센스(Oee) = 550
 컨트롤(Oee) = 650

ElseIf Oee = 641 Then
 이름(Oee) = "고강민"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 종족(Oee) = 2
 공격력(Oee) = 600
 견제(Oee) = 650
 전략(Oee) = 650
 물량(Oee) = 700
 수비력(Oee) = 650
 정찰(Oee) = 700
 센스(Oee) = 600
 컨트롤(Oee) = 550

ElseIf Oee = 642 Then
 이름(Oee) = "김대엽"
 랭크(Oee) = "Special"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 종족(Oee) = 3
 공격력(Oee) = 750
 견제(Oee) = 750
 전략(Oee) = 600
 물량(Oee) = 850
 수비력(Oee) = 750
 정찰(Oee) = 600
 센스(Oee) = 650
 컨트롤(Oee) = 650

ElseIf Oee = 643 Then
 이름(Oee) = "김재춘"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 종족(Oee) = 2
 공격력(Oee) = 700
 견제(Oee) = 700
 전략(Oee) = 650
 물량(Oee) = 600
 수비력(Oee) = 600
 정찰(Oee) = 650
 센스(Oee) = 650
 컨트롤(Oee) = 650

ElseIf Oee = 644 Then
 이름(Oee) = "남승현"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 종족(Oee) = 1
 공격력(Oee) = 650
 견제(Oee) = 600
 전략(Oee) = 650
 물량(Oee) = 600
 수비력(Oee) = 600
 정찰(Oee) = 600
 센스(Oee) = 650
 컨트롤(Oee) = 600

ElseIf Oee = 645 Then
 이름(Oee) = "박재영"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 종족(Oee) = 3
 공격력(Oee) = 650
 견제(Oee) = 600
 전략(Oee) = 600
 물량(Oee) = 800
 수비력(Oee) = 700
 정찰(Oee) = 550
 센스(Oee) = 600
 컨트롤(Oee) = 550

ElseIf Oee = 646 Then
 이름(Oee) = "박지수"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 종족(Oee) = 1
 공격력(Oee) = 800
 견제(Oee) = 600
 전략(Oee) = 600
 물량(Oee) = 800
 수비력(Oee) = 800
 정찰(Oee) = 650
 센스(Oee) = 600
 컨트롤(Oee) = 750

ElseIf Oee = 647 Then
 이름(Oee) = "배병우"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 종족(Oee) = 2
 공격력(Oee) = 700
 견제(Oee) = 600
 전략(Oee) = 650
 물량(Oee) = 700
 수비력(Oee) = 550
 정찰(Oee) = 700
 센스(Oee) = 550
 컨트롤(Oee) = 650

ElseIf Oee = 648 Then
 이름(Oee) = "우정호"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 종족(Oee) = 3
 공격력(Oee) = 600
 견제(Oee) = 600
 전략(Oee) = 600
 물량(Oee) = 750
 수비력(Oee) = 700
 정찰(Oee) = 700
 센스(Oee) = 700
 컨트롤(Oee) = 750

ElseIf Oee = 649 Then
 이름(Oee) = "이영호"
 랭크(Oee) = "Legend"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 종족(Oee) = 1
 공격력(Oee) = 850
 견제(Oee) = 800
 전략(Oee) = 800
 물량(Oee) = 950
 수비력(Oee) = 950
 정찰(Oee) = 900
 센스(Oee) = 900
 컨트롤(Oee) = 800
ElseIf Oee = 650 Then
 이름(Oee) = "최용주"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 종족(Oee) = 2
 공격력(Oee) = 550
 견제(Oee) = 600
 전략(Oee) = 500
 물량(Oee) = 550
 수비력(Oee) = 450
 정찰(Oee) = 500
 센스(Oee) = 500
 컨트롤(Oee) = 700

ElseIf Oee = 651 Then
 이름(Oee) = "황병영"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 종족(Oee) = 1
 공격력(Oee) = 650
 견제(Oee) = 550
 전략(Oee) = 500
 물량(Oee) = 750
 수비력(Oee) = 650
 정찰(Oee) = 600
 센스(Oee) = 550
 컨트롤(Oee) = 550

ElseIf Oee = 652 Then
 이름(Oee) = "김기현"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "삼성전자"
 종족(Oee) = 1
 공격력(Oee) = 500
 견제(Oee) = 600
 전략(Oee) = 500
 물량(Oee) = 600
 수비력(Oee) = 650
 정찰(Oee) = 550
 센스(Oee) = 500
 컨트롤(Oee) = 500

ElseIf Oee = 653 Then
 이름(Oee) = "박대호"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "삼성전자"
 종족(Oee) = 1
 공격력(Oee) = 550
 견제(Oee) = 600
 전략(Oee) = 650
 물량(Oee) = 600
 수비력(Oee) = 650
 정찰(Oee) = 550
 센스(Oee) = 550
 컨트롤(Oee) = 600

ElseIf Oee = 654 Then
 이름(Oee) = "송병구"
 랭크(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "삼성전자"
 종족(Oee) = 3
 공격력(Oee) = 750
 견제(Oee) = 600
 전략(Oee) = 600
 물량(Oee) = 900
 수비력(Oee) = 900
 정찰(Oee) = 700
 센스(Oee) = 800
 컨트롤(Oee) = 750

ElseIf Oee = 655 Then
 이름(Oee) = "유준희"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "삼성전자"
 종족(Oee) = 2
 공격력(Oee) = 700
 견제(Oee) = 600
 전략(Oee) = 650
 물량(Oee) = 700
 수비력(Oee) = 600
 정찰(Oee) = 650
 센스(Oee) = 600
 컨트롤(Oee) = 650

ElseIf Oee = 656 Then
 이름(Oee) = "이성은"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "삼성전자"
 종족(Oee) = 1
 공격력(Oee) = 750
 견제(Oee) = 700
 전략(Oee) = 800
 물량(Oee) = 600
 수비력(Oee) = 450
 정찰(Oee) = 600
 센스(Oee) = 700
 컨트롤(Oee) = 750

ElseIf Oee = 657 Then
 이름(Oee) = "이정현"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "삼성전자"
 종족(Oee) = 2
 공격력(Oee) = 600
 견제(Oee) = 550
 전략(Oee) = 600
 물량(Oee) = 650
 수비력(Oee) = 650
 정찰(Oee) = 650
 센스(Oee) = 550
 컨트롤(Oee) = 550

ElseIf Oee = 658 Then
 이름(Oee) = "임태규"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "삼성전자"
 종족(Oee) = 3
 공격력(Oee) = 650
 견제(Oee) = 700
 전략(Oee) = 500
 물량(Oee) = 750
 수비력(Oee) = 700
 정찰(Oee) = 600
 센스(Oee) = 500
 컨트롤(Oee) = 600

ElseIf Oee = 659 Then
 이름(Oee) = "조기석"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "삼성전자"
 종족(Oee) = 1
 공격력(Oee) = 500
 견제(Oee) = 550
 전략(Oee) = 500
 물량(Oee) = 650
 수비력(Oee) = 600
 정찰(Oee) = 650
 센스(Oee) = 500
 컨트롤(Oee) = 550

ElseIf Oee = 660 Then
 이름(Oee) = "주영달"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "삼성전자"
 종족(Oee) = 2
 공격력(Oee) = 650
 견제(Oee) = 600
 전략(Oee) = 600
 물량(Oee) = 650
 수비력(Oee) = 550
 정찰(Oee) = 600
 센스(Oee) = 650
 컨트롤(Oee) = 700

ElseIf Oee = 661 Then
 이름(Oee) = "차명환"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "삼성전자"
 종족(Oee) = 2
 공격력(Oee) = 700
 견제(Oee) = 600
 전략(Oee) = 650
 물량(Oee) = 700
 수비력(Oee) = 750
 정찰(Oee) = 700
 센스(Oee) = 700
 컨트롤(Oee) = 700

ElseIf Oee = 662 Then
 이름(Oee) = "최윤선"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "삼성전자"
 종족(Oee) = 3
 공격력(Oee) = 650
 견제(Oee) = 550
 전략(Oee) = 550
 물량(Oee) = 650
 수비력(Oee) = 500
 정찰(Oee) = 550
 센스(Oee) = 500
 컨트롤(Oee) = 700

ElseIf Oee = 663 Then
 이름(Oee) = "허영무"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "삼성전자"
 종족(Oee) = 3
 공격력(Oee) = 700
 견제(Oee) = 700
 전략(Oee) = 650
 물량(Oee) = 750
 수비력(Oee) = 550
 정찰(Oee) = 650
 센스(Oee) = 600
 컨트롤(Oee) = 750

ElseIf Oee = 664 Then
 이름(Oee) = "김경효"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 종족(Oee) = 1
 공격력(Oee) = 650
 견제(Oee) = 700
 전략(Oee) = 550
 물량(Oee) = 600
 수비력(Oee) = 650
 정찰(Oee) = 550
 센스(Oee) = 650
 컨트롤(Oee) = 650

ElseIf Oee = 665 Then
 이름(Oee) = "김동건"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 종족(Oee) = 1
 공격력(Oee) = 700
 견제(Oee) = 700
 전략(Oee) = 650
 물량(Oee) = 700
 수비력(Oee) = 700
 정찰(Oee) = 700
 센스(Oee) = 650
 컨트롤(Oee) = 600

ElseIf Oee = 666 Then
 이름(Oee) = "김성현"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 종족(Oee) = 1
 공격력(Oee) = 550
 견제(Oee) = 600
 전략(Oee) = 550
 물량(Oee) = 700
 수비력(Oee) = 700
 정찰(Oee) = 600
 센스(Oee) = 600
 컨트롤(Oee) = 650

ElseIf Oee = 667 Then
 이름(Oee) = "김구현"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 종족(Oee) = 3
 공격력(Oee) = 750
 견제(Oee) = 700
 전략(Oee) = 700
 물량(Oee) = 750
 수비력(Oee) = 650
 정찰(Oee) = 700
 센스(Oee) = 600
 컨트롤(Oee) = 600

ElseIf Oee = 668 Then
 이름(Oee) = "김윤중"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 종족(Oee) = 3
 공격력(Oee) = 650
 견제(Oee) = 700
 전략(Oee) = 650
 물량(Oee) = 900
 수비력(Oee) = 550
 정찰(Oee) = 650
 센스(Oee) = 700
 컨트롤(Oee) = 650

ElseIf Oee = 669 Then
 이름(Oee) = "김윤환1"
 랭크(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 종족(Oee) = 2
 공격력(Oee) = 700
 견제(Oee) = 750
 전략(Oee) = 850
 물량(Oee) = 750
 수비력(Oee) = 700
 정찰(Oee) = 800
 센스(Oee) = 800
 컨트롤(Oee) = 800

ElseIf Oee = 670 Then
 이름(Oee) = "김현우"
 랭크(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 종족(Oee) = 2
 공격력(Oee) = 950
 견제(Oee) = 950
 전략(Oee) = 500
 물량(Oee) = 600
 수비력(Oee) = 550
 정찰(Oee) = 550
 센스(Oee) = 950
 컨트롤(Oee) = 950

ElseIf Oee = 671 Then
 이름(Oee) = "박성준"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 종족(Oee) = 2
 공격력(Oee) = 800
 견제(Oee) = 700
 전략(Oee) = 600
 물량(Oee) = 750
 수비력(Oee) = 600
 정찰(Oee) = 600
 센스(Oee) = 650
 컨트롤(Oee) = 700

ElseIf Oee = 672 Then
 이름(Oee) = "박종수"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 종족(Oee) = 3
 공격력(Oee) = 650
 견제(Oee) = 650
 전략(Oee) = 700
 물량(Oee) = 650
 수비력(Oee) = 600
 정찰(Oee) = 650
 센스(Oee) = 600
 컨트롤(Oee) = 650

ElseIf Oee = 673 Then
 이름(Oee) = "서지수"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 종족(Oee) = 1
 공격력(Oee) = 600
 견제(Oee) = 500
 전략(Oee) = 500
 물량(Oee) = 450
 수비력(Oee) = 450
 정찰(Oee) = 500
 센스(Oee) = 450
 컨트롤(Oee) = 650

ElseIf Oee = 674 Then
 이름(Oee) = "이신형"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 종족(Oee) = 1
 공격력(Oee) = 700
 견제(Oee) = 500
 전략(Oee) = 600
 물량(Oee) = 700
 수비력(Oee) = 700
 정찰(Oee) = 650
 센스(Oee) = 650
 컨트롤(Oee) = 600

ElseIf Oee = 675 Then
 이름(Oee) = "조성호"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 종족(Oee) = 3
 공격력(Oee) = 650
 견제(Oee) = 550
 전략(Oee) = 500
 물량(Oee) = 600
 수비력(Oee) = 500
 정찰(Oee) = 550
 센스(Oee) = 550
 컨트롤(Oee) = 500

ElseIf Oee = 676 Then
 이름(Oee) = "조일장"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 종족(Oee) = 2
 공격력(Oee) = 700
 견제(Oee) = 750
 전략(Oee) = 550
 물량(Oee) = 750
 수비력(Oee) = 650
 정찰(Oee) = 650
 센스(Oee) = 600
 컨트롤(Oee) = 700

ElseIf Oee = 677 Then
 이름(Oee) = "김동주"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "웅진"
 종족(Oee) = 1
 공격력(Oee) = 700
 견제(Oee) = 700
 전략(Oee) = 600
 물량(Oee) = 650
 수비력(Oee) = 600
 정찰(Oee) = 550
 센스(Oee) = 550
 컨트롤(Oee) = 650

ElseIf Oee = 678 Then
 이름(Oee) = "김명운"
 랭크(Oee) = "Special"
 OYear(Oee) = "<10>"
 Team(Oee) = "웅진"
 종족(Oee) = 2
 공격력(Oee) = 650
 견제(Oee) = 700
 전략(Oee) = 700
 물량(Oee) = 750
 수비력(Oee) = 750
 정찰(Oee) = 750
 센스(Oee) = 700
 컨트롤(Oee) = 800

ElseIf Oee = 679 Then
 이름(Oee) = "김민철"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "웅진"
 종족(Oee) = 2
 공격력(Oee) = 700
 견제(Oee) = 600
 전략(Oee) = 550
 물량(Oee) = 700
 수비력(Oee) = 700
 정찰(Oee) = 650
 센스(Oee) = 550
 컨트롤(Oee) = 700

ElseIf Oee = 680 Then
 이름(Oee) = "김승현"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "웅진"
 종족(Oee) = 3
 공격력(Oee) = 750
 견제(Oee) = 700
 전략(Oee) = 700
 물량(Oee) = 800
 수비력(Oee) = 650
 정찰(Oee) = 650
 센스(Oee) = 650
 컨트롤(Oee) = 700

ElseIf Oee = 681 Then
 이름(Oee) = "김영진"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "웅진"
 종족(Oee) = 1
 공격력(Oee) = 750
 견제(Oee) = 650
 전략(Oee) = 650
 물량(Oee) = 600
 수비력(Oee) = 550
 정찰(Oee) = 550
 센스(Oee) = 600
 컨트롤(Oee) = 600

ElseIf Oee = 682 Then
 이름(Oee) = "노준규"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "웅진"
 종족(Oee) = 1
 공격력(Oee) = 700
 견제(Oee) = 650
 전략(Oee) = 550
 물량(Oee) = 650
 수비력(Oee) = 500
 정찰(Oee) = 550
 센스(Oee) = 550
 컨트롤(Oee) = 600

ElseIf Oee = 683 Then
 이름(Oee) = "박대만"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "웅진"
 종족(Oee) = 3
 공격력(Oee) = 700
 견제(Oee) = 650
 전략(Oee) = 600
 물량(Oee) = 700
 수비력(Oee) = 500
 정찰(Oee) = 600
 센스(Oee) = 600
 컨트롤(Oee) = 700

ElseIf Oee = 684 Then
 이름(Oee) = "윤용태"
 랭크(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "웅진"
 종족(Oee) = 3
 공격력(Oee) = 850
 견제(Oee) = 750
 전략(Oee) = 700
 물량(Oee) = 800
 수비력(Oee) = 700
 정찰(Oee) = 650
 센스(Oee) = 700
 컨트롤(Oee) = 950

ElseIf Oee = 685 Then
 이름(Oee) = "이동준"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "웅진"
 종족(Oee) = 1
 공격력(Oee) = 650
 견제(Oee) = 600
 전략(Oee) = 650
 물량(Oee) = 600
 수비력(Oee) = 700
 정찰(Oee) = 550
 센스(Oee) = 550
 컨트롤(Oee) = 550

ElseIf Oee = 686 Then
 이름(Oee) = "임정현"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "웅진"
 종족(Oee) = 2
 공격력(Oee) = 750
 견제(Oee) = 550
 전략(Oee) = 500
 물량(Oee) = 650
 수비력(Oee) = 450
 정찰(Oee) = 600
 센스(Oee) = 600
 컨트롤(Oee) = 700

ElseIf Oee = 687 Then
 이름(Oee) = "임진묵"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "웅진"
 종족(Oee) = 1
 공격력(Oee) = 700
 견제(Oee) = 650
 전략(Oee) = 650
 물량(Oee) = 650
 수비력(Oee) = 550
 정찰(Oee) = 700
 센스(Oee) = 600
 컨트롤(Oee) = 800

ElseIf Oee = 688 Then
 이름(Oee) = "정종현"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "웅진"
 종족(Oee) = 1
 공격력(Oee) = 600
 견제(Oee) = 600
 전략(Oee) = 600
 물량(Oee) = 750
 수비력(Oee) = 750
 정찰(Oee) = 700
 센스(Oee) = 700
 컨트롤(Oee) = 600

ElseIf Oee = 689 Then
 이름(Oee) = "한상봉"
 랭크(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "웅진"
 종족(Oee) = 2
 공격력(Oee) = 950
 견제(Oee) = 950
 전략(Oee) = 700
 물량(Oee) = 600
 수비력(Oee) = 550
 정찰(Oee) = 700
 센스(Oee) = 950
 컨트롤(Oee) = 950

ElseIf Oee = 690 Then
 이름(Oee) = "고인규"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 종족(Oee) = 1
 공격력(Oee) = 700
 견제(Oee) = 700
 전략(Oee) = 650
 물량(Oee) = 700
 수비력(Oee) = 600
 정찰(Oee) = 700
 센스(Oee) = 650
 컨트롤(Oee) = 700

ElseIf Oee = 691 Then
 이름(Oee) = "김택용"
 랭크(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 종족(Oee) = 3
 공격력(Oee) = 600
 견제(Oee) = 950
 전략(Oee) = 700
 물량(Oee) = 700
 수비력(Oee) = 700
 정찰(Oee) = 900
 센스(Oee) = 800
 컨트롤(Oee) = 750

ElseIf Oee = 692 Then
 이름(Oee) = "도재욱"
 랭크(Oee) = "Special"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 종족(Oee) = 3
 공격력(Oee) = 800
 견제(Oee) = 750
 전략(Oee) = 700
 물량(Oee) = 850
 수비력(Oee) = 550
 정찰(Oee) = 550
 센스(Oee) = 800
 컨트롤(Oee) = 600

ElseIf Oee = 693 Then
 이름(Oee) = "박재혁"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 종족(Oee) = 2
 공격력(Oee) = 800
 견제(Oee) = 700
 전략(Oee) = 550
 물량(Oee) = 700
 수비력(Oee) = 600
 정찰(Oee) = 600
 센스(Oee) = 750
 컨트롤(Oee) = 850

ElseIf Oee = 694 Then
 이름(Oee) = "어윤수"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 종족(Oee) = 2
 공격력(Oee) = 750
 견제(Oee) = 600
 전략(Oee) = 600
 물량(Oee) = 650
 수비력(Oee) = 600
 정찰(Oee) = 550
 센스(Oee) = 500
 컨트롤(Oee) = 700

ElseIf Oee = 695 Then
 이름(Oee) = "이승석"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 종족(Oee) = 2
 공격력(Oee) = 650
 견제(Oee) = 600
 전략(Oee) = 550
 물량(Oee) = 750
 수비력(Oee) = 700
 정찰(Oee) = 650
 센스(Oee) = 600
 컨트롤(Oee) = 650

ElseIf Oee = 696 Then
 이름(Oee) = "임요환"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 종족(Oee) = 1
 공격력(Oee) = 600
 견제(Oee) = 650
 전략(Oee) = 800
 물량(Oee) = 600
 수비력(Oee) = 450
 정찰(Oee) = 550
 센스(Oee) = 650
 컨트롤(Oee) = 700

ElseIf Oee = 697 Then
 이름(Oee) = "정명훈"
 랭크(Oee) = "Elite"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 종족(Oee) = 1
 공격력(Oee) = 700
 견제(Oee) = 950
 전략(Oee) = 800
 물량(Oee) = 950
 수비력(Oee) = 950
 정찰(Oee) = 800
 센스(Oee) = 800
 컨트롤(Oee) = 650

ElseIf Oee = 698 Then
 이름(Oee) = "정영재"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 종족(Oee) = 1
 공격력(Oee) = 600
 견제(Oee) = 600
 전략(Oee) = 550
 물량(Oee) = 600
 수비력(Oee) = 600
 정찰(Oee) = 550
 센스(Oee) = 500
 컨트롤(Oee) = 600

ElseIf Oee = 697 Then
 이름(Oee) = "정영철"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 종족(Oee) = 2
 공격력(Oee) = 650
 견제(Oee) = 650
 전략(Oee) = 650
 물량(Oee) = 700
 수비력(Oee) = 550
 정찰(Oee) = 600
 센스(Oee) = 550
 컨트롤(Oee) = 550

ElseIf Oee = 698 Then
 이름(Oee) = "정윤종"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 종족(Oee) = 3
 공격력(Oee) = 500
 견제(Oee) = 550
 전략(Oee) = 600
 물량(Oee) = 650
 수비력(Oee) = 550
 정찰(Oee) = 500
 센스(Oee) = 500
 컨트롤(Oee) = 550

ElseIf Oee = 699 Then
 이름(Oee) = "최호선"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 종족(Oee) = 1
 공격력(Oee) = 600
 견제(Oee) = 600
 전략(Oee) = 500
 물량(Oee) = 650
 수비력(Oee) = 650
 정찰(Oee) = 500
 센스(Oee) = 650
 컨트롤(Oee) = 600

ElseIf Oee = 700 Then
 이름(Oee) = "고석현"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 종족(Oee) = 2
 공격력(Oee) = 800
 견제(Oee) = 750
 전략(Oee) = 500
 물량(Oee) = 650
 수비력(Oee) = 550
 정찰(Oee) = 600
 센스(Oee) = 700
 컨트롤(Oee) = 700

ElseIf Oee = 701 Then
 이름(Oee) = "김동현"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 종족(Oee) = 2
 공격력(Oee) = 700
 견제(Oee) = 650
 전략(Oee) = 650
 물량(Oee) = 700
 수비력(Oee) = 700
 정찰(Oee) = 650
 센스(Oee) = 700
 컨트롤(Oee) = 600

ElseIf Oee = 702 Then
 이름(Oee) = "김재훈"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 종족(Oee) = 3
 공격력(Oee) = 700
 견제(Oee) = 650
 전략(Oee) = 600
 물량(Oee) = 750
 수비력(Oee) = 600
 정찰(Oee) = 600
 센스(Oee) = 650
 컨트롤(Oee) = 650

ElseIf Oee = 703 Then
 이름(Oee) = "김태훈"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 종족(Oee) = 2
 공격력(Oee) = 650
 견제(Oee) = 550
 전략(Oee) = 550
 물량(Oee) = 600
 수비력(Oee) = 600
 정찰(Oee) = 600
 센스(Oee) = 550
 컨트롤(Oee) = 700

ElseIf Oee = 704 Then
 이름(Oee) = "박수범"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 종족(Oee) = 3
 공격력(Oee) = 700
 견제(Oee) = 550
 전략(Oee) = 550
 물량(Oee) = 750
 수비력(Oee) = 600
 정찰(Oee) = 600
 센스(Oee) = 650
 컨트롤(Oee) = 600

ElseIf Oee = 705 Then
 이름(Oee) = "박지호"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 종족(Oee) = 3
 공격력(Oee) = 900
 견제(Oee) = 600
 전략(Oee) = 700
 물량(Oee) = 900
 수비력(Oee) = 500
 정찰(Oee) = 550
 센스(Oee) = 550
 컨트롤(Oee) = 550

ElseIf Oee = 706 Then
 이름(Oee) = "서경종"
 랭크(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 종족(Oee) = 2
 공격력(Oee) = 700
 견제(Oee) = 600
 전략(Oee) = 750
 물량(Oee) = 600
 수비력(Oee) = 600
 정찰(Oee) = 550
 센스(Oee) = 600
 컨트롤(Oee) = 650

ElseIf Oee = 707 Then
 이름(Oee) = "염보성"
 랭크(Oee) = "Special"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 종족(Oee) = 1
 공격력(Oee) = 700
 견제(Oee) = 750
 전략(Oee) = 700
 물량(Oee) = 850
 수비력(Oee) = 750
 정찰(Oee) = 750
 센스(Oee) = 700
 컨트롤(Oee) = 700

ElseIf Oee = 708 Then
 이름(Oee) = "이재호"
 랭크(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 종족(Oee) = 1
 공격력(Oee) = 900
 견제(Oee) = 800
 전략(Oee) = 750
 물량(Oee) = 600
 수비력(Oee) = 600
 정찰(Oee) = 800
 센스(Oee) = 750
 컨트롤(Oee) = 800
ElseIf Oee = 709 Then
 이름(Oee) = "Mystery"
 랭크(Oee) = "Unique"
 OYear(Oee) = "<11>"
 Team(Oee) = "Mystar"
 종족(Oee) = 2
 공격력(Oee) = 900
 수비력(Oee) = 750
 정찰(Oee) = 850
 물량(Oee) = 900
 전략(Oee) = 800
 컨트롤(Oee) = 750
 견제(Oee) = 600
 센스(Oee) = 850
ElseIf Oee = 710 Then
 이름(Oee) = "오늘은"
 랭크(Oee) = "Rare"
 OYear(Oee) = "<11>"
 Team(Oee) = "Mystar"
 종족(Oee) = 3
 공격력(Oee) = 750
 수비력(Oee) = 850
 정찰(Oee) = 650
 물량(Oee) = 750
 전략(Oee) = 700
 컨트롤(Oee) = 750
 견제(Oee) = 750
 센스(Oee) = 850
ElseIf Oee = 711 Then
 이름(Oee) = "Turtle"
 랭크(Oee) = "Special"
 OYear(Oee) = "<10>"
 Team(Oee) = "Mystar"
 종족(Oee) = 3
 공격력(Oee) = 950
 수비력(Oee) = 700
 정찰(Oee) = 500
 물량(Oee) = 950
 전략(Oee) = 500
 컨트롤(Oee) = 500
 센스(Oee) = 750
 견제(Oee) = 750
ElseIf Oee = 712 Then
 이름(Oee) = "플투군"
 랭크(Oee) = "Special"
 OYear(Oee) = "<11>"
 Team(Oee) = "Mystar"
 종족(Oee) = 3
 공격력(Oee) = 950
 수비력(Oee) = 600
 정찰(Oee) = 600
 물량(Oee) = 950
 전략(Oee) = 600
 컨트롤(Oee) = 600
 센스(Oee) = 750
 견제(Oee) = 600
ElseIf Oee = 713 Then
 이름(Oee) = "은비령"
 랭크(Oee) = "Unique"
 OYear(Oee) = "<11>"
 Team(Oee) = "Mystar"
 종족(Oee) = 1
 공격력(Oee) = 900
 견제(Oee) = 900
 전략(Oee) = 800
 물량(Oee) = 600
 수비력(Oee) = 750
 정찰(Oee) = 950
 센스(Oee) = 800
 컨트롤(Oee) = 700
ElseIf Oee = 714 Then
 이름(Oee) = "이성은[Ex]"
 랭크(Oee) = "Unique"
 OYear(Oee) = "<07>"
 Team(Oee) = "삼성전자"
 종족(Oee) = 1
 공격력(Oee) = 900
 견제(Oee) = 800
 전략(Oee) = 900
 물량(Oee) = 700
 수비력(Oee) = 650
 정찰(Oee) = 750
 센스(Oee) = 850
 컨트롤(Oee) = 850
ElseIf Oee = 715 Then
 이름(Oee) = "강민"
 랭크(Oee) = "Unique"
 OYear(Oee) = "<06>"
 Team(Oee) = "KTF"
 종족(Oee) = 3
 공격력(Oee) = 700
 견제(Oee) = 800
 전략(Oee) = 950
 물량(Oee) = 850
 수비력(Oee) = 850
 정찰(Oee) = 700
 센스(Oee) = 800
 컨트롤(Oee) = 750
End If

 우승(Oee) = 0
 준우승(Oee) = 0
 컨디션(Oee) = 100
 A승리(Oee) = 0
 A패배(Oee) = 0
 P승리(Oee) = 0
 P패배(Oee) = 0
 T승리(Oee) = 0
 T패배(Oee) = 0
 Z승리(Oee) = 0
 Z패배(Oee) = 0
 T연승(Oee) = 0
 Z연승(Oee) = 0
 P연승(Oee) = 0
 A연승(Oee) = 0
 T연(Oee) = "W"
 Z연(Oee) = "W"
 P연(Oee) = "W"
 A연(Oee) = "W"
Next Oee
 
Tim11.Enabled = True
Tim10.Enabled = False
End Sub

Private Sub Tim11_Timer()
For Oee = 1 To 118
If Oee = 1 Then
이름(Oee) = "한지원"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 500
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 550
ElseIf Oee = 2 Then
이름(Oee) = "서지수"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 450
수비력(Oee) = 450
정찰(Oee) = 500
센스(Oee) = 450
컨트롤(Oee) = 650
ElseIf Oee = 3 Then
이름(Oee) = "김성운"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "웅진"
종족(Oee) = 2
공격력(Oee) = 550
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 500
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 650
ElseIf Oee = 4 Then
이름(Oee) = "백동건"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "웅진"
종족(Oee) = 3
공격력(Oee) = 550
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 500
ElseIf Oee = 5 Then
이름(Oee) = "윤지용"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "웅진"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 550
ElseIf Oee = 6 Then
이름(Oee) = "홍진표"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "웅진"
종족(Oee) = 1
공격력(Oee) = 500
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 500
ElseIf Oee = 7 Then
이름(Oee) = "이호성"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 600
수비력(Oee) = 400
정찰(Oee) = 450
센스(Oee) = 500
컨트롤(Oee) = 600
ElseIf Oee = 8 Then
이름(Oee) = "김민규"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 450
물량(Oee) = 550
수비력(Oee) = 450
정찰(Oee) = 450
센스(Oee) = 500
컨트롤(Oee) = 600
ElseIf Oee = 9 Then
이름(Oee) = "김용혁"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBc"
종족(Oee) = 2
공격력(Oee) = 550
견제(Oee) = 500
전략(Oee) = 600
물량(Oee) = 500
수비력(Oee) = 450
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 600
ElseIf Oee = 10 Then
이름(Oee) = "오세기"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
종족(Oee) = 1
공격력(Oee) = 550
견제(Oee) = 550
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 500
ElseIf Oee = 11 Then
이름(Oee) = "정재우"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 500
수비력(Oee) = 550
정찰(Oee) = 500
센스(Oee) = 550
컨트롤(Oee) = 600
ElseIf Oee = 12 Then
이름(Oee) = "하재상"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 500
센스(Oee) = 550
컨트롤(Oee) = 550
ElseIf Oee = 13 Then
이름(Oee) = "백동준"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "화승"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 500
센스(Oee) = 500
컨트롤(Oee) = 650
ElseIf Oee = 14 Then
이름(Oee) = "백승혁"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "화승"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 550
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 600
ElseIf Oee = 15 Then
이름(Oee) = "송영진"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 550
견제(Oee) = 500
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 500
ElseIf Oee = 16 Then
이름(Oee) = "유영진"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
종족(Oee) = 1
공격력(Oee) = 500
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 500
ElseIf Oee = 17 Then
이름(Oee) = "강현우"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 500
물량(Oee) = 750
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 600
ElseIf Oee = 18 Then
이름(Oee) = "조기석"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "삼성전자"
종족(Oee) = 1
공격력(Oee) = 500
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 500
컨트롤(Oee) = 550
ElseIf Oee = 19 Then
이름(Oee) = "조성호"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 500
ElseIf Oee = 20 Then
이름(Oee) = "정영재"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 600
ElseIf Oee = 21 Then
이름(Oee) = "김유진"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "화승"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 500
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 500
센스(Oee) = 650
컨트롤(Oee) = 550
ElseIf Oee = 22 Then
이름(Oee) = "하늘"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "화승"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 550
수비력(Oee) = 600
정찰(Oee) = 500
센스(Oee) = 550
컨트롤(Oee) = 600
ElseIf Oee = 23 Then
이름(Oee) = "남승현"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 600
ElseIf Oee = 24 Then
이름(Oee) = "최용주"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 700
ElseIf Oee = 25 Then
이름(Oee) = "유준희"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 550
ElseIf Oee = 26 Then
이름(Oee) = "주영달"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 700
ElseIf Oee = 27 Then
이름(Oee) = "김성현"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 550
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 650
ElseIf Oee = 28 Then
이름(Oee) = "박종수"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 600
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 600
ElseIf Oee = 29 Then
이름(Oee) = "노준규"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "웅진"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 600
ElseIf Oee = 30 Then
이름(Oee) = "정경두"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 600
ElseIf Oee = 31 Then
이름(Oee) = "방태수"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "화승"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 600
ElseIf Oee = 32 Then
이름(Oee) = "손재범"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 500
ElseIf Oee = 33 Then
이름(Oee) = "정우용"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 500
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 700
ElseIf Oee = 34 Then
이름(Oee) = "한두열"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 650
ElseIf Oee = 35 Then
이름(Oee) = "권수현"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "공군"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 600
ElseIf Oee = 36 Then
이름(Oee) = "이정현"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "공군"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 550
컨트롤(Oee) = 550
ElseIf Oee = 37 Then
이름(Oee) = "강정우"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "폭스"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 550
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 550
컨트롤(Oee) = 600
ElseIf Oee = 38 Then
이름(Oee) = "김준호"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "폭스"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 550
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 650
ElseIf Oee = 39 Then
이름(Oee) = "이예훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "폭스"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 500
컨트롤(Oee) = 550
ElseIf Oee = 40 Then
이름(Oee) = "주성욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "폭스"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 650
ElseIf Oee = 41 Then
이름(Oee) = "고강민"
랭크(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 800
수비력(Oee) = 800
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 600
ElseIf Oee = 42 Then
이름(Oee) = "박재영"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 800
수비력(Oee) = 650
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 550
ElseIf Oee = 43 Then
이름(Oee) = "박정석"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 500
정찰(Oee) = 600
센스(Oee) = 550
컨트롤(Oee) = 750
ElseIf Oee = 44 Then
이름(Oee) = "우정호"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
종족(Oee) = 3
공격력(Oee) = 550
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 650
ElseIf Oee = 45 Then
이름(Oee) = "임정현"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 650
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 750
ElseIf Oee = 46 Then
이름(Oee) = "황병영"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 600
ElseIf Oee = 47 Then
이름(Oee) = "김기현"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "삼성전자"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 550
ElseIf Oee = 48 Then
이름(Oee) = "박대호"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "삼성전자"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 600
ElseIf Oee = 49 Then
이름(Oee) = "임태규"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "삼성전자"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 550
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 650
ElseIf Oee = 50 Then
이름(Oee) = "김도우"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 800
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 650
ElseIf Oee = 51 Then
이름(Oee) = "김동건"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 700
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 700
센스(Oee) = 650
컨트롤(Oee) = 600
ElseIf Oee = 52 Then
이름(Oee) = "신대근"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 700
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 600
ElseIf Oee = 53 Then
이름(Oee) = "김승현"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "웅진"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 600
ElseIf Oee = 54 Then
이름(Oee) = "신재욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "웅진"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 700
ElseIf Oee = 55 Then
이름(Oee) = "이승석"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 650
센스(Oee) = 750
컨트롤(Oee) = 750
ElseIf Oee = 56 Then
이름(Oee) = "정윤종"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 550
센스(Oee) = 600
컨트롤(Oee) = 650
ElseIf Oee = 57 Then
이름(Oee) = "최호선"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 550
센스(Oee) = 700
컨트롤(Oee) = 600
ElseIf Oee = 58 Then
이름(Oee) = "고석현"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 800
견제(Oee) = 750
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 650
ElseIf Oee = 59 Then
이름(Oee) = "김동현"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 600
ElseIf Oee = 60 Then
이름(Oee) = "민찬기"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 700
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 700
ElseIf Oee = 61 Then
이름(Oee) = "박지호"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 750
수비력(Oee) = 500
정찰(Oee) = 550
센스(Oee) = 650
컨트롤(Oee) = 600
ElseIf Oee = 62 Then
이름(Oee) = "김태균"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "화승"
종족(Oee) = 3
공격력(Oee) = 800
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 750
수비력(Oee) = 550
정찰(Oee) = 700
센스(Oee) = 600
컨트롤(Oee) = 650
ElseIf Oee = 63 Then
이름(Oee) = "손주흥"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "화승"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 600
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 650
ElseIf Oee = 64 Then
이름(Oee) = "오영종"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "화승"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 750
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 600
ElseIf Oee = 65 Then
이름(Oee) = "김정우"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 500
물량(Oee) = 700
수비력(Oee) = 650
정찰(Oee) = 650
센스(Oee) = 550
컨트롤(Oee) = 650
ElseIf Oee = 66 Then
이름(Oee) = "서지훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 700
ElseIf Oee = 67 Then
이름(Oee) = "조병세"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 750
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 750
컨트롤(Oee) = 600
ElseIf Oee = 68 Then
이름(Oee) = "고인규"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "공군"
종족(Oee) = 1
공격력(Oee) = 600
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 700
센스(Oee) = 650
컨트롤(Oee) = 600
ElseIf Oee = 69 Then
이름(Oee) = "김태훈"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "공군"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 700
ElseIf Oee = 70 Then
이름(Oee) = "박영민"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "공군"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 650
ElseIf Oee = 71 Then
이름(Oee) = "손석희"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "공군"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 700
ElseIf Oee = 72 Then
이름(Oee) = "안기효"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "공군"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 650
ElseIf Oee = 73 Then
이름(Oee) = "임진묵"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "공군"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 600
수비력(Oee) = 500
정찰(Oee) = 700
센스(Oee) = 600
컨트롤(Oee) = 800
ElseIf Oee = 74 Then
이름(Oee) = "박세정"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "폭스"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 550
ElseIf Oee = 75 Then
이름(Oee) = "이영호1"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "폭스"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 550
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 550
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 650
ElseIf Oee = 76 Then
이름(Oee) = "전상욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "폭스"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 650
컨트롤(Oee) = 650
ElseIf Oee = 77 Then
이름(Oee) = "김대엽"
랭크(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
종족(Oee) = 3
공격력(Oee) = 850
견제(Oee) = 750
전략(Oee) = 600
물량(Oee) = 800
수비력(Oee) = 850
정찰(Oee) = 650
센스(Oee) = 750
컨트롤(Oee) = 750
ElseIf Oee = 78 Then
이름(Oee) = "김성대"
랭크(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
종족(Oee) = 2
공격력(Oee) = 800
견제(Oee) = 700
전략(Oee) = 550
물량(Oee) = 800
수비력(Oee) = 800
정찰(Oee) = 700
센스(Oee) = 600
컨트롤(Oee) = 650
ElseIf Oee = 79 Then
이름(Oee) = "유병준"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "삼성전자"
종족(Oee) = 3
공격력(Oee) = 600
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 750
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 700
ElseIf Oee = 80 Then
이름(Oee) = "차명환"
랭크(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "삼성전자"
종족(Oee) = 2
공격력(Oee) = 850
견제(Oee) = 600
전략(Oee) = 700
물량(Oee) = 800
수비력(Oee) = 800
정찰(Oee) = 700
센스(Oee) = 700
컨트롤(Oee) = 850
ElseIf Oee = 81 Then
이름(Oee) = "허영무"
랭크(Oee) = "Unique"
OYear(Oee) = "<11>"
Team(Oee) = "삼성전자"
종족(Oee) = 3
공격력(Oee) = 900
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 950
수비력(Oee) = 850
정찰(Oee) = 750
센스(Oee) = 750
컨트롤(Oee) = 950
ElseIf Oee = 82 Then
이름(Oee) = "김구현"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 700
수비력(Oee) = 650
정찰(Oee) = 700
센스(Oee) = 600
컨트롤(Oee) = 650
ElseIf Oee = 83 Then
이름(Oee) = "김윤중"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
종족(Oee) = 3
공격력(Oee) = 650
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 900
수비력(Oee) = 550
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 650
ElseIf Oee = 84 Then
이름(Oee) = "김현우"
랭크(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 850
견제(Oee) = 750
전략(Oee) = 650
물량(Oee) = 600
수비력(Oee) = 500
정찰(Oee) = 600
센스(Oee) = 750
컨트롤(Oee) = 900
ElseIf Oee = 85 Then
이름(Oee) = "이신형"
랭크(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 800
수비력(Oee) = 850
정찰(Oee) = 750
센스(Oee) = 700
컨트롤(Oee) = 850
ElseIf Oee = 86 Then
이름(Oee) = "조일장"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 700
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 800
ElseIf Oee = 87 Then
이름(Oee) = "김민철"
랭크(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "웅진"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 700
전략(Oee) = 650
물량(Oee) = 750
수비력(Oee) = 800
정찰(Oee) = 700
센스(Oee) = 650
컨트롤(Oee) = 700
ElseIf Oee = 88 Then
이름(Oee) = "박상우"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "웅진"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 600
전략(Oee) = 600
물량(Oee) = 800
수비력(Oee) = 650
정찰(Oee) = 700
센스(Oee) = 700
컨트롤(Oee) = 700
ElseIf Oee = 89 Then
이름(Oee) = "윤용태"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "웅진"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 800
ElseIf Oee = 90 Then
이름(Oee) = "이재호"
랭크(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "웅진"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 750
전략(Oee) = 750
물량(Oee) = 800
수비력(Oee) = 600
정찰(Oee) = 800
센스(Oee) = 700
컨트롤(Oee) = 800
ElseIf Oee = 91 Then
이름(Oee) = "도재욱"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
종족(Oee) = 3
공격력(Oee) = 800
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 900
수비력(Oee) = 500
정찰(Oee) = 600
센스(Oee) = 750
컨트롤(Oee) = 650
ElseIf Oee = 92 Then
이름(Oee) = "박재혁"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
종족(Oee) = 2
공격력(Oee) = 800
견제(Oee) = 700
전략(Oee) = 550
물량(Oee) = 650
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 750
컨트롤(Oee) = 850
ElseIf Oee = 93 Then
이름(Oee) = "어윤수"
랭크(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 700
전략(Oee) = 600
물량(Oee) = 800
수비력(Oee) = 850
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 750
ElseIf Oee = 94 Then
이름(Oee) = "김재훈"
랭크(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 900
수비력(Oee) = 550
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 750
ElseIf Oee = 95 Then
이름(Oee) = "박수범"
랭크(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
종족(Oee) = 3
공격력(Oee) = 800
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 900
수비력(Oee) = 700
정찰(Oee) = 600
센스(Oee) = 600
컨트롤(Oee) = 750
ElseIf Oee = 96 Then
이름(Oee) = "구성훈"
랭크(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "화승"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 700
전략(Oee) = 750
물량(Oee) = 700
수비력(Oee) = 750
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 650
ElseIf Oee = 97 Then
이름(Oee) = "박준오"
랭크(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "화승"
종족(Oee) = 2
공격력(Oee) = 950
견제(Oee) = 700
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 800
컨트롤(Oee) = 950
ElseIf Oee = 98 Then
이름(Oee) = "신상문"
랭크(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 750
전략(Oee) = 800
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 700
센스(Oee) = 750
컨트롤(Oee) = 850
ElseIf Oee = 99 Then
이름(Oee) = "이경민"
랭크(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
종족(Oee) = 3
공격력(Oee) = 800
견제(Oee) = 600
전략(Oee) = 800
물량(Oee) = 900
수비력(Oee) = 700
정찰(Oee) = 550
센스(Oee) = 800
컨트롤(Oee) = 850
ElseIf Oee = 100 Then
이름(Oee) = "장윤철"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
종족(Oee) = 3
공격력(Oee) = 700
견제(Oee) = 600
전략(Oee) = 650
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 600
센스(Oee) = 700
컨트롤(Oee) = 700
ElseIf Oee = 101 Then
이름(Oee) = "진영화"
랭크(Oee) = "Unique"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
종족(Oee) = 3
공격력(Oee) = 800
견제(Oee) = 850
전략(Oee) = 700
물량(Oee) = 850
수비력(Oee) = 800
정찰(Oee) = 850
센스(Oee) = 750
컨트롤(Oee) = 800
ElseIf Oee = 102 Then
이름(Oee) = "김경모"
랭크(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "공군"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 750
전략(Oee) = 600
물량(Oee) = 800
수비력(Oee) = 800
정찰(Oee) = 700
센스(Oee) = 650
컨트롤(Oee) = 600
ElseIf Oee = 103 Then
이름(Oee) = "변형태"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "공군"
종족(Oee) = 1
공격력(Oee) = 850
견제(Oee) = 700
전략(Oee) = 600
물량(Oee) = 700
수비력(Oee) = 600
정찰(Oee) = 650
센스(Oee) = 650
컨트롤(Oee) = 650
ElseIf Oee = 104 Then
이름(Oee) = "이성은"
랭크(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "공군"
종족(Oee) = 1
공격력(Oee) = 750
견제(Oee) = 700
전략(Oee) = 800
물량(Oee) = 600
수비력(Oee) = 550
정찰(Oee) = 750
센스(Oee) = 700
컨트롤(Oee) = 750
ElseIf Oee = 105 Then
이름(Oee) = "박성균"
랭크(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "폭스"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 750
전략(Oee) = 700
물량(Oee) = 800
수비력(Oee) = 800
정찰(Oee) = 750
센스(Oee) = 650
컨트롤(Oee) = 750
ElseIf Oee = 106 Then
이름(Oee) = "신노열"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "폭스"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 750
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 700
센스(Oee) = 650
컨트롤(Oee) = 750
ElseIf Oee = 107 Then
이름(Oee) = "이영한"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "폭스"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 750
수비력(Oee) = 500
정찰(Oee) = 600
센스(Oee) = 750
컨트롤(Oee) = 700
ElseIf Oee = 108 Then
이름(Oee) = "전태양"
랭크(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "폭스"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 700
전략(Oee) = 750
물량(Oee) = 650
수비력(Oee) = 700
정찰(Oee) = 750
센스(Oee) = 800
컨트롤(Oee) = 750
ElseIf Oee = 109 Then
이름(Oee) = "이영호"
랭크(Oee) = "Legend"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 700
전략(Oee) = 850
물량(Oee) = 950
수비력(Oee) = 950
정찰(Oee) = 800
센스(Oee) = 950
컨트롤(Oee) = 800

ElseIf Oee = 110 Then
이름(Oee) = "송병구"
랭크(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "삼성전자"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 800
수비력(Oee) = 750
정찰(Oee) = 750
센스(Oee) = 800
컨트롤(Oee) = 750
ElseIf Oee = 111 Then
이름(Oee) = "김윤환1"
랭크(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 600
견제(Oee) = 750
전략(Oee) = 800
물량(Oee) = 700
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 600
컨트롤(Oee) = 600
ElseIf Oee = 112 Then
이름(Oee) = "김명운"
랭크(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "웅진"
종족(Oee) = 2
공격력(Oee) = 650
견제(Oee) = 750
전략(Oee) = 700
물량(Oee) = 800
수비력(Oee) = 800
정찰(Oee) = 850
센스(Oee) = 750
컨트롤(Oee) = 900
ElseIf Oee = 113 Then
이름(Oee) = "김택용"
랭크(Oee) = "Legend"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 950
전략(Oee) = 750
물량(Oee) = 900
수비력(Oee) = 800
정찰(Oee) = 950
센스(Oee) = 750
컨트롤(Oee) = 950
ElseIf Oee = 114 Then
이름(Oee) = "정명훈"
랭크(Oee) = "Unique"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
종족(Oee) = 1
공격력(Oee) = 700
견제(Oee) = 950
전략(Oee) = 850
물량(Oee) = 850
수비력(Oee) = 900
정찰(Oee) = 850
센스(Oee) = 800
컨트롤(Oee) = 650
ElseIf Oee = 115 Then
이름(Oee) = "염보성"
랭크(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
종족(Oee) = 1
공격력(Oee) = 650
견제(Oee) = 750
전략(Oee) = 700
물량(Oee) = 800
수비력(Oee) = 800
정찰(Oee) = 800
센스(Oee) = 750
컨트롤(Oee) = 700
ElseIf Oee = 116 Then
이름(Oee) = "이제동"
랭크(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "화승"
종족(Oee) = 2
공격력(Oee) = 850
견제(Oee) = 750
전략(Oee) = 650
물량(Oee) = 750
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 800
컨트롤(Oee) = 850
ElseIf Oee = 117 Then
이름(Oee) = "신동원"
랭크(Oee) = "Elite"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 900
견제(Oee) = 850
전략(Oee) = 600
물량(Oee) = 900
수비력(Oee) = 850
정찰(Oee) = 700
센스(Oee) = 950
컨트롤(Oee) = 950
ElseIf Oee = 118 Then
 이름(Oee) = "이기석"
 OYear(Oee) = "<99>"
 랭크(Oee) = "Legend"
 Team(Oee) = "자료없음"
 종족(Oee) = 1
 공격력(Oee) = 800
 견제(Oee) = 800
 전략(Oee) = 900
 물량(Oee) = 900
 수비력(Oee) = 800
 정찰(Oee) = 800
 센스(Oee) = 900
 컨트롤(Oee) = 900
End If
 우승(Oee) = 0
 준우승(Oee) = 0
 컨디션(Oee) = 100
 A승리(Oee) = 0
 A패배(Oee) = 0
 P승리(Oee) = 0
 P패배(Oee) = 0
 T승리(Oee) = 0
 T패배(Oee) = 0
 Z승리(Oee) = 0
 Z패배(Oee) = 0
 T연승(Oee) = 0
 Z연승(Oee) = 0
 P연승(Oee) = 0
 A연승(Oee) = 0
 T연(Oee) = "W"
 Z연(Oee) = "W"
 P연(Oee) = "W"
 A연(Oee) = "W"
Next Oee

TimElse.Enabled = True
Tim11.Enabled = False
End Sub

Private Sub Tim12_Timer()
For Oee = 724 To 800
    If Oee = 724 Then
        이름(Oee) = "염보성"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "8th"
        종족(Oee) = 1
        공격력(Oee) = 850
        견제(Oee) = 600
        전략(Oee) = 600
        물량(Oee) = 600
        수비력(Oee) = 600
        정찰(Oee) = 700
        센스(Oee) = 700
        컨트롤(Oee) = 850
    ElseIf Oee = 725 Then
        이름(Oee) = "전태양"
        랭크(Oee) = "Special"
        OYear(Oee) = "<12>"
        Team(Oee) = "8th"
        종족(Oee) = 1
        공격력(Oee) = 650
        견제(Oee) = 650
        전략(Oee) = 650
        물량(Oee) = 600
        수비력(Oee) = 700
        정찰(Oee) = 800
        센스(Oee) = 900
        컨트롤(Oee) = 650
    ElseIf Oee = 726 Then
        이름(Oee) = "김도욱"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "8th"
        종족(Oee) = 1
        공격력(Oee) = 600
    
        견제(Oee) = 550
    
        전략(Oee) = 550
    
        물량(Oee) = 650
    
        수비력(Oee) = 700
    
        정찰(Oee) = 600
    
        센스(Oee) = 500
    
        컨트롤(Oee) = 500
    
    ElseIf Oee = 727 Then
    
        이름(Oee) = "김재훈"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "8th"
    
        종족(Oee) = 3
    
        공격력(Oee) = 650
    
        견제(Oee) = 800
    
        전략(Oee) = 600
    
        물량(Oee) = 750
    
        수비력(Oee) = 650
    
        정찰(Oee) = 800
    
        센스(Oee) = 650
    
        컨트롤(Oee) = 650
    
    ElseIf Oee = 728 Then
    
        이름(Oee) = "박수범"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "8th"
    
        종족(Oee) = 3
    
        공격력(Oee) = 850
    
        견제(Oee) = 600
    
        전략(Oee) = 600
    
        물량(Oee) = 850
    
        수비력(Oee) = 700
    
        정찰(Oee) = 600
    
        센스(Oee) = 700
    
        컨트롤(Oee) = 600
    
    ElseIf Oee = 729 Then
    
        이름(Oee) = "하재상"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "8th"
    
        종족(Oee) = 3
    
        공격력(Oee) = 650
    
        견제(Oee) = 550
    
        전략(Oee) = 650
    
        물량(Oee) = 550
    
        수비력(Oee) = 600
    
        정찰(Oee) = 700
    
        센스(Oee) = 500
    
        컨트롤(Oee) = 500
    
    ElseIf Oee = 730 Then
    
        이름(Oee) = "이제동"
    
        랭크(Oee) = "Rare"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "8th"
    
        종족(Oee) = 2
    
        공격력(Oee) = 850
    
        견제(Oee) = 700
    
        전략(Oee) = 600
    
        물량(Oee) = 650
    
        수비력(Oee) = 700
    
        정찰(Oee) = 550
    
        센스(Oee) = 700
    
        컨트롤(Oee) = 850
    
    ElseIf Oee = 731 Then
    
        이름(Oee) = "백동준"
    
        랭크(Oee) = "Rare"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        종족(Oee) = 3
    
        공격력(Oee) = 950
    
        견제(Oee) = 600
    
        전략(Oee) = 600
    
        물량(Oee) = 950
    
        수비력(Oee) = 600
    
        정찰(Oee) = 800
    
        센스(Oee) = 750
    
        컨트롤(Oee) = 750
    
    ElseIf Oee = 732 Then
    
        이름(Oee) = "이영호"
    
        랭크(Oee) = "Elite"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "KT"
    
        종족(Oee) = 1
    
        공격력(Oee) = 950
    
        견제(Oee) = 700
    
        전략(Oee) = 850
    
        물량(Oee) = 800
    
        수비력(Oee) = 700
    
        정찰(Oee) = 850
    
        센스(Oee) = 950
    
        컨트롤(Oee) = 950
    
    ElseIf Oee = 733 Then
    
        이름(Oee) = "정명훈"
    
        랭크(Oee) = "Elite"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "SK"
    
        종족(Oee) = 1
    
        공격력(Oee) = 950
    
        견제(Oee) = 800
    
        전략(Oee) = 850
    
        물량(Oee) = 800
    
        수비력(Oee) = 850
    
        정찰(Oee) = 850
    
        센스(Oee) = 850
    
        컨트롤(Oee) = 850
    
    ElseIf Oee = 734 Then
    
        이름(Oee) = "송병구"
    
        랭크(Oee) = "Elite"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "삼성"
    
        종족(Oee) = 3
    
        공격력(Oee) = 700
    
        견제(Oee) = 900
    
        전략(Oee) = 850
    
        물량(Oee) = 800
    
        수비력(Oee) = 850
    
        정찰(Oee) = 800
    
        센스(Oee) = 850
    
        컨트롤(Oee) = 950
    
    ElseIf Oee = 735 Then
    
        이름(Oee) = "김민철"
    
        랭크(Oee) = "Rare"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "웅진"
    
        종족(Oee) = 2
    
        공격력(Oee) = 800
    
        견제(Oee) = 800
    
        전략(Oee) = 700
    
        물량(Oee) = 900
    
        수비력(Oee) = 900
    
        정찰(Oee) = 700
    
        센스(Oee) = 800
    
        컨트롤(Oee) = 800
    
    ElseIf Oee = 736 Then
    
        이름(Oee) = "도재욱"
    
        랭크(Oee) = "Rare"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "SK"
    
        종족(Oee) = 3
    
        공격력(Oee) = 600
    
        견제(Oee) = 950
    
        전략(Oee) = 700
    
        물량(Oee) = 600
    
        수비력(Oee) = 850
    
        정찰(Oee) = 950
    
        센스(Oee) = 850
    
        컨트롤(Oee) = 800
    
    ElseIf Oee = 737 Then
    
        이름(Oee) = "신상문"
    
        랭크(Oee) = "Rare"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        종족(Oee) = 1
    
        공격력(Oee) = 900
    
        견제(Oee) = 750
    
        전략(Oee) = 850
    
        물량(Oee) = 750
    
        수비력(Oee) = 850
    
        정찰(Oee) = 800
    
        센스(Oee) = 650
    
        컨트롤(Oee) = 850
    
    ElseIf Oee = 738 Then
    
        이름(Oee) = "김정우"
    
        랭크(Oee) = "Rare"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        종족(Oee) = 2
    
        공격력(Oee) = 700
    
        견제(Oee) = 750
    
        전략(Oee) = 850
    
        물량(Oee) = 750
    
        수비력(Oee) = 900
    
        정찰(Oee) = 700
    
        센스(Oee) = 850
    
        컨트롤(Oee) = 800
    
    ElseIf Oee = 739 Then
    
        이름(Oee) = "김택용"
    
        랭크(Oee) = "Unique"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "SK"
    
        종족(Oee) = 3
    
        공격력(Oee) = 750
    
        견제(Oee) = 900
    
        전략(Oee) = 700
    
        물량(Oee) = 750
    
        수비력(Oee) = 800
    
        정찰(Oee) = 950
    
        센스(Oee) = 800
    
        컨트롤(Oee) = 850
    
    ElseIf Oee = 740 Then
    
        이름(Oee) = "김대엽"
    
        랭크(Oee) = "Unique"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "KT"
    
        종족(Oee) = 3
    
        공격력(Oee) = 800
    
        견제(Oee) = 800
    
        전략(Oee) = 800
    
        물량(Oee) = 850
    
        수비력(Oee) = 750
    
        정찰(Oee) = 800
    
        센스(Oee) = 800
    
        컨트롤(Oee) = 800
    
    ElseIf Oee = 741 Then
    
        이름(Oee) = "김성현"
    
        랭크(Oee) = "Rare"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        종족(Oee) = 1
    
        공격력(Oee) = 600
    
        견제(Oee) = 850
    
        전략(Oee) = 750
    
        물량(Oee) = 900
    
        수비력(Oee) = 850
    
        정찰(Oee) = 800
    
        센스(Oee) = 900
    
        컨트롤(Oee) = 600
    
    ElseIf Oee = 742 Then
    
        이름(Oee) = "임정현"
    
        랭크(Oee) = "Rare"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "웅진"
    
        종족(Oee) = 2
    
        공격력(Oee) = 900
    
        견제(Oee) = 700
    
        전략(Oee) = 600
    
        물량(Oee) = 750
    
        수비력(Oee) = 800
    
        정찰(Oee) = 700
    
        센스(Oee) = 850
    
        컨트롤(Oee) = 900
    
    ElseIf Oee = 743 Then
    
        이름(Oee) = "김윤환1"
    
        랭크(Oee) = "Special"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        종족(Oee) = 2
    
        공격력(Oee) = 600
    
        견제(Oee) = 700
    
        전략(Oee) = 850
    
        물량(Oee) = 800
    
        수비력(Oee) = 850
    
        정찰(Oee) = 800
    
        센스(Oee) = 700
    
        컨트롤(Oee) = 700
    
    ElseIf Oee = 744 Then
    
        이름(Oee) = "허영무"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "삼성"
    
        종족(Oee) = 3
    
        공격력(Oee) = 700
    
        견제(Oee) = 600
    
        전략(Oee) = 650
    
        물량(Oee) = 750
    
        수비력(Oee) = 750
    
        정찰(Oee) = 650
    
        센스(Oee) = 700
    
        컨트롤(Oee) = 750
    
    ElseIf Oee = 745 Then
    
        이름(Oee) = "김유진"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "웅진"
    
        종족(Oee) = 3
    
        공격력(Oee) = 750
    
        견제(Oee) = 650
    
        전략(Oee) = 700
    
        물량(Oee) = 750
    
        수비력(Oee) = 750
    
        정찰(Oee) = 600
    
        센스(Oee) = 700
    
        컨트롤(Oee) = 650
    
    ElseIf Oee = 746 Then
    
        이름(Oee) = "김명운"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "웅진"
    
        종족(Oee) = 2
    
        공격력(Oee) = 850
    
        견제(Oee) = 700
    
        전략(Oee) = 500
    
        물량(Oee) = 650
    
        수비력(Oee) = 600
    
        정찰(Oee) = 650
    
        센스(Oee) = 800
    
        컨트롤(Oee) = 800
    
    ElseIf Oee = 747 Then
    
        이름(Oee) = "이재호"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "웅진"
    
        종족(Oee) = 1
    
        공격력(Oee) = 850
    
        견제(Oee) = 500
    
        전략(Oee) = 700
    
        물량(Oee) = 500
    
        수비력(Oee) = 600
    
        정찰(Oee) = 700
    
        센스(Oee) = 850
    
        컨트롤(Oee) = 850
    
    ElseIf Oee = 748 Then
    
        이름(Oee) = "임태규"
    
        랭크(Oee) = "Special"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "삼성"
    
        종족(Oee) = 3
    
        공격력(Oee) = 650
    
        견제(Oee) = 600
    
        전략(Oee) = 650
    
        물량(Oee) = 750
    
        수비력(Oee) = 700
    
        정찰(Oee) = 600
    
        센스(Oee) = 800
    
        컨트롤(Oee) = 850
    
    ElseIf Oee = 749 Then
    
        이름(Oee) = "어윤수"
    
        랭크(Oee) = "Special"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "SK"
    
        종족(Oee) = 2
    
        공격력(Oee) = 700
    
        견제(Oee) = 750
    
        전략(Oee) = 650
    
        물량(Oee) = 800
    
        수비력(Oee) = 700
    
        정찰(Oee) = 700
    
        센스(Oee) = 650
    
        컨트롤(Oee) = 650
    
    ElseIf Oee = 750 Then
    
        이름(Oee) = "신동원"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        종족(Oee) = 2
    
        공격력(Oee) = 850
    
        견제(Oee) = 600
    
        전략(Oee) = 600
    
        물량(Oee) = 700
    
        수비력(Oee) = 700
    
        정찰(Oee) = 700
    
        센스(Oee) = 700
    
        컨트롤(Oee) = 700
    
    ElseIf Oee = 751 Then
    
        이름(Oee) = "이신형"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        종족(Oee) = 1
    
        공격력(Oee) = 600
    
        견제(Oee) = 800
    
        전략(Oee) = 700
    
        물량(Oee) = 800
    
        수비력(Oee) = 750
    
        정찰(Oee) = 500
    
        센스(Oee) = 700
    
        컨트롤(Oee) = 700
    
    ElseIf Oee = 752 Then
    
        이름(Oee) = "이경민"
    
        랭크(Oee) = "Special"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        종족(Oee) = 3
    
        공격력(Oee) = 850
    
        견제(Oee) = 650
    
        전략(Oee) = 600
    
        물량(Oee) = 850
    
        수비력(Oee) = 600
    
        정찰(Oee) = 700
    
        센스(Oee) = 650
    
        컨트롤(Oee) = 650
    
    ElseIf Oee = 753 Then
    
        이름(Oee) = "김구현"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "공군"
    
        종족(Oee) = 3
    
        공격력(Oee) = 700
    
        견제(Oee) = 600
    
        전략(Oee) = 700
    
        물량(Oee) = 850
    
        수비력(Oee) = 700
    
        정찰(Oee) = 550
    
        센스(Oee) = 700
    
        컨트롤(Oee) = 750
    
    ElseIf Oee = 754 Then
    
        이름(Oee) = "박대호"
    
        랭크(Oee) = "Special"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "삼성"
    
        종족(Oee) = 1
    
        공격력(Oee) = 950
    
        견제(Oee) = 800
    
        전략(Oee) = 600
    
        물량(Oee) = 700
    
        수비력(Oee) = 750
    
        정찰(Oee) = 650
    
        센스(Oee) = 800
    
        컨트롤(Oee) = 550
    
    ElseIf Oee = 755 Then
    
        이름(Oee) = "고인규"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "공군"
    
        종족(Oee) = 1
    
        공격력(Oee) = 850
    
        견제(Oee) = 650
    
        전략(Oee) = 600
    
        물량(Oee) = 650
    
        수비력(Oee) = 700
    
        정찰(Oee) = 600
    
        센스(Oee) = 700
    
        컨트롤(Oee) = 800
    
    ElseIf Oee = 756 Then
    
        이름(Oee) = "주성욱"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "KT"
    
        종족(Oee) = 3
    
        공격력(Oee) = 600
    
        견제(Oee) = 800
    
        전략(Oee) = 600
    
        물량(Oee) = 600
    
        수비력(Oee) = 700
    
        정찰(Oee) = 750
    
        센스(Oee) = 700
    
        컨트롤(Oee) = 750
    
    ElseIf Oee = 757 Then
    
        이름(Oee) = "권수현"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "공군"
    
        종족(Oee) = 2
    
        공격력(Oee) = 600
    
        견제(Oee) = 700
    
        전략(Oee) = 650
    
        물량(Oee) = 750
    
        수비력(Oee) = 800
    
        정찰(Oee) = 700
    
        센스(Oee) = 650
    
        컨트롤(Oee) = 650
    
    ElseIf Oee = 758 Then
    
        이름(Oee) = "김기현"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "삼성"
    
        종족(Oee) = 1
    
        공격력(Oee) = 550
    
        견제(Oee) = 750
    
        전략(Oee) = 600
    
        물량(Oee) = 700
    
        수비력(Oee) = 700
    
        정찰(Oee) = 650
    
        센스(Oee) = 800
    
        컨트롤(Oee) = 750
    
    ElseIf Oee = 759 Then
    
        이름(Oee) = "진영화"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        종족(Oee) = 3
    
        공격력(Oee) = 600
    
        견제(Oee) = 750
    
        전략(Oee) = 650
    
        물량(Oee) = 800
    
        수비력(Oee) = 800
    
        정찰(Oee) = 600
    
        센스(Oee) = 600
    
        컨트롤(Oee) = 700
    
    ElseIf Oee = 760 Then
    
        이름(Oee) = "변현제"
    
        랭크(Oee) = "Special"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        종족(Oee) = 3
    
        공격력(Oee) = 850
    
        견제(Oee) = 800
    
        전략(Oee) = 650
    
        물량(Oee) = 700
    
        수비력(Oee) = 700
    
        정찰(Oee) = 600
    
        센스(Oee) = 750
    
        컨트롤(Oee) = 750
    
    ElseIf Oee = 761 Then
    
        이름(Oee) = "박준오"
    
        랭크(Oee) = "Special"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "8th"
    
        종족(Oee) = 2
    
        공격력(Oee) = 800
    
        견제(Oee) = 600
    
        전략(Oee) = 600
    
        물량(Oee) = 700
    
        수비력(Oee) = 700
    
        정찰(Oee) = 600
    
        센스(Oee) = 750
    
        컨트롤(Oee) = 850
    
    ElseIf Oee = 762 Then
    
        이름(Oee) = "이병렬"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "8th"
    
        종족(Oee) = 2
    
        공격력(Oee) = 650
    
        견제(Oee) = 600
    
        전략(Oee) = 600
    
        물량(Oee) = 850
    
        수비력(Oee) = 750
    
        정찰(Oee) = 650
    
        센스(Oee) = 700
    
        컨트롤(Oee) = 700
    
    ElseIf Oee = 763 Then
    
        이름(Oee) = "조병세"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        종족(Oee) = 1
    
        공격력(Oee) = 750
    
        견제(Oee) = 600
    
        전략(Oee) = 500
    
        물량(Oee) = 850
    
        수비력(Oee) = 650
    
        정찰(Oee) = 600
    
        센스(Oee) = 850
    
        컨트롤(Oee) = 700
    
    ElseIf Oee = 764 Then
    
        이름(Oee) = "유영진"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        종족(Oee) = 1
    
        공격력(Oee) = 600
    
        견제(Oee) = 700
    
        전략(Oee) = 750
    
        물량(Oee) = 550
    
        수비력(Oee) = 600
    
        정찰(Oee) = 750
    
        센스(Oee) = 550
    
        컨트롤(Oee) = 500
    
    ElseIf Oee = 765 Then
    
        이름(Oee) = "장윤철"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        종족(Oee) = 3
    
        공격력(Oee) = 800
    
        견제(Oee) = 550
    
        전략(Oee) = 600
    
        물량(Oee) = 800
    
        수비력(Oee) = 700
    
        정찰(Oee) = 600
    
        센스(Oee) = 700
    
        컨트롤(Oee) = 650
    
    ElseIf Oee = 766 Then
    
        이름(Oee) = "김준호"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        종족(Oee) = 2
    
        공격력(Oee) = 550
    
        견제(Oee) = 650
    
        전략(Oee) = 600
    
        물량(Oee) = 700
    
        수비력(Oee) = 650
    
        정찰(Oee) = 600
    
        센스(Oee) = 500
    
        컨트롤(Oee) = 500
    
    ElseIf Oee = 767 Then
    
        이름(Oee) = "한두열"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        종족(Oee) = 2
    
        공격력(Oee) = 600
    
        견제(Oee) = 600
    
        전략(Oee) = 600
    
        물량(Oee) = 600
    
        수비력(Oee) = 600
    
        정찰(Oee) = 600
    
        센스(Oee) = 600
    
        컨트롤(Oee) = 600
    
    ElseIf Oee = 768 Then
    
        이름(Oee) = "유준희"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "삼성"
    
        종족(Oee) = 2
    
        공격력(Oee) = 950
    
        견제(Oee) = 500
    
        전략(Oee) = 500
    
        물량(Oee) = 500
    
        수비력(Oee) = 500
    
        정찰(Oee) = 500
    
        센스(Oee) = 500
    
        컨트롤(Oee) = 950
    
    ElseIf Oee = 769 Then
    
        이름(Oee) = "김도우"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        종족(Oee) = 1
    
        공격력(Oee) = 550
    
        견제(Oee) = 600
    
        전략(Oee) = 650
    
        물량(Oee) = 550
    
        수비력(Oee) = 600
    
        정찰(Oee) = 650
    
        센스(Oee) = 750
    
        컨트롤(Oee) = 600
    
    ElseIf Oee = 770 Then
    
        이름(Oee) = "김윤중"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        종족(Oee) = 3
    
        공격력(Oee) = 650
    
        견제(Oee) = 600
    
        전략(Oee) = 550
    
        물량(Oee) = 750
    
        수비력(Oee) = 500
    
        정찰(Oee) = 600
    
        센스(Oee) = 650
    
        컨트롤(Oee) = 550
    
    ElseIf Oee = 771 Then
    
        이름(Oee) = "조성호"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        종족(Oee) = 3
    
        공격력(Oee) = 550
    
        견제(Oee) = 500
    
        전략(Oee) = 550
    
        물량(Oee) = 650
    
        수비력(Oee) = 550
    
        정찰(Oee) = 500
    
        센스(Oee) = 650
    
        컨트롤(Oee) = 700
    
    ElseIf Oee = 772 Then
    
        이름(Oee) = "조일장"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        종족(Oee) = 2
    
        공격력(Oee) = 650
    
        견제(Oee) = 600
    
        전략(Oee) = 600
    
        물량(Oee) = 700
    
        수비력(Oee) = 750
    
        정찰(Oee) = 600
    
        센스(Oee) = 600
    
        컨트롤(Oee) = 600
    
    ElseIf Oee = 773 Then
    
        이름(Oee) = "신대근"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        종족(Oee) = 2
    
        공격력(Oee) = 600
    
        견제(Oee) = 700
    
        전략(Oee) = 600
    
        물량(Oee) = 800
    
        수비력(Oee) = 800
    
        정찰(Oee) = 650
    
        센스(Oee) = 650
    
        컨트롤(Oee) = 600
    
    ElseIf Oee = 774 Then
    
        이름(Oee) = "김현우"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        종족(Oee) = 2
    
        공격력(Oee) = 750
    
        견제(Oee) = 600
    
        전략(Oee) = 600
    
        물량(Oee) = 550
    
        수비력(Oee) = 550
    
        정찰(Oee) = 500
    
        센스(Oee) = 700
    
        컨트롤(Oee) = 750
    
    ElseIf Oee = 775 Then
    
        이름(Oee) = "노준규"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "웅진"
    
        종족(Oee) = 1
    
        공격력(Oee) = 600
    
        견제(Oee) = 650
    
        전략(Oee) = 550
    
        물량(Oee) = 650
    
        수비력(Oee) = 550
    
        정찰(Oee) = 600
    
        센스(Oee) = 550
    
        컨트롤(Oee) = 600
    
    ElseIf Oee = 776 Then
    
        이름(Oee) = "윤용태"
    
        랭크(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "웅진"
    
        종족(Oee) = 3
    
        공격력(Oee) = 600
    
        견제(Oee) = 600
    
        전략(Oee) = 600
    
        물량(Oee) = 650
    
        수비력(Oee) = 650
    
        정찰(Oee) = 600
    
        센스(Oee) = 600
    
        컨트롤(Oee) = 750
    
    ElseIf Oee = 777 Then
        이름(Oee) = "신재욱"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "웅진"
        종족(Oee) = 3
        공격력(Oee) = 650
        견제(Oee) = 650
        전략(Oee) = 600
        물량(Oee) = 800
        수비력(Oee) = 700
        정찰(Oee) = 600
        센스(Oee) = 700
        컨트롤(Oee) = 850
    ElseIf Oee = 778 Then
        이름(Oee) = "박성균"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        종족(Oee) = 1
        공격력(Oee) = 600
        견제(Oee) = 550
        전략(Oee) = 650
        물량(Oee) = 550
        수비력(Oee) = 650
        정찰(Oee) = 550
        센스(Oee) = 650
        컨트롤(Oee) = 600
    ElseIf Oee = 779 Then
        이름(Oee) = "황병영"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        종족(Oee) = 1
        공격력(Oee) = 650
        견제(Oee) = 600
        전략(Oee) = 600
        물량(Oee) = 600
        수비력(Oee) = 600
        정찰(Oee) = 600
        센스(Oee) = 600
        컨트롤(Oee) = 600
    ElseIf Oee = 780 Then
        이름(Oee) = "김태균"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        종족(Oee) = 3
        공격력(Oee) = 600
        견제(Oee) = 600
        전략(Oee) = 600
        물량(Oee) = 600
        수비력(Oee) = 600
        정찰(Oee) = 600
        센스(Oee) = 600
        컨트롤(Oee) = 600
    ElseIf Oee = 781 Then
        이름(Oee) = "고강민"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        종족(Oee) = 2
        공격력(Oee) = 850
        견제(Oee) = 500
        전략(Oee) = 500
        물량(Oee) = 700
        수비력(Oee) = 750
        정찰(Oee) = 700
        센스(Oee) = 700
        컨트롤(Oee) = 750
    ElseIf Oee = 782 Then
        이름(Oee) = "김성대"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        종족(Oee) = 2
        공격력(Oee) = 600
        견제(Oee) = 700
        전략(Oee) = 850
        물량(Oee) = 650
        수비력(Oee) = 600
        정찰(Oee) = 650
        센스(Oee) = 650
        컨트롤(Oee) = 800
    ElseIf Oee = 783 Then
        이름(Oee) = "최용주"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        종족(Oee) = 2
        공격력(Oee) = 600
        견제(Oee) = 650
        전략(Oee) = 650
        물량(Oee) = 500
        수비력(Oee) = 600
        정찰(Oee) = 700
        센스(Oee) = 600
        컨트롤(Oee) = 500
    ElseIf Oee = 784 Then
        이름(Oee) = "최호선"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "SK"
        종족(Oee) = 1
        공격력(Oee) = 550
        견제(Oee) = 500
        전략(Oee) = 550
        물량(Oee) = 500
        수비력(Oee) = 600
        정찰(Oee) = 650
        센스(Oee) = 950
        컨트롤(Oee) = 550
    ElseIf Oee = 785 Then
        이름(Oee) = "박재혁"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "SK"
        종족(Oee) = 2
        공격력(Oee) = 750
        견제(Oee) = 500
        전략(Oee) = 500
        물량(Oee) = 500
        수비력(Oee) = 600
        정찰(Oee) = 700
        센스(Oee) = 600
        컨트롤(Oee) = 750
    ElseIf Oee = 786 Then
        이름(Oee) = "이승석"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "SK"
        종족(Oee) = 2
        공격력(Oee) = 650
        견제(Oee) = 600
        전략(Oee) = 550
        물량(Oee) = 500
        수비력(Oee) = 700
        정찰(Oee) = 600
        센스(Oee) = 600
        컨트롤(Oee) = 600
    ElseIf Oee = 787 Then
        이름(Oee) = "변형태"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "공군"
        종족(Oee) = 1
        공격력(Oee) = 850
        견제(Oee) = 550
        전략(Oee) = 500
        물량(Oee) = 550
        수비력(Oee) = 550
        정찰(Oee) = 600
        센스(Oee) = 500
        컨트롤(Oee) = 750
    ElseIf Oee = 788 Then
        이름(Oee) = "이성은"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "공군"
        종족(Oee) = 1
        공격력(Oee) = 850
        견제(Oee) = 550
        전략(Oee) = 550
        물량(Oee) = 500
        수비력(Oee) = 500
        정찰(Oee) = 500
        센스(Oee) = 500
        컨트롤(Oee) = 850
    ElseIf Oee = 789 Then
        이름(Oee) = "임진묵"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "공군"
        종족(Oee) = 1
        공격력(Oee) = 750
        견제(Oee) = 550
        전략(Oee) = 750
        물량(Oee) = 650
        수비력(Oee) = 750
        정찰(Oee) = 550
        센스(Oee) = 650
        컨트롤(Oee) = 750
    ElseIf Oee = 790 Then
        이름(Oee) = "손석희"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "공군"
        종족(Oee) = 3
        공격력(Oee) = 850
        견제(Oee) = 600
        전략(Oee) = 600
        물량(Oee) = 800
        수비력(Oee) = 700
        정찰(Oee) = 700
        센스(Oee) = 650
        컨트롤(Oee) = 650
    ElseIf Oee = 791 Then
        이름(Oee) = "김태훈"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "공군"
        종족(Oee) = 2
        공격력(Oee) = 600
        견제(Oee) = 600
        전략(Oee) = 600
        물량(Oee) = 600
        수비력(Oee) = 600
        정찰(Oee) = 600
        센스(Oee) = 600
        컨트롤(Oee) = 600
    ElseIf Oee = 792 Then
        이름(Oee) = "김경모"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "공군"
        종족(Oee) = 2
        공격력(Oee) = 600
        견제(Oee) = 600
        전략(Oee) = 600
        물량(Oee) = 600
        수비력(Oee) = 600
        정찰(Oee) = 600
        센스(Oee) = 600
        컨트롤(Oee) = 600
    ElseIf Oee = 793 Then
        이름(Oee) = "차명환"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "공군"
        종족(Oee) = 2
        공격력(Oee) = 550
        견제(Oee) = 550
        전략(Oee) = 550
        물량(Oee) = 550
        수비력(Oee) = 550
        정찰(Oee) = 550
        센스(Oee) = 550
        컨트롤(Oee) = 550
    ElseIf Oee = 794 Then
        이름(Oee) = "이정현"
        랭크(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "공군"
        종족(Oee) = 2
        공격력(Oee) = 500
        견제(Oee) = 500
        전략(Oee) = 500
        물량(Oee) = 500
        수비력(Oee) = 500
        정찰(Oee) = 500
        센스(Oee) = 500
        컨트롤(Oee) = 500
    ElseIf Oee = 795 Then
        이름(Oee) = "고강민[Ex]"
        랭크(Oee) = "Unique"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        종족(Oee) = 2
        공격력(Oee) = 950
        견제(Oee) = 700
        전략(Oee) = 700
        물량(Oee) = 850
        수비력(Oee) = 850
        정찰(Oee) = 750
        센스(Oee) = 850
        컨트롤(Oee) = 850
    ElseIf Oee = 796 Then
        이름(Oee) = "김성대[Ex]"
        랭크(Oee) = "Unique"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        종족(Oee) = 2
        공격력(Oee) = 800
        견제(Oee) = 750
        전략(Oee) = 950
        물량(Oee) = 900
        수비력(Oee) = 900
        정찰(Oee) = 800
        센스(Oee) = 700
        컨트롤(Oee) = 700
    ElseIf Oee = 797 Then
        이름(Oee) = "이영호[Ex]"
        랭크(Oee) = "Champion"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        종족(Oee) = 1
        공격력(Oee) = 950
        견제(Oee) = 900
        전략(Oee) = 900
        물량(Oee) = 950
        수비력(Oee) = 900
        정찰(Oee) = 900
        센스(Oee) = 900
        Skill(Oee) = 2
        컨트롤(Oee) = 1000
    ElseIf Oee = 798 Then
        이름(Oee) = "정명훈[Ex]"
        랭크(Oee) = "Elite"
        OYear(Oee) = "<10>"
        Team(Oee) = "SK"
        종족(Oee) = 1
        공격력(Oee) = 750
        견제(Oee) = 900
        전략(Oee) = 850
        물량(Oee) = 850
        수비력(Oee) = 900
        정찰(Oee) = 900
        센스(Oee) = 850
        컨트롤(Oee) = 650
    ElseIf Oee = 799 Then
        이름(Oee) = "박재혁[Ex]"
        랭크(Oee) = "Unique"
        OYear(Oee) = "<09>"
        Team(Oee) = "SK"
        종족(Oee) = 2
        공격력(Oee) = 950
        견제(Oee) = 850
        전략(Oee) = 650
        물량(Oee) = 800
        수비력(Oee) = 650
        정찰(Oee) = 800
        센스(Oee) = 850
        컨트롤(Oee) = 900
    ElseIf Oee = 800 Then
        이름(Oee) = "KT"
        랭크(Oee) = "Elite"
        OYear(Oee) = "<12>"
        Team(Oee) = "Mystar"
        종족(Oee) = 1
        공격력(Oee) = 950
        견제(Oee) = 900
        전략(Oee) = 650
        물량(Oee) = 800
        수비력(Oee) = 950
        정찰(Oee) = 700
        센스(Oee) = 700
        컨트롤(Oee) = 950
    End If
    우승(Oee) = 0
    준우승(Oee) = 0
    컨디션(Oee) = 100
    A승리(Oee) = 0
    A패배(Oee) = 0
    P승리(Oee) = 0
    P패배(Oee) = 0
    T승리(Oee) = 0
    T패배(Oee) = 0
    Z승리(Oee) = 0
    Z패배(Oee) = 0
    T연승(Oee) = 0
    Z연승(Oee) = 0
    P연승(Oee) = 0
    A연승(Oee) = 0
    T연(Oee) = "W"
    Z연(Oee) = "W"
    P연(Oee) = "W"
    A연(Oee) = "W"
Next

Tim12.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub TimAdd_Timer()
For Oee = 716 To 723
 If Oee = 716 Then
  이름(Oee) = "은하랑"
  랭크(Oee) = "Rare"
  OYear(Oee) = "<11>"
  Team(Oee) = "Mystar"
  종족(Oee) = 1
  공격력(Oee) = 700
  견제(Oee) = 650
  전략(Oee) = 700
  물량(Oee) = 750
  수비력(Oee) = 800
  정찰(Oee) = 800
  센스(Oee) = 1000
  컨트롤(Oee) = 600
 ElseIf Oee = 717 Then
  이름(Oee) = "박용욱"
  랭크(Oee) = "Legend"
  OYear(Oee) = "<02>"
  Team(Oee) = "IS"
  종족(Oee) = 3
  공격력(Oee) = 800
  견제(Oee) = 950
  전략(Oee) = 700
  물량(Oee) = 800
  수비력(Oee) = 800
  정찰(Oee) = 950
  센스(Oee) = 850
  컨트롤(Oee) = 950
 ElseIf Oee = 718 Then
  이름(Oee) = "서지수"
  랭크(Oee) = "Unique"
  OYear(Oee) = "<02>"
  Team(Oee) = "STX"
  종족(Oee) = 1
  공격력(Oee) = 900
  견제(Oee) = 900
  전략(Oee) = 800
  물량(Oee) = 750
  수비력(Oee) = 800
  정찰(Oee) = 650
  센스(Oee) = 850
  컨트롤(Oee) = 900
ElseIf Oee = 719 Then
 이름(Oee) = "카이"
 랭크(Oee) = "Unique"
 Team(Oee) = "Mystar"
 OYear(Oee) = "<11>"
 종족(Oee) = 2
 공격력(Oee) = 950
 견제(Oee) = 550
 전략(Oee) = 800
 물량(Oee) = 800
 수비력(Oee) = 550
 정찰(Oee) = 950
 센스(Oee) = 950
 컨트롤(Oee) = 950
ElseIf Oee = 720 Then
 이름(Oee) = "백영"
 랭크(Oee) = "Rare"
 Team(Oee) = "Mystar"
 OYear(Oee) = "<11>"
 종족(Oee) = 3
 공격력(Oee) = 900
 견제(Oee) = 600
 전략(Oee) = 950
 물량(Oee) = 950
 수비력(Oee) = 600
 정찰(Oee) = 350
 센스(Oee) = 950
 컨트롤(Oee) = 950
ElseIf Oee = 721 Then
 이름(Oee) = "월광"
 랭크(Oee) = "Rare"
 Team(Oee) = "Mystar"
 OYear(Oee) = "<11>"
 종족(Oee) = 1
 공격력(Oee) = 700
 견제(Oee) = 650
 전략(Oee) = 700
 물량(Oee) = 850
 수비력(Oee) = 800
 정찰(Oee) = 850
 센스(Oee) = 750
 컨트롤(Oee) = 600
ElseIf Oee = 722 Then
 이름(Oee) = "태양"
 랭크(Oee) = "Elite"
 Team(Oee) = "Mystar"
 OYear(Oee) = "<11>"
 종족(Oee) = 1
 공격력(Oee) = 850
 견제(Oee) = 700
 전략(Oee) = 750
 물량(Oee) = 950
 수비력(Oee) = 900
 정찰(Oee) = 800
 센스(Oee) = 950
 컨트롤(Oee) = 800
ElseIf Oee = 723 Then
 이름(Oee) = "코하이"
 랭크(Oee) = "Elite"
 Team(Oee) = "Mystar"
 OYear(Oee) = "<11>"
 종족(Oee) = 3
 공격력(Oee) = 950
 견제(Oee) = 750
 전략(Oee) = 750
 물량(Oee) = 950
 수비력(Oee) = 850
 정찰(Oee) = 850
 센스(Oee) = 800
 컨트롤(Oee) = 800
 
 End If

 우승(Oee) = 0
 준우승(Oee) = 0
 컨디션(Oee) = 100
 A승리(Oee) = 0
 A패배(Oee) = 0
 P승리(Oee) = 0
 P패배(Oee) = 0
 T승리(Oee) = 0
 T패배(Oee) = 0
 Z승리(Oee) = 0
 Z패배(Oee) = 0
 T연승(Oee) = 0
 Z연승(Oee) = 0
 P연승(Oee) = 0
 A연승(Oee) = 0
 T연(Oee) = "W"
 Z연(Oee) = "W"
 P연(Oee) = "W"
 A연(Oee) = "W"
Next
Tim12.Enabled = True
TimAdd.Enabled = False
End Sub

Private Sub TimElse_Timer()
For Oee = 540 To 575
If Oee = 540 Then
이름(Oee) = "임요환"
랭크(Oee) = "Champion"
OYear(Oee) = "<01>"
Team(Oee) = "IS"
종족(Oee) = 1
공격력(Oee) = 1000
견제(Oee) = 900
전략(Oee) = 900
물량(Oee) = 850
수비력(Oee) = 850
정찰(Oee) = 900
센스(Oee) = 900
컨트롤(Oee) = 1100
Skill(Oee) = 1


ElseIf Oee = 541 Then
이름(Oee) = "임요환"
랭크(Oee) = "Rare"
OYear(Oee) = "<02>"
Team(Oee) = "IS"
종족(Oee) = 1
공격력(Oee) = 850
견제(Oee) = 850
전략(Oee) = 800
물량(Oee) = 650
수비력(Oee) = 800
정찰(Oee) = 650
센스(Oee) = 750
컨트롤(Oee) = 850


ElseIf Oee = 542 Then
이름(Oee) = "임요환"
랭크(Oee) = "Rare"
OYear(Oee) = "<04>"
Team(Oee) = "4U"
종족(Oee) = 1
공격력(Oee) = 850
견제(Oee) = 850
전략(Oee) = 800
물량(Oee) = 500
수비력(Oee) = 650
정찰(Oee) = 600
센스(Oee) = 850
컨트롤(Oee) = 900


ElseIf Oee = 543 Then
이름(Oee) = "임요환"
랭크(Oee) = "Unique"
OYear(Oee) = "<05>"
Team(Oee) = "SK"
종족(Oee) = 1
공격력(Oee) = 850
견제(Oee) = 750
전략(Oee) = 850
물량(Oee) = 800
수비력(Oee) = 750
정찰(Oee) = 700
센스(Oee) = 850
컨트롤(Oee) = 850



ElseIf Oee = 544 Then
이름(Oee) = "홍진호"
랭크(Oee) = "Secret"
OYear(Oee) = "<01>"
Team(Oee) = "IS"
종족(Oee) = 2
공격력(Oee) = 950
견제(Oee) = 850
전략(Oee) = 800
물량(Oee) = 950
수비력(Oee) = 850
정찰(Oee) = 800
센스(Oee) = 850
컨트롤(Oee) = 950
Skill(Oee) = 3

ElseIf Oee = 545 Then
이름(Oee) = "홍진호"
랭크(Oee) = "Unique"
OYear(Oee) = "<02>"
Team(Oee) = "IS"
종족(Oee) = 2
공격력(Oee) = 850
견제(Oee) = 800
전략(Oee) = 700
물량(Oee) = 800
수비력(Oee) = 800
정찰(Oee) = 850
센스(Oee) = 800
컨트롤(Oee) = 800


ElseIf Oee = 546 Then
이름(Oee) = "홍진호"
랭크(Oee) = "Unique"
OYear(Oee) = "<03>"
Team(Oee) = "KTF"
종족(Oee) = 2
공격력(Oee) = 950
견제(Oee) = 800
전략(Oee) = 600
물량(Oee) = 850
수비력(Oee) = 850
정찰(Oee) = 600
센스(Oee) = 800
컨트롤(Oee) = 950


ElseIf Oee = 547 Then
이름(Oee) = "이윤열"
랭크(Oee) = "Secret"
OYear(Oee) = "<02>"
Team(Oee) = "IS"
종족(Oee) = 1
공격력(Oee) = 900
견제(Oee) = 800
전략(Oee) = 800
물량(Oee) = 900
수비력(Oee) = 950
정찰(Oee) = 750
센스(Oee) = 950
컨트롤(Oee) = 950
Skill(Oee) = 5

ElseIf Oee = 548 Then
이름(Oee) = "이윤열"
랭크(Oee) = "Rare"
OYear(Oee) = "<03>"
Team(Oee) = "KTF"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 700
전략(Oee) = 650
물량(Oee) = 800
수비력(Oee) = 900
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 800


ElseIf Oee = 549 Then
이름(Oee) = "이윤열"
랭크(Oee) = "Elite"
OYear(Oee) = "<04>"
Team(Oee) = "Toona"
종족(Oee) = 1
공격력(Oee) = 850
견제(Oee) = 700
전략(Oee) = 700
물량(Oee) = 950
수비력(Oee) = 950
정찰(Oee) = 700
센스(Oee) = 950
컨트롤(Oee) = 850


ElseIf Oee = 550 Then
이름(Oee) = "이윤열"
랭크(Oee) = "Unique"
OYear(Oee) = "<06>"
Team(Oee) = "Pantech"
종족(Oee) = 1
공격력(Oee) = 850
견제(Oee) = 750
전략(Oee) = 800
물량(Oee) = 850
수비력(Oee) = 750
정찰(Oee) = 700
센스(Oee) = 850
컨트롤(Oee) = 850


ElseIf Oee = 551 Then
이름(Oee) = "마재윤"
랭크(Oee) = "Unique"
OYear(Oee) = "<05>"
Team(Oee) = "GO"
종족(Oee) = 2
공격력(Oee) = 900
견제(Oee) = 700
전략(Oee) = 750
물량(Oee) = 900
수비력(Oee) = 900
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 900


ElseIf Oee = 552 Then
이름(Oee) = "마재윤"
랭크(Oee) = "Legend"
OYear(Oee) = "<06>"
Team(Oee) = "CJ"
종족(Oee) = 2
공격력(Oee) = 850
견제(Oee) = 750
전략(Oee) = 800
물량(Oee) = 950
수비력(Oee) = 950
정찰(Oee) = 750
센스(Oee) = 800
컨트롤(Oee) = 950


ElseIf Oee = 553 Then
이름(Oee) = "최연성"
랭크(Oee) = "Secret"
OYear(Oee) = "<03>"
Team(Oee) = "Orion"
종족(Oee) = 1
공격력(Oee) = 950
견제(Oee) = 800
전략(Oee) = 800
물량(Oee) = 950
수비력(Oee) = 950
정찰(Oee) = 800
센스(Oee) = 850
컨트롤(Oee) = 900
Skill(Oee) = 7

ElseIf Oee = 554 Then
이름(Oee) = "최연성"
랭크(Oee) = "Elite"
OYear(Oee) = "<04>"
Team(Oee) = "4U"
종족(Oee) = 1
공격력(Oee) = 850
견제(Oee) = 700
전략(Oee) = 800
물량(Oee) = 950
수비력(Oee) = 950
정찰(Oee) = 950
센스(Oee) = 650
컨트롤(Oee) = 750


ElseIf Oee = 555 Then
이름(Oee) = "최연성"
랭크(Oee) = "Unique"
OYear(Oee) = "<05>"
Team(Oee) = "SK"
종족(Oee) = 1
공격력(Oee) = 900
견제(Oee) = 700
전략(Oee) = 850
물량(Oee) = 950
수비력(Oee) = 900
정찰(Oee) = 650
센스(Oee) = 750
컨트롤(Oee) = 750

ElseIf Oee = 556 Then
이름(Oee) = "박태민"
랭크(Oee) = "Rare"
OYear(Oee) = "<03>"
Team(Oee) = "GO"
종족(Oee) = 2
공격력(Oee) = 700
견제(Oee) = 650
전략(Oee) = 700
물량(Oee) = 950
수비력(Oee) = 900
정찰(Oee) = 700
센스(Oee) = 700
컨트롤(Oee) = 700

ElseIf Oee = 557 Then
이름(Oee) = "박태민"
랭크(Oee) = "Unique"
OYear(Oee) = "<04>"
Team(Oee) = "GO"
종족(Oee) = 2
공격력(Oee) = 850
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 950
수비력(Oee) = 950
정찰(Oee) = 750
센스(Oee) = 750
컨트롤(Oee) = 850

ElseIf Oee = 558 Then
이름(Oee) = "김동수"
랭크(Oee) = "Unique"
OYear(Oee) = "<01>"
Team(Oee) = "한빛"
종족(Oee) = 3
공격력(Oee) = 800
견제(Oee) = 750
전략(Oee) = 650
물량(Oee) = 950
수비력(Oee) = 850
정찰(Oee) = 800
센스(Oee) = 800
컨트롤(Oee) = 800

ElseIf Oee = 559 Then
이름(Oee) = "변길섭"
랭크(Oee) = "Unique"
OYear(Oee) = "<02>"
Team(Oee) = "한빛"
종족(Oee) = 1
공격력(Oee) = 850
견제(Oee) = 800
전략(Oee) = 750
물량(Oee) = 800
수비력(Oee) = 750
정찰(Oee) = 700
센스(Oee) = 850
컨트롤(Oee) = 900

ElseIf Oee = 560 Then
이름(Oee) = "박정석"
랭크(Oee) = "Secret"
OYear(Oee) = "<02>"
Team(Oee) = "한빛"
종족(Oee) = 3
공격력(Oee) = 900
견제(Oee) = 800
전략(Oee) = 900
물량(Oee) = 950
수비력(Oee) = 900
정찰(Oee) = 800
센스(Oee) = 850
컨트롤(Oee) = 900
Skill(Oee) = 4

ElseIf Oee = 561 Then
이름(Oee) = "박정석"
랭크(Oee) = "Rare"
OYear(Oee) = "<04>"
Team(Oee) = "KTF"
종족(Oee) = 3
공격력(Oee) = 800
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 900
수비력(Oee) = 900
정찰(Oee) = 700
센스(Oee) = 750
컨트롤(Oee) = 700

ElseIf Oee = 562 Then
이름(Oee) = "강도경"
랭크(Oee) = "Rare"
OYear(Oee) = "<01>"
Team(Oee) = "한빛"
종족(Oee) = 2
공격력(Oee) = 950
견제(Oee) = 650
전략(Oee) = 600
물량(Oee) = 750
수비력(Oee) = 800
정찰(Oee) = 700
센스(Oee) = 700
컨트롤(Oee) = 850

ElseIf Oee = 563 Then
이름(Oee) = "조용호"
랭크(Oee) = "Rare"
OYear(Oee) = "<02>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 900
수비력(Oee) = 900
정찰(Oee) = 700
센스(Oee) = 700
컨트롤(Oee) = 750

ElseIf Oee = 564 Then
이름(Oee) = "조용호"
랭크(Oee) = "Unique"
OYear(Oee) = "<05>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 750
견제(Oee) = 700
전략(Oee) = 850
물량(Oee) = 950
수비력(Oee) = 900
정찰(Oee) = 700
센스(Oee) = 750
컨트롤(Oee) = 800

ElseIf Oee = 565 Then
이름(Oee) = "조용호"
랭크(Oee) = "Rare"
OYear(Oee) = "<06>"
Team(Oee) = "STX"
종족(Oee) = 2
공격력(Oee) = 900
견제(Oee) = 700
전략(Oee) = 550
물량(Oee) = 800
수비력(Oee) = 850
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 850

ElseIf Oee = 566 Then
이름(Oee) = "서지훈"
랭크(Oee) = "Unique"
OYear(Oee) = "<03>"
Team(Oee) = "GO"
종족(Oee) = 1
공격력(Oee) = 800
견제(Oee) = 900
전략(Oee) = 650
물량(Oee) = 950
수비력(Oee) = 950
정찰(Oee) = 800
센스(Oee) = 650
컨트롤(Oee) = 850

ElseIf Oee = 567 Then
이름(Oee) = "박용욱"
랭크(Oee) = "Rare"
OYear(Oee) = "<03>"
Team(Oee) = "Orion"
종족(Oee) = 3
공격력(Oee) = 750
견제(Oee) = 750
전략(Oee) = 600
물량(Oee) = 800
수비력(Oee) = 800
정찰(Oee) = 650
센스(Oee) = 900
컨트롤(Oee) = 750

ElseIf Oee = 568 Then
이름(Oee) = "강민"
랭크(Oee) = "Secret"
OYear(Oee) = "<03>"
Team(Oee) = "GO"
종족(Oee) = 3
공격력(Oee) = 850
견제(Oee) = 800
전략(Oee) = 950
물량(Oee) = 950
수비력(Oee) = 900
정찰(Oee) = 800
센스(Oee) = 950
컨트롤(Oee) = 800
Skill(Oee) = 30

ElseIf Oee = 569 Then
이름(Oee) = "박성준"
랭크(Oee) = "Legend"
OYear(Oee) = "<04>"
Team(Oee) = "POS"
종족(Oee) = 2
공격력(Oee) = 950
견제(Oee) = 950
전략(Oee) = 750
물량(Oee) = 900
수비력(Oee) = 800
정찰(Oee) = 650
센스(Oee) = 850
컨트롤(Oee) = 950

ElseIf Oee = 570 Then
이름(Oee) = "박성준"
랭크(Oee) = "Unique"
OYear(Oee) = "<05>"
Team(Oee) = "POS"
종족(Oee) = 2
공격력(Oee) = 900
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 950
수비력(Oee) = 950
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 950

ElseIf Oee = 571 Then
이름(Oee) = "박성준"
랭크(Oee) = "Rare"
OYear(Oee) = "<06>"
Team(Oee) = "MBC"
종족(Oee) = 2
공격력(Oee) = 850
견제(Oee) = 600
전략(Oee) = 550
물량(Oee) = 950
수비력(Oee) = 850
정찰(Oee) = 700
센스(Oee) = 650
컨트롤(Oee) = 850

ElseIf Oee = 572 Then
이름(Oee) = "이병민"
랭크(Oee) = "Rare"
OYear(Oee) = "<05>"
Team(Oee) = "Curitel"
종족(Oee) = 1
공격력(Oee) = 900
견제(Oee) = 650
전략(Oee) = 650
물량(Oee) = 650
수비력(Oee) = 700
정찰(Oee) = 650
센스(Oee) = 950
컨트롤(Oee) = 850

ElseIf Oee = 573 Then
이름(Oee) = "오영종"
랭크(Oee) = "Unique"
OYear(Oee) = "<04>"
Team(Oee) = "PLUS"
종족(Oee) = 3
공격력(Oee) = 900
견제(Oee) = 900
전략(Oee) = 850
물량(Oee) = 800
수비력(Oee) = 650
정찰(Oee) = 750
센스(Oee) = 750
컨트롤(Oee) = 800

ElseIf Oee = 574 Then
이름(Oee) = "한동욱"
랭크(Oee) = "Unique"
OYear(Oee) = "<06>"
Team(Oee) = "온게임넷"
종족(Oee) = 1
공격력(Oee) = 950
견제(Oee) = 950
전략(Oee) = 550
물량(Oee) = 500
수비력(Oee) = 600
정찰(Oee) = 950
센스(Oee) = 950
컨트롤(Oee) = 950

ElseIf Oee = 575 Then
이름(Oee) = "심소명"
랭크(Oee) = "Rare"
OYear(Oee) = "<06>"
Team(Oee) = "Pantech"
종족(Oee) = 2
공격력(Oee) = 800
견제(Oee) = 650
전략(Oee) = 950
물량(Oee) = 850
수비력(Oee) = 850
정찰(Oee) = 650
센스(Oee) = 700
컨트롤(Oee) = 750
End If

 우승(Oee) = 0
 준우승(Oee) = 0
 컨디션(Oee) = 100
 A승리(Oee) = 0
 A패배(Oee) = 0
 P승리(Oee) = 0
 P패배(Oee) = 0
 T승리(Oee) = 0
 T패배(Oee) = 0
 Z승리(Oee) = 0
 Z패배(Oee) = 0
 T연승(Oee) = 0
 Z연승(Oee) = 0
 P연승(Oee) = 0
 A연승(Oee) = 0
 T연(Oee) = "W"
 Z연(Oee) = "W"
 P연(Oee) = "W"
 A연(Oee) = "W"
 Next Oee
 TimElse.Enabled = False
 TimAdd.Enabled = True
End Sub

Private Sub Timer1_Timer()
For Oee = 0 To 800
 공격력(Oee) = val(공격력(Oee)) - 50
 견제(Oee) = val(견제(Oee)) - 50
 전략(Oee) = val(전략(Oee)) - 50
 물량(Oee) = val(물량(Oee)) - 50
 수비력(Oee) = val(수비력(Oee)) - 50
 정찰(Oee) = val(정찰(Oee)) - 50
 센스(Oee) = val(센스(Oee)) - 50
 컨트롤(Oee) = val(컨트롤(Oee)) - 50
Next

For Oee = 0 To 800
 NPC공격력(Oee) = 공격력(Oee)
 NPC견제(Oee) = 견제(Oee)
 NPC전략(Oee) = 전략(Oee)
 NPC물량(Oee) = 물량(Oee)
 NPC수비력(Oee) = 수비력(Oee)
 NPC정찰(Oee) = 정찰(Oee)
 NPC센스(Oee) = 센스(Oee)
 NPC컨트롤(Oee) = 컨트롤(Oee)
Next


 CmdA1.Enabled = True
 CmdA2.Enabled = True
 CmdA3.Enabled = True
 CmdB1.Enabled = True
 CmdB2.Enabled = True
 CmdB3.Enabled = True
 CmdC1.Enabled = True
 CmdC2.Enabled = True
 CmdC3.Enabled = True
 Timer4.Enabled = True
 jcbutton2.Enabled = True
 jcbutton1.Caption = "선택 완료"
 Label1 = "선택해주십시오."
 Timer2.Enabled = True
 Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Randomize Oee
End Sub

Private Sub Timer3_Timer()
If val(선택량) = 3 Then
 jcbutton1.Enabled = True
End If

Dim 마이 As Integer
For 마이 = 1 To 6
MyAW(마이) = 0
MyAL(마이) = 0
MyTW(마이) = 0
MyTL(마이) = 0
MyPW(마이) = 0
MyPL(마이) = 0
MyZW(마이) = 0
MyZL(마이) = 0
MyT연승(마이) = 0
MyZ연승(마이) = 0
MyP연승(마이) = 0
MyA연승(마이) = 0
MyT연(마이) = "W"
MyZ연(마이) = "W"
MyP연(마이) = "W"
MyA연(마이) = "W"
MySkill(마이) = 0
Turn = "OSL"
MyVic(마이) = 0
MySeVic(마이) = 0
행동력 = 0
Con = 100
MyExp(마이) = 0
MyMExp(마이) = 10
MyLev(마이) = 1
MyPoint(마이) = 0
Next 마이
End Sub

Private Sub Timer4_Timer()
Oee = Int((800 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
CR = Int((120 * Rnd) + 1)
End Sub

Private Sub Timer5_Timer()
For Oee = 0 To 800
 Skill(Oee) = "0"
Next
Timer5.Enabled = False
Tim07.Enabled = True
End Sub
