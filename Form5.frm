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
   StartUpPosition =   2  'ȭ�� ���
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "���� �Ϸ�"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      BackStyle       =   1  '�������� ����
      Height          =   855
      Left            =   0
      Top             =   5040
      Width           =   10815
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "��ø� ��ٷ� �ֽʽÿ�."
      BeginProperty Font 
         Name            =   "����"
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
      BackStyle       =   1  '�������� ����
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   10815
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  '�������� ����
      Height          =   855
      Left            =   0
      Top             =   5880
      Width           =   10815
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  '�������� ����
      Height          =   1575
      Index           =   6
      Left            =   7200
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  '�������� ����
      Height          =   1575
      Index           =   5
      Left            =   3600
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  '�������� ����
      Height          =   1575
      Index           =   4
      Left            =   0
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  '�������� ����
      Height          =   1575
      Index           =   3
      Left            =   7200
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  '�������� ����
      Height          =   1575
      Index           =   2
      Left            =   3600
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  '�������� ����
      Height          =   1575
      Index           =   1
      Left            =   0
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  '�������� ����
      Height          =   1575
      Index           =   0
      Left            =   7200
      Top             =   360
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '�������� ����
      Height          =   1575
      Left            =   3600
      Top             =   360
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '�������� ����
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
���÷� = val(���÷�) + 1
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
���÷� = val(���÷�) + 1
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
���÷� = val(���÷�) + 1
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
���÷� = val(���÷�) + 1
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
���÷� = val(���÷�) + 1
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
���÷� = val(���÷�) + 1
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
���÷� = val(���÷�) + 1
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
���÷� = val(���÷�) + 1
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
���÷� = val(���÷�) + 1
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
CodeName = InputBox("�ڵ��Է�")
If CodeName = "���ֿ���" Then
������ = 6
Oee = 540
MyName(1) = �̸�(Oee)
MyTribe(1) = 1
MyAt(1) = ���ݷ�(Oee)
MyR(1) = ����(Oee)
MySt(1) = ����(Oee)
MyAm(1) = ����(Oee)
MyDe(1) = �����(Oee)
MyPa(1) = ����(Oee)
MySe(1) = ����(Oee)
MyCo(1) = ��Ʈ��(Oee)
MyYear(1) = OYear(Oee)
MyRank(1) = ��ũ(Oee)
MyTeam(1) = Team(Oee)
MySkill(1) = Skill(Oee)
PlayNumber(1) = Oee

Oee = 93
MyName(2) = �̸�(Oee)
MyTribe(2) = 2
MyAt(2) = ���ݷ�(Oee)
MyR(2) = ����(Oee)
MySt(2) = ����(Oee)
MyAm(2) = ����(Oee)
MyDe(2) = �����(Oee)
MyPa(2) = ����(Oee)
MySe(2) = ����(Oee)
MyCo(2) = ��Ʈ��(Oee)
MyYear(2) = OYear(Oee)
MyRank(2) = ��ũ(Oee)
MyTeam(2) = Team(Oee)
MySkill(2) = Skill(Oee)
PlayNumber(2) = Oee

Oee = 113
MyName(3) = �̸�(Oee)
MyTribe(3) = 3
MyAt(3) = ���ݷ�(Oee)
MyR(3) = ����(Oee)
MySt(3) = ����(Oee)
MyAm(3) = ����(Oee)
MyDe(3) = �����(Oee)
MyPa(3) = ����(Oee)
MySe(3) = ����(Oee)
MyCo(3) = ��Ʈ��(Oee)
MyYear(3) = OYear(Oee)
MyRank(3) = ��ũ(Oee)
MyTeam(3) = Team(Oee)
MySkill(3) = Skill(Oee)
PlayNumber(3) = Oee

Oee = 114
MyName(4) = �̸�(Oee)
MyTribe(4) = 1
MyAt(4) = ���ݷ�(Oee)
MyR(4) = ����(Oee)
MySt(4) = ����(Oee)
MyAm(4) = ����(Oee)
MyDe(4) = �����(Oee)
MyPa(4) = ����(Oee)
MySe(4) = ����(Oee)
MyCo(4) = ��Ʈ��(Oee)
MyYear(4) = OYear(Oee)
MyRank(4) = ��ũ(Oee)
MyTeam(4) = Team(Oee)
MySkill(4) = Skill(Oee)
PlayNumber(4) = Oee

Oee = 175
MyName(5) = �̸�(Oee)
MyTribe(5) = 2
MyAt(5) = ���ݷ�(Oee)
MyR(5) = ����(Oee)
MySt(5) = ����(Oee)
MyAm(5) = ����(Oee)
MyDe(5) = �����(Oee)
MyPa(5) = ����(Oee)
MySe(5) = ����(Oee)
MyCo(5) = ��Ʈ��(Oee)
MyYear(5) = OYear(Oee)
MyRank(5) = ��ũ(Oee)
MyTeam(5) = Team(Oee)
MySkill(5) = Skill(Oee)
PlayNumber(5) = Oee

Oee = 320
MyName(6) = �̸�(Oee)
MyTribe(6) = 3
MyAt(6) = ���ݷ�(Oee)
MyR(6) = ����(Oee)
MySt(6) = ����(Oee)
MyAm(6) = ����(Oee)
MyDe(6) = �����(Oee)
MyPa(6) = ����(Oee)
MySe(6) = ����(Oee)
MyCo(6) = ��Ʈ��(Oee)
MyYear(6) = OYear(Oee)
MyRank(6) = ��ũ(Oee)
MyTeam(6) = Team(Oee)
MySkill(6) = Skill(Oee)
PlayNumber(6) = Oee

For ������ = 1 To 6
    If MyRank(������) = "Normal" Or MyRank(������) = "Special" Then
        MyNW(������) = "CB16"
    ElseIf MyRank(������) = "Rare" Then
        MyNW(������) = "CA1"
    ElseIf MyRank(������) = "Unique" Then
        MyNW(������) = "CA2"
    ElseIf MyRank(������) = "Elite" Then
        MyNW(������) = "CA3"
    Else
        MyNW(������) = "CS32"
    End If
Next ������

TeamName = InputBox("�г����� �Է��ϼ���.")
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
        NPC���ݷ�(Oee) = val(NPC���ݷ�(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC�����(Oee) = val(NPC�����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC��Ʈ��(Oee) = val(NPC��Ʈ��(Oee)) + 50
    Next
ElseIf Mode = "Hell" Then
    For Oee = 0 To 800
        ���ݷ�(Oee) = val(���ݷ�(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        �����(Oee) = val(�����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ��Ʈ��(Oee) = val(��Ʈ��(Oee)) + 50
    Next
End If
���� = 0
����Ƚ�� = 0
FrmMain.Show
Money = 5000
Unload Me
ElseIf CodeName = "SoulDeck" Then
������ = 6
Oee = 114
MyName(1) = �̸�(Oee)
MyTribe(1) = 1
MyAt(1) = ���ݷ�(Oee)
MyR(1) = ����(Oee)
MySt(1) = ����(Oee)
MyAm(1) = ����(Oee)
MyDe(1) = �����(Oee)
MyPa(1) = ����(Oee)
MySe(1) = ����(Oee)
MyCo(1) = ��Ʈ��(Oee)
MyYear(1) = OYear(Oee)
MyRank(1) = ��ũ(Oee)
MyTeam(1) = Team(Oee)
MySkill(1) = Skill(Oee)
PlayNumber(1) = Oee

Oee = 544
MyName(2) = �̸�(Oee)
MyTribe(2) = 2
MyAt(2) = ���ݷ�(Oee)
MyR(2) = ����(Oee)
MySt(2) = ����(Oee)
MyAm(2) = ����(Oee)
MyDe(2) = �����(Oee)
MyPa(2) = ����(Oee)
MySe(2) = ����(Oee)
MyCo(2) = ��Ʈ��(Oee)
MyYear(2) = OYear(Oee)
MyRank(2) = ��ũ(Oee)
MyTeam(2) = Team(Oee)
MySkill(2) = Skill(Oee)
PlayNumber(2) = Oee

Oee = 136
MyName(3) = �̸�(Oee)
MyTribe(3) = 3
MyAt(3) = ���ݷ�(Oee)
MyR(3) = ����(Oee)
MySt(3) = ����(Oee)
MyAm(3) = ����(Oee)
MyDe(3) = �����(Oee)
MyPa(3) = ����(Oee)
MySe(3) = ����(Oee)
MyCo(3) = ��Ʈ��(Oee)
MyYear(3) = OYear(Oee)
MyRank(3) = ��ũ(Oee)
MyTeam(3) = Team(Oee)
MySkill(3) = Skill(Oee)
PlayNumber(3) = Oee

Oee = 288
MyName(4) = �̸�(Oee)
MyTribe(4) = 3
MyAt(4) = ���ݷ�(Oee)
MyR(4) = ����(Oee)
MySt(4) = ����(Oee)
MyAm(4) = ����(Oee)
MyDe(4) = �����(Oee)
MyPa(4) = ����(Oee)
MySe(4) = ����(Oee)
MyCo(4) = ��Ʈ��(Oee)
MyYear(4) = OYear(Oee)
MyRank(4) = ��ũ(Oee)
MyTeam(4) = Team(Oee)
MySkill(4) = Skill(Oee)
PlayNumber(4) = Oee

Oee = 112
MyName(5) = �̸�(Oee)
MyTribe(5) = 2
MyAt(5) = ���ݷ�(Oee)
MyR(5) = ����(Oee)
MySt(5) = ����(Oee)
MyAm(5) = ����(Oee)
MyDe(5) = �����(Oee)
MyPa(5) = ����(Oee)
MySe(5) = ����(Oee)
MyCo(5) = ��Ʈ��(Oee)
MyYear(5) = OYear(Oee)
MyRank(5) = ��ũ(Oee)
MyTeam(5) = Team(Oee)
MySkill(5) = Skill(Oee)
PlayNumber(5) = Oee

Oee = 320
MyName(6) = �̸�(Oee)
MyTribe(6) = 3
MyAt(6) = ���ݷ�(Oee)
MyR(6) = ����(Oee)
MySt(6) = ����(Oee)
MyAm(6) = ����(Oee)
MyDe(6) = �����(Oee)
MyPa(6) = ����(Oee)
MySe(6) = ����(Oee)
MyCo(6) = ��Ʈ��(Oee)
MyYear(6) = OYear(Oee)
MyRank(6) = ��ũ(Oee)
MyTeam(6) = Team(Oee)
MySkill(6) = Skill(Oee)
PlayNumber(6) = Oee

For ������ = 1 To 6
    If MyRank(������) = "Normal" Or MyRank(������) = "Special" Then
        MyNW(������) = "CB16"
    ElseIf MyRank(������) = "Rare" Then
        MyNW(������) = "CA1"
    ElseIf MyRank(������) = "Unique" Then
        MyNW(������) = "CA2"
    ElseIf MyRank(������) = "Elite" Then
        MyNW(������) = "CA3"
    Else
        MyNW(������) = "CS32"
    End If
Next ������

TeamName = InputBox("�г����� �Է��ϼ���.")
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
        NPC���ݷ�(Oee) = val(NPC���ݷ�(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC�����(Oee) = val(NPC�����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC��Ʈ��(Oee) = val(NPC��Ʈ��(Oee)) + 50
    Next
ElseIf Mode = "Hell" Then
    For Oee = 0 To 800
        ���ݷ�(Oee) = val(���ݷ�(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        �����(Oee) = val(�����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ��Ʈ��(Oee) = val(��Ʈ��(Oee)) + 50
    Next
End If
���� = 0
����Ƚ�� = 0
FrmMain.Show
Money = 5000
Unload Me
End If

End Sub

Private Sub Command2_Click()
Dim CodeName As String
CodeName = InputBox("�ڵ��Է�")
If CodeName = "Crow" Then
Money = 10000
������ = 6
Oee = 714
MyName(1) = �̸�(Oee)
MyTribe(1) = 1
MyAt(1) = ���ݷ�(Oee)
MyR(1) = ����(Oee)
MySt(1) = ����(Oee)
MyAm(1) = ����(Oee)
MyDe(1) = �����(Oee)
MyPa(1) = ����(Oee)
MySe(1) = ����(Oee)
MyCo(1) = ��Ʈ��(Oee)
MyYear(1) = OYear(Oee)
MyRank(1) = ��ũ(Oee)
MyTeam(1) = Team(Oee)
MySkill(1) = Skill(Oee)
PlayNumber(1) = Oee

Oee = 80
MyName(2) = �̸�(Oee)
MyTribe(2) = 2
MyAt(2) = ���ݷ�(Oee)
MyR(2) = ����(Oee)
MySt(2) = ����(Oee)
MyAm(2) = ����(Oee)
MyDe(2) = �����(Oee)
MyPa(2) = ����(Oee)
MySe(2) = ����(Oee)
MyCo(2) = ��Ʈ��(Oee)
MyYear(2) = OYear(Oee)
MyRank(2) = ��ũ(Oee)
MyTeam(2) = Team(Oee)
MySkill(2) = Skill(Oee)
PlayNumber(2) = Oee

Oee = 136
MyName(3) = �̸�(Oee)
MyTribe(3) = 3
MyAt(3) = ���ݷ�(Oee)
MyR(3) = ����(Oee)
MySt(3) = ����(Oee)
MyAm(3) = ����(Oee)
MyDe(3) = �����(Oee)
MyPa(3) = ����(Oee)
MySe(3) = ����(Oee)
MyCo(3) = ��Ʈ��(Oee)
MyYear(3) = OYear(Oee)
MyRank(3) = ��ũ(Oee)
MyTeam(3) = Team(Oee)
MySkill(3) = Skill(Oee)
PlayNumber(3) = Oee

Oee = 659
MyName(4) = �̸�(Oee)
MyTribe(4) = 1
MyAt(4) = ���ݷ�(Oee)
MyR(4) = ����(Oee)
MySt(4) = ����(Oee)
MyAm(4) = ����(Oee)
MyDe(4) = �����(Oee)
MyPa(4) = ����(Oee)
MySe(4) = ����(Oee)
MyCo(4) = ��Ʈ��(Oee)
MyYear(4) = OYear(Oee)
MyRank(4) = ��ũ(Oee)
MyTeam(4) = Team(Oee)
MySkill(4) = Skill(Oee)
PlayNumber(4) = Oee

Oee = 660
MyName(5) = �̸�(Oee)
MyTribe(5) = 2
MyAt(5) = ���ݷ�(Oee)
MyR(5) = ����(Oee)
MySt(5) = ����(Oee)
MyAm(5) = ����(Oee)
MyDe(5) = �����(Oee)
MyPa(5) = ����(Oee)
MySe(5) = ����(Oee)
MyCo(5) = ��Ʈ��(Oee)
MyYear(5) = OYear(Oee)
MyRank(5) = ��ũ(Oee)
MyTeam(5) = Team(Oee)
MySkill(5) = Skill(Oee)
PlayNumber(5) = Oee

Oee = 288
MyName(6) = �̸�(Oee)
MyTribe(6) = 3
MyAt(6) = ���ݷ�(Oee)
MyR(6) = ����(Oee)
MySt(6) = ����(Oee)
MyAm(6) = ����(Oee)
MyDe(6) = �����(Oee)
MyPa(6) = ����(Oee)
MySe(6) = ����(Oee)
MyCo(6) = ��Ʈ��(Oee)
MyYear(6) = OYear(Oee)
MyRank(6) = ��ũ(Oee)
MyTeam(6) = Team(Oee)
MySkill(6) = Skill(Oee)
PlayNumber(6) = Oee

For ������ = 1 To 6
    If MyRank(������) = "Normal" Or MyRank(������) = "Special" Then
        MyNW(������) = "CB16"
    ElseIf MyRank(������) = "Rare" Then
        MyNW(������) = "CA1"
    ElseIf MyRank(������) = "Unique" Then
        MyNW(������) = "CA2"
    ElseIf MyRank(������) = "Elite" Then
        MyNW(������) = "CA3"
    Else
        MyNW(������) = "CS32"
    End If
Next ������

TeamName = InputBox("�г����� �Է��ϼ���.")
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
        NPC���ݷ�(Oee) = val(NPC���ݷ�(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC�����(Oee) = val(NPC�����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC��Ʈ��(Oee) = val(NPC��Ʈ��(Oee)) + 50
    Next
ElseIf Mode = "Hell" Then
    For Oee = 0 To 800
        ���ݷ�(Oee) = val(���ݷ�(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        �����(Oee) = val(�����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ��Ʈ��(Oee) = val(��Ʈ��(Oee)) + 50
    Next
End If
���� = 0
����Ƚ�� = 0
FrmMain.Show
Unload Me
ElseIf CodeName = "SecretDeck" Then
Money = 299792458
������ = 6
Oee = 649
MyName(1) = �̸�(Oee)
MyTribe(1) = 1
MyAt(1) = ���ݷ�(Oee)
MyR(1) = ����(Oee)
MySt(1) = ����(Oee)
MyAm(1) = ����(Oee)
MyDe(1) = �����(Oee)
MyPa(1) = ����(Oee)
MySe(1) = ����(Oee)
MyCo(1) = ��Ʈ��(Oee)
MyYear(1) = OYear(Oee)
MyRank(1) = ��ũ(Oee)
MyTeam(1) = Team(Oee)
MySkill(1) = Skill(Oee)
PlayNumber(1) = Oee

Oee = 544
MyName(2) = �̸�(Oee)
MyTribe(2) = 2
MyAt(2) = ���ݷ�(Oee)
MyR(2) = ����(Oee)
MySt(2) = ����(Oee)
MyAm(2) = ����(Oee)
MyDe(2) = �����(Oee)
MyPa(2) = ����(Oee)
MySe(2) = ����(Oee)
MyCo(2) = ��Ʈ��(Oee)
MyYear(2) = OYear(Oee)
MyRank(2) = ��ũ(Oee)
MyTeam(2) = Team(Oee)
MySkill(2) = Skill(Oee)
PlayNumber(2) = Oee

Oee = 560
MyName(3) = �̸�(Oee)
MyTribe(3) = 3
MyAt(3) = ���ݷ�(Oee)
MyR(3) = ����(Oee)
MySt(3) = ����(Oee)
MyAm(3) = ����(Oee)
MyDe(3) = �����(Oee)
MyPa(3) = ����(Oee)
MySe(3) = ����(Oee)
MyCo(3) = ��Ʈ��(Oee)
MyYear(3) = OYear(Oee)
MyRank(3) = ��ũ(Oee)
MyTeam(3) = Team(Oee)
MySkill(3) = Skill(Oee)
PlayNumber(3) = Oee

Oee = 540
MyName(4) = �̸�(Oee)
MyTribe(4) = 1
MyAt(4) = ���ݷ�(Oee)
MyR(4) = ����(Oee)
MySt(4) = ����(Oee)
MyAm(4) = ����(Oee)
MyDe(4) = �����(Oee)
MyPa(4) = ����(Oee)
MySe(4) = ����(Oee)
MyCo(4) = ��Ʈ��(Oee)
MyYear(4) = OYear(Oee)
MyRank(4) = ��ũ(Oee)
MyTeam(4) = Team(Oee)
MySkill(4) = Skill(Oee)
PlayNumber(4) = Oee

Oee = 547
MyName(5) = �̸�(Oee)
MyTribe(5) = 1
MyAt(5) = ���ݷ�(Oee)
MyR(5) = ����(Oee)
MySt(5) = ����(Oee)
MyAm(5) = ����(Oee)
MyDe(5) = �����(Oee)
MyPa(5) = ����(Oee)
MySe(5) = ����(Oee)
MyCo(5) = ��Ʈ��(Oee)
MyYear(5) = OYear(Oee)
MyRank(5) = ��ũ(Oee)
MyTeam(5) = Team(Oee)
MySkill(5) = Skill(Oee)
PlayNumber(5) = Oee

Oee = 553
MyName(6) = �̸�(Oee)
MyTribe(6) = 1
MyAt(6) = ���ݷ�(Oee)
MyR(6) = ����(Oee)
MySt(6) = ����(Oee)
MyAm(6) = ����(Oee)
MyDe(6) = �����(Oee)
MyPa(6) = ����(Oee)
MySe(6) = ����(Oee)
MyCo(6) = ��Ʈ��(Oee)
MyYear(6) = OYear(Oee)
MyRank(6) = ��ũ(Oee)
MyTeam(6) = Team(Oee)
MySkill(6) = Skill(Oee)
PlayNumber(6) = Oee

For ������ = 1 To 6
    If MyRank(������) = "Normal" Or MyRank(������) = "Special" Then
        MyNW(������) = "CB16"
    ElseIf MyRank(������) = "Rare" Then
        MyNW(������) = "CA1"
    ElseIf MyRank(������) = "Unique" Then
        MyNW(������) = "CA2"
    ElseIf MyRank(������) = "Elite" Then
        MyNW(������) = "CA3"
    Else
        MyNW(������) = "CS32"
    End If
Next ������

TeamName = InputBox("�г����� �Է��ϼ���.")
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
        NPC���ݷ�(Oee) = val(NPC���ݷ�(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC�����(Oee) = val(NPC�����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC��Ʈ��(Oee) = val(NPC��Ʈ��(Oee)) + 50
    Next
ElseIf Mode = "Hell" Then
    For Oee = 0 To 800
        ���ݷ�(Oee) = val(���ݷ�(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        �����(Oee) = val(�����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ��Ʈ��(Oee) = val(��Ʈ��(Oee)) + 50
    Next
End If
���� = 0
����Ƚ�� = 0
FrmMain.Show
Unload Me

ElseIf CodeName = "Mystar" Then
Money = 10000
������ = 6
Oee = 713
MyName(1) = �̸�(Oee)
MyTribe(1) = 1
MyAt(1) = ���ݷ�(Oee)
MyR(1) = ����(Oee)
MySt(1) = ����(Oee)
MyAm(1) = ����(Oee)
MyDe(1) = �����(Oee)
MyPa(1) = ����(Oee)
MySe(1) = ����(Oee)
MyCo(1) = ��Ʈ��(Oee)
MyYear(1) = OYear(Oee)
MyRank(1) = ��ũ(Oee)
MyTeam(1) = Team(Oee)
MySkill(1) = Skill(Oee)
PlayNumber(1) = Oee

Oee = 709
MyName(2) = �̸�(Oee)
MyTribe(2) = 2
MyAt(2) = ���ݷ�(Oee)
MyR(2) = ����(Oee)
MySt(2) = ����(Oee)
MyAm(2) = ����(Oee)
MyDe(2) = �����(Oee)
MyPa(2) = ����(Oee)
MySe(2) = ����(Oee)
MyCo(2) = ��Ʈ��(Oee)
MyYear(2) = OYear(Oee)
MyRank(2) = ��ũ(Oee)
MyTeam(2) = Team(Oee)
MySkill(2) = Skill(Oee)
PlayNumber(2) = Oee

Oee = 711
MyName(3) = �̸�(Oee)
MyTribe(3) = 3
MyAt(3) = ���ݷ�(Oee)
MyR(3) = ����(Oee)
MySt(3) = ����(Oee)
MyAm(3) = ����(Oee)
MyDe(3) = �����(Oee)
MyPa(3) = ����(Oee)
MySe(3) = ����(Oee)
MyCo(3) = ��Ʈ��(Oee)
MyYear(3) = OYear(Oee)
MyRank(3) = ��ũ(Oee)
MyTeam(3) = Team(Oee)
MySkill(3) = Skill(Oee)
PlayNumber(3) = Oee

Oee = 710
MyName(4) = �̸�(Oee)
MyTribe(4) = 1
MyAt(4) = ���ݷ�(Oee)
MyR(4) = ����(Oee)
MySt(4) = ����(Oee)
MyAm(4) = ����(Oee)
MyDe(4) = �����(Oee)
MyPa(4) = ����(Oee)
MySe(4) = ����(Oee)
MyCo(4) = ��Ʈ��(Oee)
MyYear(4) = OYear(Oee)
MyRank(4) = ��ũ(Oee)
MyTeam(4) = Team(Oee)
MySkill(4) = Skill(Oee)
PlayNumber(4) = Oee

Oee = 118
MyName(5) = �̸�(Oee)
MyTribe(5) = 1
MyAt(5) = ���ݷ�(Oee)
MyR(5) = ����(Oee)
MySt(5) = ����(Oee)
MyAm(5) = ����(Oee)
MyDe(5) = �����(Oee)
MyPa(5) = ����(Oee)
MySe(5) = ����(Oee)
MyCo(5) = ��Ʈ��(Oee)
MyYear(5) = OYear(Oee)
MyRank(5) = ��ũ(Oee)
MyTeam(5) = Team(Oee)
MySkill(5) = Skill(Oee)
PlayNumber(5) = Oee

Oee = 712
MyName(6) = �̸�(Oee)
MyTribe(6) = 1
MyAt(6) = ���ݷ�(Oee)
MyR(6) = ����(Oee)
MySt(6) = ����(Oee)
MyAm(6) = ����(Oee)
MyDe(6) = �����(Oee)
MyPa(6) = ����(Oee)
MySe(6) = ����(Oee)
MyCo(6) = ��Ʈ��(Oee)
MyYear(6) = OYear(Oee)
MyRank(6) = ��ũ(Oee)
MyTeam(6) = Team(Oee)
MySkill(6) = Skill(Oee)
PlayNumber(6) = Oee

For ������ = 1 To 6
    If MyRank(������) = "Normal" Or MyRank(������) = "Special" Then
        MyNW(������) = "CB16"
    ElseIf MyRank(������) = "Rare" Then
        MyNW(������) = "CA1"
    ElseIf MyRank(������) = "Unique" Then
        MyNW(������) = "CA2"
    ElseIf MyRank(������) = "Elite" Then
        MyNW(������) = "CA3"
    Else
        MyNW(������) = "CS32"
    End If
Next ������

TeamName = InputBox("�г����� �Է��ϼ���.")
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
        NPC���ݷ�(Oee) = val(NPC���ݷ�(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC�����(Oee) = val(NPC�����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC��Ʈ��(Oee) = val(NPC��Ʈ��(Oee)) + 50
    Next
ElseIf Mode = "Hell" Then
    For Oee = 0 To 800
        ���ݷ�(Oee) = val(���ݷ�(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        �����(Oee) = val(�����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ��Ʈ��(Oee) = val(��Ʈ��(Oee)) + 50
    Next
End If
���� = 0
����Ƚ�� = 0
FrmMain.Show
Unload Me


End If
End Sub


Private Sub Form_Load()
CmdHar_Click
For Map = 1 To 12
 If val(Map) = 1 Then
  MapName(Map) = "�۷�������"
  �����Ÿ�(Map) = 5
  �ڿ�(Map) = 5
  ���⵵(Map) = 5
  TZT(Map) = 65
  TZZ(Map) = 35
  ZPZ(Map) = 50
  ZPP(Map) = 50
  PTP(Map) = 60
  PTT(Map) = 40
 ElseIf val(Map) = 2 Then
  MapName(Map) = "�׿���Ʈ����"
  �����Ÿ�(Map) = 6
  �ڿ�(Map) = 6
  ���⵵(Map) = 4
  TZT(Map) = 55
  TZZ(Map) = 45
  ZPZ(Map) = 55
  ZPP(Map) = 45
  PTP(Map) = 45
  PTT(Map) = 55
 ElseIf val(Map) = 3 Then
  MapName(Map) = "�׿�������"
  �����Ÿ�(Map) = 8
  �ڿ�(Map) = 6
  ���⵵(Map) = 8
  TZT(Map) = 45
  TZZ(Map) = 55
  ZPZ(Map) = 50
  ZPP(Map) = 50
  PTP(Map) = 60
  PTT(Map) = 40
 ElseIf val(Map) = 4 Then
  MapName(Map) = "����"
  �����Ÿ�(Map) = 5
  �ڿ�(Map) = 5
  ���⵵(Map) = 5
  TZT(Map) = 60
  TZZ(Map) = 40
  ZPZ(Map) = 60
  ZPP(Map) = 40
  PTP(Map) = 60
  PTT(Map) = 40
 ElseIf val(Map) = 5 Then
  MapName(Map) = "�糪"
  �����Ÿ�(Map) = 5
  �ڿ�(Map) = 5
  ���⵵(Map) = 1
  TZT(Map) = 45
  TZZ(Map) = 55
  ZPZ(Map) = 55
  ZPP(Map) = 45
  PTP(Map) = 55
  PTT(Map) = 45
 ElseIf val(Map) = 6 Then
  MapName(Map) = "���¾�������"
  �����Ÿ�(Map) = 8
  �ڿ�(Map) = 5
  ���⵵(Map) = 5
  TZT(Map) = 65
  TZZ(Map) = 35
  ZPZ(Map) = 50
  ZPP(Map) = 50
  PTP(Map) = 45
  PTT(Map) = 55
 ElseIf val(Map) = 7 Then
  MapName(Map) = "�����Ǵɼ�"
  �����Ÿ�(Map) = 8
  �ڿ�(Map) = 6
  ���⵵(Map) = 9
  TZT(Map) = 40
  TZZ(Map) = 60
  ZPZ(Map) = 55
  ZPP(Map) = 45
  PTP(Map) = 60
  PTT(Map) = 40
 ElseIf val(Map) = 8 Then
  MapName(Map) = "��Ŷ�극��Ŀ"
  �����Ÿ�(Map) = 5
  �ڿ�(Map) = 5
  ���⵵(Map) = 5
  TZT(Map) = 50
  TZZ(Map) = 50
  ZPZ(Map) = 50
  ZPP(Map) = 50
  PTP(Map) = 50
  PTT(Map) = 50
 ElseIf val(Map) = 9 Then
  MapName(Map) = "���ͳ�Ƽ��"
  �����Ÿ�(Map) = 5
  �ڿ�(Map) = 5
  ���⵵(Map) = 9
  TZT(Map) = 45
  TZZ(Map) = 55
  ZPZ(Map) = 50
  ZPP(Map) = 50
  PTP(Map) = 50
  PTT(Map) = 50
 ElseIf val(Map) = 10 Then
  MapName(Map) = "��ȥ"
  �����Ÿ�(Map) = 6
  �ڿ�(Map) = 8
  ���⵵(Map) = 6
  TZT(Map) = 55
  TZZ(Map) = 45
  ZPZ(Map) = 45
  ZPP(Map) = 55
  PTP(Map) = 50
  PTT(Map) = 50
 ElseIf val(Map) = 11 Then
  MapName(Map) = "���̽�"
  �����Ÿ�(Map) = 3
  �ڿ�(Map) = 3
  ���⵵(Map) = 1
  TZT(Map) = 55
  TZZ(Map) = 45
  ZPZ(Map) = 60
  ZPP(Map) = 40
  PTP(Map) = 45
  PTT(Map) = 55
 ElseIf val(Map) = 12 Then
  MapName(Map) = "�н����δ�"
  �����Ÿ�(Map) = 4
  �ڿ�(Map) = 5
  ���⵵(Map) = 5
  TZT(Map) = 65
  TZZ(Map) = 35
  ZPZ(Map) = 65
  ZPP(Map) = 35
  PTP(Map) = 50
  PTT(Map) = 50
 End If
Next

ũ�ο���� = "No"
PL��� = 0
PL�ؿ�� = 0
�ҷ��� = False
jcbutton1.Caption = "��ø� ��ٷ��ּ���."
Dim ���� As Integer
For ���� = 1 To 9
 SubName(����) = ""
 SubTeam(����) = ""
 SubAt(����) = ""
 SubR(����) = ""
 SubSt(����) = ""
 SubAm(����) = ""
 SubDe(����) = ""
 SubPa(����) = ""
 SubSe(����) = ""
 SubCo(����) = ""
 SubLev(����) = 1
 SubExp(����) = 0
 SubMExp(����) = 50
 SubAW(����) = 0
 SubAL(����) = 0
 SubTW(����) = 0
 SubTL(����) = 0
 SubZW(����) = 0
 SubZL(����) = 0
 SubPW(����) = 0
 SubPL(����) = 0
 SubT����(����) = 0
 SubZ����(����) = 0
 SubP����(����) = 0
 SubA����(����) = 0
 SubT��(����) = "W"
 SubZ��(����) = "W"
 SubP��(����) = "W"
 SubA��(����) = "W"
 SubSkill(����) = 0
Next ����

For ���� = 1 To 6
PL������(����) = True
Next ����

PLEnd = "False"
PL�ѹ� = 1
Money = 5000
  '****���� ����
 'F


'Secretũ�ο�
Oee = 0
 �̸�(Oee) = "ũ�ο�"
 OYear(Oee) = "<10>"
 ��ũ(Oee) = "Champion"
 Team(Oee) = "MyStar"
 ����(Oee) = 1
 ���ݷ�(Oee) = 950
 ����(Oee) = 950
 ����(Oee) = 900
 ����(Oee) = 900
 �����(Oee) = 900
 ����(Oee) = 900
 ����(Oee) = 900
 ��Ʈ��(Oee) = 1000
 ���(Oee) = 0
 �ؿ��(Oee) = 0
 �����(Oee) = 100
 A�¸�(Oee) = 0
 A�й�(Oee) = 0
 P�¸�(Oee) = 0
 P�й�(Oee) = 0
 T�¸�(Oee) = 0
 T�й�(Oee) = 0
 Z�¸�(Oee) = 0
 Z�й�(Oee) = 0
 T����(Oee) = 0
 Z����(Oee) = 0
 P����(Oee) = 0
 A����(Oee) = 0
 T��(Oee) = "W"
 Z��(Oee) = "W"
 P��(Oee) = "W"
 A��(Oee) = "W"
Close #1
PL���� = 0
PL�� = 0
PL�� = 0
PL���� = "1R"
Text1.Text = "������"
End Sub


Private Sub jcbutton1_Click()
��÷��� = Int((6 * Rnd) + 1)
FrmChoice.Show
Unload Me
End Sub

Private Sub jcbutton2_Click()
Dim �׽����ڵ� As String
�׽����ڵ� = InputBox("Code")
If �׽����ڵ� = "tOdaY" Then
������ = 6
Randomize Oee

CR = Int((120 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until ��ũ(Oee) = "Legend"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
End If

Do Until (����(Oee) = 1)
CR = Int((120 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
CR = 51
If val(CR) = 2 Then
 Do Until ��ũ(Oee) = "Legend"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
End If

Loop
MyName(1) = �̸�(Oee)
MyTribe(1) = 1
MyAt(1) = ���ݷ�(Oee)
MyR(1) = ����(Oee)
MySt(1) = ����(Oee)
MyAm(1) = ����(Oee)
MyDe(1) = �����(Oee)
MyPa(1) = ����(Oee)
MySe(1) = ����(Oee)
MyCo(1) = ��Ʈ��(Oee)
MyYear(1) = OYear(Oee)
MyRank(1) = ��ũ(Oee)
MyTeam(1) = Team(Oee)
MySkill(1) = Skill(Oee)
PlayNumber(1) = Oee
Randomize Oee

CR = Int((120 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until ��ũ(Oee) = "Legend"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
End If

Do Until (����(Oee) = 2)
CR = Int((120 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
CR = 51
If val(CR) = 2 Then
 Do Until ��ũ(Oee) = "Legend"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
End If

Loop
MyName(2) = �̸�(Oee)
MyTribe(2) = 2
MyAt(2) = ���ݷ�(Oee)
MyR(2) = ����(Oee)
MySt(2) = ����(Oee)
MyAm(2) = ����(Oee)
MyDe(2) = �����(Oee)
MyPa(2) = ����(Oee)
MySe(2) = ����(Oee)
MyCo(2) = ��Ʈ��(Oee)
MyYear(2) = OYear(Oee)
MyRank(2) = ��ũ(Oee)
MyTeam(2) = Team(Oee)
MySkill(2) = Skill(Oee)
PlayNumber(2) = Oee
Randomize Oee


Oee = 710
MyName(3) = �̸�(Oee)
MyTribe(3) = 3
MyAt(3) = ���ݷ�(Oee)
MyR(3) = ����(Oee)
MySt(3) = ����(Oee)
MyAm(3) = ����(Oee)
MyDe(3) = �����(Oee)
MyPa(3) = ����(Oee)
MySe(3) = ����(Oee)
MyCo(3) = ��Ʈ��(Oee)
MyYear(3) = OYear(Oee)
MyRank(3) = ��ũ(Oee)
MyTeam(3) = Team(Oee)
MySkill(3) = Skill(Oee)
PlayNumber(3) = Oee
Randomize Oee

CR = Int((120 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until ��ũ(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Do Until (����(Oee) = 1)
CR = Int((120 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
CR = 51
If val(CR) = 2 Then
 Do Until ��ũ(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Loop
MyName(4) = �̸�(Oee)
MyTribe(4) = 1
MyAt(4) = ���ݷ�(Oee)
MyR(4) = ����(Oee)
MySt(4) = ����(Oee)
MyAm(4) = ����(Oee)
MyDe(4) = �����(Oee)
MyPa(4) = ����(Oee)
MySe(4) = ����(Oee)
MyCo(4) = ��Ʈ��(Oee)
MyYear(4) = OYear(Oee)
MyRank(4) = ��ũ(Oee)
MyTeam(4) = Team(Oee)
MySkill(4) = Skill(Oee)
PlayNumber(4) = Oee
Randomize Oee

CR = Int((120 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until ��ũ(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Do Until (����(Oee) = 2)
CR = Int((120 * Rnd) + 1)
Oee = Int((800 * Rnd) + 1)
CR = 51
If val(CR) = 2 Then
 Do Until ��ũ(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Loop
MyName(5) = �̸�(Oee)
MyTribe(5) = 2
MyAt(5) = ���ݷ�(Oee)
MyR(5) = ����(Oee)
MySt(5) = ����(Oee)
MyAm(5) = ����(Oee)
MyDe(5) = �����(Oee)
MyPa(5) = ����(Oee)
MySe(5) = ����(Oee)
MyCo(5) = ��Ʈ��(Oee)
MyYear(5) = OYear(Oee)
MyRank(5) = ��ũ(Oee)
MyTeam(5) = Team(Oee)
MySkill(5) = Skill(Oee)
PlayNumber(5) = Oee
Randomize Oee

Oee = 710
MyName(6) = �̸�(Oee)
MyTribe(6) = 3
MyAt(6) = ���ݷ�(Oee)
MyR(6) = ����(Oee)
MySt(6) = ����(Oee)
MyAm(6) = ����(Oee)
MyDe(6) = �����(Oee)
MyPa(6) = ����(Oee)
MySe(6) = ����(Oee)
MyCo(6) = ��Ʈ��(Oee)
MyYear(6) = OYear(Oee)
MyRank(6) = ��ũ(Oee)
MyTeam(6) = Team(Oee)
MySkill(6) = Skill(Oee)
PlayNumber(6) = Oee
Randomize Oee


���� = 0
For ���� = 1 To 6
 If MyRank(����) = "Normal" Then
  ���� = val(����) + 1
 ElseIf MyRank(����) = "Special" Then
  ���� = val(����) + 2
 ElseIf MyRank(����) = "Rare" Then
  ���� = val(����) + 3
 ElseIf MyRank(����) = "Unique" Then
  ���� = val(����) + 4
 ElseIf MyRank(����) = "Elite" Then
  ���� = val(����) + 5
 ElseIf MyRank(����) = "Legend" Then
  ���� = val(����) + 6
 ElseIf MyRank(����) = "Secret" Then
  ���� = val(����) + 7
 End If
Next

If val(����) = 6 Then
 Money = 25000
ElseIf (val(����) >= 7) And (val(����) <= 12) Then
 Money = 20000
ElseIf (val(����) >= 13) And (val(����) <= 18) Then
 Money = 15000
ElseIf (val(����) >= 19) And (val(����) <= 24) Then
 Money = 10000
ElseIf (val(����) >= 25) And (val(����) <= 30) Then
 Money = 5000
ElseIf (val(����) >= 31) And (val(����) <= 36) Then
 Money = 2000
ElseIf (val(����) >= 37) And (val(����) <= 42) Then
 Money = 1000
End If
Ȯ�ο�1 = val(Money) / 2

For ������ = 1 To 6
    If MyRank(������) = "Normal" Or MyRank(������) = "Special" Then
        MyNW(������) = "CB16"
    ElseIf MyRank(������) = "Rare" Then
        MyNW(������) = "CA1"
    ElseIf MyRank(������) = "Unique" Then
        MyNW(������) = "CA2"
    ElseIf MyRank(������) = "Elite" Then
        MyNW(������) = "CA3"
    Else
        MyNW(������) = "CS32"
    End If
Next ������

TeamName = InputBox("�г����� �Է��ϼ���.")
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
        NPC���ݷ�(Oee) = val(NPC���ݷ�(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC�����(Oee) = val(NPC�����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC��Ʈ��(Oee) = val(NPC��Ʈ��(Oee)) + 50
    Next
ElseIf Mode = "Hell" Then
    For Oee = 0 To 800
        ���ݷ�(Oee) = val(���ݷ�(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        �����(Oee) = val(�����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ��Ʈ��(Oee) = val(��Ʈ��(Oee)) + 50
    Next
End If
���� = 0
����Ƚ�� = 0
FrmMain.Show
Unload Me
ElseIf �׽����ڵ� = "���϶�" Then
������ = 6
Randomize Oee
Oee = 697
MyName(1) = �̸�(Oee)
MyTribe(1) = 1
MyAt(1) = ���ݷ�(Oee)
MyR(1) = ����(Oee)
MySt(1) = ����(Oee)
MyAm(1) = ����(Oee)
MyDe(1) = �����(Oee)
MyPa(1) = ����(Oee)
MySe(1) = ����(Oee)
MyCo(1) = ��Ʈ��(Oee)
MyYear(1) = OYear(Oee)
MyRank(1) = ��ũ(Oee)
MyTeam(1) = Team(Oee)
MySkill(1) = Skill(Oee)
PlayNumber(1) = Oee
Randomize Oee
Oee = 799
MyName(2) = �̸�(Oee)
MyTribe(2) = 2
MyAt(2) = ���ݷ�(Oee)
MyR(2) = ����(Oee)
MySt(2) = ����(Oee)
MyAm(2) = ����(Oee)
MyDe(2) = �����(Oee)
MyPa(2) = ����(Oee)
MySe(2) = ����(Oee)
MyCo(2) = ��Ʈ��(Oee)
MyYear(2) = OYear(Oee)
MyRank(2) = ��ũ(Oee)
MyTeam(2) = Team(Oee)
MySkill(2) = Skill(Oee)
PlayNumber(2) = Oee
Randomize Oee


Oee = 716
MyName(3) = �̸�(Oee)
MyTribe(3) = 3
MyAt(3) = ���ݷ�(Oee)
MyR(3) = ����(Oee)
MySt(3) = ����(Oee)
MyAm(3) = ����(Oee)
MyDe(3) = �����(Oee)
MyPa(3) = ����(Oee)
MySe(3) = ����(Oee)
MyCo(3) = ��Ʈ��(Oee)
MyYear(3) = OYear(Oee)
MyRank(3) = ��ũ(Oee)
MyTeam(3) = Team(Oee)
MySkill(3) = Skill(Oee)
PlayNumber(3) = Oee
Randomize Oee
Oee = 20
MyName(4) = �̸�(Oee)
MyTribe(4) = 1
MyAt(4) = ���ݷ�(Oee)
MyR(4) = ����(Oee)
MySt(4) = ����(Oee)
MyAm(4) = ����(Oee)
MyDe(4) = �����(Oee)
MyPa(4) = ����(Oee)
MySe(4) = ����(Oee)
MyCo(4) = ��Ʈ��(Oee)
MyYear(4) = OYear(Oee)
MyRank(4) = ��ũ(Oee)
MyTeam(4) = Team(Oee)
MySkill(4) = Skill(Oee)
PlayNumber(4) = Oee
Randomize Oee
Oee = 93
MyName(5) = �̸�(Oee)
MyTribe(5) = 2
MyAt(5) = ���ݷ�(Oee)
MyR(5) = ����(Oee)
MySt(5) = ����(Oee)
MyAm(5) = ����(Oee)
MyDe(5) = �����(Oee)
MyPa(5) = ����(Oee)
MySe(5) = ����(Oee)
MyCo(5) = ��Ʈ��(Oee)
MyYear(5) = OYear(Oee)
MyRank(5) = ��ũ(Oee)
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
 Do Until ��ũ(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Do Until (����(Oee) = 3)
If Mode = "Normal" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hard" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hell" Then
    CR = Int((120 * Rnd) + 1)
End If
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until ��ũ(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Loop
MyName(6) = �̸�(Oee)
MyTribe(6) = 3
MyAt(6) = ���ݷ�(Oee)
MyR(6) = ����(Oee)
MySt(6) = ����(Oee)
MyAm(6) = ����(Oee)
MyDe(6) = �����(Oee)
MyPa(6) = ����(Oee)
MySe(6) = ����(Oee)
MyCo(6) = ��Ʈ��(Oee)
MyYear(6) = OYear(Oee)
MyRank(6) = ��ũ(Oee)
MyTeam(6) = Team(Oee)
MySkill(6) = Skill(Oee)
PlayNumber(6) = Oee
Randomize Oee


���� = 0
For ���� = 1 To 6
 If MyRank(����) = "Normal" Then
  ���� = val(����) + 1
 ElseIf MyRank(����) = "Special" Then
  ���� = val(����) + 2
 ElseIf MyRank(����) = "Rare" Then
  ���� = val(����) + 3
 ElseIf MyRank(����) = "Unique" Then
  ���� = val(����) + 4
 ElseIf MyRank(����) = "Elite" Then
  ���� = val(����) + 5
 ElseIf MyRank(����) = "Legend" Then
  ���� = val(����) + 6
 ElseIf MyRank(����) = "Secret" Then
  ���� = val(����) + 7
 End If
Next

If val(����) = 6 Then
 Money = 25000
ElseIf (val(����) >= 7) And (val(����) <= 12) Then
 Money = 20000
ElseIf (val(����) >= 13) And (val(����) <= 18) Then
 Money = 15000
ElseIf (val(����) >= 19) And (val(����) <= 24) Then
 Money = 10000
ElseIf (val(����) >= 25) And (val(����) <= 30) Then
 Money = 5000
ElseIf (val(����) >= 31) And (val(����) <= 36) Then
 Money = 2000
ElseIf (val(����) >= 37) And (val(����) <= 42) Then
 Money = 1000
End If
Ȯ�ο�1 = val(Money) / 2

For ������ = 1 To 6
    If MyRank(������) = "Normal" Or MyRank(������) = "Special" Then
        MyNW(������) = "CB16"
    ElseIf MyRank(������) = "Rare" Then
        MyNW(������) = "CA1"
    ElseIf MyRank(������) = "Unique" Then
        MyNW(������) = "CA2"
    ElseIf MyRank(������) = "Elite" Then
        MyNW(������) = "CA3"
    Else
        MyNW(������) = "CS32"
    End If
Next ������

TeamName = InputBox("�г����� �Է��ϼ���.")
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
        NPC���ݷ�(Oee) = val(NPC���ݷ�(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC�����(Oee) = val(NPC�����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC��Ʈ��(Oee) = val(NPC��Ʈ��(Oee)) + 50
    Next
ElseIf Mode = "Hell" Then
    For Oee = 0 To 800
        ���ݷ�(Oee) = val(���ݷ�(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        �����(Oee) = val(�����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ��Ʈ��(Oee) = val(��Ʈ��(Oee)) + 50
    Next
End If
���� = 0
����Ƚ�� = 0
FrmMain.Show
Unload Me
ElseIf �׽����ڵ� = "moonlight" Then
������ = 6
Randomize Oee
Oee = 153
MyName(1) = �̸�(Oee)
MyTribe(1) = 1
MyAt(1) = ���ݷ�(Oee)
MyR(1) = ����(Oee)
MySt(1) = ����(Oee)
MyAm(1) = ����(Oee)
MyDe(1) = �����(Oee)
MyPa(1) = ����(Oee)
MySe(1) = ����(Oee)
MyCo(1) = ��Ʈ��(Oee)
MyYear(1) = OYear(Oee)
MyRank(1) = ��ũ(Oee)
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
 Do Until ��ũ(Oee) = "Legend"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
End If

Do Until (����(Oee) = 2)
If Mode = "Normal" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hard" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hell" Then
    CR = Int((120 * Rnd) + 1)
End If
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until ��ũ(Oee) = "Legend"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
  Oee = Int((361 * Rnd) + 1) * 2
 Loop
End If

Loop
MyName(2) = �̸�(Oee)
MyTribe(2) = 2
MyAt(2) = ���ݷ�(Oee)
MyR(2) = ����(Oee)
MySt(2) = ����(Oee)
MyAm(2) = ����(Oee)
MyDe(2) = �����(Oee)
MyPa(2) = ����(Oee)
MySe(2) = ����(Oee)
MyCo(2) = ��Ʈ��(Oee)
MyYear(2) = OYear(Oee)
MyRank(2) = ��ũ(Oee)
MyTeam(2) = Team(Oee)
MySkill(2) = Skill(Oee)
PlayNumber(2) = Oee
Randomize Oee

Oee = 311
MyName(3) = �̸�(Oee)
MyTribe(3) = 3
MyAt(3) = ���ݷ�(Oee)
MyR(3) = ����(Oee)
MySt(3) = ����(Oee)
MyAm(3) = ����(Oee)
MyDe(3) = �����(Oee)
MyPa(3) = ����(Oee)
MySe(3) = ����(Oee)
MyCo(3) = ��Ʈ��(Oee)
MyYear(3) = OYear(Oee)
MyRank(3) = ��ũ(Oee)
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
 Do Until ��ũ(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Do Until (����(Oee) = 1)
If Mode = "Normal" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hard" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hell" Then
    CR = Int((120 * Rnd) + 1)
End If
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until ��ũ(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Loop
MyName(4) = �̸�(Oee)
MyTribe(4) = 1
MyAt(4) = ���ݷ�(Oee)
MyR(4) = ����(Oee)
MySt(4) = ����(Oee)
MyAm(4) = ����(Oee)
MyDe(4) = �����(Oee)
MyPa(4) = ����(Oee)
MySe(4) = ����(Oee)
MyCo(4) = ��Ʈ��(Oee)
MyYear(4) = OYear(Oee)
MyRank(4) = ��ũ(Oee)
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
 Do Until ��ũ(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Do Until (����(Oee) = 2)
If Mode = "Normal" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hard" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hell" Then
    CR = Int((120 * Rnd) + 1)
End If
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until ��ũ(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Loop
MyName(5) = �̸�(Oee)
MyTribe(5) = 2
MyAt(5) = ���ݷ�(Oee)
MyR(5) = ����(Oee)
MySt(5) = ����(Oee)
MyAm(5) = ����(Oee)
MyDe(5) = �����(Oee)
MyPa(5) = ����(Oee)
MySe(5) = ����(Oee)
MyCo(5) = ��Ʈ��(Oee)
MyYear(5) = OYear(Oee)
MyRank(5) = ��ũ(Oee)
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
 Do Until ��ũ(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Do Until (����(Oee) = 3)
If Mode = "Normal" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hard" Then
    CR = Int((118 * Rnd) + 3)
ElseIf Mode = "Hell" Then
    CR = Int((120 * Rnd) + 1)
End If
Oee = Int((800 * Rnd) + 1)
If val(CR) = 2 Then
 Do Until ��ũ(Oee) = "Legend"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) = 3 Or val(CR) = 4 Then
 Do Until ��ũ(Oee) = "Elite"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 5 And val(CR) <= 10 Then
 Do Until ��ũ(Oee) = "Unique"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 11 And val(CR) <= 20 Then
 Do Until ��ũ(Oee) = "Rare"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
ElseIf val(CR) >= 21 And val(CR) <= 50 Then
 Do Until ��ũ(Oee) = "Special"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
Else
 Do Until ��ũ(Oee) = "Normal"
 Oee = (Int((362 * Rnd) + 1) * 2) - 1
 Loop
End If

Loop
MyName(6) = �̸�(Oee)
MyTribe(6) = 3
MyAt(6) = ���ݷ�(Oee)
MyR(6) = ����(Oee)
MySt(6) = ����(Oee)
MyAm(6) = ����(Oee)
MyDe(6) = �����(Oee)
MyPa(6) = ����(Oee)
MySe(6) = ����(Oee)
MyCo(6) = ��Ʈ��(Oee)
MyYear(6) = OYear(Oee)
MyRank(6) = ��ũ(Oee)
MyTeam(6) = Team(Oee)
MySkill(6) = Skill(Oee)
PlayNumber(6) = Oee
Randomize Oee


���� = 0
For ���� = 1 To 6
 If MyRank(����) = "Normal" Then
  ���� = val(����) + 1
 ElseIf MyRank(����) = "Special" Then
  ���� = val(����) + 2
 ElseIf MyRank(����) = "Rare" Then
  ���� = val(����) + 3
 ElseIf MyRank(����) = "Unique" Then
  ���� = val(����) + 4
 ElseIf MyRank(����) = "Elite" Then
  ���� = val(����) + 5
 ElseIf MyRank(����) = "Legend" Then
  ���� = val(����) + 6
 ElseIf MyRank(����) = "Secret" Then
  ���� = val(����) + 7
 End If
Next

If val(����) = 6 Then
 Money = 25000
ElseIf (val(����) >= 7) And (val(����) <= 12) Then
 Money = 20000
ElseIf (val(����) >= 13) And (val(����) <= 18) Then
 Money = 15000
ElseIf (val(����) >= 19) And (val(����) <= 24) Then
 Money = 10000
ElseIf (val(����) >= 25) And (val(����) <= 30) Then
 Money = 5000
ElseIf (val(����) >= 31) And (val(����) <= 36) Then
 Money = 2000
ElseIf (val(����) >= 37) And (val(����) <= 42) Then
 Money = 1000
End If
Ȯ�ο�1 = val(Money) / 2

For ������ = 1 To 6
    If MyRank(������) = "Normal" Or MyRank(������) = "Special" Then
        MyNW(������) = "CB16"
    ElseIf MyRank(������) = "Rare" Then
        MyNW(������) = "CA1"
    ElseIf MyRank(������) = "Unique" Then
        MyNW(������) = "CA2"
    ElseIf MyRank(������) = "Elite" Then
        MyNW(������) = "CA3"
    Else
        MyNW(������) = "CS32"
    End If
Next ������

TeamName = InputBox("�г����� �Է��ϼ���.")
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
        NPC���ݷ�(Oee) = val(NPC���ݷ�(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC�����(Oee) = val(NPC�����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC����(Oee) = val(NPC����(Oee)) + 50
        NPC��Ʈ��(Oee) = val(NPC��Ʈ��(Oee)) + 50
    Next
ElseIf Mode = "Hell" Then
    For Oee = 0 To 800
        ���ݷ�(Oee) = val(���ݷ�(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        �����(Oee) = val(�����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ����(Oee) = val(����(Oee)) + 50
        ��Ʈ��(Oee) = val(��Ʈ��(Oee)) + 50
    Next
End If
���� = 0
����Ƚ�� = 0
FrmMain.Show
Unload Me
Else

MsgBox "�ڵ����"
End If
End Sub

Private Sub Label1_Click()
Dim ��ü�ڵ� As String
��ü�ڵ� = InputBox("�ڵ��Է�")
If ��ü�ڵ� = "55886248" Then
 Command1.Enabled = True
 Command2.Enabled = True
 Command1.Visible = True
 Command2.Visible = True
ElseIf ��ü�ڵ� = "Data ������" Then
    Open App.Path & "\Data\Data������.txt" For Output As #1
        For i = 1 To 800
            Print #1, "����" & i & "��° ����"
            Print #1, "�̸�(" & i & ") = ū����ǥ" & �̸�(i) & "ū����ǥ"
            Print #1, "��ũ(" & i & ") = ū����ǥ" & ��ũ(i) & "ū����ǥ"
            Print #1, "OYear(" & i & ") = ū����ǥ" & OYear(i) & "ū����ǥ"
            Print #1, "Team(" & i & ") = ū����ǥ" & Team(i) & "ū����ǥ"
            Print #1, "����(" & i & ") = ū����ǥ" & ����(i) & "ū����ǥ"
            Print #1, "���ݷ�(" & i & ") = ū����ǥ" & ���ݷ�(i) & "ū����ǥ"
            Print #1, "����(" & i & ") = ū����ǥ" & ����(i) & "ū����ǥ"
            Print #1, "����(" & i & ") = ū����ǥ" & ����(i) & "ū����ǥ"
            Print #1, "����(" & i & ") = ū����ǥ" & ����(i) & "ū����ǥ"
            Print #1, "�����(" & i & ") = ū����ǥ" & �����(i) & "ū����ǥ"
            Print #1, "����(" & i & ") = ū����ǥ" & ����(i) & "ū����ǥ"
            Print #1, "����(" & i & ") = ū����ǥ" & ����(i) & "ū����ǥ"
            Print #1, "��Ʈ��(" & i & ") = ū����ǥ" & ��Ʈ��(i) & "ū����ǥ"
            Print #1, ""
        Next
    Close #1
    MsgBox "�Ϸ�"
ElseIf ��ü�ڵ� = "��üSetting" Then
    Open App.Path & "\��ü �ɷ�ġ.txt" For Output As #1
        For i = 1 To 800
            Print #1, "������ȣ : " & i
            Print #1, OYear(i) & �̸�(i)
            Print #1, ��ũ(i)
            Print #1, ���ݷ�(i)
            Print #1, ����(i)
            Print #1, ����(i)
            Print #1, ����(i)
            Print #1, �����(i)
            Print #1, ����(i)
            Print #1, ����(i)
            Print #1, ��Ʈ��(i)
            Print #1, "----------------"
            Print #1, "----------------"
        Next
    Close #1
ElseIf ��ü�ڵ� = "Setting" Then
    For i = 1 To 800
        ���(1, i) = ���ݷ�(i)
        ���(2, i) = ����(i)
        ���(3, i) = ����(i)
        ���(4, i) = ����(i)
        ���(5, i) = �����(i)
        ���(6, i) = ����(i)
        ���(7, i) = ����(i)
        ���(8, i) = ��Ʈ��(i)
        �̸����(i) = �̸�(i)
        ��ũ���(i) = ��ũ(i)
        �⵵���(i) = OYear(i)
    Next
    
    For M�켼 = 1 To 800
        For O�켼 = 1 To 8
            ����(M�켼) = ����(M�켼) + ���(O�켼, M�켼)
        Next
    Next
    
    For M�켼 = 1 To 1000
        For i = 1 To 799
            If ����(i) < ����(i + 1) Then
                ��赵��(4) = ����(i)
                ����(i) = ����(i + 1)
                ����(i + 1) = val(��赵��(4))
                
                ��赵��(1) = �̸����(i)
                �̸����(i) = �̸����(i + 1)
                �̸����(i + 1) = ��赵��(1)
                
                ��赵��(2) = ��ũ���(i)
                ��ũ���(i) = ��ũ���(i + 1)
                ��ũ���(i + 1) = ��赵��(2)
                
                ��赵��(3) = �⵵���(i)
                �⵵���(i) = �⵵���(i + 1)
                �⵵���(i + 1) = ��赵��(3)
            End If
        Next
    Next

    Open App.Path & "\�ɷ�ġ ���.txt" For Output As #1
        For i = 1 To 800
                Print #1, i & ". " & �⵵���(i) & �̸����(i) & " �� " & ��ũ���(i) & ", ���� : " & ����(i)
        Next
    Close #1

    MsgBox "�Ϸ�"
ElseIf ��ü�ڵ� = "Say" Then
    Open App.Path & "\�ɷ�ġ�߸��Ⱦֵ�.txt" For Output As #1
    '�ɷ�ġ Ȯ��
    ''Normal
    Print #1, "��Normal"
    For Oee = 1 To 800
        If ��ũ(Oee) = "Normal" Then
            Print #1, OYear(Oee) & �̸�(Oee) & " �� " & ��ũ(Oee)
        End If
    Next
    
    ''Special
    Print #1, "��Special"
    For Oee = 1 To 800
        If ��ũ(Oee) = "Special" Then
            Print #1, OYear(Oee) & �̸�(Oee) & " �� " & ��ũ(Oee)
        End If
    Next
    
    ''Rare
    Print #1, "��Rare"
    For Oee = 1 To 800
        If ��ũ(Oee) = "Rare" Then
            Print #1, OYear(Oee) & �̸�(Oee) & " �� " & ��ũ(Oee)
        End If
    Next
    
    ''Unique
    Print #1, "��Unique"
    For Oee = 1 To 800
        If ��ũ(Oee) = "Unique" Then
            Print #1, OYear(Oee) & �̸�(Oee) & " �� " & ��ũ(Oee)
        End If
    Next
    
    ''Elite
    Print #1, "��Elite"
    For Oee = 1 To 800
        If ��ũ(Oee) = "Elite" Then
            Print #1, OYear(Oee) & �̸�(Oee) & " �� " & ��ũ(Oee)
        End If
    Next
    
    ''Legend
    Print #1, "��Legend"
    For Oee = 1 To 800
        If ��ũ(Oee) = "Legend" Then
            Print #1, OYear(Oee) & �̸�(Oee) & " �� " & ��ũ(Oee)
        End If
    Next
    
    ''Secret
    Print #1, "��Secret"
    For Oee = 1 To 800
        If ��ũ(Oee) = "Secret" Then
            Print #1, OYear(Oee) & �̸�(Oee) & " �� " & ��ũ(Oee)
        End If
    Next
    
    ''Champion
    Print #1, "��Champion"
    For Oee = 1 To 800
        If ��ũ(Oee) = "Champion" Then
            Print #1, OYear(Oee) & �̸�(Oee) & " �� " & ��ũ(Oee)
        End If
    Next
    DoEvents
    MsgBox "�Ϸ�"
ElseIf ��ü�ڵ� = "���" Then
    '��ũ�� �ο� üũ & ��ũ�� ���� ��� üũ
    Dim k As Integer, sum As Integer
    sum = 0
    
    Open App.Path & "\���.txt" For Output As #1
    For i = 1 To 8
        Select Case i
        Case 1
            For k = 0 To 800
                If ��ũ(k) = "Normal" Then
                    sum = sum + 1
                End If
            Next
            Print #1, "��Normal : " & sum & "���"
            For k = 0 To 800
                If ��ũ(k) = "Normal" Then
                    Print #1, "<" & OYear(k) & ">  " & �̸�(k)
                End If
            Next
        Case 2
            For k = 0 To 800
                If ��ũ(k) = "Special" Then
                    sum = sum + 1
                End If
            Next
            Print #1, "��Special : " & sum & "���"
            For k = 0 To 800
                If ��ũ(k) = "Special" Then
                    Print #1, "<" & OYear(k) & ">  " & �̸�(k)
                End If
            Next
        Case 3
            For k = 0 To 800
                If ��ũ(k) = "Rare" Then
                    sum = sum + 1
                End If
            Next
            Print #1, "��Rare : " & sum & "���"
            For k = 0 To 800
                If ��ũ(k) = "Rare" Then
                    Print #1, "<" & OYear(k) & ">  " & �̸�(k)
                End If
            Next
        Case 4
            For k = 0 To 800
                If ��ũ(k) = "Unique" Then
                    sum = sum + 1
                End If
            Next
            Print #1, "��Unique : " & sum & "���"
            For k = 0 To 800
                If ��ũ(k) = "Unique" Then
                    Print #1, "<" & OYear(k) & ">  " & �̸�(k)
                End If
            Next
        Case 5
            For k = 0 To 800
                If ��ũ(k) = "Elite" Then
                    sum = sum + 1
                End If
            Next
            Print #1, "��Elite : " & sum & "���"
            For k = 0 To 800
                If ��ũ(k) = "Elite" Then
                    Print #1, "<" & OYear(k) & ">  " & �̸�(k)
                End If
            Next
        Case 6
            For k = 0 To 800
                If ��ũ(k) = "Legend" Then
                    sum = sum + 1
                End If
            Next
            Print #1, "��Legend : " & sum & "���"
            For k = 0 To 800
                If ��ũ(k) = "Legend" Then
                    Print #1, "<" & OYear(k) & ">  " & �̸�(k)
                End If
            Next
        Case 7
            For k = 0 To 800
                If ��ũ(k) = "Secret" Then
                    sum = sum + 1
                End If
            Next
            Print #1, "��Secret : " & sum & "���"
            For k = 0 To 800
                If ��ũ(k) = "Secret" Then
                    Print #1, "<" & OYear(k) & ">  " & �̸�(k)
                End If
            Next
        Case 8
            For k = 0 To 800
                If ��ũ(k) = "Champion" Then
                    sum = sum + 1
                End If
            Next
            Print #1, "��Champion : " & sum & "���"
            For k = 0 To 800
                If ��ũ(k) = "Champion" Then
                    Print #1, "<" & OYear(k) & ">  " & �̸�(k)
                End If
            Next
        End Select
    sum = 0
    Next
    MsgBox "�Ϸ�"
    Close #1
Else
 MsgBox "�ڵ� �����Դϴ�."
End If
End Sub

Private Sub Text1_Change()
Timer5.Enabled = True
End Sub


Private Sub Tim07_Timer()
For Oee = 119 To 260

If Oee = 119 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 700
����(Oee) = 750
����(Oee) = 600
�����(Oee) = 750
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 120 Then
�̸�(Oee) = "�赿��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 600
����(Oee) = 600
�����(Oee) = 400
����(Oee) = 450
����(Oee) = 450
��Ʈ��(Oee) = 500



ElseIf Oee = 121 Then
�̸�(Oee) = "����ȯ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 600
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 122 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 500
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 850



ElseIf Oee = 123 Then
�̸�(Oee) = "�躴��"
��ũ(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
����(Oee) = 2
���ݷ�(Oee) = 900
����(Oee) = 600
����(Oee) = 600
����(Oee) = 850
�����(Oee) = 850
����(Oee) = 700
����(Oee) = 700
��Ʈ��(Oee) = 800



ElseIf Oee = 124 Then
�̸�(Oee) = "���漷"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
����(Oee) = 1
���ݷ�(Oee) = 850
����(Oee) = 500
����(Oee) = 500
����(Oee) = 550
�����(Oee) = 500
����(Oee) = 400
����(Oee) = 400
��Ʈ��(Oee) = 650



ElseIf Oee = 125 Then
�̸�(Oee) = "�̺���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
����(Oee) = 1
���ݷ�(Oee) = 550
����(Oee) = 500
����(Oee) = 500
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 550



ElseIf Oee = 126 Then
�̸�(Oee) = "�̿�ȣ"
��ũ(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
����(Oee) = 1
���ݷ�(Oee) = 900
����(Oee) = 700
����(Oee) = 700
����(Oee) = 850
�����(Oee) = 750
����(Oee) = 600
����(Oee) = 750
��Ʈ��(Oee) = 750



ElseIf Oee = 127 Then
�̸�(Oee) = "�̿�ȣ1"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 700
����(Oee) = 750
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 700



ElseIf Oee = 128 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 600



ElseIf Oee = 129 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 600
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 750
��Ʈ��(Oee) = 750



ElseIf Oee = 130 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 500
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 500
��Ʈ��(Oee) = 600



ElseIf Oee = 131 Then
�̸�(Oee) = "ȫ��ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "KTF"
����(Oee) = 2
���ݷ�(Oee) = 850
����(Oee) = 600
����(Oee) = 600
����(Oee) = 500
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 700
��Ʈ��(Oee) = 600



ElseIf Oee = 132 Then
�̸�(Oee) = "�赿��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ｚ����"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 133 Then
�̸�(Oee) = "�ڼ���1"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 800
����(Oee) = 600
����(Oee) = 500
����(Oee) = 700
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 500
��Ʈ��(Oee) = 600



ElseIf Oee = 134 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ｚ����"
����(Oee) = 3
���ݷ�(Oee) = 550
����(Oee) = 650
����(Oee) = 800
����(Oee) = 500
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 600


ElseIf Oee = 135 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 800
����(Oee) = 600
����(Oee) = 450
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 500
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 136 Then
�̸�(Oee) = "�ۺ���"
��ũ(Oee) = "Legend"
OYear(Oee) = "<07>"
Team(Oee) = "�Ｚ����"
����(Oee) = 3
���ݷ�(Oee) = 900
����(Oee) = 750
����(Oee) = 800
����(Oee) = 950
�����(Oee) = 700
����(Oee) = 800
����(Oee) = 950
��Ʈ��(Oee) = 950



ElseIf Oee = 137 Then
�̸�(Oee) = "�̼���"
��ũ(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "�Ｚ����"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 800
����(Oee) = 850
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 650
����(Oee) = 850
��Ʈ��(Oee) = 700



ElseIf Oee = 138 Then
�̸�(Oee) = "����Ȳ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 139 Then
�̸�(Oee) = "��â��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 500
����(Oee) = 650
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 140 Then
�̸�(Oee) = "��ä��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ｚ����"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 650
����(Oee) = 650
����(Oee) = 600
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 700



ElseIf Oee = 141 Then
�̸�(Oee) = "��뼮"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ｚ����"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 650
����(Oee) = 700
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 750



ElseIf Oee = 142 Then
�̸�(Oee) = "�ֿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 800
����(Oee) = 600
����(Oee) = 600
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 800



ElseIf Oee = 142 Then
�̸�(Oee) = "�㿵��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ｚ����"
����(Oee) = 3
���ݷ�(Oee) = 850
����(Oee) = 600
����(Oee) = 600
����(Oee) = 850
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 143 Then
�̸�(Oee) = "�豸��"
��ũ(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
����(Oee) = 3
���ݷ�(Oee) = 800
����(Oee) = 900
����(Oee) = 600
����(Oee) = 900
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 800



ElseIf Oee = 144 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 500
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 600



ElseIf Oee = 145 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 500
����(Oee) = 600
����(Oee) = 450
��Ʈ��(Oee) = 550



ElseIf Oee = 146 Then
�̸�(Oee) = "����ȯ1"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 600
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 650



ElseIf Oee = 147 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 500
����(Oee) = 550
�����(Oee) = 700
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 500



ElseIf Oee = 148 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 500
����(Oee) = 800
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 149 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 700
����(Oee) = 650
����(Oee) = 600
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 150 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 400
����(Oee) = 350
�����(Oee) = 350
����(Oee) = 350
����(Oee) = 400
��Ʈ��(Oee) = 500



ElseIf Oee = 151 Then
�̸�(Oee) = "��ö��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 600
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 152 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 550
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 153 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Unique"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 850
����(Oee) = 950
����(Oee) = 700
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 900
����(Oee) = 850
��Ʈ��(Oee) = 900



ElseIf Oee = 154 Then
�̸�(Oee) = "�ֿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 700
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 450
��Ʈ��(Oee) = 550



ElseIf Oee = 155 Then
�̸�(Oee) = "�賲��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ѻ�"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 700
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 156 Then
�̸�(Oee) = "�赿��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ѻ�"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 157 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ѻ�"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 500
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 800



ElseIf Oee = 158 Then
�̸�(Oee) = "�躴��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ѻ�"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 550



ElseIf Oee = 159 Then
�̸�(Oee) = "���α�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ѻ�"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 650
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 650



ElseIf Oee = 160 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ѻ�"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 550
����(Oee) = 650
����(Oee) = 750
�����(Oee) = 550
����(Oee) = 500
����(Oee) = 550
��Ʈ��(Oee) = 700



ElseIf Oee = 161 Then
�̸�(Oee) = "���ؿ�"
��ũ(Oee) = "Unique"
OYear(Oee) = "<07>"
Team(Oee) = "�Ѻ�"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 700
����(Oee) = 700
����(Oee) = 950
�����(Oee) = 950
����(Oee) = 900
����(Oee) = 750
��Ʈ��(Oee) = 750



ElseIf Oee = 162 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ѻ�"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 600
����(Oee) = 600
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 550



ElseIf Oee = 163 Then
�̸�(Oee) = "�ڰ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ѻ�"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 850
����(Oee) = 700
����(Oee) = 500
�����(Oee) = 550
����(Oee) = 500
����(Oee) = 550
��Ʈ��(Oee) = 550



ElseIf Oee = 164 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ѻ�"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 550
����(Oee) = 700
����(Oee) = 650
�����(Oee) = 450
����(Oee) = 500
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 165 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ѻ�"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 450
����(Oee) = 800
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 550



ElseIf Oee = 166 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "�Ѻ�"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 750
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 750



ElseIf Oee = 167 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�Ѻ�"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 550
����(Oee) = 600
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 500
����(Oee) = 550
��Ʈ��(Oee) = 700



ElseIf Oee = 168 Then
�̸�(Oee) = "���α�"
��ũ(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 600
����(Oee) = 850
�����(Oee) = 750
����(Oee) = 700
����(Oee) = 650
��Ʈ��(Oee) = 750



ElseIf Oee = 169 Then
�̸�(Oee) = "�ǿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 500
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 500
��Ʈ��(Oee) = 550



ElseIf Oee = 170 Then
�̸�(Oee) = "�輺��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
����(Oee) = 3
���ݷ�(Oee) = 400
����(Oee) = 800
����(Oee) = 600
����(Oee) = 400
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 550
��Ʈ��(Oee) = 800



ElseIf Oee = 171 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
����(Oee) = 3
���ݷ�(Oee) = 950
����(Oee) = 750
����(Oee) = 500
����(Oee) = 950
�����(Oee) = 750
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 650



ElseIf Oee = 172 Then
�̸�(Oee) = "�ڴ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
����(Oee) = 3
���ݷ�(Oee) = 850
����(Oee) = 750
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 500
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 173 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
����(Oee) = 2
���ݷ�(Oee) = 800
����(Oee) = 900
����(Oee) = 600
����(Oee) = 800
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 800



ElseIf Oee = 174 Then
�̸�(Oee) = "�ڿ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 650
����(Oee) = 600
����(Oee) = 550
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 750



ElseIf Oee = 175 Then
�̸�(Oee) = "���¹�"
��ũ(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 600
����(Oee) = 900
�����(Oee) = 900
����(Oee) = 850
����(Oee) = 650
��Ʈ��(Oee) = 750



ElseIf Oee = 176 Then
�̸�(Oee) = "�ս���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 450
��Ʈ��(Oee) = 550



ElseIf Oee = 177 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 600
����(Oee) = 800
�����(Oee) = 800
����(Oee) = 650
����(Oee) = 500
��Ʈ��(Oee) = 700



ElseIf Oee = 178 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 450
����(Oee) = 750
�����(Oee) = 750
����(Oee) = 700
����(Oee) = 450
��Ʈ��(Oee) = 600



ElseIf Oee = 179 Then
�̸�(Oee) = "�̰���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 700



ElseIf Oee = 180 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 700
����(Oee) = 850
�����(Oee) = 900
����(Oee) = 850
����(Oee) = 600
��Ʈ��(Oee) = 700



ElseIf Oee = 181 Then
�̸�(Oee) = "�ֿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "SK"
����(Oee) = 1
���ݷ�(Oee) = 550
����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 750
����(Oee) = 650
����(Oee) = 750
��Ʈ��(Oee) = 750



ElseIf Oee = 182 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 183 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 850
����(Oee) = 600
����(Oee) = 800
����(Oee) = 800
�����(Oee) = 850
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 900



ElseIf Oee = 184 Then
�̸�(Oee) = "�赿��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 500
��Ʈ��(Oee) = 600



ElseIf Oee = 185 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 600
����(Oee) = 800
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 186 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 500
����(Oee) = 500
����(Oee) = 550
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 450
��Ʈ��(Oee) = 600



ElseIf Oee = 187 Then
�̸�(Oee) = "���ÿ�"
��ũ(Oee) = "Elite"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
����(Oee) = 3
���ݷ�(Oee) = 800
����(Oee) = 950
����(Oee) = 700
����(Oee) = 850
�����(Oee) = 700
����(Oee) = 950
����(Oee) = 950
��Ʈ��(Oee) = 800



ElseIf Oee = 188 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 550
����(Oee) = 550
�����(Oee) = 550
����(Oee) = 500
����(Oee) = 550
��Ʈ��(Oee) = 750



ElseIf Oee = 189 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
����(Oee) = 3
���ݷ�(Oee) = 900
����(Oee) = 700
����(Oee) = 650
����(Oee) = 900
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 600



ElseIf Oee = 190 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 700
����(Oee) = 500
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 600
��Ʈ��(Oee) = 800



ElseIf Oee = 191 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 800
����(Oee) = 600
����(Oee) = 800
�����(Oee) = 850
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 750



ElseIf Oee = 192 Then
�̸�(Oee) = "����ö"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 500
��Ʈ��(Oee) = 550



ElseIf Oee = 193 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "MBC"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 700
����(Oee) = 750
����(Oee) = 900
�����(Oee) = 850
����(Oee) = 700
����(Oee) = 700
��Ʈ��(Oee) = 650



ElseIf Oee = 194 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "������"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 650
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 195 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "������"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 500
��Ʈ��(Oee) = 600



ElseIf Oee = 196 Then
�̸�(Oee) = "�輺��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "������"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 500
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 550



ElseIf Oee = 197 Then
�̸�(Oee) = "����ȯ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "������"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 700
����(Oee) = 700
����(Oee) = 500
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 700



ElseIf Oee = 198 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "������"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 500
����(Oee) = 900
�����(Oee) = 900
����(Oee) = 800
����(Oee) = 700
��Ʈ��(Oee) = 600



ElseIf Oee = 199 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "������"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 850
����(Oee) = 850
����(Oee) = 900
�����(Oee) = 550
����(Oee) = 750
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 200 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "������"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 500
��Ʈ��(Oee) = 500



ElseIf Oee = 201 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Unique"
OYear(Oee) = "<07>"
Team(Oee) = "������"
����(Oee) = 2
���ݷ�(Oee) = 900
����(Oee) = 950
����(Oee) = 700
����(Oee) = 850
�����(Oee) = 650
����(Oee) = 750
����(Oee) = 800
��Ʈ��(Oee) = 950



ElseIf Oee = 202 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "������"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 500
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 550
��Ʈ��(Oee) = 500



ElseIf Oee = 203 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "������"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 700
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 600



ElseIf Oee = 204 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "������"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 800
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 700
����(Oee) = 600
��Ʈ��(Oee) = 700



ElseIf Oee = 205 Then
�̸�(Oee) = "�ְ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "������"
����(Oee) = 2
���ݷ�(Oee) = 900
����(Oee) = 700
����(Oee) = 600
����(Oee) = 400
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 750



ElseIf Oee = 206 Then
�̸�(Oee) = "�Ǽ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 207 Then
�̸�(Oee) = "���ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 550
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 450
��Ʈ��(Oee) = 600



ElseIf Oee = 208 Then
�̸�(Oee) = "�輺��"
��ũ(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
����(Oee) = 1
���ݷ�(Oee) = 900
����(Oee) = 600
����(Oee) = 550
����(Oee) = 900
�����(Oee) = 800
����(Oee) = 700
����(Oee) = 600
��Ʈ��(Oee) = 550



ElseIf Oee = 209 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Secret"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 850
����(Oee) = 850
����(Oee) = 850
����(Oee) = 950
�����(Oee) = 950
����(Oee) = 900
����(Oee) = 850
��Ʈ��(Oee) = 800
Skill(Oee) = 6


ElseIf Oee = 210 Then
�̸�(Oee) = "�ڿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 600
����(Oee) = 800
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 700
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 211 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
����(Oee) = 1
���ݷ�(Oee) = 950
����(Oee) = 800
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 800
��Ʈ��(Oee) = 800



ElseIf Oee = 212 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 700
����(Oee) = 750
��Ʈ��(Oee) = 800



ElseIf Oee = 213 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 550



ElseIf Oee = 214 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 700
����(Oee) = 700
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 700



ElseIf Oee = 215 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 500
����(Oee) = 450
��Ʈ��(Oee) = 600



ElseIf Oee = 216 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 550
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 217 Then
�̸�(Oee) = "�ѻ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 900
����(Oee) = 750
����(Oee) = 700
����(Oee) = 600
�����(Oee) = 500
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 750



ElseIf Oee = 218 Then
�̸�(Oee) = "�豤��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 2
���ݷ�(Oee) = 550
����(Oee) = 650
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 219 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 700
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 700
����(Oee) = 550
��Ʈ��(Oee) = 700



ElseIf Oee = 220 Then
�̸�(Oee) = "�Ż�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 221 Then
�̸�(Oee) = "�Ȼ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 550
��Ʈ��(Oee) = 550



ElseIf Oee = 222 Then
�̸�(Oee) = "�̽���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 800
����(Oee) = 700
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 800
��Ʈ��(Oee) = 700



ElseIf Oee = 223 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 2
���ݷ�(Oee) = 350
����(Oee) = 300
����(Oee) = 350
����(Oee) = 400
�����(Oee) = 400
����(Oee) = 400
����(Oee) = 350
��Ʈ��(Oee) = 350



ElseIf Oee = 224 Then
�̸�(Oee) = "�ӿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 700



ElseIf Oee = 225 Then
�̸�(Oee) = "���±�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 500
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 750
����(Oee) = 650
��Ʈ��(Oee) = 500



ElseIf Oee = 226 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 550
����(Oee) = 550
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 227 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
����(Oee) = 2
���ݷ�(Oee) = 550
����(Oee) = 500
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 550



ElseIf Oee = 228 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 500
��Ʈ��(Oee) = 500



ElseIf Oee = 229 Then
�̸�(Oee) = "��α�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 550
����(Oee) = 550
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 550
��Ʈ��(Oee) = 700



ElseIf Oee = 230 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 650
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 750



ElseIf Oee = 231 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 550
����(Oee) = 550
����(Oee) = 500
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 600



ElseIf Oee = 232 Then
�̸�(Oee) = "�ڹ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 750
����(Oee) = 500
����(Oee) = 800
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 750



ElseIf Oee = 233 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
����(Oee) = 3
���ݷ�(Oee) = 850
����(Oee) = 600
����(Oee) = 600
����(Oee) = 850
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 650



ElseIf Oee = 234 Then
�̸�(Oee) = "�Ż�ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 600
����(Oee) = 800
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 235 Then
�̸�(Oee) = "���뼺"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 600



ElseIf Oee = 236 Then
�̸�(Oee) = "�ֿ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "eSTRO"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 500
����(Oee) = 650
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 500
��Ʈ��(Oee) = 550



ElseIf Oee = 237 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 500
����(Oee) = 550
�����(Oee) = 450
����(Oee) = 450
����(Oee) = 500
��Ʈ��(Oee) = 500



ElseIf Oee = 238 Then
�̸�(Oee) = "�輱��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 450
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 500



ElseIf Oee = 239 Then
�̸�(Oee) = "��ȯ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 500
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 550



ElseIf Oee = 240 Then
�̸�(Oee) = "�ڴ븸"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 500
����(Oee) = 450
����(Oee) = 750
�����(Oee) = 500
����(Oee) = 600
����(Oee) = 500
��Ʈ��(Oee) = 550



ElseIf Oee = 241 Then
�̸�(Oee) = "���н�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 550
����(Oee) = 500
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 600



ElseIf Oee = 242 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 450
����(Oee) = 450
����(Oee) = 700
�����(Oee) = 500
����(Oee) = 600
����(Oee) = 450
��Ʈ��(Oee) = 600



ElseIf Oee = 243 Then
�̸�(Oee) = "���ֿ�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 600
����(Oee) = 900
�����(Oee) = 900
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 550



ElseIf Oee = 244 Then
�̸�(Oee) = "�ӿ�ȯ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 800
����(Oee) = 750
����(Oee) = 500
�����(Oee) = 550
����(Oee) = 700
����(Oee) = 900
��Ʈ��(Oee) = 650



ElseIf Oee = 245 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 700
����(Oee) = 500
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 550



ElseIf Oee = 246 Then
�̸�(Oee) = "���α�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 500
����(Oee) = 500
����(Oee) = 500
����(Oee) = 500
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 550



ElseIf Oee = 247 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�����̵�"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 550
����(Oee) = 550
����(Oee) = 500
�����(Oee) = 450
����(Oee) = 500
����(Oee) = 450
��Ʈ��(Oee) = 650



ElseIf Oee = 248 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Special"
OYear(Oee) = "<07>"
Team(Oee) = "�����̵�"
����(Oee) = 2
���ݷ�(Oee) = 800
����(Oee) = 600
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 850
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 850



ElseIf Oee = 249 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�����̵�"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 650
����(Oee) = 650
����(Oee) = 400
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 600
��Ʈ��(Oee) = 750



ElseIf Oee = 250 Then
�̸�(Oee) = "�ڿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�����̵�"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 600
����(Oee) = 600
����(Oee) = 450
�����(Oee) = 450
����(Oee) = 500
����(Oee) = 450
��Ʈ��(Oee) = 650



ElseIf Oee = 251 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�����̵�"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 800
����(Oee) = 650
����(Oee) = 550
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 252 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Unique"
OYear(Oee) = "<07>"
Team(Oee) = "�����̵�"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 650
����(Oee) = 650
����(Oee) = 900
�����(Oee) = 950
����(Oee) = 850
����(Oee) = 800
��Ʈ��(Oee) = 800



ElseIf Oee = 253 Then
�̸�(Oee) = "�տ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�����̵�"
����(Oee) = 3
���ݷ�(Oee) = 550
����(Oee) = 700
����(Oee) = 700
����(Oee) = 550
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 550
��Ʈ��(Oee) = 550



ElseIf Oee = 254 Then
�̸�(Oee) = "�ɼҸ�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�����̵�"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 550
����(Oee) = 800
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 550



ElseIf Oee = 255 Then
�̸�(Oee) = "�ȱ�ȿ"
��ũ(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "�����̵�"
����(Oee) = 3
���ݷ�(Oee) = 850
����(Oee) = 600
����(Oee) = 700
����(Oee) = 950
�����(Oee) = 800
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 900



ElseIf Oee = 256 Then
�̸�(Oee) = "�ӵ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�����̵�"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 700
����(Oee) = 500
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 450
��Ʈ��(Oee) = 550



ElseIf Oee = 257 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Rare"
OYear(Oee) = "<07>"
Team(Oee) = "�����̵�"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 750
����(Oee) = 850
�����(Oee) = 800
����(Oee) = 700
����(Oee) = 800
��Ʈ��(Oee) = 800



ElseIf Oee = 258 Then
�̸�(Oee) = "���¾�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�����̵�"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 500
����(Oee) = 700
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 450
��Ʈ��(Oee) = 550



ElseIf Oee = 259 Then
�̸�(Oee) = "�ѵ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�����̵�"
����(Oee) = 1
���ݷ�(Oee) = 850
����(Oee) = 850
����(Oee) = 600
����(Oee) = 450
�����(Oee) = 450
����(Oee) = 500
����(Oee) = 600
��Ʈ��(Oee) = 950



ElseIf Oee = 260 Then
�̸�(Oee) = "�ѵ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<07>"
Team(Oee) = "�����̵�"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 700
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 700
End If
 ���(Oee) = 0
 �ؿ��(Oee) = 0
 �����(Oee) = 100
 A�¸�(Oee) = 0
 A�й�(Oee) = 0
 P�¸�(Oee) = 0
 P�й�(Oee) = 0
 T�¸�(Oee) = 0
 T�й�(Oee) = 0
 Z�¸�(Oee) = 0
 Z�й�(Oee) = 0
 T����(Oee) = 0
 Z����(Oee) = 0
 P����(Oee) = 0
 A����(Oee) = 0
 T��(Oee) = "W"
 Z��(Oee) = "W"
 P��(Oee) = "W"
 A��(Oee) = "W"
Next Oee

Tim08.Enabled = True
Tim07.Enabled = False
End Sub

Private Sub Tim08_Timer()
For Oee = 261 To 407
If Oee = 261 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
����(Oee) = 2
���ݷ�(Oee) = 550
����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 550



ElseIf Oee = 262 Then
�̸�(Oee) = "��뿱"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 700
��Ʈ��(Oee) = 550



ElseIf Oee = 263 Then
�̸�(Oee) = "�迵��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 650
����(Oee) = 700
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 264 Then
�̸�(Oee) = "����ȯ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 650



ElseIf Oee = 265 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 650
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 700



ElseIf Oee = 266 Then
�̸�(Oee) = "���翵"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 700
����(Oee) = 900
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 650
��Ʈ��(Oee) = 550



ElseIf Oee = 267 Then
�̸�(Oee) = "���ؿ�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 700
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 500
����(Oee) = 450
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 268 Then
�̸�(Oee) = "�躴��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 600
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 269 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 650
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 270 Then
�̸�(Oee) = "�̿�ȣ"
��ũ(Oee) = "Unique"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 750
����(Oee) = 850
����(Oee) = 900
�����(Oee) = 900
����(Oee) = 750
����(Oee) = 900
��Ʈ��(Oee) = 650



ElseIf Oee = 271 Then
�̸�(Oee) = "�̿�ȣ1"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 272 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 600



ElseIf Oee = 273 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 450
����(Oee) = 650
����(Oee) = 500
��Ʈ��(Oee) = 500



ElseIf Oee = 274 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "KTF"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 275 Then
�̸�(Oee) = "�赿��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�Ｚ����"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 600



ElseIf Oee = 276 Then
�̸�(Oee) = "�ڵ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�Ｚ����"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 500
����(Oee) = 500
�����(Oee) = 550
����(Oee) = 500
����(Oee) = 450
��Ʈ��(Oee) = 700



ElseIf Oee = 277 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 550
����(Oee) = 600
����(Oee) = 750
����(Oee) = 600
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 278 Then
�̸�(Oee) = "�ռ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�Ｚ����"
����(Oee) = 3
���ݷ�(Oee) = 550
����(Oee) = 500
����(Oee) = 600
����(Oee) = 500
�����(Oee) = 650
����(Oee) = 500
����(Oee) = 600
��Ʈ��(Oee) = 700



ElseIf Oee = 279 Then
�̸�(Oee) = "�ۺ���"
��ũ(Oee) = "Elite"
OYear(Oee) = "<08>"
Team(Oee) = "�Ｚ����"
����(Oee) = 3
���ݷ�(Oee) = 900
����(Oee) = 850
����(Oee) = 850
����(Oee) = 850
�����(Oee) = 850
����(Oee) = 850
����(Oee) = 700
��Ʈ��(Oee) = 850



ElseIf Oee = 280 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 281 Then
�̸�(Oee) = "�̼���"
��ũ(Oee) = "Rare"
OYear(Oee) = "<08>"
Team(Oee) = "�Ｚ����"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 800
����(Oee) = 750
����(Oee) = 800
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 900
��Ʈ��(Oee) = 800



ElseIf Oee = 282 Then
�̸�(Oee) = "����Ȳ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 283 Then
�̸�(Oee) = "��ä��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�Ｚ����"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 550
����(Oee) = 550
�����(Oee) = 450
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 700



ElseIf Oee = 284 Then
�̸�(Oee) = "���±�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�Ｚ����"
����(Oee) = 3
���ݷ�(Oee) = 550
����(Oee) = 600
����(Oee) = 500
����(Oee) = 550
�����(Oee) = 450
����(Oee) = 450
����(Oee) = 450
��Ʈ��(Oee) = 550



ElseIf Oee = 285 Then
�̸�(Oee) = "�ֿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 700



ElseIf Oee = 286 Then
�̸�(Oee) = "����ȯ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 650
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 700
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 287 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�Ｚ����"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 500
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 650



ElseIf Oee = 288 Then
�̸�(Oee) = "�㿵��"
��ũ(Oee) = "Elite"
OYear(Oee) = "<08>"
Team(Oee) = "�Ｚ����"
����(Oee) = 3
���ݷ�(Oee) = 850
����(Oee) = 850
����(Oee) = 750
����(Oee) = 900
�����(Oee) = 700
����(Oee) = 750
����(Oee) = 800
��Ʈ��(Oee) = 950



ElseIf Oee = 289 Then
�̸�(Oee) = "���ȿ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 290 Then
�̸�(Oee) = "�豸��"
��ũ(Oee) = "Rare"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
����(Oee) = 3
���ݷ�(Oee) = 800
����(Oee) = 900
����(Oee) = 850
����(Oee) = 800
�����(Oee) = 650
����(Oee) = 750
����(Oee) = 750
��Ʈ��(Oee) = 800



ElseIf Oee = 291 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
��Ʈ��(Oee) = 600



ElseIf Oee = 292 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 500
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 293 Then
�̸�(Oee) = "����ȯ1"
��ũ(Oee) = "Rare"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 850
����(Oee) = 750
����(Oee) = 700
����(Oee) = 900
�����(Oee) = 750
����(Oee) = 650
����(Oee) = 750
��Ʈ��(Oee) = 650



ElseIf Oee = 294 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 550
����(Oee) = 550
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 650
��Ʈ��(Oee) = 700



ElseIf Oee = 295 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Unique"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 900
����(Oee) = 800
����(Oee) = 600
����(Oee) = 900
�����(Oee) = 800
����(Oee) = 650
����(Oee) = 850
��Ʈ��(Oee) = 900



ElseIf Oee = 296 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 297 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 500
����(Oee) = 450
�����(Oee) = 450
����(Oee) = 500
����(Oee) = 450
��Ʈ��(Oee) = 650



ElseIf Oee = 298 Then
�̸�(Oee) = "�̽���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 550



ElseIf Oee = 299 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 500



ElseIf Oee = 300 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 600



ElseIf Oee = 301 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 800
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 700
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 302 Then
�̸�(Oee) = "���α�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 500
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 750
����(Oee) = 750
��Ʈ��(Oee) = 800



ElseIf Oee = 303 Then
�̸�(Oee) = "�賲��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 304 Then
�̸�(Oee) = "�赿��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 305 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 700
����(Oee) = 650
����(Oee) = 750
�����(Oee) = 750
����(Oee) = 700
����(Oee) = 700
��Ʈ��(Oee) = 750



ElseIf Oee = 306 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 700

ElseIf Oee = 307 Then
�̸�(Oee) = "���α�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 650



ElseIf Oee = 308 Then
�̸�(Oee) = "���ؿ�"
��ũ(Oee) = "Special"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 600
����(Oee) = 850
�����(Oee) = 800
����(Oee) = 800
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 309 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 550
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 310 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 450
����(Oee) = 800
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 550



ElseIf Oee = 311 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Special"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 750
����(Oee) = 800
����(Oee) = 650
��Ʈ��(Oee) = 750



ElseIf Oee = 312 Then
�̸�(Oee) = "�̵���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 650
����(Oee) = 600
�����(Oee) = 700
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 550



ElseIf Oee = 313 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 500



ElseIf Oee = 314 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 650



ElseIf Oee = 315 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 650
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 316 Then
�̸�(Oee) = "���α�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 700
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 700
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 317 Then
�̸�(Oee) = "�ǿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 318 Then
�̸�(Oee) = "�輺��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
����(Oee) = 3
���ݷ�(Oee) = 500
����(Oee) = 800
����(Oee) = 600
����(Oee) = 500
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 750



ElseIf Oee = 319 Then
�̸�(Oee) = "���ÿ�"
��ũ(Oee) = "Unique"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 900
����(Oee) = 750
����(Oee) = 850
�����(Oee) = 750
����(Oee) = 850
����(Oee) = 900
��Ʈ��(Oee) = 850



ElseIf Oee = 320 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Unique"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
����(Oee) = 3
���ݷ�(Oee) = 950
����(Oee) = 700
����(Oee) = 700
����(Oee) = 950
�����(Oee) = 700
����(Oee) = 750
����(Oee) = 900
��Ʈ��(Oee) = 750



ElseIf Oee = 321 Then
�̸�(Oee) = "�ڴ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 600
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 322 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 323 Then
�̸�(Oee) = "���¹�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
����(Oee) = 2
���ݷ�(Oee) = 550
����(Oee) = 600
����(Oee) = 550
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 700
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 324 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
����(Oee) = 1
���ݷ�(Oee) = 550
����(Oee) = 550
����(Oee) = 500
����(Oee) = 750
�����(Oee) = 750
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 550



ElseIf Oee = 325 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 326 Then
�̸�(Oee) = "�̽¼�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
����(Oee) = 2
���ݷ�(Oee) = 550
����(Oee) = 600
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 327 Then
�̸�(Oee) = "�ӿ�ȯ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 800
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 328 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 329 Then
�̸�(Oee) = "����ö"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
����(Oee) = 2
���ݷ�(Oee) = 550
����(Oee) = 550
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 700



ElseIf Oee = 330 Then
�̸�(Oee) = "�ֿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
����(Oee) = 1
���ݷ�(Oee) = 500
����(Oee) = 600
����(Oee) = 550
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 331 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 600
����(Oee) = 750
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 450
����(Oee) = 700
��Ʈ��(Oee) = 600



ElseIf Oee = 332 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 800
����(Oee) = 750
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 333 Then
�̸�(Oee) = "�赿��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 650



ElseIf Oee = 334 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 600
����(Oee) = 800
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 600



ElseIf Oee = 335 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 500
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 450
��Ʈ��(Oee) = 650



ElseIf Oee = 336 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 750
����(Oee) = 700
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 800



ElseIf Oee = 337 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 500
����(Oee) = 700
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 600



ElseIf Oee = 338 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Special"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
����(Oee) = 3
���ݷ�(Oee) = 900
����(Oee) = 650
����(Oee) = 700
����(Oee) = 800
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 600



ElseIf Oee = 339 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 750
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 700



ElseIf Oee = 340 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Special"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 750
����(Oee) = 650
����(Oee) = 800
�����(Oee) = 800
����(Oee) = 750
����(Oee) = 700
��Ʈ��(Oee) = 750



ElseIf Oee = 341 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 700
����(Oee) = 750
����(Oee) = 600
�����(Oee) = 650
����(Oee) = 700
����(Oee) = 700
��Ʈ��(Oee) = 650



ElseIf Oee = 342 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "MBC"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 600



ElseIf Oee = 343 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Special"
OYear(Oee) = "<08>"
Team(Oee) = "������"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 750
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 344 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "������"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 345 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "������"
����(Oee) = 1
���ݷ�(Oee) = 500
����(Oee) = 500
����(Oee) = 500
����(Oee) = 500
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 500



ElseIf Oee = 346 Then
�̸�(Oee) = "����ȯ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "������"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 650
����(Oee) = 550
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 347 Then
�̸�(Oee) = "���±�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "������"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 550
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 500



ElseIf Oee = 348 Then
�̸�(Oee) = "�뿵��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "������"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 550
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 500
����(Oee) = 600
��Ʈ��(Oee) = 600


ElseIf Oee = 349 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Unique"
OYear(Oee) = "<08>"
Team(Oee) = "������"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 800
����(Oee) = 800
����(Oee) = 850
�����(Oee) = 900
����(Oee) = 700
����(Oee) = 800
��Ʈ��(Oee) = 750

 

ElseIf Oee = 350 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Special"
OYear(Oee) = "<08>"
Team(Oee) = "������"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 650
����(Oee) = 700
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 800



ElseIf Oee = 351 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "������"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 700
����(Oee) = 800
�����(Oee) = 500
����(Oee) = 700
����(Oee) = 800
��Ʈ��(Oee) = 650



ElseIf Oee = 352 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Unique"
OYear(Oee) = "<08>"
Team(Oee) = "������"
����(Oee) = 2
���ݷ�(Oee) = 900
����(Oee) = 850
����(Oee) = 750
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 900
��Ʈ��(Oee) = 900



ElseIf Oee = 353 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "������"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 354 Then
�̸�(Oee) = "Ȳ���ǿ�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "������"
����(Oee) = 2
���ݷ�(Oee) = 500
����(Oee) = 550
����(Oee) = 500
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 500
��Ʈ��(Oee) = 500



ElseIf Oee = 355 Then
�̸�(Oee) = "�Ǽ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 356 Then
�̸�(Oee) = "�豹��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 550
����(Oee) = 550
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 700



ElseIf Oee = 357 Then
�̸�(Oee) = "���ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 500
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 450
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 358 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Rare"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 750
����(Oee) = 700
����(Oee) = 800
�����(Oee) = 800
����(Oee) = 700
����(Oee) = 750
��Ʈ��(Oee) = 750



ElseIf Oee = 359 Then
�̸�(Oee) = "�ڿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 750
����(Oee) = 750
�����(Oee) = 650
����(Oee) = 700
����(Oee) = 700
��Ʈ��(Oee) = 750



ElseIf Oee = 360 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 800
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 361 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 362 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
��Ʈ��(Oee) = 500



ElseIf Oee = 363 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 650
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 364 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 700
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 365 Then
�̸�(Oee) = "����ȭ"
��ũ(Oee) = "Rare"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
����(Oee) = 3
���ݷ�(Oee) = 800
����(Oee) = 850
����(Oee) = 650
����(Oee) = 800
�����(Oee) = 850
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 366 Then
�̸�(Oee) = "�ѻ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 900
����(Oee) = 700
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 600
��Ʈ��(Oee) = 800



ElseIf Oee = 367 Then
�̸�(Oee) = "�豤��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 368 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 650
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 650
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 369 Then
�̸�(Oee) = "���м�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 550
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 550



ElseIf Oee = 370 Then
�̸�(Oee) = "�Ż�"
��ũ(Oee) = "Legend"
OYear(Oee) = "<08>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 1
���ݷ�(Oee) = 950
����(Oee) = 800
����(Oee) = 750
����(Oee) = 900
�����(Oee) = 750
����(Oee) = 750
����(Oee) = 950
��Ʈ��(Oee) = 950



ElseIf Oee = 371 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Rare"
OYear(Oee) = "<08>"
Team(Oee) = "SKT"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 800
����(Oee) = 800
����(Oee) = 850
�����(Oee) = 800
����(Oee) = 650
����(Oee) = 850
��Ʈ��(Oee) = 650



ElseIf Oee = 372 Then
�̸�(Oee) = "�Ȼ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 650
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 750
����(Oee) = 700
��Ʈ��(Oee) = 650



ElseIf Oee = 373 Then
�̸�(Oee) = "�̰��"
��ũ(Oee) = "Special"
OYear(Oee) = "<08>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 800
����(Oee) = 750
����(Oee) = 850
�����(Oee) = 750
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 374 Then
�̸�(Oee) = "�̽���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 750
����(Oee) = 700
����(Oee) = 700
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 375 Then
�̸�(Oee) = "�ӿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 700
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 700
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 376 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 650
��Ʈ��(Oee) = 700



ElseIf Oee = 377 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
����(Oee) = 2
���ݷ�(Oee) = 500
����(Oee) = 550
����(Oee) = 500
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 500
��Ʈ��(Oee) = 550



ElseIf Oee = 378 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 650
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 600



ElseIf Oee = 379 Then
�̸�(Oee) = "�ڹ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 750
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 380 Then
�̸�(Oee) = "�ڻ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 750
����(Oee) = 700
����(Oee) = 750
��Ʈ��(Oee) = 650



ElseIf Oee = 381 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 382 Then
�̸�(Oee) = "�Ŵ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 383 Then
�̸�(Oee) = "�Ż�ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
����(Oee) = 3
���ݷ�(Oee) = 850
����(Oee) = 700
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 650



ElseIf Oee = 384 Then
�̸�(Oee) = "�ȼ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 500
����(Oee) = 500
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 500



ElseIf Oee = 385 Then
�̸�(Oee) = "��ȣ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
����(Oee) = 1
���ݷ�(Oee) = 550
����(Oee) = 500
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 750
����(Oee) = 700
����(Oee) = 600
��Ʈ��(Oee) = 500



ElseIf Oee = 386 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "eSTRO"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 650
����(Oee) = 500
����(Oee) = 500
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 650



ElseIf Oee = 387 Then
�̸�(Oee) = "�輱��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 388 Then
�̸�(Oee) = "��ȯ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 389 Then
�̸�(Oee) = "�ڴ븸"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 700



ElseIf Oee = 389 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 750



ElseIf Oee = 390 Then
�̸�(Oee) = "���н�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 650
����(Oee) = 550
����(Oee) = 750
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 600


ElseIf Oee = 391 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 700
����(Oee) = 650
����(Oee) = 800
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 700



ElseIf Oee = 392 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 700



ElseIf Oee = 393 Then
�̸�(Oee) = "���ֿ�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 750



ElseIf Oee = 394 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 550
����(Oee) = 550
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 550



ElseIf Oee = 395 Then
�̸�(Oee) = "�ѵ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 750
����(Oee) = 650
����(Oee) = 600
�����(Oee) = 450
����(Oee) = 500
����(Oee) = 650
��Ʈ��(Oee) = 800



ElseIf Oee = 396 Then
�̸�(Oee) = "ȫ��ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 800
����(Oee) = 750
����(Oee) = 650
����(Oee) = 500
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 650
��Ʈ��(Oee) = 600



ElseIf Oee = 397 Then
�̸�(Oee) = "�輺��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�����̵�"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 500
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 398 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Special"
OYear(Oee) = "<08>"
Team(Oee) = "�����̵�"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 750
����(Oee) = 700
����(Oee) = 800
�����(Oee) = 800
����(Oee) = 750
����(Oee) = 800
��Ʈ��(Oee) = 700



ElseIf Oee = 399 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�����̵�"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 650



ElseIf Oee = 400 Then
�̸�(Oee) = "�տ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�����̵�"
����(Oee) = 3
���ݷ�(Oee) = 550
����(Oee) = 650
����(Oee) = 650
����(Oee) = 550
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 550
��Ʈ��(Oee) = 550



ElseIf Oee = 401 Then
�̸�(Oee) = "�ų뿭"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�����̵�"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 402 Then
�̸�(Oee) = "�ȱ�ȿ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�����̵�"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 650
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 403 Then
�̸�(Oee) = "�̿���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�����̵�"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 700
����(Oee) = 650
����(Oee) = 600
�����(Oee) = 500
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 700



ElseIf Oee = 404 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�����̵�"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 700
����(Oee) = 600
��Ʈ��(Oee) = 700



ElseIf Oee = 405 Then
�̸�(Oee) = "�ӵ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�����̵�"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 700



ElseIf Oee = 406 Then
�̸�(Oee) = "���¾�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�����̵�"
����(Oee) = 1
���ݷ�(Oee) = 550
����(Oee) = 600
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 750
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 600



ElseIf Oee = 407 Then
�̸�(Oee) = "�ѵ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<08>"
Team(Oee) = "�����̵�"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 650
End If
 ���(Oee) = 0
 �ؿ��(Oee) = 0
 �����(Oee) = 100
 A�¸�(Oee) = 0
 A�й�(Oee) = 0
 P�¸�(Oee) = 0
 P�й�(Oee) = 0
 T�¸�(Oee) = 0
 T�й�(Oee) = 0
 Z�¸�(Oee) = 0
 Z�й�(Oee) = 0
 T����(Oee) = 0
 Z����(Oee) = 0
 P����(Oee) = 0
 A����(Oee) = 0
 T��(Oee) = "W"
 Z��(Oee) = "W"
 P��(Oee) = "W"
 A��(Oee) = "W"

Next Oee
Tim09.Enabled = True
Tim08.Enabled = False
End Sub

Private Sub Tim09_Timer()
For Oee = 408 To 539
If Oee = 408 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 700
����(Oee) = 700
����(Oee) = 700
�����(Oee) = 650
����(Oee) = 700
����(Oee) = 600
��Ʈ��(Oee) = 550



ElseIf Oee = 409 Then
�̸�(Oee) = "��뿱"
��ũ(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
����(Oee) = 3
���ݷ�(Oee) = 850
����(Oee) = 700
����(Oee) = 600
����(Oee) = 850
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 410 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 700
����(Oee) = 700
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 700



ElseIf Oee = 411 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 650
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 600



ElseIf Oee = 412 Then
�̸�(Oee) = "���翵"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
����(Oee) = 3
���ݷ�(Oee) = 550
����(Oee) = 650
����(Oee) = 600
����(Oee) = 800
�����(Oee) = 700
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 413 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 650
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 800
����(Oee) = 700
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 414 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Unique"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
����(Oee) = 2
���ݷ�(Oee) = 900
����(Oee) = 800
����(Oee) = 650
����(Oee) = 750
�����(Oee) = 800
����(Oee) = 750
����(Oee) = 800
��Ʈ��(Oee) = 950



ElseIf Oee = 415 Then
�̸�(Oee) = "�躴��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 650
����(Oee) = 750
�����(Oee) = 550
����(Oee) = 700
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 416 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Elite"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
����(Oee) = 3
���ݷ�(Oee) = 900
����(Oee) = 800
����(Oee) = 750
����(Oee) = 850
�����(Oee) = 850
����(Oee) = 850
����(Oee) = 750
��Ʈ��(Oee) = 850



ElseIf Oee = 417 Then
�̸�(Oee) = "�̿�ȣ"
��ũ(Oee) = "Unique"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 700
����(Oee) = 750
����(Oee) = 950
�����(Oee) = 900
����(Oee) = 800
����(Oee) = 950
��Ʈ��(Oee) = 700



ElseIf Oee = 418 Then
�̸�(Oee) = "Ȳ����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "KT"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 550
����(Oee) = 500
����(Oee) = 750
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 550



ElseIf Oee = 419 Then
�̸�(Oee) = "�ڴ�ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�Ｚ����"
����(Oee) = 1
���ݷ�(Oee) = 550
����(Oee) = 600
����(Oee) = 650
����(Oee) = 600
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 420 Then
�̸�(Oee) = "�ڵ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�Ｚ����"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 500
����(Oee) = 550
��Ʈ��(Oee) = 700



ElseIf Oee = 421 Then
�̸�(Oee) = "�ռ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�Ｚ����"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 750
����(Oee) = 700
����(Oee) = 600
��Ʈ��(Oee) = 750



ElseIf Oee = 422 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 423 Then
�̸�(Oee) = "�̼���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�Ｚ����"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 750
����(Oee) = 800
����(Oee) = 550
�����(Oee) = 500
����(Oee) = 650
����(Oee) = 800
��Ʈ��(Oee) = 800



ElseIf Oee = 424 Then
�̸�(Oee) = "����Ȳ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 500
��Ʈ��(Oee) = 550



ElseIf Oee = 425 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 500
����(Oee) = 550
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 550



ElseIf Oee = 426 Then
�̸�(Oee) = "��ä��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�Ｚ����"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 550
����(Oee) = 550
�����(Oee) = 450
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 700



ElseIf Oee = 427 Then
�̸�(Oee) = "���±�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�Ｚ����"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 550
����(Oee) = 500
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 450
��Ʈ��(Oee) = 550



ElseIf Oee = 428 Then
�̸�(Oee) = "�ֿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 700



ElseIf Oee = 429 Then
�̸�(Oee) = "����ȯ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 750
����(Oee) = 700
�����(Oee) = 750
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 430 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�Ｚ����"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 550
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 700



ElseIf Oee = 431 Then
�̸�(Oee) = "�㿵��"
��ũ(Oee) = "Rare"
OYear(Oee) = "<09>"
Team(Oee) = "�Ｚ����"
����(Oee) = 3
���ݷ�(Oee) = 800
����(Oee) = 800
����(Oee) = 650
����(Oee) = 800
�����(Oee) = 800
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 900



ElseIf Oee = 432 Then
�̸�(Oee) = "���ȿ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 433 Then
�̸�(Oee) = "�豸��"
��ũ(Oee) = "Rare"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
����(Oee) = 3
���ݷ�(Oee) = 850
����(Oee) = 800
����(Oee) = 700
����(Oee) = 850
�����(Oee) = 650
����(Oee) = 700
����(Oee) = 800
��Ʈ��(Oee) = 750



ElseIf Oee = 434 Then
�̸�(Oee) = "�赿��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 600



ElseIf Oee = 435 Then
�̸�(Oee) = "�輺��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 550
����(Oee) = 600
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 436 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 650
����(Oee) = 800
�����(Oee) = 550
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 437 Then
�̸�(Oee) = "����ȯ1"
��ũ(Oee) = "Elite"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 850
����(Oee) = 750
����(Oee) = 850
����(Oee) = 800
�����(Oee) = 850
����(Oee) = 800
����(Oee) = 900
��Ʈ��(Oee) = 900



ElseIf Oee = 438 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 650
��Ʈ��(Oee) = 750



ElseIf Oee = 439 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 900
����(Oee) = 700
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 750



ElseIf Oee = 440 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 700
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 441 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 500
����(Oee) = 450
�����(Oee) = 450
����(Oee) = 500
����(Oee) = 450
��Ʈ��(Oee) = 650



ElseIf Oee = 442 Then
�̸�(Oee) = "�̽���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 500
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 600



ElseIf Oee = 443 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 550
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 500



ElseIf Oee = 444 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Rare"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 800
����(Oee) = 750
����(Oee) = 500
����(Oee) = 900
�����(Oee) = 900
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 800



ElseIf Oee = 445 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 850
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 750
����(Oee) = 750
��Ʈ��(Oee) = 800



ElseIf Oee = 446 Then
�̸�(Oee) = "���α�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 500
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 447 Then
�̸�(Oee) = "�赿��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 700
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 448 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 750
����(Oee) = 800
�����(Oee) = 800
����(Oee) = 800
����(Oee) = 700
��Ʈ��(Oee) = 800



ElseIf Oee = 449 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 800
����(Oee) = 700
����(Oee) = 700
����(Oee) = 800
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 700



ElseIf Oee = 450 Then
�̸�(Oee) = "�迵��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 650
����(Oee) = 650
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 451 Then
�̸�(Oee) = "���ر�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 452 Then
�̸�(Oee) = "�ڴ븸"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 500
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 700



ElseIf Oee = 453 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 550
����(Oee) = 550
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 700
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 650



ElseIf Oee = 454 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 800
����(Oee) = 750
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 900



ElseIf Oee = 455 Then
�̸�(Oee) = "�̵���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 650
����(Oee) = 600
�����(Oee) = 700
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 550



ElseIf Oee = 456 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 550
����(Oee) = 500
����(Oee) = 650
�����(Oee) = 450
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 700



ElseIf Oee = 457 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 700
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 750
����(Oee) = 600
��Ʈ��(Oee) = 850



ElseIf Oee = 458 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 459 Then
�̸�(Oee) = "�ѻ��"
��ũ(Oee) = "Rare"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 900
����(Oee) = 850
����(Oee) = 700
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 700
����(Oee) = 750
��Ʈ��(Oee) = 950



ElseIf Oee = 460 Then
�̸�(Oee) = "���α�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 750
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 700
����(Oee) = 700
��Ʈ��(Oee) = 750



ElseIf Oee = 461 Then
�̸�(Oee) = "�ǿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
����(Oee) = 3
���ݷ�(Oee) = 550
����(Oee) = 600
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 462 Then
�̸�(Oee) = "���ÿ�"
��ũ(Oee) = "Unique"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 950
����(Oee) = 750
����(Oee) = 800
�����(Oee) = 750
����(Oee) = 900
����(Oee) = 850
��Ʈ��(Oee) = 900



ElseIf Oee = 463 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
����(Oee) = 3
���ݷ�(Oee) = 800
����(Oee) = 700
����(Oee) = 650
����(Oee) = 900
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 650



ElseIf Oee = 464 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 700
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 750
��Ʈ��(Oee) = 800



ElseIf Oee = 465 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 550
����(Oee) = 600
����(Oee) = 500
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 550



ElseIf Oee = 466 Then
�̸�(Oee) = "�̽¼�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 467 Then
�̸�(Oee) = "�ӿ�ȯ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 800
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 468 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Rare"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 800
����(Oee) = 800
����(Oee) = 800
�����(Oee) = 900
����(Oee) = 750
����(Oee) = 850
��Ʈ��(Oee) = 600



ElseIf Oee = 469 Then
�̸�(Oee) = "����ö"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 650
����(Oee) = 750
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 470 Then
�̸�(Oee) = "�ֿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
����(Oee) = 1
���ݷ�(Oee) = 550
����(Oee) = 600
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 471 Then
�̸�(Oee) = "��ȣ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "SKT"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 500
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 500
����(Oee) = 650
��Ʈ��(Oee) = 600



ElseIf Oee = 472 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 800
����(Oee) = 750
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 700



ElseIf Oee = 473 Then
�̸�(Oee) = "�赿��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 650



ElseIf Oee = 474 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 475 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 500
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 476 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 500
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 600



ElseIf Oee = 477 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
����(Oee) = 3
���ݷ�(Oee) = 850
����(Oee) = 650
����(Oee) = 700
����(Oee) = 800
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 478 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 750
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 700



ElseIf Oee = 479 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 700
����(Oee) = 850
�����(Oee) = 800
����(Oee) = 800
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 480 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Rare"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 800
����(Oee) = 750
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 800
����(Oee) = 750
��Ʈ��(Oee) = 800



ElseIf Oee = 481 Then
�̸�(Oee) = "�Ӽ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 700
����(Oee) = 600
����(Oee) = 500
�����(Oee) = 500
����(Oee) = 650
����(Oee) = 500
��Ʈ��(Oee) = 500



ElseIf Oee = 482 Then
�̸�(Oee) = "���ö"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 650



ElseIf Oee = 483 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 600



ElseIf Oee = 484 Then
�̸�(Oee) = "���켭"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 485 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "ȭ��"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 450
����(Oee) = 550
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 450
��Ʈ��(Oee) = 650



ElseIf Oee = 486 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "ȭ��"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 750
����(Oee) = 750
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 600



ElseIf Oee = 487 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "ȭ��"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 488 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "ȭ��"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 500
����(Oee) = 550
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 489 Then
�̸�(Oee) = "���±�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "ȭ��"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 650
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 490 Then
�̸�(Oee) = "�뿵��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "ȭ��"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 550
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 450
��Ʈ��(Oee) = 500



ElseIf Oee = 491 Then
�̸�(Oee) = "���ؿ�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "ȭ��"
����(Oee) = 2
���ݷ�(Oee) = 800
����(Oee) = 700
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 700
����(Oee) = 550
��Ʈ��(Oee) = 750



ElseIf Oee = 492 Then
�̸�(Oee) = "���¼�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "ȭ��"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 500
����(Oee) = 500
�����(Oee) = 450
����(Oee) = 450
����(Oee) = 450
��Ʈ��(Oee) = 550



ElseIf Oee = 493 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "ȭ��"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 700
����(Oee) = 700
�����(Oee) = 750
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 750



ElseIf Oee = 494 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "ȭ��"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 900
����(Oee) = 700
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 750



ElseIf Oee = 495 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Legend"
OYear(Oee) = "<09>"
Team(Oee) = "ȭ��"
����(Oee) = 2
���ݷ�(Oee) = 950
����(Oee) = 950
����(Oee) = 750
����(Oee) = 800
�����(Oee) = 800
����(Oee) = 750
����(Oee) = 950
��Ʈ��(Oee) = 950



ElseIf Oee = 496 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "ȭ��"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 497 Then
�̸�(Oee) = "�ӿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "ȭ��"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 700
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 700
����(Oee) = 550
��Ʈ��(Oee) = 650



ElseIf Oee = 498 Then
�̸�(Oee) = "�Ǽ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 499 Then
�̸�(Oee) = "���ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 500
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 500 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Elite"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 900
����(Oee) = 800
����(Oee) = 650
����(Oee) = 900
�����(Oee) = 900
����(Oee) = 850
����(Oee) = 800
��Ʈ��(Oee) = 800


ElseIf Oee = 501 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
����(Oee) = 1
���ݷ�(Oee) = 850
����(Oee) = 750
����(Oee) = 650
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 650



ElseIf Oee = 502 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
��Ʈ��(Oee) = 500



ElseIf Oee = 503 Then
�̸�(Oee) = "�ŵ���"
��ũ(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 800
����(Oee) = 600
����(Oee) = 600
����(Oee) = 900
�����(Oee) = 750
����(Oee) = 650
����(Oee) = 550
��Ʈ��(Oee) = 750



ElseIf Oee = 504 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 700



ElseIf Oee = 505 Then
�̸�(Oee) = "���ֿ�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 750
��Ʈ��(Oee) = 750



ElseIf Oee = 506 Then
�̸�(Oee) = "����ö"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 550



ElseIf Oee = 507 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
����(Oee) = 1
���ݷ�(Oee) = 900
����(Oee) = 750
����(Oee) = 650
����(Oee) = 750
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 750
��Ʈ��(Oee) = 700



ElseIf Oee = 508 Then
�̸�(Oee) = "����ȭ"
��ũ(Oee) = "Rare"
OYear(Oee) = "<09>"
Team(Oee) = "CJ"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 750
����(Oee) = 750
����(Oee) = 800
�����(Oee) = 850
����(Oee) = 650
����(Oee) = 750
��Ʈ��(Oee) = 700



ElseIf Oee = 509 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����Ʈ"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 700
����(Oee) = 600
����(Oee) = 800
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 600



ElseIf Oee = 510 Then
�̸�(Oee) = "���м�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����Ʈ"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 550
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 550



ElseIf Oee = 511 Then
�̸�(Oee) = "�Ż�"
��ũ(Oee) = "Rare"
OYear(Oee) = "<09>"
Team(Oee) = "����Ʈ"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 800
����(Oee) = 750
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 700
����(Oee) = 850
��Ʈ��(Oee) = 800



ElseIf Oee = 512 Then
�̸�(Oee) = "���ؿ�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����Ʈ"
����(Oee) = 2
���ݷ�(Oee) = 550
����(Oee) = 550
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 650



ElseIf Oee = 513 Then
�̸�(Oee) = "�̰��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����Ʈ"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 850
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 514 Then
�̸�(Oee) = "�赵��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 650
����(Oee) = 500
����(Oee) = 450
��Ʈ��(Oee) = 550



ElseIf Oee = 515 Then
�̸�(Oee) = "�輺��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 500
����(Oee) = 700
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 516 Then
�̸�(Oee) = "�ڻ��"
��ũ(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 650
����(Oee) = 850
�����(Oee) = 800
����(Oee) = 750
����(Oee) = 750
��Ʈ��(Oee) = 650



ElseIf Oee = 517 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 600



ElseIf Oee = 518 Then
�̸�(Oee) = "�Ŵ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 750
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 700
����(Oee) = 700
��Ʈ��(Oee) = 700



ElseIf Oee = 519 Then
�̸�(Oee) = "�Ż�ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
����(Oee) = 3
���ݷ�(Oee) = 850
����(Oee) = 700
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 650



ElseIf Oee = 520 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 650
����(Oee) = 650
����(Oee) = 750
�����(Oee) = 800
����(Oee) = 700
����(Oee) = 600
��Ʈ��(Oee) = 700



ElseIf Oee = 521 Then
�̸�(Oee) = "�ȼ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 500
����(Oee) = 500
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 500



ElseIf Oee = 522 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 523 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "eSTRO"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 600
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 650



ElseIf Oee = 524 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 750
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 750



ElseIf Oee = 525 Then
�̸�(Oee) = "�ڿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 750
����(Oee) = 800
�����(Oee) = 650
����(Oee) = 700
����(Oee) = 600
��Ʈ��(Oee) = 700



ElseIf Oee = 526 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 650
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 500
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 750



ElseIf Oee = 527 Then
�̸�(Oee) = "���¹�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 700
����(Oee) = 650
��Ʈ��(Oee) = 700



ElseIf Oee = 527 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 528 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 750
����(Oee) = 650
����(Oee) = 800
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 700



ElseIf Oee = 529 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 550
����(Oee) = 550
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 550



ElseIf Oee = 530 Then
�̸�(Oee) = "�ѵ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 750
����(Oee) = 650
����(Oee) = 600
�����(Oee) = 450
����(Oee) = 500
����(Oee) = 600
��Ʈ��(Oee) = 800



ElseIf Oee = 531 Then
�̸�(Oee) = "ȫ��ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 850
����(Oee) = 750
����(Oee) = 750
����(Oee) = 500
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 531 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�����̵�"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 550
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 600



ElseIf Oee = 532 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�����̵�"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 700
����(Oee) = 650
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 750
����(Oee) = 700
��Ʈ��(Oee) = 650



ElseIf Oee = 533 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�����̵�"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 800
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 534 Then
�̸�(Oee) = "�ų뿭"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�����̵�"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 750
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 700
����(Oee) = 700
����(Oee) = 600
��Ʈ��(Oee) = 750



ElseIf Oee = 535 Then
�̸�(Oee) = "�ȱ�ȿ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�����̵�"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 650



ElseIf Oee = 536 Then
�̸�(Oee) = "�̿���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�����̵�"
����(Oee) = 2
���ݷ�(Oee) = 850
����(Oee) = 750
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 750



ElseIf Oee = 537 Then
�̸�(Oee) = "�̿�ȣ1"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�����̵�"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 550
����(Oee) = 650
��Ʈ��(Oee) = 650



ElseIf Oee = 538 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Special"
OYear(Oee) = "<09>"
Team(Oee) = "�����̵�"
����(Oee) = 1
���ݷ�(Oee) = 850
����(Oee) = 650
����(Oee) = 650
����(Oee) = 850
�����(Oee) = 800
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 600



ElseIf Oee = 539 Then
�̸�(Oee) = "���¾�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<09>"
Team(Oee) = "�����̵�"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 800
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 650
End If
 ���(Oee) = 0
 �ؿ��(Oee) = 0
 �����(Oee) = 100
 A�¸�(Oee) = 0
 A�й�(Oee) = 0
 P�¸�(Oee) = 0
 P�й�(Oee) = 0
 T�¸�(Oee) = 0
 T�й�(Oee) = 0
 Z�¸�(Oee) = 0
 Z�й�(Oee) = 0
 T����(Oee) = 0
 Z����(Oee) = 0
 P����(Oee) = 0
 A����(Oee) = 0
 T��(Oee) = "W"
 Z��(Oee) = "W"
 P��(Oee) = "W"
 A��(Oee) = "W"

Next Oee
Tim10.Enabled = True
Tim09.Enabled = False
End Sub

Private Sub Tim10_Timer()
For Oee = 576 To 715
If Oee = 576 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 700
 ����(Oee) = 550
 ����(Oee) = 600
 ����(Oee) = 700
 �����(Oee) = 550
 ����(Oee) = 550
 ����(Oee) = 550
 ��Ʈ��(Oee) = 600

ElseIf Oee = 577 Then
 �̸�(Oee) = "�ڼ���"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 550
 ����(Oee) = 700
 �����(Oee) = 800
 ����(Oee) = 700
 ����(Oee) = 700
 ��Ʈ��(Oee) = 550

ElseIf Oee = 578 Then
 �̸�(Oee) = "�ڼ���"
 ��ũ(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 3
 ���ݷ�(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 800
 �����(Oee) = 700
 ����(Oee) = 800
 ����(Oee) = 800
 ��Ʈ��(Oee) = 850

ElseIf Oee = 579 Then
 �̸�(Oee) = "�ų뿭"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 2
 ���ݷ�(Oee) = 650
 ����(Oee) = 750
 ����(Oee) = 550
 ����(Oee) = 750
 �����(Oee) = 650
 ����(Oee) = 700
 ����(Oee) = 800
 ��Ʈ��(Oee) = 750

ElseIf Oee = 580 Then
 �̸�(Oee) = "�̿�ȣ1"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 3
 ���ݷ�(Oee) = 700
 ����(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 650
 �����(Oee) = 550
 ����(Oee) = 550
 ����(Oee) = 650
 ��Ʈ��(Oee) = 650

ElseIf Oee = 581 Then
 �̸�(Oee) = "�̿���"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 2
 ���ݷ�(Oee) = 850
 ����(Oee) = 750
 ����(Oee) = 650
 ����(Oee) = 700
 �����(Oee) = 550
 ����(Oee) = 600
 ����(Oee) = 650
 ��Ʈ��(Oee) = 750

ElseIf Oee = 582 Then
 �̸�(Oee) = "�̿���"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 2
 ���ݷ�(Oee) = 700
 ����(Oee) = 600
 ����(Oee) = 500
 ����(Oee) = 650
 �����(Oee) = 600
 ����(Oee) = 550
 ����(Oee) = 550
 ��Ʈ��(Oee) = 550

ElseIf Oee = 583 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 700
 ����(Oee) = 700
 �����(Oee) = 750
 ����(Oee) = 700
 ����(Oee) = 650
 ��Ʈ��(Oee) = 700

ElseIf Oee = 584 Then
 �̸�(Oee) = "�����"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 750
 ����(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 700
 �����(Oee) = 750
 ����(Oee) = 700
 ����(Oee) = 700
 ��Ʈ��(Oee) = 650

ElseIf Oee = 585 Then
 �̸�(Oee) = "���¾�"
 ��ũ(Oee) = "Elite"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 850
 ����(Oee) = 850
 ����(Oee) = 750
 ����(Oee) = 850
 �����(Oee) = 900
 ����(Oee) = 750
 ����(Oee) = 800
 ��Ʈ��(Oee) = 850

ElseIf Oee = 586 Then
 �̸�(Oee) = "����"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 2
 ���ݷ�(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 500
 ����(Oee) = 600
 �����(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 550
 ��Ʈ��(Oee) = 650

ElseIf Oee = 587 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 750
 ����(Oee) = 750
 ����(Oee) = 650
 ����(Oee) = 650
 �����(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 600
 ��Ʈ��(Oee) = 800

ElseIf Oee = 588 Then
 �̸�(Oee) = "�ڿ���"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 3
 ���ݷ�(Oee) = 700
 ����(Oee) = 600
 ����(Oee) = 700
 ����(Oee) = 800
 �����(Oee) = 600
 ����(Oee) = 700
 ����(Oee) = 600
 ��Ʈ��(Oee) = 700

ElseIf Oee = 589 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 3
 ���ݷ�(Oee) = 750
 ����(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 700
 �����(Oee) = 500
 ����(Oee) = 600
 ����(Oee) = 550
 ��Ʈ��(Oee) = 750

ElseIf Oee = 590 Then
 �̸�(Oee) = "���¹�"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 2
 ���ݷ�(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 750
 �����(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 650
 ��Ʈ��(Oee) = 700

ElseIf Oee = 591 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 650
 �����(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 650
 ��Ʈ��(Oee) = 650

ElseIf Oee = 592 Then
 �̸�(Oee) = "�ռ���"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 3
 ���ݷ�(Oee) = 650
 ����(Oee) = 500
 ����(Oee) = 600
 ����(Oee) = 600
 �����(Oee) = 550
 ����(Oee) = 500
 ����(Oee) = 550
 ��Ʈ��(Oee) = 650

ElseIf Oee = 593 Then
 �̸�(Oee) = "�ȱ�ȿ"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 3
 ���ݷ�(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 700
 �����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 600
 ��Ʈ��(Oee) = 650

ElseIf Oee = 594 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 3
 ���ݷ�(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 800
 �����(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 650
 ��Ʈ��(Oee) = 700

ElseIf Oee = 595 Then
 �̸�(Oee) = "�����"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 600
 ����(Oee) = 500
 ����(Oee) = 550
 ����(Oee) = 550
 �����(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 600
 ��Ʈ��(Oee) = 550

ElseIf Oee = 596 Then
 �̸�(Oee) = "�ѵ���"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 800
 ����(Oee) = 700
 ����(Oee) = 600
 ����(Oee) = 600
 �����(Oee) = 450
 ����(Oee) = 500
 ����(Oee) = 600
 ��Ʈ��(Oee) = 800

ElseIf Oee = 597 Then
 �̸�(Oee) = "ȫ��ȣ"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 2
 ���ݷ�(Oee) = 850
 ����(Oee) = 750
 ����(Oee) = 750
 ����(Oee) = 500
 �����(Oee) = 500
 ����(Oee) = 500
 ����(Oee) = 600
 ��Ʈ��(Oee) = 650

ElseIf Oee = 598 Then
 �̸�(Oee) = "�赵��"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 ����(Oee) = 1
 ���ݷ�(Oee) = 700
 ����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 650
 �����(Oee) = 750
 ����(Oee) = 600
 ����(Oee) = 650
 ��Ʈ��(Oee) = 700

ElseIf Oee = 599 Then
 �̸�(Oee) = "�輺��"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 ����(Oee) = 2
 ���ݷ�(Oee) = 650
 ����(Oee) = 700
 ����(Oee) = 550
 ����(Oee) = 800
 �����(Oee) = 800
 ����(Oee) = 700
 ����(Oee) = 650
 ��Ʈ��(Oee) = 700

ElseIf Oee = 600 Then
 �̸�(Oee) = "�ڻ��"
 ��ũ(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 ����(Oee) = 1
 ���ݷ�(Oee) = 800
 ����(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 950
 �����(Oee) = 850
 ����(Oee) = 750
 ����(Oee) = 800
 ��Ʈ��(Oee) = 600

ElseIf Oee = 601 Then
 �̸�(Oee) = "�Ŵ��"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 ����(Oee) = 2
 ���ݷ�(Oee) = 750
 ����(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 700
 �����(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 650
 ��Ʈ��(Oee) = 550

ElseIf Oee = 602 Then
 �̸�(Oee) = "�Ż�ȣ"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 ����(Oee) = 3
 ���ݷ�(Oee) = 800
 ����(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 700
 �����(Oee) = 600
 ����(Oee) = 550
 ����(Oee) = 550
 ��Ʈ��(Oee) = 550

ElseIf Oee = 603 Then
 �̸�(Oee) = "�����"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 ����(Oee) = 3
 ���ݷ�(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 750
 ����(Oee) = 650
 �����(Oee) = 800
 ����(Oee) = 700
 ����(Oee) = 650
 ��Ʈ��(Oee) = 750

ElseIf Oee = 604 Then
 �̸�(Oee) = "�ȼ���"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 ����(Oee) = 3
 ���ݷ�(Oee) = 650
 ����(Oee) = 500
 ����(Oee) = 500
 ����(Oee) = 650
 �����(Oee) = 500
 ����(Oee) = 500
 ����(Oee) = 500
 ��Ʈ��(Oee) = 500

ElseIf Oee = 605 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 ����(Oee) = 3
 ���ݷ�(Oee) = 550
 ����(Oee) = 550
 ����(Oee) = 500
 ����(Oee) = 600
 �����(Oee) = 600
 ����(Oee) = 500
 ����(Oee) = 500
 ��Ʈ��(Oee) = 550

ElseIf Oee = 606 Then
 �̸�(Oee) = "����ȣ"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 ����(Oee) = 2
 ���ݷ�(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 550
 ����(Oee) = 650
 �����(Oee) = 600
 ����(Oee) = 550
 ����(Oee) = 600
 ��Ʈ��(Oee) = 650

ElseIf Oee = 607 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "eSTRO"
 ����(Oee) = 1
 ���ݷ�(Oee) = 700
 ����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 600
 �����(Oee) = 550
 ����(Oee) = 500
 ����(Oee) = 500
 ��Ʈ��(Oee) = 650

ElseIf Oee = 608 Then
 �̸�(Oee) = "����"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����Ʈ"
 ����(Oee) = 2
 ���ݷ�(Oee) = 650
 ����(Oee) = 550
 ����(Oee) = 500
 ����(Oee) = 600
 �����(Oee) = 600
 ����(Oee) = 500
 ����(Oee) = 550
 ��Ʈ��(Oee) = 600

ElseIf Oee = 609 Then
 �̸�(Oee) = "�����"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����Ʈ"
 ����(Oee) = 3
 ���ݷ�(Oee) = 550
 ����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 750
 �����(Oee) = 500
 ����(Oee) = 500
 ����(Oee) = 550
 ��Ʈ��(Oee) = 600

ElseIf Oee = 610 Then
 �̸�(Oee) = "����"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����Ʈ"
 ����(Oee) = 2
 ���ݷ�(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 800
 �����(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 600
 ��Ʈ��(Oee) = 650

ElseIf Oee = 611 Then
 �̸�(Oee) = "���м�"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����Ʈ"
 ����(Oee) = 3
 ���ݷ�(Oee) = 700
 ����(Oee) = 550
 ����(Oee) = 650
 ����(Oee) = 650
 �����(Oee) = 550
 ����(Oee) = 550
 ����(Oee) = 550
 ��Ʈ��(Oee) = 550

ElseIf Oee = 612 Then
 �̸�(Oee) = "�Ż�"
 ��ũ(Oee) = "Special"
 OYear(Oee) = "<10>"
 Team(Oee) = "����Ʈ"
 ����(Oee) = 1
 ���ݷ�(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 750
 ����(Oee) = 750
 �����(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 700
 ��Ʈ��(Oee) = 700

ElseIf Oee = 613 Then
 �̸�(Oee) = "���ؿ�"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����Ʈ"
 ����(Oee) = 2
 ���ݷ�(Oee) = 550
 ����(Oee) = 550
 ����(Oee) = 550
 ����(Oee) = 650
 �����(Oee) = 550
 ����(Oee) = 550
 ����(Oee) = 500
 ��Ʈ��(Oee) = 650

ElseIf Oee = 614 Then
 �̸�(Oee) = "�̰��"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����Ʈ"
 ����(Oee) = 3
 ���ݷ�(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 750
 ����(Oee) = 850
 �����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 700
 ��Ʈ��(Oee) = 750

ElseIf Oee = 615 Then
 �̸�(Oee) = "��ȣ��"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����Ʈ"
 ����(Oee) = 1
 ���ݷ�(Oee) = 600
 ����(Oee) = 550
 ����(Oee) = 600
 ����(Oee) = 750
 �����(Oee) = 750
 ����(Oee) = 750
 ����(Oee) = 600
 ��Ʈ��(Oee) = 550

ElseIf Oee = 616 Then
 �̸�(Oee) = "�����"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����Ʈ"
 ����(Oee) = 3
 ���ݷ�(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 550
 ����(Oee) = 650
 �����(Oee) = 550
 ����(Oee) = 550
 ����(Oee) = 650
 ��Ʈ��(Oee) = 650

ElseIf Oee = 617 Then
 �̸�(Oee) = "��ȫ��"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����Ʈ"
 ����(Oee) = 3
 ���ݷ�(Oee) = 550
 ����(Oee) = 550
 ����(Oee) = 500
 ����(Oee) = 600
 �����(Oee) = 550
 ����(Oee) = 500
 ����(Oee) = 500
 ��Ʈ��(Oee) = 500

ElseIf Oee = 618 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����Ʈ"
 ����(Oee) = 3
 ���ݷ�(Oee) = 650
 ����(Oee) = 550
 ����(Oee) = 600
 ����(Oee) = 650
 �����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 550
 ��Ʈ��(Oee) = 550

ElseIf Oee = 619 Then
 �̸�(Oee) = "�Ǽ���"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 ����(Oee) = 2
 ���ݷ�(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 550
 ����(Oee) = 600
 �����(Oee) = 700
 ����(Oee) = 600
 ����(Oee) = 550
 ��Ʈ��(Oee) = 600

ElseIf Oee = 620 Then
 �̸�(Oee) = "���ȣ"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 ����(Oee) = 2
 ���ݷ�(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 500
 ����(Oee) = 650
 �����(Oee) = 500
 ����(Oee) = 500
 ����(Oee) = 600
 ��Ʈ��(Oee) = 650

ElseIf Oee = 621 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Unique"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 ����(Oee) = 2
 ���ݷ�(Oee) = 850
 ����(Oee) = 800
 ����(Oee) = 750
 ����(Oee) = 850
 �����(Oee) = 800
 ����(Oee) = 800
 ����(Oee) = 750
 ��Ʈ��(Oee) = 850

ElseIf Oee = 622 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 ����(Oee) = 1
 ���ݷ�(Oee) = 850
 ����(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 700
 �����(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 700
 ��Ʈ��(Oee) = 650

ElseIf Oee = 623 Then
 �̸�(Oee) = "�����"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 ����(Oee) = 3
 ���ݷ�(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 550
 ����(Oee) = 700
 �����(Oee) = 650
 ����(Oee) = 550
 ����(Oee) = 600
 ��Ʈ��(Oee) = 500

ElseIf Oee = 624 Then
 �̸�(Oee) = "�ŵ���"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 ����(Oee) = 2
 ���ݷ�(Oee) = 750
 ����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 700
 �����(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 650
 ��Ʈ��(Oee) = 750

ElseIf Oee = 625 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 ����(Oee) = 1
 ���ݷ�(Oee) = 950
 ����(Oee) = 750
 ����(Oee) = 700
 ����(Oee) = 750
 �����(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 950
 ��Ʈ��(Oee) = 800

ElseIf Oee = 626 Then
 �̸�(Oee) = "����ö"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 ����(Oee) = 3
 ���ݷ�(Oee) = 750
 ����(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 800
 �����(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 700
 ��Ʈ��(Oee) = 850

ElseIf Oee = 627 Then
 �̸�(Oee) = "�����"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 ����(Oee) = 1
 ���ݷ�(Oee) = 650
 ����(Oee) = 550
 ����(Oee) = 550
 ����(Oee) = 650
 �����(Oee) = 500
 ����(Oee) = 600
 ����(Oee) = 650
 ��Ʈ��(Oee) = 600

ElseIf Oee = 628 Then
 �̸�(Oee) = "����ȭ"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 ����(Oee) = 3
 ���ݷ�(Oee) = 600
 ����(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 750
 �����(Oee) = 750
 ����(Oee) = 750
 ����(Oee) = 650
 ��Ʈ��(Oee) = 700

ElseIf Oee = 629 Then
 �̸�(Oee) = "�ѵο�"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "CJ"
 ����(Oee) = 2
 ���ݷ�(Oee) = 550
 ����(Oee) = 600
 ����(Oee) = 550
 ����(Oee) = 600
 �����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 550
 ��Ʈ��(Oee) = 600

ElseIf Oee = 630 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "ȭ��"
 ����(Oee) = 2
 ���ݷ�(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 450
 ����(Oee) = 550
 �����(Oee) = 500
 ����(Oee) = 550
 ����(Oee) = 450
 ��Ʈ��(Oee) = 650

ElseIf Oee = 631 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "ȭ��"
 ����(Oee) = 1
 ���ݷ�(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 650
 �����(Oee) = 750
 ����(Oee) = 700
 ����(Oee) = 750
 ��Ʈ��(Oee) = 650

ElseIf Oee = 632 Then
 �̸�(Oee) = "�����"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "ȭ��"
 ����(Oee) = 1
 ���ݷ�(Oee) = 600
 ����(Oee) = 500
 ����(Oee) = 500
 ����(Oee) = 550
 �����(Oee) = 500
 ����(Oee) = 550
 ����(Oee) = 600
 ��Ʈ��(Oee) = 600

ElseIf Oee = 633 Then
 �̸�(Oee) = "���±�"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "ȭ��"
 ����(Oee) = 3
 ���ݷ�(Oee) = 750
 ����(Oee) = 650
 ����(Oee) = 550
 ����(Oee) = 700
 �����(Oee) = 550
 ����(Oee) = 600
 ����(Oee) = 600
 ��Ʈ��(Oee) = 600

ElseIf Oee = 634 Then
 �̸�(Oee) = "���ؿ�"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "ȭ��"
 ����(Oee) = 2
 ���ݷ�(Oee) = 750
 ����(Oee) = 700
 ����(Oee) = 550
 ����(Oee) = 650
 �����(Oee) = 600
 ����(Oee) = 700
 ����(Oee) = 600
 ��Ʈ��(Oee) = 750

ElseIf Oee = 635 Then
 �̸�(Oee) = "���¼�"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "ȭ��"
 ����(Oee) = 2
 ���ݷ�(Oee) = 600
 ����(Oee) = 500
 ����(Oee) = 500
 ����(Oee) = 500
 �����(Oee) = 450
 ����(Oee) = 450
 ����(Oee) = 450
 ��Ʈ��(Oee) = 550

ElseIf Oee = 636 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "ȭ��"
 ����(Oee) = 1
 ���ݷ�(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 700
 ����(Oee) = 700
 �����(Oee) = 650
 ����(Oee) = 700
 ����(Oee) = 600
 ��Ʈ��(Oee) = 750

ElseIf Oee = 637 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "ȭ��"
 ����(Oee) = 3
 ���ݷ�(Oee) = 700
 ����(Oee) = 900
 ����(Oee) = 700
 ����(Oee) = 600
 �����(Oee) = 550
 ����(Oee) = 600
 ����(Oee) = 600
 ��Ʈ��(Oee) = 750

ElseIf Oee = 638 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Elite"
 OYear(Oee) = "<10>"
 Team(Oee) = "ȭ��"
 ����(Oee) = 2
 ���ݷ�(Oee) = 950
 ����(Oee) = 950
 ����(Oee) = 800
 ����(Oee) = 800
 �����(Oee) = 750
 ����(Oee) = 750
 ����(Oee) = 800
 ��Ʈ��(Oee) = 800

ElseIf Oee = 639 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "ȭ��"
 ����(Oee) = 1
 ���ݷ�(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 550
 ����(Oee) = 650
 �����(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 600
 ��Ʈ��(Oee) = 650

ElseIf Oee = 640 Then
 �̸�(Oee) = "�ӿ���"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "ȭ��"
 ����(Oee) = 3
 ���ݷ�(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 700
 ����(Oee) = 650
 �����(Oee) = 500
 ����(Oee) = 700
 ����(Oee) = 550
 ��Ʈ��(Oee) = 650

ElseIf Oee = 641 Then
 �̸�(Oee) = "����"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 ����(Oee) = 2
 ���ݷ�(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 700
 �����(Oee) = 650
 ����(Oee) = 700
 ����(Oee) = 600
 ��Ʈ��(Oee) = 550

ElseIf Oee = 642 Then
 �̸�(Oee) = "��뿱"
 ��ũ(Oee) = "Special"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 ����(Oee) = 3
 ���ݷ�(Oee) = 750
 ����(Oee) = 750
 ����(Oee) = 600
 ����(Oee) = 850
 �����(Oee) = 750
 ����(Oee) = 600
 ����(Oee) = 650
 ��Ʈ��(Oee) = 650

ElseIf Oee = 643 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 ����(Oee) = 2
 ���ݷ�(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 600
 �����(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 650
 ��Ʈ��(Oee) = 650

ElseIf Oee = 644 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 ����(Oee) = 1
 ���ݷ�(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 600
 �����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 650
 ��Ʈ��(Oee) = 600

ElseIf Oee = 645 Then
 �̸�(Oee) = "���翵"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 ����(Oee) = 3
 ���ݷ�(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 800
 �����(Oee) = 700
 ����(Oee) = 550
 ����(Oee) = 600
 ��Ʈ��(Oee) = 550

ElseIf Oee = 646 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 ����(Oee) = 1
 ���ݷ�(Oee) = 800
 ����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 800
 �����(Oee) = 800
 ����(Oee) = 650
 ����(Oee) = 600
 ��Ʈ��(Oee) = 750

ElseIf Oee = 647 Then
 �̸�(Oee) = "�躴��"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 ����(Oee) = 2
 ���ݷ�(Oee) = 700
 ����(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 700
 �����(Oee) = 550
 ����(Oee) = 700
 ����(Oee) = 550
 ��Ʈ��(Oee) = 650

ElseIf Oee = 648 Then
 �̸�(Oee) = "����ȣ"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 ����(Oee) = 3
 ���ݷ�(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 750
 �����(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 700
 ��Ʈ��(Oee) = 750

ElseIf Oee = 649 Then
 �̸�(Oee) = "�̿�ȣ"
 ��ũ(Oee) = "Legend"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 ����(Oee) = 1
 ���ݷ�(Oee) = 850
 ����(Oee) = 800
 ����(Oee) = 800
 ����(Oee) = 950
 �����(Oee) = 950
 ����(Oee) = 900
 ����(Oee) = 900
 ��Ʈ��(Oee) = 800
ElseIf Oee = 650 Then
 �̸�(Oee) = "�ֿ���"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 ����(Oee) = 2
 ���ݷ�(Oee) = 550
 ����(Oee) = 600
 ����(Oee) = 500
 ����(Oee) = 550
 �����(Oee) = 450
 ����(Oee) = 500
 ����(Oee) = 500
 ��Ʈ��(Oee) = 700

ElseIf Oee = 651 Then
 �̸�(Oee) = "Ȳ����"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "KT"
 ����(Oee) = 1
 ���ݷ�(Oee) = 650
 ����(Oee) = 550
 ����(Oee) = 500
 ����(Oee) = 750
 �����(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 550
 ��Ʈ��(Oee) = 550

ElseIf Oee = 652 Then
 �̸�(Oee) = "�����"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "�Ｚ����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 500
 ����(Oee) = 600
 ����(Oee) = 500
 ����(Oee) = 600
 �����(Oee) = 650
 ����(Oee) = 550
 ����(Oee) = 500
 ��Ʈ��(Oee) = 500

ElseIf Oee = 653 Then
 �̸�(Oee) = "�ڴ�ȣ"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "�Ｚ����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 550
 ����(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 600
 �����(Oee) = 650
 ����(Oee) = 550
 ����(Oee) = 550
 ��Ʈ��(Oee) = 600

ElseIf Oee = 654 Then
 �̸�(Oee) = "�ۺ���"
 ��ũ(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "�Ｚ����"
 ����(Oee) = 3
 ���ݷ�(Oee) = 750
 ����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 900
 �����(Oee) = 900
 ����(Oee) = 700
 ����(Oee) = 800
 ��Ʈ��(Oee) = 750

ElseIf Oee = 655 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "�Ｚ����"
 ����(Oee) = 2
 ���ݷ�(Oee) = 700
 ����(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 700
 �����(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 600
 ��Ʈ��(Oee) = 650

ElseIf Oee = 656 Then
 �̸�(Oee) = "�̼���"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "�Ｚ����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 750
 ����(Oee) = 700
 ����(Oee) = 800
 ����(Oee) = 600
 �����(Oee) = 450
 ����(Oee) = 600
 ����(Oee) = 700
 ��Ʈ��(Oee) = 750

ElseIf Oee = 657 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "�Ｚ����"
 ����(Oee) = 2
 ���ݷ�(Oee) = 600
 ����(Oee) = 550
 ����(Oee) = 600
 ����(Oee) = 650
 �����(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 550
 ��Ʈ��(Oee) = 550

ElseIf Oee = 658 Then
 �̸�(Oee) = "���±�"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "�Ｚ����"
 ����(Oee) = 3
 ���ݷ�(Oee) = 650
 ����(Oee) = 700
 ����(Oee) = 500
 ����(Oee) = 750
 �����(Oee) = 700
 ����(Oee) = 600
 ����(Oee) = 500
 ��Ʈ��(Oee) = 600

ElseIf Oee = 659 Then
 �̸�(Oee) = "���⼮"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "�Ｚ����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 500
 ����(Oee) = 550
 ����(Oee) = 500
 ����(Oee) = 650
 �����(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 500
 ��Ʈ��(Oee) = 550

ElseIf Oee = 660 Then
 �̸�(Oee) = "�ֿ���"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "�Ｚ����"
 ����(Oee) = 2
 ���ݷ�(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 650
 �����(Oee) = 550
 ����(Oee) = 600
 ����(Oee) = 650
 ��Ʈ��(Oee) = 700

ElseIf Oee = 661 Then
 �̸�(Oee) = "����ȯ"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "�Ｚ����"
 ����(Oee) = 2
 ���ݷ�(Oee) = 700
 ����(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 700
 �����(Oee) = 750
 ����(Oee) = 700
 ����(Oee) = 700
 ��Ʈ��(Oee) = 700

ElseIf Oee = 662 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "�Ｚ����"
 ����(Oee) = 3
 ���ݷ�(Oee) = 650
 ����(Oee) = 550
 ����(Oee) = 550
 ����(Oee) = 650
 �����(Oee) = 500
 ����(Oee) = 550
 ����(Oee) = 500
 ��Ʈ��(Oee) = 700

ElseIf Oee = 663 Then
 �̸�(Oee) = "�㿵��"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "�Ｚ����"
 ����(Oee) = 3
 ���ݷ�(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 750
 �����(Oee) = 550
 ����(Oee) = 650
 ����(Oee) = 600
 ��Ʈ��(Oee) = 750

ElseIf Oee = 664 Then
 �̸�(Oee) = "���ȿ"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 ����(Oee) = 1
 ���ݷ�(Oee) = 650
 ����(Oee) = 700
 ����(Oee) = 550
 ����(Oee) = 600
 �����(Oee) = 650
 ����(Oee) = 550
 ����(Oee) = 650
 ��Ʈ��(Oee) = 650

ElseIf Oee = 665 Then
 �̸�(Oee) = "�赿��"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 ����(Oee) = 1
 ���ݷ�(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 700
 �����(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 650
 ��Ʈ��(Oee) = 600

ElseIf Oee = 666 Then
 �̸�(Oee) = "�輺��"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 ����(Oee) = 1
 ���ݷ�(Oee) = 550
 ����(Oee) = 600
 ����(Oee) = 550
 ����(Oee) = 700
 �����(Oee) = 700
 ����(Oee) = 600
 ����(Oee) = 600
 ��Ʈ��(Oee) = 650

ElseIf Oee = 667 Then
 �̸�(Oee) = "�豸��"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 ����(Oee) = 3
 ���ݷ�(Oee) = 750
 ����(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 750
 �����(Oee) = 650
 ����(Oee) = 700
 ����(Oee) = 600
 ��Ʈ��(Oee) = 600

ElseIf Oee = 668 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 ����(Oee) = 3
 ���ݷ�(Oee) = 650
 ����(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 900
 �����(Oee) = 550
 ����(Oee) = 650
 ����(Oee) = 700
 ��Ʈ��(Oee) = 650

ElseIf Oee = 669 Then
 �̸�(Oee) = "����ȯ1"
 ��ũ(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 ����(Oee) = 2
 ���ݷ�(Oee) = 700
 ����(Oee) = 750
 ����(Oee) = 850
 ����(Oee) = 750
 �����(Oee) = 700
 ����(Oee) = 800
 ����(Oee) = 800
 ��Ʈ��(Oee) = 800

ElseIf Oee = 670 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 ����(Oee) = 2
 ���ݷ�(Oee) = 950
 ����(Oee) = 950
 ����(Oee) = 500
 ����(Oee) = 600
 �����(Oee) = 550
 ����(Oee) = 550
 ����(Oee) = 950
 ��Ʈ��(Oee) = 950

ElseIf Oee = 671 Then
 �̸�(Oee) = "�ڼ���"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 ����(Oee) = 2
 ���ݷ�(Oee) = 800
 ����(Oee) = 700
 ����(Oee) = 600
 ����(Oee) = 750
 �����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 650
 ��Ʈ��(Oee) = 700

ElseIf Oee = 672 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 ����(Oee) = 3
 ���ݷ�(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 700
 ����(Oee) = 650
 �����(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 600
 ��Ʈ��(Oee) = 650

ElseIf Oee = 673 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 ����(Oee) = 1
 ���ݷ�(Oee) = 600
 ����(Oee) = 500
 ����(Oee) = 500
 ����(Oee) = 450
 �����(Oee) = 450
 ����(Oee) = 500
 ����(Oee) = 450
 ��Ʈ��(Oee) = 650

ElseIf Oee = 674 Then
 �̸�(Oee) = "�̽���"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 ����(Oee) = 1
 ���ݷ�(Oee) = 700
 ����(Oee) = 500
 ����(Oee) = 600
 ����(Oee) = 700
 �����(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 650
 ��Ʈ��(Oee) = 600

ElseIf Oee = 675 Then
 �̸�(Oee) = "����ȣ"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 ����(Oee) = 3
 ���ݷ�(Oee) = 650
 ����(Oee) = 550
 ����(Oee) = 500
 ����(Oee) = 600
 �����(Oee) = 500
 ����(Oee) = 550
 ����(Oee) = 550
 ��Ʈ��(Oee) = 500

ElseIf Oee = 676 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "STX"
 ����(Oee) = 2
 ���ݷ�(Oee) = 700
 ����(Oee) = 750
 ����(Oee) = 550
 ����(Oee) = 750
 �����(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 600
 ��Ʈ��(Oee) = 700

ElseIf Oee = 677 Then
 �̸�(Oee) = "�赿��"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 600
 ����(Oee) = 650
 �����(Oee) = 600
 ����(Oee) = 550
 ����(Oee) = 550
 ��Ʈ��(Oee) = 650

ElseIf Oee = 678 Then
 �̸�(Oee) = "����"
 ��ũ(Oee) = "Special"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 2
 ���ݷ�(Oee) = 650
 ����(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 750
 �����(Oee) = 750
 ����(Oee) = 750
 ����(Oee) = 700
 ��Ʈ��(Oee) = 800

ElseIf Oee = 679 Then
 �̸�(Oee) = "���ö"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 2
 ���ݷ�(Oee) = 700
 ����(Oee) = 600
 ����(Oee) = 550
 ����(Oee) = 700
 �����(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 550
 ��Ʈ��(Oee) = 700

ElseIf Oee = 680 Then
 �̸�(Oee) = "�����"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 3
 ���ݷ�(Oee) = 750
 ����(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 800
 �����(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 650
 ��Ʈ��(Oee) = 700

ElseIf Oee = 681 Then
 �̸�(Oee) = "�迵��"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 750
 ����(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 600
 �����(Oee) = 550
 ����(Oee) = 550
 ����(Oee) = 600
 ��Ʈ��(Oee) = 600

ElseIf Oee = 682 Then
 �̸�(Oee) = "���ر�"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 550
 ����(Oee) = 650
 �����(Oee) = 500
 ����(Oee) = 550
 ����(Oee) = 550
 ��Ʈ��(Oee) = 600

ElseIf Oee = 683 Then
 �̸�(Oee) = "�ڴ븸"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 3
 ���ݷ�(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 700
 �����(Oee) = 500
 ����(Oee) = 600
 ����(Oee) = 600
 ��Ʈ��(Oee) = 700

ElseIf Oee = 684 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 3
 ���ݷ�(Oee) = 850
 ����(Oee) = 750
 ����(Oee) = 700
 ����(Oee) = 800
 �����(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 700
 ��Ʈ��(Oee) = 950

ElseIf Oee = 685 Then
 �̸�(Oee) = "�̵���"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 600
 �����(Oee) = 700
 ����(Oee) = 550
 ����(Oee) = 550
 ��Ʈ��(Oee) = 550

ElseIf Oee = 686 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 2
 ���ݷ�(Oee) = 750
 ����(Oee) = 550
 ����(Oee) = 500
 ����(Oee) = 650
 �����(Oee) = 450
 ����(Oee) = 600
 ����(Oee) = 600
 ��Ʈ��(Oee) = 700

ElseIf Oee = 687 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 650
 �����(Oee) = 550
 ����(Oee) = 700
 ����(Oee) = 600
 ��Ʈ��(Oee) = 800

ElseIf Oee = 688 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 750
 �����(Oee) = 750
 ����(Oee) = 700
 ����(Oee) = 700
 ��Ʈ��(Oee) = 600

ElseIf Oee = 689 Then
 �̸�(Oee) = "�ѻ��"
 ��ũ(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "����"
 ����(Oee) = 2
 ���ݷ�(Oee) = 950
 ����(Oee) = 950
 ����(Oee) = 700
 ����(Oee) = 600
 �����(Oee) = 550
 ����(Oee) = 700
 ����(Oee) = 950
 ��Ʈ��(Oee) = 950

ElseIf Oee = 690 Then
 �̸�(Oee) = "���α�"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 ����(Oee) = 1
 ���ݷ�(Oee) = 700
 ����(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 700
 �����(Oee) = 600
 ����(Oee) = 700
 ����(Oee) = 650
 ��Ʈ��(Oee) = 700

ElseIf Oee = 691 Then
 �̸�(Oee) = "���ÿ�"
 ��ũ(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 ����(Oee) = 3
 ���ݷ�(Oee) = 600
 ����(Oee) = 950
 ����(Oee) = 700
 ����(Oee) = 700
 �����(Oee) = 700
 ����(Oee) = 900
 ����(Oee) = 800
 ��Ʈ��(Oee) = 750

ElseIf Oee = 692 Then
 �̸�(Oee) = "�����"
 ��ũ(Oee) = "Special"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 ����(Oee) = 3
 ���ݷ�(Oee) = 800
 ����(Oee) = 750
 ����(Oee) = 700
 ����(Oee) = 850
 �����(Oee) = 550
 ����(Oee) = 550
 ����(Oee) = 800
 ��Ʈ��(Oee) = 600

ElseIf Oee = 693 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 ����(Oee) = 2
 ���ݷ�(Oee) = 800
 ����(Oee) = 700
 ����(Oee) = 550
 ����(Oee) = 700
 �����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 750
 ��Ʈ��(Oee) = 850

ElseIf Oee = 694 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 ����(Oee) = 2
 ���ݷ�(Oee) = 750
 ����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 650
 �����(Oee) = 600
 ����(Oee) = 550
 ����(Oee) = 500
 ��Ʈ��(Oee) = 700

ElseIf Oee = 695 Then
 �̸�(Oee) = "�̽¼�"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 ����(Oee) = 2
 ���ݷ�(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 550
 ����(Oee) = 750
 �����(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 600
 ��Ʈ��(Oee) = 650

ElseIf Oee = 696 Then
 �̸�(Oee) = "�ӿ�ȯ"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 ����(Oee) = 1
 ���ݷ�(Oee) = 600
 ����(Oee) = 650
 ����(Oee) = 800
 ����(Oee) = 600
 �����(Oee) = 450
 ����(Oee) = 550
 ����(Oee) = 650
 ��Ʈ��(Oee) = 700

ElseIf Oee = 697 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Elite"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 ����(Oee) = 1
 ���ݷ�(Oee) = 700
 ����(Oee) = 950
 ����(Oee) = 800
 ����(Oee) = 950
 �����(Oee) = 950
 ����(Oee) = 800
 ����(Oee) = 800
 ��Ʈ��(Oee) = 650

ElseIf Oee = 698 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 ����(Oee) = 1
 ���ݷ�(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 550
 ����(Oee) = 600
 �����(Oee) = 600
 ����(Oee) = 550
 ����(Oee) = 500
 ��Ʈ��(Oee) = 600

ElseIf Oee = 697 Then
 �̸�(Oee) = "����ö"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 ����(Oee) = 2
 ���ݷ�(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 700
 �����(Oee) = 550
 ����(Oee) = 600
 ����(Oee) = 550
 ��Ʈ��(Oee) = 550

ElseIf Oee = 698 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 ����(Oee) = 3
 ���ݷ�(Oee) = 500
 ����(Oee) = 550
 ����(Oee) = 600
 ����(Oee) = 650
 �����(Oee) = 550
 ����(Oee) = 500
 ����(Oee) = 500
 ��Ʈ��(Oee) = 550

ElseIf Oee = 699 Then
 �̸�(Oee) = "��ȣ��"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "SK"
 ����(Oee) = 1
 ���ݷ�(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 500
 ����(Oee) = 650
 �����(Oee) = 650
 ����(Oee) = 500
 ����(Oee) = 650
 ��Ʈ��(Oee) = 600

ElseIf Oee = 700 Then
 �̸�(Oee) = "����"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 ����(Oee) = 2
 ���ݷ�(Oee) = 800
 ����(Oee) = 750
 ����(Oee) = 500
 ����(Oee) = 650
 �����(Oee) = 550
 ����(Oee) = 600
 ����(Oee) = 700
 ��Ʈ��(Oee) = 700

ElseIf Oee = 701 Then
 �̸�(Oee) = "�赿��"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 ����(Oee) = 2
 ���ݷ�(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 650
 ����(Oee) = 700
 �����(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 700
 ��Ʈ��(Oee) = 600

ElseIf Oee = 702 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 ����(Oee) = 3
 ���ݷ�(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 600
 ����(Oee) = 750
 �����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 650
 ��Ʈ��(Oee) = 650

ElseIf Oee = 703 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 ����(Oee) = 2
 ���ݷ�(Oee) = 650
 ����(Oee) = 550
 ����(Oee) = 550
 ����(Oee) = 600
 �����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 550
 ��Ʈ��(Oee) = 700

ElseIf Oee = 704 Then
 �̸�(Oee) = "�ڼ���"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 ����(Oee) = 3
 ���ݷ�(Oee) = 700
 ����(Oee) = 550
 ����(Oee) = 550
 ����(Oee) = 750
 �����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 650
 ��Ʈ��(Oee) = 600

ElseIf Oee = 705 Then
 �̸�(Oee) = "����ȣ"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 ����(Oee) = 3
 ���ݷ�(Oee) = 900
 ����(Oee) = 600
 ����(Oee) = 700
 ����(Oee) = 900
 �����(Oee) = 500
 ����(Oee) = 550
 ����(Oee) = 550
 ��Ʈ��(Oee) = 550

ElseIf Oee = 706 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Normal"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 ����(Oee) = 2
 ���ݷ�(Oee) = 700
 ����(Oee) = 600
 ����(Oee) = 750
 ����(Oee) = 600
 �����(Oee) = 600
 ����(Oee) = 550
 ����(Oee) = 600
 ��Ʈ��(Oee) = 650

ElseIf Oee = 707 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Special"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 ����(Oee) = 1
 ���ݷ�(Oee) = 700
 ����(Oee) = 750
 ����(Oee) = 700
 ����(Oee) = 850
 �����(Oee) = 750
 ����(Oee) = 750
 ����(Oee) = 700
 ��Ʈ��(Oee) = 700

ElseIf Oee = 708 Then
 �̸�(Oee) = "����ȣ"
 ��ũ(Oee) = "Rare"
 OYear(Oee) = "<10>"
 Team(Oee) = "MBC"
 ����(Oee) = 1
 ���ݷ�(Oee) = 900
 ����(Oee) = 800
 ����(Oee) = 750
 ����(Oee) = 600
 �����(Oee) = 600
 ����(Oee) = 800
 ����(Oee) = 750
 ��Ʈ��(Oee) = 800
ElseIf Oee = 709 Then
 �̸�(Oee) = "Mystery"
 ��ũ(Oee) = "Unique"
 OYear(Oee) = "<11>"
 Team(Oee) = "Mystar"
 ����(Oee) = 2
 ���ݷ�(Oee) = 900
 �����(Oee) = 750
 ����(Oee) = 850
 ����(Oee) = 900
 ����(Oee) = 800
 ��Ʈ��(Oee) = 750
 ����(Oee) = 600
 ����(Oee) = 850
ElseIf Oee = 710 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Rare"
 OYear(Oee) = "<11>"
 Team(Oee) = "Mystar"
 ����(Oee) = 3
 ���ݷ�(Oee) = 750
 �����(Oee) = 850
 ����(Oee) = 650
 ����(Oee) = 750
 ����(Oee) = 700
 ��Ʈ��(Oee) = 750
 ����(Oee) = 750
 ����(Oee) = 850
ElseIf Oee = 711 Then
 �̸�(Oee) = "Turtle"
 ��ũ(Oee) = "Special"
 OYear(Oee) = "<10>"
 Team(Oee) = "Mystar"
 ����(Oee) = 3
 ���ݷ�(Oee) = 950
 �����(Oee) = 700
 ����(Oee) = 500
 ����(Oee) = 950
 ����(Oee) = 500
 ��Ʈ��(Oee) = 500
 ����(Oee) = 750
 ����(Oee) = 750
ElseIf Oee = 712 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Special"
 OYear(Oee) = "<11>"
 Team(Oee) = "Mystar"
 ����(Oee) = 3
 ���ݷ�(Oee) = 950
 �����(Oee) = 600
 ����(Oee) = 600
 ����(Oee) = 950
 ����(Oee) = 600
 ��Ʈ��(Oee) = 600
 ����(Oee) = 750
 ����(Oee) = 600
ElseIf Oee = 713 Then
 �̸�(Oee) = "�����"
 ��ũ(Oee) = "Unique"
 OYear(Oee) = "<11>"
 Team(Oee) = "Mystar"
 ����(Oee) = 1
 ���ݷ�(Oee) = 900
 ����(Oee) = 900
 ����(Oee) = 800
 ����(Oee) = 600
 �����(Oee) = 750
 ����(Oee) = 950
 ����(Oee) = 800
 ��Ʈ��(Oee) = 700
ElseIf Oee = 714 Then
 �̸�(Oee) = "�̼���[Ex]"
 ��ũ(Oee) = "Unique"
 OYear(Oee) = "<07>"
 Team(Oee) = "�Ｚ����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 900
 ����(Oee) = 800
 ����(Oee) = 900
 ����(Oee) = 700
 �����(Oee) = 650
 ����(Oee) = 750
 ����(Oee) = 850
 ��Ʈ��(Oee) = 850
ElseIf Oee = 715 Then
 �̸�(Oee) = "����"
 ��ũ(Oee) = "Unique"
 OYear(Oee) = "<06>"
 Team(Oee) = "KTF"
 ����(Oee) = 3
 ���ݷ�(Oee) = 700
 ����(Oee) = 800
 ����(Oee) = 950
 ����(Oee) = 850
 �����(Oee) = 850
 ����(Oee) = 700
 ����(Oee) = 800
 ��Ʈ��(Oee) = 750
End If

 ���(Oee) = 0
 �ؿ��(Oee) = 0
 �����(Oee) = 100
 A�¸�(Oee) = 0
 A�й�(Oee) = 0
 P�¸�(Oee) = 0
 P�й�(Oee) = 0
 T�¸�(Oee) = 0
 T�й�(Oee) = 0
 Z�¸�(Oee) = 0
 Z�й�(Oee) = 0
 T����(Oee) = 0
 Z����(Oee) = 0
 P����(Oee) = 0
 A����(Oee) = 0
 T��(Oee) = "W"
 Z��(Oee) = "W"
 P��(Oee) = "W"
 A��(Oee) = "W"
Next Oee
 
Tim11.Enabled = True
Tim10.Enabled = False
End Sub

Private Sub Tim11_Timer()
For Oee = 1 To 118
If Oee = 1 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 500
����(Oee) = 500
����(Oee) = 500
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 550
ElseIf Oee = 2 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 500
����(Oee) = 450
�����(Oee) = 450
����(Oee) = 500
����(Oee) = 450
��Ʈ��(Oee) = 650
ElseIf Oee = 3 Then
�̸�(Oee) = "�輺��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 550
����(Oee) = 550
����(Oee) = 500
����(Oee) = 500
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 650
ElseIf Oee = 4 Then
�̸�(Oee) = "�鵿��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 550
����(Oee) = 550
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 500
ElseIf Oee = 5 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 500
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 550
ElseIf Oee = 6 Then
�̸�(Oee) = "ȫ��ǥ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 500
����(Oee) = 550
����(Oee) = 500
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 500
ElseIf Oee = 7 Then
�̸�(Oee) = "��ȣ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 650
����(Oee) = 600
����(Oee) = 600
�����(Oee) = 400
����(Oee) = 450
����(Oee) = 500
��Ʈ��(Oee) = 600
ElseIf Oee = 8 Then
�̸�(Oee) = "��α�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 450
����(Oee) = 550
�����(Oee) = 450
����(Oee) = 450
����(Oee) = 500
��Ʈ��(Oee) = 600
ElseIf Oee = 9 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBc"
����(Oee) = 2
���ݷ�(Oee) = 550
����(Oee) = 500
����(Oee) = 600
����(Oee) = 500
�����(Oee) = 450
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 600
ElseIf Oee = 10 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
����(Oee) = 1
���ݷ�(Oee) = 550
����(Oee) = 550
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 500
ElseIf Oee = 11 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 500
����(Oee) = 500
�����(Oee) = 550
����(Oee) = 500
����(Oee) = 550
��Ʈ��(Oee) = 600
ElseIf Oee = 12 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 500
����(Oee) = 550
��Ʈ��(Oee) = 550
ElseIf Oee = 13 Then
�̸�(Oee) = "�鵿��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "ȭ��"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 500
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 500
����(Oee) = 500
��Ʈ��(Oee) = 650
ElseIf Oee = 14 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "ȭ��"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 500
����(Oee) = 500
����(Oee) = 550
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 600
ElseIf Oee = 15 Then
�̸�(Oee) = "�ۿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 550
����(Oee) = 500
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 500
ElseIf Oee = 16 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
����(Oee) = 1
���ݷ�(Oee) = 500
����(Oee) = 550
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 500
ElseIf Oee = 17 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 500
����(Oee) = 750
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 600
ElseIf Oee = 18 Then
�̸�(Oee) = "���⼮"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "�Ｚ����"
����(Oee) = 1
���ݷ�(Oee) = 500
����(Oee) = 550
����(Oee) = 500
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 500
��Ʈ��(Oee) = 550
ElseIf Oee = 19 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 550
����(Oee) = 500
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 500
ElseIf Oee = 20 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 600
ElseIf Oee = 21 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "ȭ��"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 500
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 500
����(Oee) = 650
��Ʈ��(Oee) = 550
ElseIf Oee = 22 Then
�̸�(Oee) = "�ϴ�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "ȭ��"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 650
����(Oee) = 550
�����(Oee) = 600
����(Oee) = 500
����(Oee) = 550
��Ʈ��(Oee) = 600
ElseIf Oee = 23 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 650
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 600
ElseIf Oee = 24 Then
�̸�(Oee) = "�ֿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 700
ElseIf Oee = 25 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 550
ElseIf Oee = 26 Then
�̸�(Oee) = "�ֿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 700
ElseIf Oee = 27 Then
�̸�(Oee) = "�輺��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 550
����(Oee) = 600
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 650
ElseIf Oee = 28 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 650
����(Oee) = 600
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 600
ElseIf Oee = 29 Then
�̸�(Oee) = "���ر�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 600
ElseIf Oee = 30 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 600
ElseIf Oee = 31 Then
�̸�(Oee) = "���¼�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "ȭ��"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 600
ElseIf Oee = 32 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 500
ElseIf Oee = 33 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 600
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 500
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 700
ElseIf Oee = 34 Then
�̸�(Oee) = "�ѵο�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 650
ElseIf Oee = 35 Then
�̸�(Oee) = "�Ǽ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 600
ElseIf Oee = 36 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 550
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 550
��Ʈ��(Oee) = 550
ElseIf Oee = 37 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 550
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 550
��Ʈ��(Oee) = 600
ElseIf Oee = 38 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 550
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 650
ElseIf Oee = 39 Then
�̸�(Oee) = "�̿���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 500
��Ʈ��(Oee) = 550
ElseIf Oee = 40 Then
�̸�(Oee) = "�ּ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 650
ElseIf Oee = 41 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 650
����(Oee) = 700
����(Oee) = 800
�����(Oee) = 800
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 600
ElseIf Oee = 42 Then
�̸�(Oee) = "���翵"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 650
����(Oee) = 800
�����(Oee) = 650
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 550
ElseIf Oee = 43 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 650
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 500
����(Oee) = 600
����(Oee) = 550
��Ʈ��(Oee) = 750
ElseIf Oee = 44 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
����(Oee) = 3
���ݷ�(Oee) = 550
����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 650
ElseIf Oee = 45 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 650
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 750
ElseIf Oee = 46 Then
�̸�(Oee) = "Ȳ����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 550
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 600
ElseIf Oee = 47 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "�Ｚ����"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 550
ElseIf Oee = 48 Then
�̸�(Oee) = "�ڴ�ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "�Ｚ����"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 700
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 600
ElseIf Oee = 49 Then
�̸�(Oee) = "���±�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "�Ｚ����"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 550
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 650
ElseIf Oee = 50 Then
�̸�(Oee) = "�赵��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 800
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 650
ElseIf Oee = 51 Then
�̸�(Oee) = "�赿��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 700
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 700
����(Oee) = 650
��Ʈ��(Oee) = 600
ElseIf Oee = 52 Then
�̸�(Oee) = "�Ŵ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 700
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 600
ElseIf Oee = 53 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 600
ElseIf Oee = 54 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 700
ElseIf Oee = 55 Then
�̸�(Oee) = "�̽¼�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 650
����(Oee) = 750
��Ʈ��(Oee) = 750
ElseIf Oee = 56 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 550
����(Oee) = 600
��Ʈ��(Oee) = 650
ElseIf Oee = 57 Then
�̸�(Oee) = "��ȣ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 600
����(Oee) = 550
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 550
����(Oee) = 700
��Ʈ��(Oee) = 600
ElseIf Oee = 58 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 800
����(Oee) = 750
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 650
ElseIf Oee = 59 Then
�̸�(Oee) = "�赿��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 650
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 600
ElseIf Oee = 60 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 700
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 700
ElseIf Oee = 61 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 650
����(Oee) = 750
�����(Oee) = 500
����(Oee) = 550
����(Oee) = 650
��Ʈ��(Oee) = 600
ElseIf Oee = 62 Then
�̸�(Oee) = "���±�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "ȭ��"
����(Oee) = 3
���ݷ�(Oee) = 800
����(Oee) = 600
����(Oee) = 650
����(Oee) = 750
�����(Oee) = 550
����(Oee) = 700
����(Oee) = 600
��Ʈ��(Oee) = 650
ElseIf Oee = 63 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "ȭ��"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 700
����(Oee) = 600
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 650
ElseIf Oee = 64 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "ȭ��"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 650
����(Oee) = 650
����(Oee) = 750
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 600
ElseIf Oee = 65 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 500
����(Oee) = 700
�����(Oee) = 650
����(Oee) = 650
����(Oee) = 550
��Ʈ��(Oee) = 650
ElseIf Oee = 66 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 700
ElseIf Oee = 67 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 750
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 750
��Ʈ��(Oee) = 600
ElseIf Oee = 68 Then
�̸�(Oee) = "���α�"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 600
����(Oee) = 650
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 700
����(Oee) = 650
��Ʈ��(Oee) = 600
ElseIf Oee = 69 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 700
ElseIf Oee = 70 Then
�̸�(Oee) = "�ڿ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 650
ElseIf Oee = 71 Then
�̸�(Oee) = "�ռ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 650
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 700
ElseIf Oee = 72 Then
�̸�(Oee) = "�ȱ�ȿ"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 650
ElseIf Oee = 73 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 650
����(Oee) = 600
�����(Oee) = 500
����(Oee) = 700
����(Oee) = 600
��Ʈ��(Oee) = 800
ElseIf Oee = 74 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 700
����(Oee) = 700
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 550
ElseIf Oee = 75 Then
�̸�(Oee) = "�̿�ȣ1"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 550
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 550
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 650
ElseIf Oee = 76 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 650
��Ʈ��(Oee) = 650
ElseIf Oee = 77 Then
�̸�(Oee) = "��뿱"
��ũ(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
����(Oee) = 3
���ݷ�(Oee) = 850
����(Oee) = 750
����(Oee) = 600
����(Oee) = 800
�����(Oee) = 850
����(Oee) = 650
����(Oee) = 750
��Ʈ��(Oee) = 750
ElseIf Oee = 78 Then
�̸�(Oee) = "�輺��"
��ũ(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
����(Oee) = 2
���ݷ�(Oee) = 800
����(Oee) = 700
����(Oee) = 550
����(Oee) = 800
�����(Oee) = 800
����(Oee) = 700
����(Oee) = 600
��Ʈ��(Oee) = 650
ElseIf Oee = 79 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "�Ｚ����"
����(Oee) = 3
���ݷ�(Oee) = 600
����(Oee) = 650
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 750
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 700
ElseIf Oee = 80 Then
�̸�(Oee) = "����ȯ"
��ũ(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "�Ｚ����"
����(Oee) = 2
���ݷ�(Oee) = 850
����(Oee) = 600
����(Oee) = 700
����(Oee) = 800
�����(Oee) = 800
����(Oee) = 700
����(Oee) = 700
��Ʈ��(Oee) = 850
ElseIf Oee = 81 Then
�̸�(Oee) = "�㿵��"
��ũ(Oee) = "Unique"
OYear(Oee) = "<11>"
Team(Oee) = "�Ｚ����"
����(Oee) = 3
���ݷ�(Oee) = 900
����(Oee) = 700
����(Oee) = 700
����(Oee) = 950
�����(Oee) = 850
����(Oee) = 750
����(Oee) = 750
��Ʈ��(Oee) = 950
ElseIf Oee = 82 Then
�̸�(Oee) = "�豸��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 700
����(Oee) = 700
����(Oee) = 700
�����(Oee) = 650
����(Oee) = 700
����(Oee) = 600
��Ʈ��(Oee) = 650
ElseIf Oee = 83 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
����(Oee) = 3
���ݷ�(Oee) = 650
����(Oee) = 650
����(Oee) = 650
����(Oee) = 900
�����(Oee) = 550
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 650
ElseIf Oee = 84 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 850
����(Oee) = 750
����(Oee) = 650
����(Oee) = 600
�����(Oee) = 500
����(Oee) = 600
����(Oee) = 750
��Ʈ��(Oee) = 900
ElseIf Oee = 85 Then
�̸�(Oee) = "�̽���"
��ũ(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 600
����(Oee) = 650
����(Oee) = 800
�����(Oee) = 850
����(Oee) = 750
����(Oee) = 700
��Ʈ��(Oee) = 850
ElseIf Oee = 86 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 700
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 800
ElseIf Oee = 87 Then
�̸�(Oee) = "���ö"
��ũ(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 700
����(Oee) = 650
����(Oee) = 750
�����(Oee) = 800
����(Oee) = 700
����(Oee) = 650
��Ʈ��(Oee) = 700
ElseIf Oee = 88 Then
�̸�(Oee) = "�ڻ��"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 600
����(Oee) = 600
����(Oee) = 800
�����(Oee) = 650
����(Oee) = 700
����(Oee) = 700
��Ʈ��(Oee) = 700
ElseIf Oee = 89 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 800
ElseIf Oee = 90 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 750
����(Oee) = 750
����(Oee) = 800
�����(Oee) = 600
����(Oee) = 800
����(Oee) = 700
��Ʈ��(Oee) = 800
ElseIf Oee = 91 Then
�̸�(Oee) = "�����"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
����(Oee) = 3
���ݷ�(Oee) = 800
����(Oee) = 650
����(Oee) = 650
����(Oee) = 900
�����(Oee) = 500
����(Oee) = 600
����(Oee) = 750
��Ʈ��(Oee) = 650
ElseIf Oee = 92 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
����(Oee) = 2
���ݷ�(Oee) = 800
����(Oee) = 700
����(Oee) = 550
����(Oee) = 650
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 750
��Ʈ��(Oee) = 850
ElseIf Oee = 93 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 700
����(Oee) = 600
����(Oee) = 800
�����(Oee) = 850
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 750
ElseIf Oee = 94 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 650
����(Oee) = 650
����(Oee) = 900
�����(Oee) = 550
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 750
ElseIf Oee = 95 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
����(Oee) = 3
���ݷ�(Oee) = 800
����(Oee) = 650
����(Oee) = 600
����(Oee) = 900
�����(Oee) = 700
����(Oee) = 600
����(Oee) = 600
��Ʈ��(Oee) = 750
ElseIf Oee = 96 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "ȭ��"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 700
����(Oee) = 750
����(Oee) = 700
�����(Oee) = 750
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 650
ElseIf Oee = 97 Then
�̸�(Oee) = "���ؿ�"
��ũ(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "ȭ��"
����(Oee) = 2
���ݷ�(Oee) = 950
����(Oee) = 700
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 800
��Ʈ��(Oee) = 950
ElseIf Oee = 98 Then
�̸�(Oee) = "�Ż�"
��ũ(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 750
����(Oee) = 800
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 700
����(Oee) = 750
��Ʈ��(Oee) = 850
ElseIf Oee = 99 Then
�̸�(Oee) = "�̰��"
��ũ(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
����(Oee) = 3
���ݷ�(Oee) = 800
����(Oee) = 600
����(Oee) = 800
����(Oee) = 900
�����(Oee) = 700
����(Oee) = 550
����(Oee) = 800
��Ʈ��(Oee) = 850
ElseIf Oee = 100 Then
�̸�(Oee) = "����ö"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
����(Oee) = 3
���ݷ�(Oee) = 700
����(Oee) = 600
����(Oee) = 650
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 600
����(Oee) = 700
��Ʈ��(Oee) = 700
ElseIf Oee = 101 Then
�̸�(Oee) = "����ȭ"
��ũ(Oee) = "Unique"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
����(Oee) = 3
���ݷ�(Oee) = 800
����(Oee) = 850
����(Oee) = 700
����(Oee) = 850
�����(Oee) = 800
����(Oee) = 850
����(Oee) = 750
��Ʈ��(Oee) = 800
ElseIf Oee = 102 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 750
����(Oee) = 600
����(Oee) = 800
�����(Oee) = 800
����(Oee) = 700
����(Oee) = 650
��Ʈ��(Oee) = 600
ElseIf Oee = 103 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 850
����(Oee) = 700
����(Oee) = 600
����(Oee) = 700
�����(Oee) = 600
����(Oee) = 650
����(Oee) = 650
��Ʈ��(Oee) = 650
ElseIf Oee = 104 Then
�̸�(Oee) = "�̼���"
��ũ(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 750
����(Oee) = 700
����(Oee) = 800
����(Oee) = 600
�����(Oee) = 550
����(Oee) = 750
����(Oee) = 700
��Ʈ��(Oee) = 750
ElseIf Oee = 105 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 750
����(Oee) = 700
����(Oee) = 800
�����(Oee) = 800
����(Oee) = 750
����(Oee) = 650
��Ʈ��(Oee) = 750
ElseIf Oee = 106 Then
�̸�(Oee) = "�ų뿭"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 750
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 700
����(Oee) = 650
��Ʈ��(Oee) = 750
ElseIf Oee = 107 Then
�̸�(Oee) = "�̿���"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 650
����(Oee) = 700
����(Oee) = 750
�����(Oee) = 500
����(Oee) = 600
����(Oee) = 750
��Ʈ��(Oee) = 700
ElseIf Oee = 108 Then
�̸�(Oee) = "���¾�"
��ũ(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 700
����(Oee) = 750
����(Oee) = 650
�����(Oee) = 700
����(Oee) = 750
����(Oee) = 800
��Ʈ��(Oee) = 750
ElseIf Oee = 109 Then
�̸�(Oee) = "�̿�ȣ"
��ũ(Oee) = "Legend"
OYear(Oee) = "<11>"
Team(Oee) = "KT"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 700
����(Oee) = 850
����(Oee) = 950
�����(Oee) = 950
����(Oee) = 800
����(Oee) = 950
��Ʈ��(Oee) = 800

ElseIf Oee = 110 Then
�̸�(Oee) = "�ۺ���"
��ũ(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "�Ｚ����"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 700
����(Oee) = 700
����(Oee) = 800
�����(Oee) = 750
����(Oee) = 750
����(Oee) = 800
��Ʈ��(Oee) = 750
ElseIf Oee = 111 Then
�̸�(Oee) = "����ȯ1"
��ũ(Oee) = "Normal"
OYear(Oee) = "<11>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 600
����(Oee) = 750
����(Oee) = 800
����(Oee) = 700
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 600
��Ʈ��(Oee) = 600
ElseIf Oee = 112 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "����"
����(Oee) = 2
���ݷ�(Oee) = 650
����(Oee) = 750
����(Oee) = 700
����(Oee) = 800
�����(Oee) = 800
����(Oee) = 850
����(Oee) = 750
��Ʈ��(Oee) = 900
ElseIf Oee = 113 Then
�̸�(Oee) = "���ÿ�"
��ũ(Oee) = "Legend"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 950
����(Oee) = 750
����(Oee) = 900
�����(Oee) = 800
����(Oee) = 950
����(Oee) = 750
��Ʈ��(Oee) = 950
ElseIf Oee = 114 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Unique"
OYear(Oee) = "<11>"
Team(Oee) = "SK"
����(Oee) = 1
���ݷ�(Oee) = 700
����(Oee) = 950
����(Oee) = 850
����(Oee) = 850
�����(Oee) = 900
����(Oee) = 850
����(Oee) = 800
��Ʈ��(Oee) = 650
ElseIf Oee = 115 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Special"
OYear(Oee) = "<11>"
Team(Oee) = "MBC"
����(Oee) = 1
���ݷ�(Oee) = 650
����(Oee) = 750
����(Oee) = 700
����(Oee) = 800
�����(Oee) = 800
����(Oee) = 800
����(Oee) = 750
��Ʈ��(Oee) = 700
ElseIf Oee = 116 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Rare"
OYear(Oee) = "<11>"
Team(Oee) = "ȭ��"
����(Oee) = 2
���ݷ�(Oee) = 850
����(Oee) = 750
����(Oee) = 650
����(Oee) = 750
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 800
��Ʈ��(Oee) = 850
ElseIf Oee = 117 Then
�̸�(Oee) = "�ŵ���"
��ũ(Oee) = "Elite"
OYear(Oee) = "<11>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 900
����(Oee) = 850
����(Oee) = 600
����(Oee) = 900
�����(Oee) = 850
����(Oee) = 700
����(Oee) = 950
��Ʈ��(Oee) = 950
ElseIf Oee = 118 Then
 �̸�(Oee) = "�̱⼮"
 OYear(Oee) = "<99>"
 ��ũ(Oee) = "Legend"
 Team(Oee) = "�ڷ����"
 ����(Oee) = 1
 ���ݷ�(Oee) = 800
 ����(Oee) = 800
 ����(Oee) = 900
 ����(Oee) = 900
 �����(Oee) = 800
 ����(Oee) = 800
 ����(Oee) = 900
 ��Ʈ��(Oee) = 900
End If
 ���(Oee) = 0
 �ؿ��(Oee) = 0
 �����(Oee) = 100
 A�¸�(Oee) = 0
 A�й�(Oee) = 0
 P�¸�(Oee) = 0
 P�й�(Oee) = 0
 T�¸�(Oee) = 0
 T�й�(Oee) = 0
 Z�¸�(Oee) = 0
 Z�й�(Oee) = 0
 T����(Oee) = 0
 Z����(Oee) = 0
 P����(Oee) = 0
 A����(Oee) = 0
 T��(Oee) = "W"
 Z��(Oee) = "W"
 P��(Oee) = "W"
 A��(Oee) = "W"
Next Oee

TimElse.Enabled = True
Tim11.Enabled = False
End Sub

Private Sub Tim12_Timer()
For Oee = 724 To 800
    If Oee = 724 Then
        �̸�(Oee) = "������"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "8th"
        ����(Oee) = 1
        ���ݷ�(Oee) = 850
        ����(Oee) = 600
        ����(Oee) = 600
        ����(Oee) = 600
        �����(Oee) = 600
        ����(Oee) = 700
        ����(Oee) = 700
        ��Ʈ��(Oee) = 850
    ElseIf Oee = 725 Then
        �̸�(Oee) = "���¾�"
        ��ũ(Oee) = "Special"
        OYear(Oee) = "<12>"
        Team(Oee) = "8th"
        ����(Oee) = 1
        ���ݷ�(Oee) = 650
        ����(Oee) = 650
        ����(Oee) = 650
        ����(Oee) = 600
        �����(Oee) = 700
        ����(Oee) = 800
        ����(Oee) = 900
        ��Ʈ��(Oee) = 650
    ElseIf Oee = 726 Then
        �̸�(Oee) = "�赵��"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "8th"
        ����(Oee) = 1
        ���ݷ�(Oee) = 600
    
        ����(Oee) = 550
    
        ����(Oee) = 550
    
        ����(Oee) = 650
    
        �����(Oee) = 700
    
        ����(Oee) = 600
    
        ����(Oee) = 500
    
        ��Ʈ��(Oee) = 500
    
    ElseIf Oee = 727 Then
    
        �̸�(Oee) = "������"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "8th"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 650
    
        ����(Oee) = 800
    
        ����(Oee) = 600
    
        ����(Oee) = 750
    
        �����(Oee) = 650
    
        ����(Oee) = 800
    
        ����(Oee) = 650
    
        ��Ʈ��(Oee) = 650
    
    ElseIf Oee = 728 Then
    
        �̸�(Oee) = "�ڼ���"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "8th"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 850
    
        ����(Oee) = 600
    
        ����(Oee) = 600
    
        ����(Oee) = 850
    
        �����(Oee) = 700
    
        ����(Oee) = 600
    
        ����(Oee) = 700
    
        ��Ʈ��(Oee) = 600
    
    ElseIf Oee = 729 Then
    
        �̸�(Oee) = "�����"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "8th"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 650
    
        ����(Oee) = 550
    
        ����(Oee) = 650
    
        ����(Oee) = 550
    
        �����(Oee) = 600
    
        ����(Oee) = 700
    
        ����(Oee) = 500
    
        ��Ʈ��(Oee) = 500
    
    ElseIf Oee = 730 Then
    
        �̸�(Oee) = "������"
    
        ��ũ(Oee) = "Rare"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "8th"
    
        ����(Oee) = 2
    
        ���ݷ�(Oee) = 850
    
        ����(Oee) = 700
    
        ����(Oee) = 600
    
        ����(Oee) = 650
    
        �����(Oee) = 700
    
        ����(Oee) = 550
    
        ����(Oee) = 700
    
        ��Ʈ��(Oee) = 850
    
    ElseIf Oee = 731 Then
    
        �̸�(Oee) = "�鵿��"
    
        ��ũ(Oee) = "Rare"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 950
    
        ����(Oee) = 600
    
        ����(Oee) = 600
    
        ����(Oee) = 950
    
        �����(Oee) = 600
    
        ����(Oee) = 800
    
        ����(Oee) = 750
    
        ��Ʈ��(Oee) = 750
    
    ElseIf Oee = 732 Then
    
        �̸�(Oee) = "�̿�ȣ"
    
        ��ũ(Oee) = "Elite"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "KT"
    
        ����(Oee) = 1
    
        ���ݷ�(Oee) = 950
    
        ����(Oee) = 700
    
        ����(Oee) = 850
    
        ����(Oee) = 800
    
        �����(Oee) = 700
    
        ����(Oee) = 850
    
        ����(Oee) = 950
    
        ��Ʈ��(Oee) = 950
    
    ElseIf Oee = 733 Then
    
        �̸�(Oee) = "������"
    
        ��ũ(Oee) = "Elite"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "SK"
    
        ����(Oee) = 1
    
        ���ݷ�(Oee) = 950
    
        ����(Oee) = 800
    
        ����(Oee) = 850
    
        ����(Oee) = 800
    
        �����(Oee) = 850
    
        ����(Oee) = 850
    
        ����(Oee) = 850
    
        ��Ʈ��(Oee) = 850
    
    ElseIf Oee = 734 Then
    
        �̸�(Oee) = "�ۺ���"
    
        ��ũ(Oee) = "Elite"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "�Ｚ"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 700
    
        ����(Oee) = 900
    
        ����(Oee) = 850
    
        ����(Oee) = 800
    
        �����(Oee) = 850
    
        ����(Oee) = 800
    
        ����(Oee) = 850
    
        ��Ʈ��(Oee) = 950
    
    ElseIf Oee = 735 Then
    
        �̸�(Oee) = "���ö"
    
        ��ũ(Oee) = "Rare"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "����"
    
        ����(Oee) = 2
    
        ���ݷ�(Oee) = 800
    
        ����(Oee) = 800
    
        ����(Oee) = 700
    
        ����(Oee) = 900
    
        �����(Oee) = 900
    
        ����(Oee) = 700
    
        ����(Oee) = 800
    
        ��Ʈ��(Oee) = 800
    
    ElseIf Oee = 736 Then
    
        �̸�(Oee) = "�����"
    
        ��ũ(Oee) = "Rare"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "SK"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 600
    
        ����(Oee) = 950
    
        ����(Oee) = 700
    
        ����(Oee) = 600
    
        �����(Oee) = 850
    
        ����(Oee) = 950
    
        ����(Oee) = 850
    
        ��Ʈ��(Oee) = 800
    
    ElseIf Oee = 737 Then
    
        �̸�(Oee) = "�Ż�"
    
        ��ũ(Oee) = "Rare"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        ����(Oee) = 1
    
        ���ݷ�(Oee) = 900
    
        ����(Oee) = 750
    
        ����(Oee) = 850
    
        ����(Oee) = 750
    
        �����(Oee) = 850
    
        ����(Oee) = 800
    
        ����(Oee) = 650
    
        ��Ʈ��(Oee) = 850
    
    ElseIf Oee = 738 Then
    
        �̸�(Oee) = "������"
    
        ��ũ(Oee) = "Rare"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        ����(Oee) = 2
    
        ���ݷ�(Oee) = 700
    
        ����(Oee) = 750
    
        ����(Oee) = 850
    
        ����(Oee) = 750
    
        �����(Oee) = 900
    
        ����(Oee) = 700
    
        ����(Oee) = 850
    
        ��Ʈ��(Oee) = 800
    
    ElseIf Oee = 739 Then
    
        �̸�(Oee) = "���ÿ�"
    
        ��ũ(Oee) = "Unique"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "SK"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 750
    
        ����(Oee) = 900
    
        ����(Oee) = 700
    
        ����(Oee) = 750
    
        �����(Oee) = 800
    
        ����(Oee) = 950
    
        ����(Oee) = 800
    
        ��Ʈ��(Oee) = 850
    
    ElseIf Oee = 740 Then
    
        �̸�(Oee) = "��뿱"
    
        ��ũ(Oee) = "Unique"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "KT"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 800
    
        ����(Oee) = 800
    
        ����(Oee) = 800
    
        ����(Oee) = 850
    
        �����(Oee) = 750
    
        ����(Oee) = 800
    
        ����(Oee) = 800
    
        ��Ʈ��(Oee) = 800
    
    ElseIf Oee = 741 Then
    
        �̸�(Oee) = "�輺��"
    
        ��ũ(Oee) = "Rare"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        ����(Oee) = 1
    
        ���ݷ�(Oee) = 600
    
        ����(Oee) = 850
    
        ����(Oee) = 750
    
        ����(Oee) = 900
    
        �����(Oee) = 850
    
        ����(Oee) = 800
    
        ����(Oee) = 900
    
        ��Ʈ��(Oee) = 600
    
    ElseIf Oee = 742 Then
    
        �̸�(Oee) = "������"
    
        ��ũ(Oee) = "Rare"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "����"
    
        ����(Oee) = 2
    
        ���ݷ�(Oee) = 900
    
        ����(Oee) = 700
    
        ����(Oee) = 600
    
        ����(Oee) = 750
    
        �����(Oee) = 800
    
        ����(Oee) = 700
    
        ����(Oee) = 850
    
        ��Ʈ��(Oee) = 900
    
    ElseIf Oee = 743 Then
    
        �̸�(Oee) = "����ȯ1"
    
        ��ũ(Oee) = "Special"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        ����(Oee) = 2
    
        ���ݷ�(Oee) = 600
    
        ����(Oee) = 700
    
        ����(Oee) = 850
    
        ����(Oee) = 800
    
        �����(Oee) = 850
    
        ����(Oee) = 800
    
        ����(Oee) = 700
    
        ��Ʈ��(Oee) = 700
    
    ElseIf Oee = 744 Then
    
        �̸�(Oee) = "�㿵��"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "�Ｚ"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 700
    
        ����(Oee) = 600
    
        ����(Oee) = 650
    
        ����(Oee) = 750
    
        �����(Oee) = 750
    
        ����(Oee) = 650
    
        ����(Oee) = 700
    
        ��Ʈ��(Oee) = 750
    
    ElseIf Oee = 745 Then
    
        �̸�(Oee) = "������"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "����"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 750
    
        ����(Oee) = 650
    
        ����(Oee) = 700
    
        ����(Oee) = 750
    
        �����(Oee) = 750
    
        ����(Oee) = 600
    
        ����(Oee) = 700
    
        ��Ʈ��(Oee) = 650
    
    ElseIf Oee = 746 Then
    
        �̸�(Oee) = "����"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "����"
    
        ����(Oee) = 2
    
        ���ݷ�(Oee) = 850
    
        ����(Oee) = 700
    
        ����(Oee) = 500
    
        ����(Oee) = 650
    
        �����(Oee) = 600
    
        ����(Oee) = 650
    
        ����(Oee) = 800
    
        ��Ʈ��(Oee) = 800
    
    ElseIf Oee = 747 Then
    
        �̸�(Oee) = "����ȣ"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "����"
    
        ����(Oee) = 1
    
        ���ݷ�(Oee) = 850
    
        ����(Oee) = 500
    
        ����(Oee) = 700
    
        ����(Oee) = 500
    
        �����(Oee) = 600
    
        ����(Oee) = 700
    
        ����(Oee) = 850
    
        ��Ʈ��(Oee) = 850
    
    ElseIf Oee = 748 Then
    
        �̸�(Oee) = "���±�"
    
        ��ũ(Oee) = "Special"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "�Ｚ"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 650
    
        ����(Oee) = 600
    
        ����(Oee) = 650
    
        ����(Oee) = 750
    
        �����(Oee) = 700
    
        ����(Oee) = 600
    
        ����(Oee) = 800
    
        ��Ʈ��(Oee) = 850
    
    ElseIf Oee = 749 Then
    
        �̸�(Oee) = "������"
    
        ��ũ(Oee) = "Special"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "SK"
    
        ����(Oee) = 2
    
        ���ݷ�(Oee) = 700
    
        ����(Oee) = 750
    
        ����(Oee) = 650
    
        ����(Oee) = 800
    
        �����(Oee) = 700
    
        ����(Oee) = 700
    
        ����(Oee) = 650
    
        ��Ʈ��(Oee) = 650
    
    ElseIf Oee = 750 Then
    
        �̸�(Oee) = "�ŵ���"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        ����(Oee) = 2
    
        ���ݷ�(Oee) = 850
    
        ����(Oee) = 600
    
        ����(Oee) = 600
    
        ����(Oee) = 700
    
        �����(Oee) = 700
    
        ����(Oee) = 700
    
        ����(Oee) = 700
    
        ��Ʈ��(Oee) = 700
    
    ElseIf Oee = 751 Then
    
        �̸�(Oee) = "�̽���"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        ����(Oee) = 1
    
        ���ݷ�(Oee) = 600
    
        ����(Oee) = 800
    
        ����(Oee) = 700
    
        ����(Oee) = 800
    
        �����(Oee) = 750
    
        ����(Oee) = 500
    
        ����(Oee) = 700
    
        ��Ʈ��(Oee) = 700
    
    ElseIf Oee = 752 Then
    
        �̸�(Oee) = "�̰��"
    
        ��ũ(Oee) = "Special"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 850
    
        ����(Oee) = 650
    
        ����(Oee) = 600
    
        ����(Oee) = 850
    
        �����(Oee) = 600
    
        ����(Oee) = 700
    
        ����(Oee) = 650
    
        ��Ʈ��(Oee) = 650
    
    ElseIf Oee = 753 Then
    
        �̸�(Oee) = "�豸��"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "����"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 700
    
        ����(Oee) = 600
    
        ����(Oee) = 700
    
        ����(Oee) = 850
    
        �����(Oee) = 700
    
        ����(Oee) = 550
    
        ����(Oee) = 700
    
        ��Ʈ��(Oee) = 750
    
    ElseIf Oee = 754 Then
    
        �̸�(Oee) = "�ڴ�ȣ"
    
        ��ũ(Oee) = "Special"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "�Ｚ"
    
        ����(Oee) = 1
    
        ���ݷ�(Oee) = 950
    
        ����(Oee) = 800
    
        ����(Oee) = 600
    
        ����(Oee) = 700
    
        �����(Oee) = 750
    
        ����(Oee) = 650
    
        ����(Oee) = 800
    
        ��Ʈ��(Oee) = 550
    
    ElseIf Oee = 755 Then
    
        �̸�(Oee) = "���α�"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "����"
    
        ����(Oee) = 1
    
        ���ݷ�(Oee) = 850
    
        ����(Oee) = 650
    
        ����(Oee) = 600
    
        ����(Oee) = 650
    
        �����(Oee) = 700
    
        ����(Oee) = 600
    
        ����(Oee) = 700
    
        ��Ʈ��(Oee) = 800
    
    ElseIf Oee = 756 Then
    
        �̸�(Oee) = "�ּ���"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "KT"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 600
    
        ����(Oee) = 800
    
        ����(Oee) = 600
    
        ����(Oee) = 600
    
        �����(Oee) = 700
    
        ����(Oee) = 750
    
        ����(Oee) = 700
    
        ��Ʈ��(Oee) = 750
    
    ElseIf Oee = 757 Then
    
        �̸�(Oee) = "�Ǽ���"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "����"
    
        ����(Oee) = 2
    
        ���ݷ�(Oee) = 600
    
        ����(Oee) = 700
    
        ����(Oee) = 650
    
        ����(Oee) = 750
    
        �����(Oee) = 800
    
        ����(Oee) = 700
    
        ����(Oee) = 650
    
        ��Ʈ��(Oee) = 650
    
    ElseIf Oee = 758 Then
    
        �̸�(Oee) = "�����"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "�Ｚ"
    
        ����(Oee) = 1
    
        ���ݷ�(Oee) = 550
    
        ����(Oee) = 750
    
        ����(Oee) = 600
    
        ����(Oee) = 700
    
        �����(Oee) = 700
    
        ����(Oee) = 650
    
        ����(Oee) = 800
    
        ��Ʈ��(Oee) = 750
    
    ElseIf Oee = 759 Then
    
        �̸�(Oee) = "����ȭ"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 600
    
        ����(Oee) = 750
    
        ����(Oee) = 650
    
        ����(Oee) = 800
    
        �����(Oee) = 800
    
        ����(Oee) = 600
    
        ����(Oee) = 600
    
        ��Ʈ��(Oee) = 700
    
    ElseIf Oee = 760 Then
    
        �̸�(Oee) = "������"
    
        ��ũ(Oee) = "Special"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 850
    
        ����(Oee) = 800
    
        ����(Oee) = 650
    
        ����(Oee) = 700
    
        �����(Oee) = 700
    
        ����(Oee) = 600
    
        ����(Oee) = 750
    
        ��Ʈ��(Oee) = 750
    
    ElseIf Oee = 761 Then
    
        �̸�(Oee) = "���ؿ�"
    
        ��ũ(Oee) = "Special"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "8th"
    
        ����(Oee) = 2
    
        ���ݷ�(Oee) = 800
    
        ����(Oee) = 600
    
        ����(Oee) = 600
    
        ����(Oee) = 700
    
        �����(Oee) = 700
    
        ����(Oee) = 600
    
        ����(Oee) = 750
    
        ��Ʈ��(Oee) = 850
    
    ElseIf Oee = 762 Then
    
        �̸�(Oee) = "�̺���"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "8th"
    
        ����(Oee) = 2
    
        ���ݷ�(Oee) = 650
    
        ����(Oee) = 600
    
        ����(Oee) = 600
    
        ����(Oee) = 850
    
        �����(Oee) = 750
    
        ����(Oee) = 650
    
        ����(Oee) = 700
    
        ��Ʈ��(Oee) = 700
    
    ElseIf Oee = 763 Then
    
        �̸�(Oee) = "������"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        ����(Oee) = 1
    
        ���ݷ�(Oee) = 750
    
        ����(Oee) = 600
    
        ����(Oee) = 500
    
        ����(Oee) = 850
    
        �����(Oee) = 650
    
        ����(Oee) = 600
    
        ����(Oee) = 850
    
        ��Ʈ��(Oee) = 700
    
    ElseIf Oee = 764 Then
    
        �̸�(Oee) = "������"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        ����(Oee) = 1
    
        ���ݷ�(Oee) = 600
    
        ����(Oee) = 700
    
        ����(Oee) = 750
    
        ����(Oee) = 550
    
        �����(Oee) = 600
    
        ����(Oee) = 750
    
        ����(Oee) = 550
    
        ��Ʈ��(Oee) = 500
    
    ElseIf Oee = 765 Then
    
        �̸�(Oee) = "����ö"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 800
    
        ����(Oee) = 550
    
        ����(Oee) = 600
    
        ����(Oee) = 800
    
        �����(Oee) = 700
    
        ����(Oee) = 600
    
        ����(Oee) = 700
    
        ��Ʈ��(Oee) = 650
    
    ElseIf Oee = 766 Then
    
        �̸�(Oee) = "����ȣ"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        ����(Oee) = 2
    
        ���ݷ�(Oee) = 550
    
        ����(Oee) = 650
    
        ����(Oee) = 600
    
        ����(Oee) = 700
    
        �����(Oee) = 650
    
        ����(Oee) = 600
    
        ����(Oee) = 500
    
        ��Ʈ��(Oee) = 500
    
    ElseIf Oee = 767 Then
    
        �̸�(Oee) = "�ѵο�"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "CJ"
    
        ����(Oee) = 2
    
        ���ݷ�(Oee) = 600
    
        ����(Oee) = 600
    
        ����(Oee) = 600
    
        ����(Oee) = 600
    
        �����(Oee) = 600
    
        ����(Oee) = 600
    
        ����(Oee) = 600
    
        ��Ʈ��(Oee) = 600
    
    ElseIf Oee = 768 Then
    
        �̸�(Oee) = "������"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "�Ｚ"
    
        ����(Oee) = 2
    
        ���ݷ�(Oee) = 950
    
        ����(Oee) = 500
    
        ����(Oee) = 500
    
        ����(Oee) = 500
    
        �����(Oee) = 500
    
        ����(Oee) = 500
    
        ����(Oee) = 500
    
        ��Ʈ��(Oee) = 950
    
    ElseIf Oee = 769 Then
    
        �̸�(Oee) = "�赵��"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        ����(Oee) = 1
    
        ���ݷ�(Oee) = 550
    
        ����(Oee) = 600
    
        ����(Oee) = 650
    
        ����(Oee) = 550
    
        �����(Oee) = 600
    
        ����(Oee) = 650
    
        ����(Oee) = 750
    
        ��Ʈ��(Oee) = 600
    
    ElseIf Oee = 770 Then
    
        �̸�(Oee) = "������"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 650
    
        ����(Oee) = 600
    
        ����(Oee) = 550
    
        ����(Oee) = 750
    
        �����(Oee) = 500
    
        ����(Oee) = 600
    
        ����(Oee) = 650
    
        ��Ʈ��(Oee) = 550
    
    ElseIf Oee = 771 Then
    
        �̸�(Oee) = "����ȣ"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 550
    
        ����(Oee) = 500
    
        ����(Oee) = 550
    
        ����(Oee) = 650
    
        �����(Oee) = 550
    
        ����(Oee) = 500
    
        ����(Oee) = 650
    
        ��Ʈ��(Oee) = 700
    
    ElseIf Oee = 772 Then
    
        �̸�(Oee) = "������"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        ����(Oee) = 2
    
        ���ݷ�(Oee) = 650
    
        ����(Oee) = 600
    
        ����(Oee) = 600
    
        ����(Oee) = 700
    
        �����(Oee) = 750
    
        ����(Oee) = 600
    
        ����(Oee) = 600
    
        ��Ʈ��(Oee) = 600
    
    ElseIf Oee = 773 Then
    
        �̸�(Oee) = "�Ŵ��"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        ����(Oee) = 2
    
        ���ݷ�(Oee) = 600
    
        ����(Oee) = 700
    
        ����(Oee) = 600
    
        ����(Oee) = 800
    
        �����(Oee) = 800
    
        ����(Oee) = 650
    
        ����(Oee) = 650
    
        ��Ʈ��(Oee) = 600
    
    ElseIf Oee = 774 Then
    
        �̸�(Oee) = "������"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "STX"
    
        ����(Oee) = 2
    
        ���ݷ�(Oee) = 750
    
        ����(Oee) = 600
    
        ����(Oee) = 600
    
        ����(Oee) = 550
    
        �����(Oee) = 550
    
        ����(Oee) = 500
    
        ����(Oee) = 700
    
        ��Ʈ��(Oee) = 750
    
    ElseIf Oee = 775 Then
    
        �̸�(Oee) = "���ر�"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "����"
    
        ����(Oee) = 1
    
        ���ݷ�(Oee) = 600
    
        ����(Oee) = 650
    
        ����(Oee) = 550
    
        ����(Oee) = 650
    
        �����(Oee) = 550
    
        ����(Oee) = 600
    
        ����(Oee) = 550
    
        ��Ʈ��(Oee) = 600
    
    ElseIf Oee = 776 Then
    
        �̸�(Oee) = "������"
    
        ��ũ(Oee) = "Normal"
    
        OYear(Oee) = "<12>"
    
        Team(Oee) = "����"
    
        ����(Oee) = 3
    
        ���ݷ�(Oee) = 600
    
        ����(Oee) = 600
    
        ����(Oee) = 600
    
        ����(Oee) = 650
    
        �����(Oee) = 650
    
        ����(Oee) = 600
    
        ����(Oee) = 600
    
        ��Ʈ��(Oee) = 750
    
    ElseIf Oee = 777 Then
        �̸�(Oee) = "�����"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "����"
        ����(Oee) = 3
        ���ݷ�(Oee) = 650
        ����(Oee) = 650
        ����(Oee) = 600
        ����(Oee) = 800
        �����(Oee) = 700
        ����(Oee) = 600
        ����(Oee) = 700
        ��Ʈ��(Oee) = 850
    ElseIf Oee = 778 Then
        �̸�(Oee) = "�ڼ���"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        ����(Oee) = 1
        ���ݷ�(Oee) = 600
        ����(Oee) = 550
        ����(Oee) = 650
        ����(Oee) = 550
        �����(Oee) = 650
        ����(Oee) = 550
        ����(Oee) = 650
        ��Ʈ��(Oee) = 600
    ElseIf Oee = 779 Then
        �̸�(Oee) = "Ȳ����"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        ����(Oee) = 1
        ���ݷ�(Oee) = 650
        ����(Oee) = 600
        ����(Oee) = 600
        ����(Oee) = 600
        �����(Oee) = 600
        ����(Oee) = 600
        ����(Oee) = 600
        ��Ʈ��(Oee) = 600
    ElseIf Oee = 780 Then
        �̸�(Oee) = "���±�"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        ����(Oee) = 3
        ���ݷ�(Oee) = 600
        ����(Oee) = 600
        ����(Oee) = 600
        ����(Oee) = 600
        �����(Oee) = 600
        ����(Oee) = 600
        ����(Oee) = 600
        ��Ʈ��(Oee) = 600
    ElseIf Oee = 781 Then
        �̸�(Oee) = "����"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        ����(Oee) = 2
        ���ݷ�(Oee) = 850
        ����(Oee) = 500
        ����(Oee) = 500
        ����(Oee) = 700
        �����(Oee) = 750
        ����(Oee) = 700
        ����(Oee) = 700
        ��Ʈ��(Oee) = 750
    ElseIf Oee = 782 Then
        �̸�(Oee) = "�輺��"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        ����(Oee) = 2
        ���ݷ�(Oee) = 600
        ����(Oee) = 700
        ����(Oee) = 850
        ����(Oee) = 650
        �����(Oee) = 600
        ����(Oee) = 650
        ����(Oee) = 650
        ��Ʈ��(Oee) = 800
    ElseIf Oee = 783 Then
        �̸�(Oee) = "�ֿ���"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        ����(Oee) = 2
        ���ݷ�(Oee) = 600
        ����(Oee) = 650
        ����(Oee) = 650
        ����(Oee) = 500
        �����(Oee) = 600
        ����(Oee) = 700
        ����(Oee) = 600
        ��Ʈ��(Oee) = 500
    ElseIf Oee = 784 Then
        �̸�(Oee) = "��ȣ��"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "SK"
        ����(Oee) = 1
        ���ݷ�(Oee) = 550
        ����(Oee) = 500
        ����(Oee) = 550
        ����(Oee) = 500
        �����(Oee) = 600
        ����(Oee) = 650
        ����(Oee) = 950
        ��Ʈ��(Oee) = 550
    ElseIf Oee = 785 Then
        �̸�(Oee) = "������"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "SK"
        ����(Oee) = 2
        ���ݷ�(Oee) = 750
        ����(Oee) = 500
        ����(Oee) = 500
        ����(Oee) = 500
        �����(Oee) = 600
        ����(Oee) = 700
        ����(Oee) = 600
        ��Ʈ��(Oee) = 750
    ElseIf Oee = 786 Then
        �̸�(Oee) = "�̽¼�"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "SK"
        ����(Oee) = 2
        ���ݷ�(Oee) = 650
        ����(Oee) = 600
        ����(Oee) = 550
        ����(Oee) = 500
        �����(Oee) = 700
        ����(Oee) = 600
        ����(Oee) = 600
        ��Ʈ��(Oee) = 600
    ElseIf Oee = 787 Then
        �̸�(Oee) = "������"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "����"
        ����(Oee) = 1
        ���ݷ�(Oee) = 850
        ����(Oee) = 550
        ����(Oee) = 500
        ����(Oee) = 550
        �����(Oee) = 550
        ����(Oee) = 600
        ����(Oee) = 500
        ��Ʈ��(Oee) = 750
    ElseIf Oee = 788 Then
        �̸�(Oee) = "�̼���"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "����"
        ����(Oee) = 1
        ���ݷ�(Oee) = 850
        ����(Oee) = 550
        ����(Oee) = 550
        ����(Oee) = 500
        �����(Oee) = 500
        ����(Oee) = 500
        ����(Oee) = 500
        ��Ʈ��(Oee) = 850
    ElseIf Oee = 789 Then
        �̸�(Oee) = "������"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "����"
        ����(Oee) = 1
        ���ݷ�(Oee) = 750
        ����(Oee) = 550
        ����(Oee) = 750
        ����(Oee) = 650
        �����(Oee) = 750
        ����(Oee) = 550
        ����(Oee) = 650
        ��Ʈ��(Oee) = 750
    ElseIf Oee = 790 Then
        �̸�(Oee) = "�ռ���"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "����"
        ����(Oee) = 3
        ���ݷ�(Oee) = 850
        ����(Oee) = 600
        ����(Oee) = 600
        ����(Oee) = 800
        �����(Oee) = 700
        ����(Oee) = 700
        ����(Oee) = 650
        ��Ʈ��(Oee) = 650
    ElseIf Oee = 791 Then
        �̸�(Oee) = "������"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "����"
        ����(Oee) = 2
        ���ݷ�(Oee) = 600
        ����(Oee) = 600
        ����(Oee) = 600
        ����(Oee) = 600
        �����(Oee) = 600
        ����(Oee) = 600
        ����(Oee) = 600
        ��Ʈ��(Oee) = 600
    ElseIf Oee = 792 Then
        �̸�(Oee) = "����"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "����"
        ����(Oee) = 2
        ���ݷ�(Oee) = 600
        ����(Oee) = 600
        ����(Oee) = 600
        ����(Oee) = 600
        �����(Oee) = 600
        ����(Oee) = 600
        ����(Oee) = 600
        ��Ʈ��(Oee) = 600
    ElseIf Oee = 793 Then
        �̸�(Oee) = "����ȯ"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "����"
        ����(Oee) = 2
        ���ݷ�(Oee) = 550
        ����(Oee) = 550
        ����(Oee) = 550
        ����(Oee) = 550
        �����(Oee) = 550
        ����(Oee) = 550
        ����(Oee) = 550
        ��Ʈ��(Oee) = 550
    ElseIf Oee = 794 Then
        �̸�(Oee) = "������"
        ��ũ(Oee) = "Normal"
        OYear(Oee) = "<12>"
        Team(Oee) = "����"
        ����(Oee) = 2
        ���ݷ�(Oee) = 500
        ����(Oee) = 500
        ����(Oee) = 500
        ����(Oee) = 500
        �����(Oee) = 500
        ����(Oee) = 500
        ����(Oee) = 500
        ��Ʈ��(Oee) = 500
    ElseIf Oee = 795 Then
        �̸�(Oee) = "����[Ex]"
        ��ũ(Oee) = "Unique"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        ����(Oee) = 2
        ���ݷ�(Oee) = 950
        ����(Oee) = 700
        ����(Oee) = 700
        ����(Oee) = 850
        �����(Oee) = 850
        ����(Oee) = 750
        ����(Oee) = 850
        ��Ʈ��(Oee) = 850
    ElseIf Oee = 796 Then
        �̸�(Oee) = "�輺��[Ex]"
        ��ũ(Oee) = "Unique"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        ����(Oee) = 2
        ���ݷ�(Oee) = 800
        ����(Oee) = 750
        ����(Oee) = 950
        ����(Oee) = 900
        �����(Oee) = 900
        ����(Oee) = 800
        ����(Oee) = 700
        ��Ʈ��(Oee) = 700
    ElseIf Oee = 797 Then
        �̸�(Oee) = "�̿�ȣ[Ex]"
        ��ũ(Oee) = "Champion"
        OYear(Oee) = "<12>"
        Team(Oee) = "KT"
        ����(Oee) = 1
        ���ݷ�(Oee) = 950
        ����(Oee) = 900
        ����(Oee) = 900
        ����(Oee) = 950
        �����(Oee) = 900
        ����(Oee) = 900
        ����(Oee) = 900
        Skill(Oee) = 2
        ��Ʈ��(Oee) = 1000
    ElseIf Oee = 798 Then
        �̸�(Oee) = "������[Ex]"
        ��ũ(Oee) = "Elite"
        OYear(Oee) = "<10>"
        Team(Oee) = "SK"
        ����(Oee) = 1
        ���ݷ�(Oee) = 750
        ����(Oee) = 900
        ����(Oee) = 850
        ����(Oee) = 850
        �����(Oee) = 900
        ����(Oee) = 900
        ����(Oee) = 850
        ��Ʈ��(Oee) = 650
    ElseIf Oee = 799 Then
        �̸�(Oee) = "������[Ex]"
        ��ũ(Oee) = "Unique"
        OYear(Oee) = "<09>"
        Team(Oee) = "SK"
        ����(Oee) = 2
        ���ݷ�(Oee) = 950
        ����(Oee) = 850
        ����(Oee) = 650
        ����(Oee) = 800
        �����(Oee) = 650
        ����(Oee) = 800
        ����(Oee) = 850
        ��Ʈ��(Oee) = 900
    ElseIf Oee = 800 Then
        �̸�(Oee) = "KT"
        ��ũ(Oee) = "Elite"
        OYear(Oee) = "<12>"
        Team(Oee) = "Mystar"
        ����(Oee) = 1
        ���ݷ�(Oee) = 950
        ����(Oee) = 900
        ����(Oee) = 650
        ����(Oee) = 800
        �����(Oee) = 950
        ����(Oee) = 700
        ����(Oee) = 700
        ��Ʈ��(Oee) = 950
    End If
    ���(Oee) = 0
    �ؿ��(Oee) = 0
    �����(Oee) = 100
    A�¸�(Oee) = 0
    A�й�(Oee) = 0
    P�¸�(Oee) = 0
    P�й�(Oee) = 0
    T�¸�(Oee) = 0
    T�й�(Oee) = 0
    Z�¸�(Oee) = 0
    Z�й�(Oee) = 0
    T����(Oee) = 0
    Z����(Oee) = 0
    P����(Oee) = 0
    A����(Oee) = 0
    T��(Oee) = "W"
    Z��(Oee) = "W"
    P��(Oee) = "W"
    A��(Oee) = "W"
Next

Tim12.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub TimAdd_Timer()
For Oee = 716 To 723
 If Oee = 716 Then
  �̸�(Oee) = "���϶�"
  ��ũ(Oee) = "Rare"
  OYear(Oee) = "<11>"
  Team(Oee) = "Mystar"
  ����(Oee) = 1
  ���ݷ�(Oee) = 700
  ����(Oee) = 650
  ����(Oee) = 700
  ����(Oee) = 750
  �����(Oee) = 800
  ����(Oee) = 800
  ����(Oee) = 1000
  ��Ʈ��(Oee) = 600
 ElseIf Oee = 717 Then
  �̸�(Oee) = "�ڿ��"
  ��ũ(Oee) = "Legend"
  OYear(Oee) = "<02>"
  Team(Oee) = "IS"
  ����(Oee) = 3
  ���ݷ�(Oee) = 800
  ����(Oee) = 950
  ����(Oee) = 700
  ����(Oee) = 800
  �����(Oee) = 800
  ����(Oee) = 950
  ����(Oee) = 850
  ��Ʈ��(Oee) = 950
 ElseIf Oee = 718 Then
  �̸�(Oee) = "������"
  ��ũ(Oee) = "Unique"
  OYear(Oee) = "<02>"
  Team(Oee) = "STX"
  ����(Oee) = 1
  ���ݷ�(Oee) = 900
  ����(Oee) = 900
  ����(Oee) = 800
  ����(Oee) = 750
  �����(Oee) = 800
  ����(Oee) = 650
  ����(Oee) = 850
  ��Ʈ��(Oee) = 900
ElseIf Oee = 719 Then
 �̸�(Oee) = "ī��"
 ��ũ(Oee) = "Unique"
 Team(Oee) = "Mystar"
 OYear(Oee) = "<11>"
 ����(Oee) = 2
 ���ݷ�(Oee) = 950
 ����(Oee) = 550
 ����(Oee) = 800
 ����(Oee) = 800
 �����(Oee) = 550
 ����(Oee) = 950
 ����(Oee) = 950
 ��Ʈ��(Oee) = 950
ElseIf Oee = 720 Then
 �̸�(Oee) = "�鿵"
 ��ũ(Oee) = "Rare"
 Team(Oee) = "Mystar"
 OYear(Oee) = "<11>"
 ����(Oee) = 3
 ���ݷ�(Oee) = 900
 ����(Oee) = 600
 ����(Oee) = 950
 ����(Oee) = 950
 �����(Oee) = 600
 ����(Oee) = 350
 ����(Oee) = 950
 ��Ʈ��(Oee) = 950
ElseIf Oee = 721 Then
 �̸�(Oee) = "����"
 ��ũ(Oee) = "Rare"
 Team(Oee) = "Mystar"
 OYear(Oee) = "<11>"
 ����(Oee) = 1
 ���ݷ�(Oee) = 700
 ����(Oee) = 650
 ����(Oee) = 700
 ����(Oee) = 850
 �����(Oee) = 800
 ����(Oee) = 850
 ����(Oee) = 750
 ��Ʈ��(Oee) = 600
ElseIf Oee = 722 Then
 �̸�(Oee) = "�¾�"
 ��ũ(Oee) = "Elite"
 Team(Oee) = "Mystar"
 OYear(Oee) = "<11>"
 ����(Oee) = 1
 ���ݷ�(Oee) = 850
 ����(Oee) = 700
 ����(Oee) = 750
 ����(Oee) = 950
 �����(Oee) = 900
 ����(Oee) = 800
 ����(Oee) = 950
 ��Ʈ��(Oee) = 800
ElseIf Oee = 723 Then
 �̸�(Oee) = "������"
 ��ũ(Oee) = "Elite"
 Team(Oee) = "Mystar"
 OYear(Oee) = "<11>"
 ����(Oee) = 3
 ���ݷ�(Oee) = 950
 ����(Oee) = 750
 ����(Oee) = 750
 ����(Oee) = 950
 �����(Oee) = 850
 ����(Oee) = 850
 ����(Oee) = 800
 ��Ʈ��(Oee) = 800
 
 End If

 ���(Oee) = 0
 �ؿ��(Oee) = 0
 �����(Oee) = 100
 A�¸�(Oee) = 0
 A�й�(Oee) = 0
 P�¸�(Oee) = 0
 P�й�(Oee) = 0
 T�¸�(Oee) = 0
 T�й�(Oee) = 0
 Z�¸�(Oee) = 0
 Z�й�(Oee) = 0
 T����(Oee) = 0
 Z����(Oee) = 0
 P����(Oee) = 0
 A����(Oee) = 0
 T��(Oee) = "W"
 Z��(Oee) = "W"
 P��(Oee) = "W"
 A��(Oee) = "W"
Next
Tim12.Enabled = True
TimAdd.Enabled = False
End Sub

Private Sub TimElse_Timer()
For Oee = 540 To 575
If Oee = 540 Then
�̸�(Oee) = "�ӿ�ȯ"
��ũ(Oee) = "Champion"
OYear(Oee) = "<01>"
Team(Oee) = "IS"
����(Oee) = 1
���ݷ�(Oee) = 1000
����(Oee) = 900
����(Oee) = 900
����(Oee) = 850
�����(Oee) = 850
����(Oee) = 900
����(Oee) = 900
��Ʈ��(Oee) = 1100
Skill(Oee) = 1


ElseIf Oee = 541 Then
�̸�(Oee) = "�ӿ�ȯ"
��ũ(Oee) = "Rare"
OYear(Oee) = "<02>"
Team(Oee) = "IS"
����(Oee) = 1
���ݷ�(Oee) = 850
����(Oee) = 850
����(Oee) = 800
����(Oee) = 650
�����(Oee) = 800
����(Oee) = 650
����(Oee) = 750
��Ʈ��(Oee) = 850


ElseIf Oee = 542 Then
�̸�(Oee) = "�ӿ�ȯ"
��ũ(Oee) = "Rare"
OYear(Oee) = "<04>"
Team(Oee) = "4U"
����(Oee) = 1
���ݷ�(Oee) = 850
����(Oee) = 850
����(Oee) = 800
����(Oee) = 500
�����(Oee) = 650
����(Oee) = 600
����(Oee) = 850
��Ʈ��(Oee) = 900


ElseIf Oee = 543 Then
�̸�(Oee) = "�ӿ�ȯ"
��ũ(Oee) = "Unique"
OYear(Oee) = "<05>"
Team(Oee) = "SK"
����(Oee) = 1
���ݷ�(Oee) = 850
����(Oee) = 750
����(Oee) = 850
����(Oee) = 800
�����(Oee) = 750
����(Oee) = 700
����(Oee) = 850
��Ʈ��(Oee) = 850



ElseIf Oee = 544 Then
�̸�(Oee) = "ȫ��ȣ"
��ũ(Oee) = "Secret"
OYear(Oee) = "<01>"
Team(Oee) = "IS"
����(Oee) = 2
���ݷ�(Oee) = 950
����(Oee) = 850
����(Oee) = 800
����(Oee) = 950
�����(Oee) = 850
����(Oee) = 800
����(Oee) = 850
��Ʈ��(Oee) = 950
Skill(Oee) = 3

ElseIf Oee = 545 Then
�̸�(Oee) = "ȫ��ȣ"
��ũ(Oee) = "Unique"
OYear(Oee) = "<02>"
Team(Oee) = "IS"
����(Oee) = 2
���ݷ�(Oee) = 850
����(Oee) = 800
����(Oee) = 700
����(Oee) = 800
�����(Oee) = 800
����(Oee) = 850
����(Oee) = 800
��Ʈ��(Oee) = 800


ElseIf Oee = 546 Then
�̸�(Oee) = "ȫ��ȣ"
��ũ(Oee) = "Unique"
OYear(Oee) = "<03>"
Team(Oee) = "KTF"
����(Oee) = 2
���ݷ�(Oee) = 950
����(Oee) = 800
����(Oee) = 600
����(Oee) = 850
�����(Oee) = 850
����(Oee) = 600
����(Oee) = 800
��Ʈ��(Oee) = 950


ElseIf Oee = 547 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Secret"
OYear(Oee) = "<02>"
Team(Oee) = "IS"
����(Oee) = 1
���ݷ�(Oee) = 900
����(Oee) = 800
����(Oee) = 800
����(Oee) = 900
�����(Oee) = 950
����(Oee) = 750
����(Oee) = 950
��Ʈ��(Oee) = 950
Skill(Oee) = 5

ElseIf Oee = 548 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Rare"
OYear(Oee) = "<03>"
Team(Oee) = "KTF"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 700
����(Oee) = 650
����(Oee) = 800
�����(Oee) = 900
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 800


ElseIf Oee = 549 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Elite"
OYear(Oee) = "<04>"
Team(Oee) = "Toona"
����(Oee) = 1
���ݷ�(Oee) = 850
����(Oee) = 700
����(Oee) = 700
����(Oee) = 950
�����(Oee) = 950
����(Oee) = 700
����(Oee) = 950
��Ʈ��(Oee) = 850


ElseIf Oee = 550 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Unique"
OYear(Oee) = "<06>"
Team(Oee) = "Pantech"
����(Oee) = 1
���ݷ�(Oee) = 850
����(Oee) = 750
����(Oee) = 800
����(Oee) = 850
�����(Oee) = 750
����(Oee) = 700
����(Oee) = 850
��Ʈ��(Oee) = 850


ElseIf Oee = 551 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Unique"
OYear(Oee) = "<05>"
Team(Oee) = "GO"
����(Oee) = 2
���ݷ�(Oee) = 900
����(Oee) = 700
����(Oee) = 750
����(Oee) = 900
�����(Oee) = 900
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 900


ElseIf Oee = 552 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Legend"
OYear(Oee) = "<06>"
Team(Oee) = "CJ"
����(Oee) = 2
���ݷ�(Oee) = 850
����(Oee) = 750
����(Oee) = 800
����(Oee) = 950
�����(Oee) = 950
����(Oee) = 750
����(Oee) = 800
��Ʈ��(Oee) = 950


ElseIf Oee = 553 Then
�̸�(Oee) = "�ֿ���"
��ũ(Oee) = "Secret"
OYear(Oee) = "<03>"
Team(Oee) = "Orion"
����(Oee) = 1
���ݷ�(Oee) = 950
����(Oee) = 800
����(Oee) = 800
����(Oee) = 950
�����(Oee) = 950
����(Oee) = 800
����(Oee) = 850
��Ʈ��(Oee) = 900
Skill(Oee) = 7

ElseIf Oee = 554 Then
�̸�(Oee) = "�ֿ���"
��ũ(Oee) = "Elite"
OYear(Oee) = "<04>"
Team(Oee) = "4U"
����(Oee) = 1
���ݷ�(Oee) = 850
����(Oee) = 700
����(Oee) = 800
����(Oee) = 950
�����(Oee) = 950
����(Oee) = 950
����(Oee) = 650
��Ʈ��(Oee) = 750


ElseIf Oee = 555 Then
�̸�(Oee) = "�ֿ���"
��ũ(Oee) = "Unique"
OYear(Oee) = "<05>"
Team(Oee) = "SK"
����(Oee) = 1
���ݷ�(Oee) = 900
����(Oee) = 700
����(Oee) = 850
����(Oee) = 950
�����(Oee) = 900
����(Oee) = 650
����(Oee) = 750
��Ʈ��(Oee) = 750

ElseIf Oee = 556 Then
�̸�(Oee) = "���¹�"
��ũ(Oee) = "Rare"
OYear(Oee) = "<03>"
Team(Oee) = "GO"
����(Oee) = 2
���ݷ�(Oee) = 700
����(Oee) = 650
����(Oee) = 700
����(Oee) = 950
�����(Oee) = 900
����(Oee) = 700
����(Oee) = 700
��Ʈ��(Oee) = 700

ElseIf Oee = 557 Then
�̸�(Oee) = "���¹�"
��ũ(Oee) = "Unique"
OYear(Oee) = "<04>"
Team(Oee) = "GO"
����(Oee) = 2
���ݷ�(Oee) = 850
����(Oee) = 650
����(Oee) = 650
����(Oee) = 950
�����(Oee) = 950
����(Oee) = 750
����(Oee) = 750
��Ʈ��(Oee) = 850

ElseIf Oee = 558 Then
�̸�(Oee) = "�赿��"
��ũ(Oee) = "Unique"
OYear(Oee) = "<01>"
Team(Oee) = "�Ѻ�"
����(Oee) = 3
���ݷ�(Oee) = 800
����(Oee) = 750
����(Oee) = 650
����(Oee) = 950
�����(Oee) = 850
����(Oee) = 800
����(Oee) = 800
��Ʈ��(Oee) = 800

ElseIf Oee = 559 Then
�̸�(Oee) = "���漷"
��ũ(Oee) = "Unique"
OYear(Oee) = "<02>"
Team(Oee) = "�Ѻ�"
����(Oee) = 1
���ݷ�(Oee) = 850
����(Oee) = 800
����(Oee) = 750
����(Oee) = 800
�����(Oee) = 750
����(Oee) = 700
����(Oee) = 850
��Ʈ��(Oee) = 900

ElseIf Oee = 560 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Secret"
OYear(Oee) = "<02>"
Team(Oee) = "�Ѻ�"
����(Oee) = 3
���ݷ�(Oee) = 900
����(Oee) = 800
����(Oee) = 900
����(Oee) = 950
�����(Oee) = 900
����(Oee) = 800
����(Oee) = 850
��Ʈ��(Oee) = 900
Skill(Oee) = 4

ElseIf Oee = 561 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Rare"
OYear(Oee) = "<04>"
Team(Oee) = "KTF"
����(Oee) = 3
���ݷ�(Oee) = 800
����(Oee) = 650
����(Oee) = 600
����(Oee) = 900
�����(Oee) = 900
����(Oee) = 700
����(Oee) = 750
��Ʈ��(Oee) = 700

ElseIf Oee = 562 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Rare"
OYear(Oee) = "<01>"
Team(Oee) = "�Ѻ�"
����(Oee) = 2
���ݷ�(Oee) = 950
����(Oee) = 650
����(Oee) = 600
����(Oee) = 750
�����(Oee) = 800
����(Oee) = 700
����(Oee) = 700
��Ʈ��(Oee) = 850

ElseIf Oee = 563 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Rare"
OYear(Oee) = "<02>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 650
����(Oee) = 650
����(Oee) = 900
�����(Oee) = 900
����(Oee) = 700
����(Oee) = 700
��Ʈ��(Oee) = 750

ElseIf Oee = 564 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Unique"
OYear(Oee) = "<05>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 750
����(Oee) = 700
����(Oee) = 850
����(Oee) = 950
�����(Oee) = 900
����(Oee) = 700
����(Oee) = 750
��Ʈ��(Oee) = 800

ElseIf Oee = 565 Then
�̸�(Oee) = "����ȣ"
��ũ(Oee) = "Rare"
OYear(Oee) = "<06>"
Team(Oee) = "STX"
����(Oee) = 2
���ݷ�(Oee) = 900
����(Oee) = 700
����(Oee) = 550
����(Oee) = 800
�����(Oee) = 850
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 850

ElseIf Oee = 566 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Unique"
OYear(Oee) = "<03>"
Team(Oee) = "GO"
����(Oee) = 1
���ݷ�(Oee) = 800
����(Oee) = 900
����(Oee) = 650
����(Oee) = 950
�����(Oee) = 950
����(Oee) = 800
����(Oee) = 650
��Ʈ��(Oee) = 850

ElseIf Oee = 567 Then
�̸�(Oee) = "�ڿ��"
��ũ(Oee) = "Rare"
OYear(Oee) = "<03>"
Team(Oee) = "Orion"
����(Oee) = 3
���ݷ�(Oee) = 750
����(Oee) = 750
����(Oee) = 600
����(Oee) = 800
�����(Oee) = 800
����(Oee) = 650
����(Oee) = 900
��Ʈ��(Oee) = 750

ElseIf Oee = 568 Then
�̸�(Oee) = "����"
��ũ(Oee) = "Secret"
OYear(Oee) = "<03>"
Team(Oee) = "GO"
����(Oee) = 3
���ݷ�(Oee) = 850
����(Oee) = 800
����(Oee) = 950
����(Oee) = 950
�����(Oee) = 900
����(Oee) = 800
����(Oee) = 950
��Ʈ��(Oee) = 800
Skill(Oee) = 30

ElseIf Oee = 569 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Legend"
OYear(Oee) = "<04>"
Team(Oee) = "POS"
����(Oee) = 2
���ݷ�(Oee) = 950
����(Oee) = 950
����(Oee) = 750
����(Oee) = 900
�����(Oee) = 800
����(Oee) = 650
����(Oee) = 850
��Ʈ��(Oee) = 950

ElseIf Oee = 570 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Unique"
OYear(Oee) = "<05>"
Team(Oee) = "POS"
����(Oee) = 2
���ݷ�(Oee) = 900
����(Oee) = 650
����(Oee) = 650
����(Oee) = 950
�����(Oee) = 950
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 950

ElseIf Oee = 571 Then
�̸�(Oee) = "�ڼ���"
��ũ(Oee) = "Rare"
OYear(Oee) = "<06>"
Team(Oee) = "MBC"
����(Oee) = 2
���ݷ�(Oee) = 850
����(Oee) = 600
����(Oee) = 550
����(Oee) = 950
�����(Oee) = 850
����(Oee) = 700
����(Oee) = 650
��Ʈ��(Oee) = 850

ElseIf Oee = 572 Then
�̸�(Oee) = "�̺���"
��ũ(Oee) = "Rare"
OYear(Oee) = "<05>"
Team(Oee) = "Curitel"
����(Oee) = 1
���ݷ�(Oee) = 900
����(Oee) = 650
����(Oee) = 650
����(Oee) = 650
�����(Oee) = 700
����(Oee) = 650
����(Oee) = 950
��Ʈ��(Oee) = 850

ElseIf Oee = 573 Then
�̸�(Oee) = "������"
��ũ(Oee) = "Unique"
OYear(Oee) = "<04>"
Team(Oee) = "PLUS"
����(Oee) = 3
���ݷ�(Oee) = 900
����(Oee) = 900
����(Oee) = 850
����(Oee) = 800
�����(Oee) = 650
����(Oee) = 750
����(Oee) = 750
��Ʈ��(Oee) = 800

ElseIf Oee = 574 Then
�̸�(Oee) = "�ѵ���"
��ũ(Oee) = "Unique"
OYear(Oee) = "<06>"
Team(Oee) = "�°��ӳ�"
����(Oee) = 1
���ݷ�(Oee) = 950
����(Oee) = 950
����(Oee) = 550
����(Oee) = 500
�����(Oee) = 600
����(Oee) = 950
����(Oee) = 950
��Ʈ��(Oee) = 950

ElseIf Oee = 575 Then
�̸�(Oee) = "�ɼҸ�"
��ũ(Oee) = "Rare"
OYear(Oee) = "<06>"
Team(Oee) = "Pantech"
����(Oee) = 2
���ݷ�(Oee) = 800
����(Oee) = 650
����(Oee) = 950
����(Oee) = 850
�����(Oee) = 850
����(Oee) = 650
����(Oee) = 700
��Ʈ��(Oee) = 750
End If

 ���(Oee) = 0
 �ؿ��(Oee) = 0
 �����(Oee) = 100
 A�¸�(Oee) = 0
 A�й�(Oee) = 0
 P�¸�(Oee) = 0
 P�й�(Oee) = 0
 T�¸�(Oee) = 0
 T�й�(Oee) = 0
 Z�¸�(Oee) = 0
 Z�й�(Oee) = 0
 T����(Oee) = 0
 Z����(Oee) = 0
 P����(Oee) = 0
 A����(Oee) = 0
 T��(Oee) = "W"
 Z��(Oee) = "W"
 P��(Oee) = "W"
 A��(Oee) = "W"
 Next Oee
 TimElse.Enabled = False
 TimAdd.Enabled = True
End Sub

Private Sub Timer1_Timer()
For Oee = 0 To 800
 ���ݷ�(Oee) = val(���ݷ�(Oee)) - 50
 ����(Oee) = val(����(Oee)) - 50
 ����(Oee) = val(����(Oee)) - 50
 ����(Oee) = val(����(Oee)) - 50
 �����(Oee) = val(�����(Oee)) - 50
 ����(Oee) = val(����(Oee)) - 50
 ����(Oee) = val(����(Oee)) - 50
 ��Ʈ��(Oee) = val(��Ʈ��(Oee)) - 50
Next

For Oee = 0 To 800
 NPC���ݷ�(Oee) = ���ݷ�(Oee)
 NPC����(Oee) = ����(Oee)
 NPC����(Oee) = ����(Oee)
 NPC����(Oee) = ����(Oee)
 NPC�����(Oee) = �����(Oee)
 NPC����(Oee) = ����(Oee)
 NPC����(Oee) = ����(Oee)
 NPC��Ʈ��(Oee) = ��Ʈ��(Oee)
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
 jcbutton1.Caption = "���� �Ϸ�"
 Label1 = "�������ֽʽÿ�."
 Timer2.Enabled = True
 Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Randomize Oee
End Sub

Private Sub Timer3_Timer()
If val(���÷�) = 3 Then
 jcbutton1.Enabled = True
End If

Dim ���� As Integer
For ���� = 1 To 6
MyAW(����) = 0
MyAL(����) = 0
MyTW(����) = 0
MyTL(����) = 0
MyPW(����) = 0
MyPL(����) = 0
MyZW(����) = 0
MyZL(����) = 0
MyT����(����) = 0
MyZ����(����) = 0
MyP����(����) = 0
MyA����(����) = 0
MyT��(����) = "W"
MyZ��(����) = "W"
MyP��(����) = "W"
MyA��(����) = "W"
MySkill(����) = 0
Turn = "OSL"
MyVic(����) = 0
MySeVic(����) = 0
�ൿ�� = 0
Con = 100
MyExp(����) = 0
MyMExp(����) = 10
MyLev(����) = 1
MyPoint(����) = 0
Next ����
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
