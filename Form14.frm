VERSION 5.00
Begin VB.Form FrmGameInfo 
   Caption         =   "����â"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form14"
   ScaleHeight     =   5775
   ScaleWidth      =   9255
   StartUpPosition =   2  'ȭ�� ���
   Begin CSO.jcbutton jcbutton1 
      Height          =   375
      Left            =   3120
      TabIndex        =   18
      Top             =   3600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Go"
      CaptionEffects  =   0
   End
   Begin VB.Label lblO���� 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label7"
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
      Left            =   6240
      TabIndex        =   25
      Top             =   5400
      Width           =   3015
   End
   Begin VB.Label lblM���� 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label6"
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
      TabIndex        =   24
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label lblMapTri 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   23
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label lblOTeam 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label6"
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
      Left            =   6240
      TabIndex        =   22
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label lblMTeam 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label6"
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
      TabIndex        =   21
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label lblOrank 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label7"
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
      Left            =   6240
      TabIndex        =   20
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label lblMrank 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label6"
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
      TabIndex        =   19
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label lblSAO 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "<����>"
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
      Left            =   6240
      TabIndex        =   17
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label lblSA 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "<����>"
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
      TabIndex        =   16
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label Label5 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "-Special Ability-"
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
      Left            =   6240
      TabIndex        =   15
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "-Special Ability-"
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
      TabIndex        =   14
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label lblOTT 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "[10�� 0��]"
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
      Left            =   6240
      TabIndex        =   13
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label lblMTT 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "[10�� 0��]"
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
      TabIndex        =   12
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Label lblMT 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "[Vs T]"
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
      Left            =   6240
      TabIndex        =   11
      Top             =   4920
      Width           =   3015
   End
   Begin VB.Label lblOT 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "[Vs Z]"
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
      TabIndex        =   10
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "<��� ������>"
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
      Left            =   3120
      TabIndex        =   9
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Label lblOSt 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Stats : 6500"
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
      Left            =   6240
      TabIndex        =   8
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label lblOR 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Rank : C-"
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
      Left            =   6240
      TabIndex        =   7
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label lblMSt 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Stats : 6500"
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
      TabIndex        =   6
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label lblMR 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Rank : C-"
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
      TabIndex        =   5
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "<Card Rank>"
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
      Left            =   3120
      TabIndex        =   4
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '�������� ����
      Height          =   1095
      Index           =   5
      Left            =   6240
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '�������� ����
      Height          =   1095
      Index           =   4
      Left            =   3120
      Top             =   4680
      Width           =   3135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '�������� ����
      Height          =   1095
      Index           =   3
      Left            =   0
      Top             =   4680
      Width           =   3135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '�������� ����
      Height          =   735
      Index           =   2
      Left            =   6240
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '�������� ����
      Height          =   735
      Index           =   1
      Left            =   3120
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '�������� ����
      Height          =   735
      Index           =   0
      Left            =   0
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label lblMapName 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "[�׿����۷��̺�]"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "<Map Order>"
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
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   3135
   End
   Begin VB.Image ImgMa 
      Height          =   1500
      Left            =   3960
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label lblON 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�̿�ȣ"
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
      Left            =   6240
      TabIndex        =   1
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label lblMN 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "���±�"
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
      TabIndex        =   0
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Image ImgOp 
      Height          =   1500
      Left            =   6960
      Top             =   360
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      Height          =   4000
      Index           =   2
      Left            =   6240
      Top             =   0
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      Height          =   4000
      Index           =   1
      Left            =   3120
      Top             =   0
      Width           =   3135
   End
   Begin VB.Image ImgMe 
      Height          =   1500
      Left            =   720
      Top             =   360
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      Height          =   4005
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "FrmGameInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

If val(MyTribe(����)) = 1 Then
 If val(����(Oee)) = 1 Then
  lblMapTri = "T vs T = 50 : 50"
 ElseIf val(����(Oee)) = 2 Then
  lblMapTri = "T vs Z = " & TZT(Map) & " : " & TZZ(Map)
 ElseIf val(����(Oee)) = 3 Then
  lblMapTri = "T vs P = " & PTT(Map) & " : " & PTP(Map)
 End If
ElseIf val(MyTribe(����)) = 2 Then
 If val(����(Oee)) = 1 Then
  lblMapTri = "Z vs T = " & TZZ(Map) & " : " & TZT(Map)
 ElseIf val(����(Oee)) = 2 Then
  lblMapTri = "Z vs Z = 50 : 50"
 ElseIf val(����(Oee)) = 3 Then
  lblMapTri = "Z vs P = " & ZPZ(Map) & " : " & ZPP(Map)
 End If
ElseIf val(MyTribe(����)) = 3 Then
 If val(����(Oee)) = 1 Then
  lblMapTri = "P vs T = " & PTP(Map) & " : " & PTT(Map)
 ElseIf val(����(Oee)) = 2 Then
  lblMapTri = "P vs Z = " & ZPP(Map) & " : " & ZPZ(Map)
 ElseIf val(����(Oee)) = 3 Then
  lblMapTri = "P vs P = 50 : 50"
 End If
End If


Dim ���̷�ũ As String, ��뷩ũ As String
lblMN = MyYear(����) & " " & MyName(����)
lblON = OYear(Oee) & " " & �̸�(Oee)
lblMTeam = MyTeam(����)
lblOTeam = Team(Oee)

If MySkill(����) = 1 Then
    lblSA = "���"
ElseIf MySkill(����) = 2 Then
    lblSA = "����ܲ��"
ElseIf MySkill(����) = 3 Then
    lblSA = "����"
ElseIf MySkill(����) = 4 Then
    lblSA = "���"
ElseIf MySkill(����) = 5 Then
    lblSA = "���"
ElseIf MySkill(����) = 6 Then
    lblSA = "Maestro"
ElseIf MySkill(����) = 7 Then
    lblSA = "��ڪ"
ElseIf MySkill(����) = 8 Then
    lblSA = "ޫף"
ElseIf MySkill(����) = 9 Then
    lblSA = "��ף"
ElseIf MySkill(����) = 10 Then
    lblSA = "��ף"
ElseIf MySkill(����) = 11 Then
    lblSA = "�ף"
ElseIf MySkill(����) = 12 Then
    lblSA = "��ף"
ElseIf MySkill(����) = 13 Then
    lblSA = "��ף"
ElseIf MySkill(����) = 14 Then
    lblSA = "����"
ElseIf MySkill(����) = 15 Then
    lblSA = "����"
ElseIf MySkill(����) = 16 Then
    lblSA = "����"
ElseIf MySkill(����) = 17 Then
    lblSA = "����"
ElseIf MySkill(����) = 18 Then
    lblSA = "����"
ElseIf MySkill(����) = 19 Then
    lblSA = "����"
ElseIf MySkill(����) = 20 Then
    lblSA = "���"
ElseIf MySkill(����) = 21 Then
    lblSA = "rEd sNipeR"
ElseIf MySkill(����) = 22 Then
    lblSA = "������"
ElseIf MySkill(����) = 23 Then
    lblSA = "Sun"
ElseIf MySkill(����) = 24 Then
    lblSA = "pErfecT tErraN"
ElseIf MySkill(����) = 25 Then
    lblSA = "Brain"
ElseIf MySkill(����) = 26 Then
    lblSA = "zErg sPeicaL kILLeR"
ElseIf MySkill(����) = 27 Then
    lblSA = "�����"
ElseIf MySkill(����) = 28 Then
    lblSA = "����"
ElseIf MySkill(����) = 29 Then
    lblSA = "�����"
ElseIf MySkill(����) = 30 Then
    lblSA = "����ʫ"
End If

If Skill(Oee) = 1 Then
    lblSAO = "���"
ElseIf Skill(Oee) = 2 Then
    lblSAO = "����ܲ��"
ElseIf Skill(Oee) = 3 Then
    lblSAO = "����"
ElseIf Skill(Oee) = 4 Then
    lblSAO = "���"
ElseIf Skill(Oee) = 5 Then
    lblSAO = "���"
ElseIf Skill(Oee) = 6 Then
    lblSAO = "Maestro"
ElseIf Skill(Oee) = 7 Then
    lblSAO = "��ڪ"
ElseIf Skill(Oee) = 8 Then
    lblSAO = "ޫף"
ElseIf Skill(Oee) = 9 Then
    lblSAO = "��ף"
ElseIf Skill(Oee) = 10 Then
    lblSAO = "��ף"
ElseIf Skill(Oee) = 11 Then
    lblSAO = "�ף"
ElseIf Skill(Oee) = 12 Then
    lblSAO = "��ף"
ElseIf Skill(Oee) = 13 Then
    lblSAO = "��ף"
ElseIf Skill(Oee) = 14 Then
    lblSAO = "����"
ElseIf Skill(Oee) = 15 Then
    lblSAO = "����"
ElseIf Skill(Oee) = 16 Then
    lblSAO = "����"
ElseIf Skill(Oee) = 17 Then
    lblSAO = "����"
ElseIf Skill(Oee) = 18 Then
    lblSAO = "����"
ElseIf Skill(Oee) = 19 Then
    lblSAO = "����"
ElseIf Skill(Oee) = 20 Then
    lblSAO = "���"
ElseIf Skill(Oee) = 21 Then
    lblSAO = "rEd sNipeR"
ElseIf Skill(Oee) = 22 Then
    lblSAO = "������"
ElseIf Skill(Oee) = 23 Then
    lblSAO = "Sun"
ElseIf Skill(Oee) = 24 Then
    lblSAO = "pErfecT tErraN"
ElseIf Skill(Oee) = 25 Then
    lblSAO = "Brain"
ElseIf Skill(Oee) = 26 Then
    lblSAO = "zErg sPeicaL kILLeR"
ElseIf Skill(Oee) = 27 Then
    lblSAO = "�����"
ElseIf Skill(Oee) = 28 Then
    lblSAO = "����"
ElseIf Skill(Oee) = 29 Then
    lblSAO = "�����"
ElseIf Skill(Oee) = 30 Then
    lblSAO = "����ʫ"
End If


If ����(Oee) = 1 Then
 lblOT = "[Vs T]"
 lblMTT = "[" & MyTW(����) & "�� " & MyTL(����) & "��]"
 If MyT��(����) = "W" Then
  lblM���� = "[" & MyT����(����) & "������" & "]"
  lblM����.ForeColor = RGB(0, 255, 255)
 ElseIf MyT��(����) = "L" Then
  lblM���� = "[" & MyT����(����) & "������" & "]"
  lblM����.ForeColor = RGB(255, 0, 0)
 End If
ElseIf ����(Oee) = 2 Then
 lblOT = "[Vs Z]"
 lblMTT = "[" & MyZW(����) & "�� " & MyZL(����) & "��]"
 If MyZ��(����) = "W" Then
  lblM���� = "[" & MyZ����(����) & "������" & "]"
  lblM����.ForeColor = RGB(0, 255, 255)
 ElseIf MyZ��(����) = "L" Then
  lblM���� = "[" & MyZ����(����) & "������" & "]"
  lblM����.ForeColor = RGB(255, 0, 0)
 End If
ElseIf ����(Oee) = 3 Then
 lblOT = "[Vs P]"
 lblMTT = "[" & MyPW(����) & "�� " & MyPL(����) & "��]"
 If MyP��(����) = "W" Then
  lblM���� = "[" & MyP����(����) & "������" & "]"
  lblM����.ForeColor = RGB(0, 255, 255)
 ElseIf MyP��(����) = "L" Then
  lblM���� = "[" & MyP����(����) & "������" & "]"
  lblM����.ForeColor = RGB(255, 0, 0)
 End If
End If

If MyTribe(����) = 1 Then
 lblMT = "[Vs T]"
 lblOTT = "[" & T�¸�(Oee) & "�� " & T�й�(Oee) & "��]"
 If T��(Oee) = "W" Then
  lblO���� = "[" & T����(Oee) & "������" & "]"
  lblO����.ForeColor = RGB(0, 255, 255)
 ElseIf T��(Oee) = "L" Then
  lblO���� = "[" & T����(Oee) & "������" & "]"
  lblO����.ForeColor = RGB(255, 0, 0)
 End If
ElseIf MyTribe(����) = 2 Then
 lblMT = "[Vs Z]"
 lblOTT = "[" & Z�¸�(Oee) & "�� " & Z�й�(Oee) & "��]"
 If Z��(Oee) = "W" Then
  lblO���� = "[" & Z����(Oee) & "������" & "]"
  lblO����.ForeColor = RGB(0, 255, 255)
 ElseIf Z��(Oee) = "L" Then
  lblO���� = "[" & Z����(Oee) & "������" & "]"
  lblO����.ForeColor = RGB(255, 0, 0)
 End If
ElseIf MyTribe(����) = 3 Then
 lblMT = "[Vs P]"
 lblOTT = "[" & P�¸�(Oee) & "�� " & P�й�(Oee) & "��]"
 If P��(Oee) = "W" Then
  lblO���� = "[" & P����(Oee) & "������" & "]"
  lblO����.ForeColor = RGB(0, 255, 255)
 ElseIf P��(Oee) = "L" Then
  lblO���� = "[" & P����(Oee) & "������" & "]"
  lblO����.ForeColor = RGB(255, 0, 0)
 End If
End If

If Len(Dir(App.Path & "\img\����\[" & Mid(MyYear(����), 2, 2) & "]" & MN & ".gif")) <> 0 Then
 ImgMe = LoadPicture(App.Path & "\img\����\[" & Mid(MyYear(����), 2, 2) & "]" & MN & ".gif")
Else
 ImgMe = LoadPicture(App.Path & "\img\����\" & MyName(����) & ".gif")
End If

If Len(Dir(App.Path & "\img\����\[" & Mid(OYear(Oee), 2, 2) & "]" & �̸�(Oee) & ".gif")) <> 0 Then
 ImgOp = LoadPicture(App.Path & "\img\����\[" & Mid(OYear(Oee), 2, 2) & "]" & �̸�(Oee) & ".gif")
Else
 ImgOp = LoadPicture(App.Path & "\img\����\" & �̸�(Oee) & ".gif")
End If

If Len(Dir(App.Path & "\img\��\" & MapName(Map) & ".gif")) <> 0 Then
 ImgMa = LoadPicture(App.Path & "\img\��\" & MapName(Map) & ".gif")
Else
 ImgMa = Nothing
End If


AA = val(AT) + val(R) + val(St) + val(Am) + val(De) + val(Pa) + val(SE) + val(Co)
If val(AA) < 4500 Then
���̷�ũ = "F"
'&H4B4B4B
ElseIf val(AA) >= 4500 And val(AA) < 4700 Then
���̷�ũ = "E"
'&HB0B0B0
ElseIf val(AA) >= 4700 And val(AA) < 4800 Then
���̷�ũ = "D-"
'&HFF3232
ElseIf val(AA) >= 4800 And val(AA) < 4900 Then
���̷�ũ = "D"
'&HFF3232
ElseIf val(AA) >= 4900 And val(AA) < 5000 Then
���̷�ũ = "D+"
'&HFF3232
ElseIf val(AA) >= 5000 And val(AA) < 5100 Then
���̷�ũ = "C-"
'&HFF00&
ElseIf val(AA) >= 5100 And val(AA) < 5200 Then
���̷�ũ = "C"
'&HFF00&
ElseIf val(AA) >= 5200 And val(AA) < 5400 Then
���̷�ũ = "C+"
'&HFF00&
ElseIf val(AA) >= 5400 And val(AA) < 5600 Then
���̷�ũ = "B-"
'&HFFFD&
ElseIf val(AA) >= 5600 And val(AA) < 5800 Then
���̷�ũ = "B"
'&HFFFD&
ElseIf val(AA) >= 5800 And val(AA) < 6000 Then
���̷�ũ = "B+"
'&HFFFD&
ElseIf val(AA) >= 6000 And val(AA) < 6200 Then
���̷�ũ = "A-"
'&H6663FF
ElseIf val(AA) >= 6200 And val(AA) < 6400 Then
���̷�ũ = "A"
'&H6663FF
ElseIf val(AA) >= 6400 And val(AA) < 6600 Then
���̷�ũ = "A+"
'&H6663FF
ElseIf val(AA) >= 6600 And val(AA) < 6800 Then
���̷�ũ = "S"
ElseIf val(AA) >= 6800 And val(AA) < 7000 Then
���̷�ũ = "SS"
ElseIf val(AA) >= 7000 Then
���̷�ũ = "SSS"
End If

AAO = val(���ݷ�(Oee)) + val(����(Oee)) + val(����(Oee)) + val(����(Oee)) + val(�����(Oee)) + val(����(Oee)) + val(����(Oee)) + val(��Ʈ��(Oee))
If val(AAO) < 4500 Then
��뷩ũ = "F"
'&H4B4B4B
ElseIf val(AAO) >= 4500 And val(AAO) < 4700 Then
��뷩ũ = "E"
'&HB0B0B0
ElseIf val(AAO) >= 4700 And val(AAO) < 4800 Then
��뷩ũ = "D-"
'&HFF3232
ElseIf val(AAO) >= 4800 And val(AAO) < 4900 Then
��뷩ũ = "D"
'&HFF3232
ElseIf val(AAO) >= 4900 And val(AAO) < 5000 Then
��뷩ũ = "D+"
'&HFF3232
ElseIf val(AAO) >= 5000 And val(AAO) < 5100 Then
��뷩ũ = "C-"
'&HFF00&
ElseIf val(AAO) >= 5100 And val(AAO) < 5200 Then
��뷩ũ = "C"
'&HFF00&
ElseIf val(AAO) >= 5200 And val(AAO) < 5400 Then
��뷩ũ = "C+"
'&HFF00&
ElseIf val(AAO) >= 5400 And val(AAO) < 5600 Then
��뷩ũ = "B-"
'&HFFFD&
ElseIf val(AAO) >= 5600 And val(AAO) < 5800 Then
��뷩ũ = "B"
'&HFFFD&
ElseIf val(AAO) >= 5800 And val(AAO) < 6000 Then
��뷩ũ = "B+"
'&HFFFD&
ElseIf val(AAO) >= 6000 And val(AAO) < 6200 Then
��뷩ũ = "A-"
'&H6663FF
ElseIf val(AAO) >= 6200 And val(AAO) < 6400 Then
��뷩ũ = "A"
'&H6663FF
ElseIf val(AAO) >= 6400 And val(AAO) < 6600 Then
��뷩ũ = "A+"
'&H6663FF
ElseIf val(AAO) >= 6600 And val(AAO) < 6800 Then
��뷩ũ = "S"
ElseIf val(AAO) >= 6800 And val(AAO) < 7000 Then
��뷩ũ = "SS"
ElseIf val(AAO) >= 7000 Then
��뷩ũ = "SSS"
End If

lblMrank = MyRank(����)
lblOrank = ��ũ(Oee)
If MyRank(����) = "Normal" Then
 lblMrank.ForeColor = RGB(0, 0, 0)
ElseIf MyRank(����) = "Special" Then
 lblMrank.ForeColor = RGB(0, 255, 0)
ElseIf MyRank(����) = "Rare" Then
 lblMrank.ForeColor = &HFF80FF
ElseIf MyRank(����) = "Unique" Then
 lblMrank.ForeColor = &HFF8080
ElseIf MyRank(����) = "Elite" Then
 lblMrank.ForeColor = &H800080
ElseIf MyRank(����) = "Legend" Then
 lblMrank.ForeColor = &H80FF&
ElseIf MyRank(����) = "Secret" Then
 lblMrank.ForeColor = &HFFC0C0
ElseIf MyRank(����) = "Champion" Then
 lblMrank.ForeColor = RGB(255, 0, 0)
End If

If ��ũ(Oee) = "Normal" Then
 lblOrank.ForeColor = RGB(0, 0, 0)
ElseIf ��ũ(Oee) = "Special" Then
 lblOrank.ForeColor = RGB(0, 255, 0)
ElseIf ��ũ(Oee) = "Rare" Then
 lblOrank.ForeColor = &HFF80FF
ElseIf ��ũ(Oee) = "Unique" Then
 lblOrank.ForeColor = &HFF8080
ElseIf ��ũ(Oee) = "Elite" Then
 lblOrank.ForeColor = &H800080
ElseIf ��ũ(Oee) = "Legend" Then
 lblOrank.ForeColor = &H80FF&
ElseIf ��ũ(Oee) = "Secret" Then
 lblOrank.ForeColor = &HFFC0C0
ElseIf ��ũ(Oee) = "Champion" Then
 lblOrank.ForeColor = RGB(255, 0, 0)
End If



lblMapName = MapName(Map)
lblMR = "Rank : " & ���̷�ũ
lblOR = "Rank : " & ��뷩ũ
lblMSt = "Stats : " & AA
lblOSt = "Stats : " & AAO
End Sub

Private Sub jcbutton1_Click()
OStyle = Int((5 * Rnd) + 1)
If val(OStyle) = 1 Then
 OStyle = "������"
ElseIf val(OStyle) = 2 Then
 OStyle = "������"
ElseIf val(OStyle) = 3 Then
 OStyle = "������"
ElseIf val(OStyle) = 4 Then
 OStyle = "���"
ElseIf val(OStyle) = 5 Then
 OStyle = "�����"
End If
FrmPickSt.Visible = True
Unload Me
End Sub
