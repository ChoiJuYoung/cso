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
   StartUpPosition =   2  'ȭ�� ���
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
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "���Ⱥй�"
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
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "����ġ %"
   End
   Begin VB.Label lblLeDe 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label2"
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
      Top             =   3480
      Width           =   5895
   End
   Begin VB.Label lbl���� 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label2"
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
      TabIndex        =   23
      Top             =   2640
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label1"
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
      BackStyle       =   0  '����
      Caption         =   "����Ʈ :"
      BeginProperty Font 
         Name            =   "����"
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
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BackStyle       =   0  '����
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
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
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "vsP : 0�� 0��"
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
      Left            =   3840
      TabIndex        =   17
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label lblZ 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "vsZ : 0�� 0��"
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
      Left            =   1800
      TabIndex        =   16
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label lblT 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "vsT : 0�� 0��"
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
      TabIndex        =   15
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblA 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "vsA : 0�� 0��"
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
      Left            =   3840
      TabIndex        =   14
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label lblVic 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Ư�̻��� : ����"
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
      Top             =   3000
      Width           =   5895
   End
   Begin VB.Label lblTri 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Tribe : Protoss"
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
      Top             =   2400
      Width           =   5895
   End
   Begin VB.Label lblNa 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Name : �谡��"
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
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "��   �� : 1000"
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
      Left            =   1800
      TabIndex        =   10
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label lblAm 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "��   �� : 1000"
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
      TabIndex        =   9
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label lblTeam 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�Ҽ��� : �Ｚ����"
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
      Left            =   3840
      TabIndex        =   8
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label lblR 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "��   �� : 1000"
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
      Left            =   1800
      TabIndex        =   7
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label lblPa 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "��   �� : 1000"
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
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lblExp 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "����ġ : 0 %"
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
      Left            =   3840
      TabIndex        =   5
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label lblCo 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "��Ʈ�� : 1000"
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
      Left            =   1800
      TabIndex        =   4
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label lblDe 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "����� : 1000"
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
      TabIndex        =   3
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lblLv 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Level : 1"
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
      Left            =   3840
      TabIndex        =   2
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lblSt 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "��   �� : 1000"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lblAt 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "���ݷ� : 1000"
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
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00EAE2F1&
      BackStyle       =   1  '�������� ����
      Height          =   1575
      Left            =   0
      Top             =   3840
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00AE73E5&
      BackStyle       =   1  '�������� ����
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
 If MyNW(����) = "CB16" Then
  lblLeDe = "Code B 16��"
 ElseIf MyNW(����) = "CB8" Then
  lblLeDe = "Code B 8��"
 ElseIf MyNW(����) = "CB4" Then
  lblLeDe = "Code B 4��"
 ElseIf MyNW(����) = "CBFin" Then
  lblLeDe = "Code B �����"
 ElseIf MyNW(����) = "CA1" Then
  lblLeDe = "Code A 1Round"
 ElseIf MyNW(����) = "CA2" Then
  lblLeDe = "Code A 2Round"
 ElseIf MyNW(����) = "CA3" Then
  lblLeDe = "Code A 3Round"
 ElseIf MyNW(����) = "CS32" Then
  lblLeDe = "Code S 32��"
 ElseIf MyNW(����) = "CS16" Then
  lblLeDe = "Code S 16��"
 ElseIf MyNW(����) = "CS8" Then
  lblLeDe = "Code S 8��"
 ElseIf MyNW(����) = "CS4" Then
  lblLeDe = "Code S 4��"
 ElseIf MyNW(����) = "CSFin" Then
  lblLeDe = "Code S �����"
 ElseIf MyNW(����) = "UpADo" Then
  lblLeDe = "�°� ������"
 End If
 
Dim ���� As Integer
For ���� = 1 To 6
 MyMExp(����) = val(MyLev(����)) * 50
Next
lbl���� = MyA����(����)
If MyA��(����) = "W" Then
 lbl���� = "���� : " & lbl���� & "����"
Else
 lbl���� = "���� : " & lbl���� & "����"
End If

If MySkill(����) = 1 Then
    Label1 = "��ų :���"
ElseIf MySkill(����) = 2 Then
    Label1 = "��ų :����ܲ��"
ElseIf MySkill(����) = 3 Then
    Label1 = "��ų :����"
ElseIf MySkill(����) = 4 Then
    Label1 = "��ų :���"
ElseIf MySkill(����) = 5 Then
    Label1 = "��ų :���"
ElseIf MySkill(����) = 6 Then
    Label1 = "��ų :Maestro"
ElseIf MySkill(����) = 7 Then
    Label1 = "��ų :��ڪ"
ElseIf MySkill(����) = 8 Then
    Label1 = "��ų :ޫף"
ElseIf MySkill(����) = 9 Then
    Label1 = "��ų :��ף"
ElseIf MySkill(����) = 10 Then
    Label1 = "��ų :��ף"
ElseIf MySkill(����) = 11 Then
    Label1 = "��ų :�ף"
ElseIf MySkill(����) = 12 Then
    Label1 = "��ų :��ף"
ElseIf MySkill(����) = 13 Then
    Label1 = "��ų :��ף"
ElseIf MySkill(����) = 14 Then
    Label1 = "��ų :����"
ElseIf MySkill(����) = 15 Then
    Label1 = "��ų :����"
ElseIf MySkill(����) = 16 Then
    Label1 = "��ų :����"
ElseIf MySkill(����) = 17 Then
    Label1 = "��ų :����"
ElseIf MySkill(����) = 18 Then
    Label1 = "��ų :����"
ElseIf MySkill(����) = 19 Then
    Label1 = "��ų :����"
ElseIf MySkill(����) = 20 Then
    Label1 = "��ų :���"
ElseIf MySkill(����) = 21 Then
    Label1 = "��ų :rEd sNipeR"
ElseIf MySkill(����) = 22 Then
    Label1 = "��ų :������"
ElseIf MySkill(����) = 23 Then
    Label1 = "��ų :Sun"
ElseIf MySkill(����) = 24 Then
    Label1 = "��ų :pErfecT tErraN"
ElseIf MySkill(����) = 25 Then
    Label1 = "��ų :Brain"
ElseIf MySkill(����) = 26 Then
    Label1 = "��ų :zErg sPeicaL kILLeR"
ElseIf MySkill(����) = 27 Then
    Label1 = "��ų :�����"
ElseIf MySkill(����) = 28 Then
    Label1 = "��ų :����"
ElseIf MySkill(����) = 29 Then
    Label1 = "��ų :�����"
ElseIf MySkill(����) = 30 Then
    Label1 = "��ų :����ʫ"
ElseIf MySkill(����) = 31 Then
    Label1 = "��ų :���ݷ�߾"
ElseIf MySkill(����) = 32 Then
    Label1 = "��ų :����߾"
ElseIf MySkill(����) = 33 Then
    Label1 = "��ų :����߾"
ElseIf MySkill(����) = 34 Then
    Label1 = "��ų :����߾"
ElseIf MySkill(����) = 35 Then
    Label1 = "��ų :�����߾"
ElseIf MySkill(����) = 36 Then
    Label1 = "��ų :����߾"
ElseIf MySkill(����) = 37 Then
    Label1 = "��ų :����߾"
ElseIf MySkill(����) = 38 Then
    Label1 = "��ų :��Ʈ��߾"
Else
    Label1 = "��ų :����"
End If


Label15 = MyPoint(����)
For ���� = 1 To 9
 SubMExp(����) = val(SubMExp(����)) * 50
Next
lblVic = "��� : " & MyVic(����) & "   " & "�ؿ�� : " & MySeVic(����)
lblAt = "���ݷ� : " & MyAt(����)
lblSt = "��   �� : " & MySt(����)
lblAm = "��   �� : " & MyAm(����)
lblR = "��   �� : " & MyR(����)
lblDe = "����� : " & MyDe(����)
lblPa = "��   �� : " & MyPa(����)
lblCo = "��Ʈ�� : " & MyCo(����)
lblSe = "��   �� : " & MySe(����)
lblNa = "Name : " & MyYear(����) & MyName(����)
lblTeam = "�Ҽ��� : " & MyTeam(����)
lblLV = "Level : " & MyLev(����)
lblExp = "����ġ : " & Int(val(MyExp(����)) * 100 / val(MyMExp(����))) & " %"
lblA = "vsA : " + MyAW(����) + "�� " + MyAL(����) + "��"
lblT = "vsT : " + MyTW(����) + "�� " + MyTL(����) + "��"
lblZ = "vsZ : " + MyZW(����) + "�� " + MyZL(����) + "��"
lblP = "vsP : " + MyPW(����) + "�� " + MyPL(����) + "��"
If MyTribe(����) = 1 Then
lblTri = "Tribe : Terran"
ElseIf MyTribe(����) = 2 Then
lblTri = "Tribe : Zerg"
ElseIf MyTribe(����) = 3 Then
lblTri = "Tribe : Protoss"
End If
lblSpe = "Ư�̻��� : ����"
If Len(Dir(App.Path & "\img\����\[" & Mid(MyYear(����), 2, 2) & "]" & MyName(����) & ".gif")) <> 0 Then
 Img = LoadPicture(App.Path & "\img\����\[" & Mid(MyYear(����), 2, 2) & "]" & MyName(����) & ".gif")
Else
 Img = LoadPicture(App.Path & "\img\����\" & MyName(����) & ".gif")
End If
ProgressBar1.Value = Int(val(MyExp(����)) * 100 / val(MyMExp(����)))
ProgressBar1.Text = Int(val(MyExp(����))) & " / " & Int(val(MyMExp(����)))


If MyRank(����) = "Normal" Then
 Shape1.BackColor = RGB(255, 255, 255)
ElseIf MyRank(����) = "Special" Then
 Shape1.BackColor = RGB(0, 255, 0)
ElseIf MyRank(����) = "Rare" Then
 Shape1.BackColor = &HFF80FF
ElseIf MyRank(����) = "Unique" Then
 Shape1.BackColor = &HFF8080
ElseIf MyRank(����) = "Elite" Then
 Shape1.BackColor = &H800080
ElseIf MyRank(����) = "Legend" Then
 Shape1.BackColor = &H80FF&
ElseIf MyRank(����) = "Secret" Then
 Shape1.BackColor = &HFFC0C0
ElseIf MyRank(����) = "Champion" Then
 Shape1.BackColor = RGB(255, 0, 0)
End If
End Sub

Private Sub Img_Click()
Dim AdCode As String
AdCode = InputBox("Code�Է�")
If AdCode = "sEtting" Then
    ���� = ����
    FrmBug.Show
End If
End Sub

Private Sub jcbutton1_Click()
FrmStat.Show
Unload Me
End Sub

