VERSION 5.00
Begin VB.Form FrmPlayGame 
   BackColor       =   &H00000000&
   Caption         =   "����â"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   8580
   ScaleWidth      =   11250
   StartUpPosition =   2  'ȭ�� ���
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
      Caption         =   "�ɷ�ġ ��"
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
         Name            =   "����"
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
         Name            =   "����"
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
      BackStyle       =   1  '�������� ����
      Height          =   3135
      Left            =   3000
      Top             =   5040
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�ʻ� :"
      BeginProperty Font 
         Name            =   "����"
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
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      BackStyle       =   0  '����
      Caption         =   "���¾�������"
      BeginProperty Font 
         Name            =   "����"
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
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
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
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
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
      Alignment       =   1  '������ ����
      BackColor       =   &H00000000&
      Caption         =   "(Z)"
      BeginProperty Font 
         Name            =   "����"
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
      Alignment       =   1  '������ ����
      BackColor       =   &H00000000&
      Caption         =   "<11> �̿�ȣ"
      BeginProperty Font 
         Name            =   "����"
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
      Alignment       =   1  '������ ����
      BackColor       =   &H00000000&
      Caption         =   "(Z)"
      BeginProperty Font 
         Name            =   "����"
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
      Alignment       =   1  '������ ����
      BackColor       =   &H00000000&
      Caption         =   "<11> �̿�ȣ"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      Caption         =   "VS"
      BeginProperty Font 
         Name            =   "����"
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
If ����(Oee) = 1 Then
    lblOTr = "(T)"
ElseIf ����(Oee) = 2 Then
    lblOTr = "(Z)"
Else
    lblOTr = "(P)"
End If

If MyTribe(����) = ����(Oee) Then
    lblMyTr = lblOTr
    PGB.Value = 50
    PGB.Text = MyYear(����) & MyName(����) & "50 : 50" & OYear(����) & �̸�(Oee)
Else
    If MyTribe(����) = 1 Then
        lblMyTr = "(T)"
        If ����(Oee) = 2 Then
            PGB.Value = TZT(Map)
            PGB.Text = MyYear(����) & MyName(����) & TZT(Map) & " : " & TZZ(Map) & OYear(Oee) & �̸�(Oee)
        Else
            PGB.Value = PTT(Map)
            PGB.Text = MyYear(����) & MyName(����) & PTT(Map) & " : " & PTP(Map) & OYear(Oee) & �̸�(Oee)
        End If
    ElseIf MyTribe(����) = 2 Then
        lblMyTr = "(Z)"
        If ����(Oee) = 1 Then
            PGB.Value = TZZ(Map)
            PGB.Text = MyYear(����) & MyName(����) & TZZ(Map) & " : " & TZT(Map) & OYear(Oee) & �̸�(Oee)
        Else
            PGB.Value = ZPZ(Map)
            PGB.Text = MyYear(����) & MyName(����) & ZPZ(Map) & " : " & ZPP(Map) & OYear(Oee) & �̸�(Oee)
        End If
    Else
        lblMyTr = "(P)"
        If ����(Oee) = 1 Then
            PGB.Value = PTP(Map)
            PGB.Text = MyYear(����) & MyName(����) & PTP(Map) & " : " & PTT(Map) & OYear(Oee) & �̸�(Oee)
        Else
            PGB.Value = ZPP(Map)
            PGB.Text = MyYear(����) & MyName(����) & ZPP(Map) & " : " & ZPP(Map) & OYear(Oee) & �̸�(Oee)
        End If
    End If
End If

If MyRank(����) = "Normal" Then
    My��ũ�� = 1
ElseIf MyRank(����) = "Special" Then
    My��ũ�� = 2
ElseIf MyRank(����) = "Rare" Then
    My��ũ�� = 3
ElseIf MyRank(����) = "Unique" Then
    My��ũ�� = 4
ElseIf MyRank(����) = "Elite" Then
    My��ũ�� = 5
ElseIf MyRank(����) = "Legend" Then
    My��ũ�� = 6
ElseIf MyRank(����) = "Secret" Then
    My��ũ�� = 7
Else
    My��ũ�� = 10
End If

If ��ũ(Oee) = "Normal" Then
    O��ũ�� = 1
ElseIf ��ũ(Oee) = "Special" Then
    O��ũ�� = 2
ElseIf ��ũ(Oee) = "Rare" Then
    O��ũ�� = 3
ElseIf ��ũ(Oee) = "Unique" Then
    O��ũ�� = 4
ElseIf ��ũ(Oee) = "Elite" Then
    O��ũ�� = 5
ElseIf ��ũ(Oee) = "Legend" Then
    O��ũ�� = 6
ElseIf ��ũ(Oee) = "Secret" Then
    O��ũ�� = 7
Else
    O��ũ�� = 10
End If

If Turn = "OSL" Then
    MW = 0
    OW = 0
    AAA = 0
    lblLe = "MystarCraft�� ��Ÿ����"
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
ElseIf Turn = "PL" Then
    lblLe = "MystarCraft�� ���θ���"
    lblLeDe = ""
End If

If Len(Dir(App.Path & "\img\����\[" & Mid(OYear(Oee), 2, 2) & "]" & �̸�(Oee) & ".gif")) <> 0 Then
    ImgOp = LoadPicture(App.Path & "\img\����\[" & Mid(OYear(Oee), 2, 2) & "]" & �̸�(Oee) & ".gif")
Else
    ImgOp = LoadPicture(App.Path & "\img\����\" & �̸�(Oee) & ".gif")
End If

If Len(Dir(App.Path & "\img\����\[" & Mid(MyYear(����), 2, 2) & "]" & MyName(����) & ".gif")) <> 0 Then
    ImgP = LoadPicture(App.Path & "\img\����\[" & Mid(MyYear(����), 2, 2) & "]" & MyName(����) & ".gif")
Else
    ImgP = LoadPicture(App.Path & "\img\����\" & MyName(����) & ".gif")
End If


ImgM.Picture = LoadPicture(App.Path & "\img\��\" & MapName(Map) & ".gif")


lblM = MapName(Map)
lblMyName = MyYear(����) & " " & MyName(����)
lblOName = OYear(Oee) & " " & �̸�(Oee)
lblMW = MW
lblOW = OW
End Sub

Private Sub jcbutton1_Click()
i = 1

L1 = MyAt(����) / 100
L2 = MyR(����) / 100
L3 = MySt(����) / 100
L4 = MyAm(����) / 100
L5 = MyDe(����) / 100
L6 = MyPa(����) / 100
L7 = MySe(����) / 100
L8 = MyCo(����) / 100

r1 = val(���ݷ�(Oee)) / 100
r2 = val(����(Oee)) / 100
r3 = val(����(Oee)) / 100
R4 = val(����(Oee)) / 100
R5 = val(�����(Oee)) / 100
R6 = val(����(Oee)) / 100
R7 = val(����(Oee)) / 100
R8 = val(��Ʈ��(Oee)) / 100

jcbutton1.Visible = False
Timer1.Enabled = True
End Sub

Private Sub lblle_click()
Dim CodeS As String
CodeS = InputBox("")
If CodeS = "��" Then
    Winer = "��"
    If val(Text2) <> 1 Then
        Text2 = 1
    Else
        Text2 = 2
    End If
ElseIf CodeS = "���" Then
    Winer = "���"
    If val(Text2) <> 1 Then
        Text2 = 1
    Else
        Text2 = 2
    End If
End If
End Sub

Private Sub CmdGo_Click()
Randomize ������
Randomize AP

If MyTribe(����) = ����(Oee) Then
    MP = 1: OP = 1
ElseIf MyTribe(����) = 1 Then
    If ����(Oee) = 2 Then
        MP = TZT(Map)
        OP = TZZ(Map)
    Else
        MP = PTT(Map)
        OP = PTP(Map)
    End If
ElseIf MyTribe(����) = 2 Then
    If ����(Oee) = 1 Then
        MP = TZZ(Map)
        OP = TZT(Map)
    Else
        MP = ZPZ(Map)
        OP = ZPP(Map)
    End If
ElseIf MyTribe(����) = 3 Then
    If ����(Oee) = 1 Then
        MP = PTP(Map)
        OP = PTT(Map)
    Else
        MP = ZPP(Map)
        OP = ZPZ(Map)
    End If
End If


AT = MyAt(����): ATO = ���ݷ�(Oee)
R = MyR(����): RO = ����(Oee)
St = MySt(����): StO = ����(Oee)
Am = MyAm(����): AmO = ����(Oee)
De = MyDe(����): DeO = �����(Oee)
Pa = MyPa(����): PaO = ����(Oee)
SE = MySe(����): SeO = ����(Oee)
Co = MyCo(����): CoO = ��Ʈ��(Oee)

If MySkill(����) = 1 Then
    If ����(Oee) = 2 Then
        Co = val(MyCo(����)) + 150
    End If
ElseIf MySkill(����) = 2 Then
    If val(MW) + val(OW) >= 5 Then
        AT = val(MyAt(����)) + 50
        R = val(MyR(����)) + 50
        St = val(MySt(����)) + 50
        Am = val(MyAm(����)) + 50
        De = val(MyDe(����)) + 50
        Pa = val(MyPa(����)) + 50
        SE = val(MySe(����)) + 50
        Co = val(MyCo(����)) + 50
    End If
ElseIf MySkill(����) = 3 Then
    If ����(Oee) = 2 Then
        De = val(MyDe(����)) - 25
    ElseIf ����(Oee) = 3 Then
        AT = val(MyAt(����)) + 150
    End If
ElseIf MySkill(����) = 4 Then
    If ����(Oee) = 3 Then
        Am = val(MyAm(����)) + 125
    End If
ElseIf MySkill(����) = 5 Then
    If ����(Oee) = 1 Then
        AT = val(MyAt(����)) + 25
    Else
        Am = val(MyAm(����)) + 50
    End If
ElseIf MySkill(����) = 6 Then
    If ����(Oee) = 2 Then
        AT = val(MyAt(����)) - 75
    ElseIf ����(Oee) = 3 Then
        Am = val(MyAm(����)) + 200
    End If
ElseIf MySkill(����) = 7 Then
    Am = val(MyAm(����)) + 50
    If ����(Oee) = 1 Then
        SE = val(MySe(����)) - 25
    End If
ElseIf MySkill(����) = 8 Then
    If ����(Oee) = 2 Then
        R = val(MyR(����)) + 50
        Co = val(MyCo(����)) + 50
    End If
ElseIf MySkill(����) = 9 Then
    If ����(Oee) = 1 Or ����(Oee) = 3 Then
        Am = val(MyAm(����)) + 50
    End If
ElseIf MySkill(����) = 10 Then
    R = val(MyR(����)) + 30
    If ����(Oee) = 1 Then
        Am = val(MyAm(����)) + 10
    End If
ElseIf MySkill(����) = 11 Then
    If ����(Oee) = 3 Then
        Co = val(MyCo(����)) + 100
    End If
ElseIf MySkill(����) = 12 Then
    If ����(Oee) = 1 Then
        Am = val(MyAm(����)) + 100
        AT = val(MyAt(����)) + 100
    ElseIf ����(Oee) = 2 Then
        Co = val(MyCo(����)) - 100
    End If
ElseIf MySkill(����) = 13 Then
    If ����(Oee) = 3 Then
        Co = val(MyCo(����)) + 50
        Am = val(MyAm(����)) + 50
    End If
ElseIf MySkill(����) = 14 Then
    If ����(Oee) = 3 Then
        R = val(MyR(����)) + 100
    End If
ElseIf MySkill(����) = 15 Then
    If ����(Oee) = 3 Then
        AT = val(MyAt(����)) + 100
    End If
ElseIf MySkill(����) = 16 Then
    If ����(Oee) = 2 Then
        AT = val(MyAt(����)) + 50
        Co = val(MyCo(����)) + 50
    End If
ElseIf MySkill(����) = 17 Then
    Am = val(MyAm(����)) + 25
ElseIf MySkill(����) = 18 Then
    If ����(Oee) = 2 Then
        R = val(MyR(����)) + 75
    End If
ElseIf MySkill(����) = 19 Then
    Am = val(MyAm(����)) + 50
    Co = val(MyCo(����)) - 25
ElseIf MySkill(����) = 20 Then
    If ����(Oee) = 2 Then
        Co = val(MyCo(����)) + 75
    End If
ElseIf MySkill(����) = 21 Then
    If ����(Oee) = 2 Then
        Co = val(MyCo(����)) + 100
        AT = val(MyAt(����)) + 50
    ElseIf ����(Oee) = 3 Then
        R = val(MyR(����)) - 75
    End If
ElseIf MySkill(����) = 22 Then
    If val(MW) < val(OW) Then
        AT = val(MyAt(����)) + 30
        R = val(MyR(����)) + 30
        St = val(MySt(����)) + 30
        Am = val(MyAm(����)) + 30
        De = val(MyDe(����)) + 30
        Pa = val(MyPa(����)) + 30
        SE = val(MySe(����)) + 30
        Co = val(MyCo(����)) + 30
    End If
ElseIf MySkill(����) = 23 Then
    If ����(Oee) = 1 Then
        SE = val(MySe(����)) + 75
    End If
ElseIf MySkill(����) = 24 Then
    If ����(Oee) = 3 Then
        Am = val(MyAm(����)) + 75
    End If
ElseIf MySkill(����) = 25 Then
    If ����(Oee) = 2 Then
        Co = val(MyCo(����)) - 50
    ElseIf ����(Oee) = 3 Then
        Am = val(MyAm(����)) + 125
    End If
ElseIf MySkill(����) = 26 Then
    If ����(Oee) = 2 Then
        AT = val(MyAt(����)) + 100
    ElseIf ����(Oee) = 3 Then
        Am = val(MyAm(����)) - 25
    End If
ElseIf MySkill(����) = 27 Then
    If ����(Oee) = 1 Then
        Am = val(MyAm(����)) + 75
    End If
ElseIf MySkill(����) = 28 Then
    De = val(MyDe(����)) + 25
ElseIf MySkill(����) = 29 Then
    If ����(Oee) = 2 Then
        Co = val(MyCo(����)) + 200
    ElseIf ����(Oee) = 3 Then
        Am = val(MyAm(����)) - 125
    End If
ElseIf MySkill(����) = 30 Then
    If ����(Oee) = 2 Then
        R = val(MyR(����)) + 125
    End If
ElseIf MySkill(����) = 31 Then
    AT = val(MyAt(����)) + 25
ElseIf MySkill(����) = 32 Then
    R = val(MyR(����)) + 25
ElseIf MySkill(����) = 33 Then
    St = val(MySt(����)) + 25
ElseIf MySkill(����) = 34 Then
    Am = val(MyAm(����)) + 25
ElseIf MySkill(����) = 35 Then
    De = val(MyDe(����)) + 25
ElseIf MySkill(����) = 36 Then
    Pa = val(MyPa(����)) + 25
ElseIf MySkill(����) = 37 Then
    SE = val(MySe(����)) + 25
ElseIf MySkill(����) = 38 Then
    Co = val(MyCo(����)) + 25
End If


If Skill(Oee) = 1 Then
    If MyTribe(����) = 2 Then
        CoO = val(��Ʈ��(Oee)) + 150
    End If
ElseIf Skill(Oee) = 2 Then
    If val(MW) + val(OW) >= 5 Then
        RAT = val(���ݷ�(Oee)) + 50
        RO = val(����(Oee)) + 50
        StO = val(����(Oee)) + 50
        AmO = val(����(Oee)) + 50
        DeO = val(�����(Oee)) + 50
        PaO = val(����(Oee)) + 50
        SeO = val(����(Oee)) + 50
        CoO = val(��Ʈ��(Oee)) + 50
    End If
ElseIf Skill(Oee) = 3 Then
    If MyTribe(����) = 2 Then
        DeO = val(�����(Oee)) - 25
    ElseIf MyTribe(����) = 3 Then
        RAT = val(���ݷ�(Oee)) + 150
    End If
ElseIf Skill(Oee) = 4 Then
    If MyTribe(����) = 3 Then
        AmO = val(����(Oee)) + 125
    End If
ElseIf Skill(Oee) = 5 Then
    If MyTribe(����) = 1 Then
        RAT = val(���ݷ�(Oee)) + 25
    Else
        AmO = val(����(Oee)) + 50
    End If
ElseIf Skill(Oee) = 6 Then
    If MyTribe(����) = 2 Then
        RAT = val(���ݷ�(Oee)) - 75
    ElseIf MyTribe(����) = 3 Then
        AmO = val(����(Oee)) + 200
    End If
ElseIf Skill(Oee) = 7 Then
    AmO = val(����(Oee)) + 50
    If MyTribe(����) = 1 Then
        SeO = val(����(Oee)) - 25
    End If
ElseIf Skill(Oee) = 8 Then
    If MyTribe(����) = 2 Then
        RO = val(����(Oee)) + 50
        CoO = val(��Ʈ��(Oee)) + 50
    End If
ElseIf Skill(Oee) = 9 Then
    If MyTribe(����) = 1 Or MyTribe(����) = 3 Then
        AmO = val(����(Oee)) + 50
    End If
ElseIf Skill(Oee) = 10 Then
    RO = val(����(Oee)) + 30
    If MyTribe(����) = 1 Then
        AmO = val(����(Oee)) + 10
    End If
ElseIf Skill(Oee) = 11 Then
    If MyTribe(����) = 3 Then
        CoO = val(��Ʈ��(Oee)) + 100
    End If
ElseIf Skill(Oee) = 12 Then
    If MyTribe(����) = 1 Then
        AmO = val(����(Oee)) + 100
        RAT = val(���ݷ�(Oee)) + 100
    ElseIf MyTribe(����) = 2 Then
        CoO = val(��Ʈ��(Oee)) - 100
    End If
ElseIf Skill(Oee) = 13 Then
    If MyTribe(����) = 3 Then
        CoO = val(��Ʈ��(Oee)) + 50
        AmO = val(����(Oee)) + 50
    End If
ElseIf Skill(Oee) = 14 Then
    If MyTribe(����) = 3 Then
        RO = val(����(Oee)) + 100
    End If
ElseIf Skill(Oee) = 15 Then
    If MyTribe(����) = 3 Then
        RAT = val(���ݷ�(Oee)) + 100
    End If
ElseIf Skill(Oee) = 16 Then
    If MyTribe(����) = 2 Then
        RAT = val(���ݷ�(Oee)) + 50
        CoO = val(��Ʈ��(Oee)) + 50
    End If
ElseIf Skill(Oee) = 17 Then
    AmO = val(����(Oee)) + 25
ElseIf Skill(Oee) = 18 Then
    If MyTribe(����) = 2 Then
        RO = val(����(Oee)) + 75
    End If
ElseIf Skill(Oee) = 19 Then
    Am = val(����(Oee)) + 50
    Co = val(��Ʈ��(Oee)) - 25
ElseIf Skill(Oee) = 20 Then
    If MyTribe(����) = 2 Then
        CoO = val(��Ʈ��(Oee)) + 75
    End If
ElseIf Skill(Oee) = 21 Then
    If MyTribe(����) = 2 Then
        CoO = val(��Ʈ��(Oee)) + 100
        RAT = val(���ݷ�(Oee)) + 50
    ElseIf MyTribe(����) = 3 Then
        RO = val(����(Oee)) - 75
    End If
ElseIf Skill(Oee) = 22 Then
    If val(OW) < val(MW) Then
        RAT = val(���ݷ�(Oee)) + 30
        RO = val(����(Oee)) + 30
        StO = val(����(Oee)) + 30
        AmO = val(����(Oee)) + 30
        DeO = val(�����(Oee)) + 30
        PaO = val(����(Oee)) + 30
        SeO = val(����(Oee)) + 30
        CoO = val(��Ʈ��(Oee)) + 30
    End If
ElseIf Skill(Oee) = 23 Then
    If MyTribe(����) = 1 Then
        SeO = val(����(Oee)) + 75
    End If
ElseIf Skill(Oee) = 24 Then
    If MyTribe(����) = 3 Then
        AmO = val(����(Oee)) + 75
    End If
ElseIf Skill(Oee) = 25 Then
    If MyTribe(����) = 2 Then
        CoO = val(��Ʈ��(Oee)) - 50
    ElseIf MyTribe(����) = 3 Then
        AmO = val(����(Oee)) + 125
    End If
ElseIf Skill(Oee) = 26 Then
    If MyTribe(����) = 2 Then
        RAT = val(���ݷ�(Oee)) + 100
    ElseIf MyTribe(����) = 3 Then
        AmO = val(����(Oee)) - 25
    End If
ElseIf Skill(Oee) = 27 Then
    If MyTribe(����) = 1 Then
        AmO = val(����(Oee)) + 75
    End If
ElseIf Skill(Oee) = 28 Then
    DeO = val(�����(Oee)) + 25
ElseIf Skill(Oee) = 29 Then
    If MyTribe(����) = 2 Then
        CoO = val(��Ʈ��(Oee)) + 200
    ElseIf MyTribe(����) = 3 Then
        AmO = val(����(Oee)) - 125
    End If
ElseIf Skill(Oee) = 30 Then
    If MyTribe(����) = 2 Then
        RO = val(����(Oee)) + 125
    End If
ElseIf Skill(Oee) = 31 Then
    RAT = val(���ݷ�(Oee)) + 25
ElseIf Skill(Oee) = 32 Then
    RO = val(����(Oee)) + 25
ElseIf Skill(Oee) = 33 Then
    StO = val(����(Oee)) + 25
ElseIf Skill(Oee) = 34 Then
    AmO = val(����(Oee)) + 25
ElseIf Skill(Oee) = 35 Then
    DeO = val(�����(Oee)) + 25
ElseIf Skill(Oee) = 36 Then
    PaO = val(����(Oee)) + 25
ElseIf Skill(Oee) = 37 Then
    SeO = val(����(Oee)) + 25
ElseIf Skill(Oee) = 38 Then
    CoO = val(��Ʈ��(Oee)) + 25
End If

If Deck <> "" Then
    If Deck�⵵ = False Then
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

������ = Int((Oee * Rnd) + 1)
For ������ = 1 To ������
    AP = val((101 * Rnd) + 0)
Next

If MyTribe(����) = 1 Then
    If ����(Oee) = 1 Then
        MP = val(SE) * val(Co) * val(Am) * 20 / 1000000
        OP = val(SeO) * val(CoO) * val(AmO) * 20 / 1000000
    ElseIf ����(Oee) = 2 Then
        MP = val(AT) * val(Co) * val(St) * val(R) * 20 / 100000000
        OP = val(AmO) * val(DeO) * val(StO) * val(ATO) * 20 / 100000000
    Else
        MP = val(Am) * val(De) * val(R) * 20 / 100000000
        OP = val(ATO) * val(AmO) * val(DeO) * 20 / 100000000
    End If
ElseIf MyTribe(����) = 2 Then
    If ����(Oee) = 1 Then
        MP = val(Am) * val(De) * val(St) * val(AT) * 20 / 100000000
        OP = val(ATO) * val(CoO) * val(StO) * val(RO) * 20 / 100000000
    ElseIf ����(Oee) = 2 Then
        MP = ((val(AT) * val(Co) * val(SE) / 1000000) ^ 2)
        OP = ((val(ATO) * val(CoO) * val(SeO) / 1000000) ^ 2)
    Else
        MP = val(Am) * val(De) * val(Co) * 20 / 1000000
        OP = val(PaO) * val(RO) * val(CoO) * 20 / 1000000
    End If
Else
    If ����(Oee) = 1 Then
        MP = val(De) * val(AT) * val(Am) * 20 / 1000000
        OP = val(AmO) * val(DeO) * val(RO) * 20 / 1000000
    ElseIf ����(Oee) = 2 Then
        MP = val(Pa) * val(R) * val(Co) * 20 / 1000000
        OP = val(AmO) * val(DeO) * val(CoO) * 20 / 1000000
    Else
        MP = val(Am) * val(Co) * val(SE) * val(R) * 20 / 1000000
        OP = val(AmO) * val(CoO) * val(SeO) * val(RO) * 20 / 1000000
    End If
End If
MP = (val(MP) / 100) * val(Pa)
OP = (val(OP) / 100) * val(PaO)

If val(�����Ÿ�(Map)) = 1 Then
    MP = MP + val(AT) * 5
    OP = OP + val(���ݷ�(Oee)) * 5
ElseIf val(�����Ÿ�(Map)) = 2 Then
    MP = MP + val(AT) * 4
    OP = OP + val(���ݷ�(Oee)) * 4
ElseIf val(�����Ÿ�(Map)) = 3 Then
    MP = MP + val(AT) * 3
    OP = OP + val(���ݷ�(Oee)) * 3
ElseIf val(�����Ÿ�(Map)) = 4 Then
    MP = MP + val(AT) * 2
    OP = OP + val(���ݷ�(Oee)) * 2
ElseIf val(�����Ÿ�(Map)) = 5 Then
    MP = MP + (val(AT) + val(De)) * 1
    OP = OP + (val(���ݷ�(Oee)) + val(�����(Oee))) * 1
ElseIf val(�����Ÿ�(Map)) = 6 Then
    MP = MP + val(De) * 2
    OP = OP + val(�����(Oee)) * 2
ElseIf val(�����Ÿ�(Map)) = 7 Then
    MP = MP + val(De) * 3
    OP = OP + val(�����(Oee)) * 3
ElseIf val(�����Ÿ�(Map)) = 8 Then
    MP = MP + val(De) * 4
    OP = OP + val(�����(Oee)) * 4
ElseIf val(�����Ÿ�(Map)) = 9 Then
    MP = MP + val(De) * 5
    OP = OP + val(�����(Oee)) * 5
End If

MP = val(MP) + val(Am) * val(�ڿ�(Map))
OP = val(OP) + val(����(Oee)) * val(�ڿ�(Map))

MP = val(MP) + (val(St) + val(Pa)) * val(���⵵(Map))
OP = val(OP) + (val(����(Oee)) + val(����(Oee))) * val(���⵵(Map))

MP = Int(val(MP) / 100)
OP = Int(val(OP) / 100)
If val(My��ũ��) > val(O��ũ��) Then
    MP = val(MP) * 2 * val(val(My��ũ��) - val(O��ũ��))
ElseIf val(My��ũ��) < val(O��ũ��) Then
    OP = val(OP) * 2 * val(val(O��ũ��) - val(My��ũ��))
End If

If val(RAA) > val(RAAO) Then
    MP = val(MP) + val(RAA) * 200
ElseIf val(RAA) < val(RAAO) Then
    OP = val(OP) + val(RAAO) * 200
End If

���� = val(MP) * 100 / val(val(MP) + val(OP))


If val(����) <= 1 Then
    ���� = 4
ElseIf val(����) >= 99 Then
    ���� = 95
End If
If 0 <= val(AP) And val(AP) <= val(����) Then
    Winer = "��"
ElseIf val(����) < val(AP) And val(AP) <= 100 Then
    Winer = "���"
Else
    MsgBox "�����Դϴ�. �ٽ� �����ּ���"
End If


Dim ���� As Long
Randomize ����
���� = val((100 * Rnd) + 1)
If 1 <= ���� And 3 >= ���� Then
    If Winer = "��" Then
        Winer = "���"
    Else
        Winer = "��"
    End If
End If

If Text2.Text <> "������" Then
    Text2 = "������"
Else
    Text2 = "����"
End If
End Sub

Private Sub text2_change()
If Winer = "��" Then
    Money = val(Money) + val((Int(val(RAAO) / 1000) + 1) * 15)
    MW = val(MW) + 1
    MW2 = val(MW2) + 1
    MyExp(����) = val(MyExp(����)) + val(RAAO) / 1000 + 1
    If Mode = "Hell" Then
        MyExp(����) = val(MyExp(����) + 3)
    End If
    MyAW(����) = val(MyAW(����)) + 1
    A�й�(Oee) = val(A�й�(Oee)) + 1
    If MT = 1 Then
        T�й�(Oee) = val(T�й�(Oee)) + 1
        If T��(Oee) = "W" Then
            T��(Oee) = "L"
            T����(Oee) = 1
        Else
            T����(Oee) = val(T����(Oee)) + 1
        End If
    ElseIf MT = 2 Then
        Z�й�(Oee) = val(Z�й�(Oee)) + 1
        If Z��(Oee) = "W" Then
            Z��(Oee) = "L"
            Z����(Oee) = 1
        Else
            Z����(Oee) = val(Z����(Oee)) + 1
        End If
    ElseIf MT = 3 Then
        P�й�(Oee) = val(P�й�(Oee)) + 1
        If P��(Oee) = "W" Then
            P��(Oee) = "L"
            P����(Oee) = 1
        Else
            P����(Oee) = val(P����(Oee)) + 1
        End If
    End If
    If ����(Oee) = 1 Then
        MyTW(����) = val(MyTW(����)) + 1
    ElseIf ����(Oee) = 2 Then
        MyZW(����) = val(MyZW(����)) + 1
    ElseIf ����(Oee) = 3 Then
        MyPW(����) = val(MyPW(����)) + 1
    End If
    If MyA��(����) = "L" Then
        MyA��(����) = "W"
        MyA����(����) = 1
    Else
        MyA����(����) = val(MyA����(����)) + 1
    End If
    If ����(Oee) = 1 Then
        If MyT��(����) = "L" Then
            MyT��(����) = "W"
            MyT����(����) = 1
        Else
            MyT����(����) = val(MyT����(����)) + 1
        End If
    ElseIf ����(Oee) = 2 Then
        If MyZ��(����) = "L" Then
            MyZ��(����) = "W"
            MyZ����(����) = 1
        Else
            MyZ����(����) = val(MyZ����(����)) + 1
        End If
    ElseIf ����(Oee) = 3 Then
        If MyP��(����) = "L" Then
            MyP��(����) = "W"
            MyP����(����) = 1
        Else
            MyP����(����) = val(MyP����(����)) + 1
        End If
    End If
ElseIf Winer = "���" Then
    OW = val(OW) + 1
    OW2 = val(OW2) + 1
    MyExp(����) = val(MyExp(����)) - val(Int(val(RAA) / 1500) + 1)
    MyAL(����) = val(MyAL(����)) + 1
    A�¸�(Oee) = val(A�¸�(Oee)) + 1
    If MT = 1 Then
        T�¸�(Oee) = val(T�¸�(Oee)) + 1
        If T��(Oee) = "L" Then
            T��(Oee) = "W"
            T����(Oee) = 1
        Else
            T����(Oee) = val(T����(Oee)) + 1
        End If
    ElseIf MT = 2 Then
        Z�¸�(Oee) = val(Z�¸�(Oee)) + 1
        If Z��(Oee) = "L" Then
            Z��(Oee) = "W"
            Z����(Oee) = 1
        Else
            Z����(Oee) = val(Z����(Oee)) + 1
        End If
    ElseIf MT = 3 Then
        P�¸�(Oee) = val(P�¸�(Oee)) + 1
        If P��(Oee) = "L" Then
            P��(Oee) = "W"
            P����(Oee) = 1
        Else
            P����(Oee) = val(P����(Oee)) + 1
        End If
    End If
    If ����(Oee) = 1 Then
        MyTL(����) = val(MyTL(����)) + 1
    ElseIf ����(Oee) = 2 Then
        MyZL(����) = val(MyZL(����)) + 1
    ElseIf ����(Oee) = 3 Then
        MyPL(����) = val(MyPL(����)) + 1
    End If
    If MyA��(����) = "W" Then
        MyA��(����) = "L"
        MyA����(����) = 1
    Else
        MyA����(����) = val(MyA����(����)) + 1
    End If
    If ����(Oee) = 1 Then
        If MyT��(����) = "W" Then
            MyT��(����) = "L"
            MyT����(����) = 1
        Else
            MyT����(����) = val(MyT����(����)) + 1
        End If
    ElseIf ����(Oee) = 2 Then
        If MyZ��(����) = "W" Then
            MyZ��(����) = "L"
            MyZ����(����) = 1
        Else
            MyZ����(����) = val(MyZ����(����)) + 1
        End If
    ElseIf ����(Oee) = 3 Then
        If MyP��(����) = "W" Then
            MyP��(����) = "L"
            MyP����(����) = 1
        Else
            MyP����(����) = val(MyP����(����)) + 1
        End If
    End If
End If

lblMW = val(MW)
lblOW = val(OW)
Map = Int((12 * Rnd) + 1)
ImgM.Picture = LoadPicture(App.Path & "\img\��\" & MapName(Map) & ".gif")
lblM = MapName(Map)

If Turn = "OSL" Then
    If val(val(MW) + val(OW)) >= val(SetA) Then
        If MyNW(����) = "CB16" Then
            If val(MW) = 1 Then
                MyNW(����) = "CB8"
                AAA = 1
            ElseIf val(OW) = 1 Then
                MyNW(����) = "CB16"
                AAA = 1
            End If
        ElseIf MyNW(����) = "CB8" Then
            If val(MW) = 1 Then
                MyNW(����) = "CB4"
                AAA = 1
            ElseIf val(OW) = 1 Then
                MyNW(����) = "CB16"
                AAA = 1
            End If
        ElseIf MyNW(����) = "CB4" Then
            If val(MW) = 1 Then
                MyNW(����) = "CBFin"
                AAA = 1
            ElseIf val(OW) = 1 Then
                MyNW(����) = "CB16"
                AAA = 1
            End If
        ElseIf MyNW(����) = "CBFin" Then
            If val(MW) = 1 Then
                MyNW(����) = "CA1"
                AAA = 1
            ElseIf val(OW) = 1 Then
                MyNW(����) = "CB16"
                AAA = 1
            End If
        ElseIf MyNW(����) = "CA1" Then
            If val(MW) = 1 Then
                MyNW(����) = "CA2"
                AAA = 1
            ElseIf val(OW) = 1 Then
                MyNW(����) = "CB16"
                AAA = 1
            End If
        ElseIf MyNW(����) = "CA2" Then
            If val(MW) = 1 Then
                MyNW(����) = "CA3"
                AAA = 1
            ElseIf val(OW) = 1 Then
                MyNW(����) = "UpADo"
                AAA = 1
            End If
        ElseIf MyNW(����) = "CA3" Then
            If val(MW) = 1 Then
                MyNW(����) = "CS32"
                AAA = 1
            ElseIf val(OW) = 1 Then
                MyNW(����) = "UpADo"
                AAA = 1
            End If
        ElseIf MyNW(����) = "UpADo" Then
            If val(MW) = 3 Then
                MyNW(����) = "CS32"
                AAA = 1
            ElseIf val(OW) = 3 Then
                MyNW(����) = "CA1"
                AAA = 1
            End If
        ElseIf MyNW(����) = "CS32" Then
            If val(MW) = 2 Then
                MyNW(����) = "CS16"
                AAA = 1
            ElseIf val(OW) = 2 Then
                MyNW(����) = "CA1"
                AAA = 1
            End If
        ElseIf MyNW(����) = "CS16" Then
            If val(MW) = 2 Then
                MyNW(����) = "CS8"
                AAA = 1
            ElseIf val(OW) = 2 Then
                MyNW(����) = "CA2"
                AAA = 1
            End If
        ElseIf MyNW(����) = "CS8" Then
            If val(MW) = 3 Then
                MyNW(����) = "CS4"
                AAA = 1
            ElseIf val(OW) = 3 Then
                MyNW(����) = "CA3"
                AAA = 1
            End If
        ElseIf MyNW(����) = "CS4" Then
            If val(MW) = 3 Then
                MyNW(����) = "CSFin"
                AAA = 1
            ElseIf val(OW) = 3 Then
                MyNW(����) = "CS32"
                AAA = 1
            End If
        ElseIf MyNW(����) = "CSFin" Then
            If val(MW) >= 4 Then
                MyNW(����) = "CS32"
                MyVic(����) = val(MyVic(����)) + 1
                �ؿ��(Oee) = val(�ؿ��(Oee)) + 1
                MsgBox "Code S���� ����ϼ̽��ϴ�! �����մϴ�!! �� + 20000"
                Money = val(Money) + 20000
                If Mode = "Normal" Then
                    StatPlusFin = 5
                ElseIf Mode = "Hard" Then
                    StatPlusFin = 7
                Else
                    StatPlusFin = 15
                End If
                For i = 0 To 800
                    ���ݷ�(i) = val(���ݷ�(i)) + val(StatPlusFin)
                    ����(i) = val(����(i)) + val(StatPlusFin)
                    ����(i) = val(����(i)) + val(StatPlusFin)
                    ����(i) = val(����(i)) + val(StatPlusFin)
                    �����(i) = val(�����(i)) + val(StatPlusFin)
                    ����(i) = val(����(i)) + val(StatPlusFin)
                    ����(i) = val(����(i)) + val(StatPlusFin)
                    ��Ʈ��(i) = val(��Ʈ��(i)) + val(StatPlusFin)
                Next
                AAA = 1
            ElseIf val(OW) = 4 Then
                MyNW(����) = "CS32"
                MySeVic(����) = val(MySeVic(����)) + 1
                ���(Oee) = val(���(Oee)) + 1
                MsgBox "Code S���� �ؿ���ϼ̽��ϴ�! ���ϵ����! �� + 7500"
                Money = val(Money) + 7500
                If Mode = "Normal" Then
                    StatPlusFin = 4
                ElseIf Mode = "Hard" Then
                    StatPlusFin = 6
                Else
                    StatPlusFin = 10
                End If
                For i = 0 To 800
                    ���ݷ�(i) = val(���ݷ�(i)) + val(StatPlusFin)
                    ����(i) = val(����(i)) + val(StatPlusFin)
                    ����(i) = val(����(i)) + val(StatPlusFin)
                    ����(i) = val(����(i)) + val(StatPlusFin)
                    �����(i) = val(�����(i)) + val(StatPlusFin)
                    ����(i) = val(����(i)) + val(StatPlusFin)
                    ����(i) = val(����(i)) + val(StatPlusFin)
                    ��Ʈ��(i) = val(��Ʈ��(i)) + val(StatPlusFin)
                Next
                AAA = 1
            End If
        End If
    End If
Else
    If PL���� = "1R" Or PL���� = "2R" Or PL���� = "3R" Then
        If val(MW) + val(OW) < 3 Then
            PL������(����) = False
            FrmResult.Show
            Unload Me
        Else
            If val(MW) >= 3 Or val(OW) >= 3 Then
                For i = 1 To 6
                    PL������(i) = True
                Next
                If val(MW) >= 3 Then
                    PL�� = val(PL��) + 1
                    For i = 1 To 6
                    MyExp(i) = val(MyExp(i)) + 7
                    Next
                Else
                    PL�� = val(PL��) + 1
                    For i = 1 To 6
                        MyExp(i) = val(MyExp(i)) - 5
                    Next
                End If
                
                PL���� = val(PL����) + 1
                
                If val(PL����) >= 12 Then
                    If PL���� = "1R" Then
                        PL���� = "2R"
                        MsgBox "������ �����մϴ�."
                        PL���� = 0
                        FrmMain.CmdSa.Visible = True
                        VisibleȮ�� = True
                    ElseIf PL���� = "2R" Then
                        PL���� = "3R"
                        MsgBox "�����̰����մϴ�."
                        PL���� = 0
                        FrmMain.CmdSa.Visible = True
                        VisibleȮ�� = True
                    Else
                        PL���� = Int((12 * Rnd) + 0)
                        If val(PL��) >= 33 Then
                            PL���� = "Final"
                            MsgBox "Proleague, ����� ����!"
                        ElseIf val(PL��) >= 30 Then
                            PL���� = "PO"
                            MsgBox "Proleague, �÷��̿��� ����!"
                        ElseIf val(PL��) >= 25 Then
                            PL���� = "6��"
                            MsgBox "Proleague, 6�� ����!"
                        Else
                            PL���� = "1R"
                            PL�ѹ� = 2
                            PL���� = 0
                            MsgBox "����Ʈ���� Ż��"
                        End If
                    End If
                End If
                MW = 0
                OW = 0
                FrmResult.Show
                PLEnd = "True"
                Unload Me
            Else
                PL������(����) = False
                FrmResult.Show
                Unload Me
            End If
        End If
    Else
        If val(MW) + val(OW) < 4 Then
            PL������(����) = False
            FrmResult.Show
            Unload Me
        ElseIf val(MW) + val(OW) >= 4 Then
            If val(MW) >= 4 Or val(OW) >= 4 Then
            PLEnd = "True"
                For i = 1 To 6
                    PL������(i) = True
                Next
            PL�� = 0
            PL�� = 0
            PL�ѹ� = 2
                If val(MW) >= 4 Then
                    PL���� = Int((12 * Rnd) + 0)
                    If PL���� = "6��" Then
                        PL���� = "SPO"
                    ElseIf PL���� = "SPO" Then
                        PL���� = "PO"
                    ElseIf PL���� = "PO" Then
                        PL���� = "Final"
                    Else
                        PL��� = val(PL���) + 1
                        PL���� = "1R"
                        PL���� = 0
                        Money = val(Money) + 10000
                        MsgBox "���θ��� ���! ��Ű�" & ������ & "���� �������� �ϱų� ���! �� + 10000"
                        If Mode = "Normal" Then
                            StatPlusFin = 2
                        ElseIf Mode = "Hard" Then
                            StatPlusFin = 5
                        Else
                            StatPlusFin = 7
                        End If
                        For i = 0 To 800
                            ���ݷ�(i) = val(���ݷ�(i)) + StatPlusFin
                            ����(i) = val(����(i)) + val(StatPlusFin)
                            ����(i) = val(����(i)) + val(StatPlusFin)
                            ����(i) = val(����(i)) + val(StatPlusFin)
                            �����(i) = val(�����(i)) + val(StatPlusFin)
                            ����(i) = val(����(i)) + val(StatPlusFin)
                            ����(i) = val(����(i)) + val(StatPlusFin)
                            ��Ʈ��(i) = val(��Ʈ��(i)) + val(StatPlusFin)
                        Next
                        MsgBox "������ �����մϴ�."
                        FrmMain.CmdSa.Visible = True
                        VisibleȮ�� = True
                    End If
                Else
                    PL���� = "1R"
                    PL���� = 0
                    If PL���� = "Final" Then
                        MsgBox "�ƽ��� �ؿ��! �� + 7000"
                        PL�ؿ�� = val(PL�ؿ��) + 1
                        Money = val(Money) + 7000
                        If Mode = "Normal" Then
                            StatPlusFin = 1
                        ElseIf Mode = "Hard" Then
                            StatPlusFin = 4
                        Else
                            StatPlusFin = 6
                        End If
                        For i = 0 To 800
                            ���ݷ�(i) = val(���ݷ�(i)) + StatPlusFin
                            ����(i) = val(����(i)) + val(StatPlusFin)
                            ����(i) = val(����(i)) + val(StatPlusFin)
                            ����(i) = val(����(i)) + val(StatPlusFin)
                            �����(i) = val(�����(i)) + val(StatPlusFin)
                            ����(i) = val(����(i)) + val(StatPlusFin)
                            ����(i) = val(����(i)) + val(StatPlusFin)
                            ��Ʈ��(i) = val(��Ʈ��(i)) + val(StatPlusFin)
                        Next
                    End If
                    FrmMain.CmdSa.Visible = True
                    MsgBox "������ �����մϴ�."
                    VisibleȮ�� = True
                End If
                MW = 0
                OW = 0
            Else
                PL������(����) = False
                PLEnd = "False"
                If val(MW) = 3 And val(OW) = 3 Then
                    For i = 1 To 6
                        PL������(i) = True
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
