VERSION 5.00
Begin VB.Form FrmCoupon 
   BackColor       =   &H00FFFFFF&
   Caption         =   "���� �޴�"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4815
   Icon            =   "Form26.frx":0000
   LinkTopic       =   "Form26"
   ScaleHeight     =   1575
   ScaleWidth      =   4815
   StartUpPosition =   2  'ȭ�� ���
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
      Caption         =   "���� ���"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   49152
         Caption         =   "�ڷ� ����"
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
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   49152
         Caption         =   "���� ��������"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   49152
         Caption         =   "���� ����"
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
            Name            =   "����"
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
      Caption         =   "ȯ���մϴ�."
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16576
         Caption         =   "���� ���"
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
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16576
         Caption         =   "ī�� ����"
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
 MsgBox �ζ� & "Cro ��÷"
 Money = val(Money) - 50000
 Money = val(Money) + val(�ζ�)
 MsgBox �ζ� & "�� ��÷."
Else
 MsgBox "50000Cro�� �ʿ��ؿ�"
End If
End Sub

Private Sub jcbutton3_Click()
If val(����) >= 1 Then
 xFrame1.Visible = False
 xFrame2.Visible = True
Else
 MsgBox "������ �����ϴ� �Ф�...."
End If
End Sub

Private Sub jcbutton4_Click()
If val(����) >= 1 Then
    ���� = val(����) - 1
    Money = val(Money) + val(�ζ�)
    MsgBox �ζ� & "�� ��÷"
Else
    MsgBox "���� ���ڳ���; �ֱ׷��� �Ƹ��߾��;"
End If
End Sub

Private Sub jcbutton5_Click()
If val(����) >= 1 Then
    If val(������) <= 14 Then
        ���Ű��� = "Yes"
        ���� = val(����) - 1
        Do Until ��ũ(����NPC) = "Elite" Or ��ũ(����NPC) = "Normal" Or ��ũ(����NPC) = "Legend" Or ��ũ(����NPC) = "Secret"
            ����NPC = Int((800 * Rnd) + 1)
        Loop
        Dim ����NPC���� As Integer
        ����NPC���� = val(������) - 5
        If ���Ű��� = "Yes" Then
            SubName(����NPC����) = �̸�(����NPC)
            SubTeam(����NPC����) = Team(����NPC)
            SubAt(����NPC����) = NPC���ݷ�(����NPC)
            SubR(����NPC����) = NPC����(����NPC)
            SubSt(����NPC����) = NPC����(����NPC)
            SubAm(����NPC����) = NPC����(����NPC)
            SubDe(����NPC����) = NPC�����(����NPC)
            SubPa(����NPC����) = NPC����(����NPC)
            SubSe(����NPC����) = NPC����(����NPC)
            SubCo(����NPC����) = NPC��Ʈ��(����NPC)
            SubRank(����NPC����) = ��ũ(����NPC)
            SubYear(����NPC����) = OYear(����NPC)
            SubTribe(����NPC����) = ����(����NPC)
            SubLev(����NPC����) = 1
            SubExp(����NPC����) = 0
            SubMExp(����NPC����) = 50
            SubPoint(����NPC����) = 0
            SubNum(����NPC����) = val(����NPC)
            SubAW(����NPC����) = 0
            SubAL(����NPC����) = 0
            SubTW(����NPC����) = 0
            SubTL(����NPC����) = 0
            SubZW(����NPC����) = 0
            SubZL(����NPC����) = 0
            SubPW(����NPC����) = 0
            SubPL(����NPC����) = 0
            SubVic(����NPC����) = 0
            SubSeVic(����NPC����) = 0
            SubCode(����NPC����) = "B"
            SubSkill(����NPC����) = Skill(����NPC)
            ������ = val(������) + 1
            ���Ű��� = "No"
            If SubRank(����NPC����) = "Normal" Or SubRank(����NPC����) = "Special" Then
                SubNW(����NPC����) = "CB16"
            ElseIf SubRank(����NPC����) = "Rare" Then
                SubNW(����NPC����) = "CA1"
            ElseIf SubRank(����NPC����) = "Unique" Then
                SubNW(����NPC����) = "CA2"
            ElseIf SubRank(����NPC����) = "Elite" Then
                SubNW(����NPC����) = "CA3"
            Else
                SubNW(����NPC����) = "CS32"
            End If
                If 9.Text1 = "������" Then
                    9.Text1 = "������"
                Else
                    9.Text1 = "������"
                End If
                Unload FrmCoupon
                Unload Me
        End If
    Else
        MsgBox "�������� �ִ��Դϴ�."
    End If
Else
    MsgBox "���� ���ڳ���; �ֱ׷��� �Ƹ��߾��;"
End If
End Sub

Private Sub jcbutton6_Click()
Dim SKillChange As String
SKillChange = 0
If val(����) >= 1 Then
    Do Until val(SKillChange) >= 1 And val(SKillChange) <= 6
        SKillChange = InputBox("������ȣ�� �Է��ϼ���. 1 = " & MyYear(1) & MyName(1) & " 2 = " & MyYear(2) & MyName(2) & " 3 = " & MyYear(3) & MyName(3) & " 4 = " & MyYear(4) & MyName(4) & " 5 = " & MyYear(5) & MyName(5) & " 6 = " & MyYear(6) & MyName(6))
    Loop
    
    MySkill(val(SKillChange)) = Int((38 * Rnd) + 1)
    ���� = val(����) - 1
Else
    MsgBox "���� ���ڳ���; �ֱ׷��� �Ƹ��߾��;"
End If

End Sub

Private Sub jcbutton7_Click()
If val(����) >= 1 Then
    ���� = val(����) - 1
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
    
    For i = 1 To val(������ - 5)
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
    MsgBox "���� ���ڳ���; �ֱ׷��� �Ƹ��߾��;"
End If
End Sub

Private Sub jcbutton8_Click()
If val(����) >= 1 Then
    If val(������) <= 14 Then
        Dim ������ As String
        ������ = InputBox("���ϴ� �� �̸��� �����ּ���. �Ｚ���� eSTRO ���� MBC POS CJ GO �°��ӳ� ����Ʈ STX ������ ȭ�� PLUS Mystar 8th ���� �Ѻ� SK Orion IS 4U ���� Toona Pantech Curitel�߿� ���� �մϴ�. ��ҹ��� ������ ��Ȯ�� ���ּ���.")
        ���Ű��� = "Yes"
        ���� = val(����) - 1
        Do Until (Team(����NPC) = ������) And (��ũ(����NPC) <> "Champion") And (��ũ(����NPC) <> "Normal")
            ����NPC = Int((800 * Rnd) + 1)
        Loop
        Dim ����NPC���� As Integer
        ����NPC���� = val(������) - 5
        If ���Ű��� = "Yes" Then
            SubName(����NPC����) = �̸�(����NPC)
            SubTeam(����NPC����) = Team(����NPC)
            SubAt(����NPC����) = NPC���ݷ�(����NPC)
            SubR(����NPC����) = NPC����(����NPC)
            SubSt(����NPC����) = NPC����(����NPC)
            SubAm(����NPC����) = NPC����(����NPC)
            SubDe(����NPC����) = NPC�����(����NPC)
            SubPa(����NPC����) = NPC����(����NPC)
            SubSe(����NPC����) = NPC����(����NPC)
            SubCo(����NPC����) = NPC��Ʈ��(����NPC)
            SubRank(����NPC����) = ��ũ(����NPC)
            SubYear(����NPC����) = OYear(����NPC)
            SubTribe(����NPC����) = ����(����NPC)
            SubLev(����NPC����) = 1
            SubExp(����NPC����) = 0
            SubMExp(����NPC����) = 50
            SubPoint(����NPC����) = 0
            SubNum(����NPC����) = val(����NPC)
            SubAW(����NPC����) = 0
            SubAL(����NPC����) = 0
            SubTW(����NPC����) = 0
            SubTL(����NPC����) = 0
            SubZW(����NPC����) = 0
            SubZL(����NPC����) = 0
            SubPW(����NPC����) = 0
            SubPL(����NPC����) = 0
            SubVic(����NPC����) = 0
            SubSeVic(����NPC����) = 0
            SubCode(����NPC����) = "B"
            SubSkill(����NPC����) = Skill(����NPC)
            ������ = val(������) + 1
            ���Ű��� = "No"
            If SubRank(����NPC����) = "Normal" Or SubRank(����NPC����) = "Special" Then
                SubNW(����NPC����) = "CB16"
            ElseIf SubRank(����NPC����) = "Rare" Then
                SubNW(����NPC����) = "CA1"
            ElseIf SubRank(����NPC����) = "Unique" Then
                SubNW(����NPC����) = "CA2"
            ElseIf SubRank(����NPC����) = "Elite" Then
                SubNW(����NPC����) = "CA3"
            Else
                SubNW(����NPC����) = "CS32"
            End If
                If 9.Text1 = "������" Then
                    9.Text1 = "������"
                Else
                    9.Text1 = "������"
                End If
                Unload FrmCoupon
                Unload Me
        End If
    Else
        MsgBox "�������� �ִ��Դϴ�."
    End If
End If
End Sub

Private Sub jcbutton9_Click()
xFrame1.Visible = True
xFrame2.Visible = False
End Sub
