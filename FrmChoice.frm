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
   StartUpPosition =   2  'ȭ�� ���
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
         Name            =   "����"
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
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      BackStyle       =   0  '����
      Caption         =   "NAme"
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
      Index           =   5
      Left            =   8160
      TabIndex        =   13
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  '����
      Caption         =   "NAme"
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
      Index           =   4
      Left            =   6480
      TabIndex        =   12
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  '����
      Caption         =   "NAme"
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
      Index           =   3
      Left            =   9840
      TabIndex        =   11
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  '����
      Caption         =   "NAme"
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
      Index           =   2
      Left            =   8160
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  '����
      Caption         =   "NAme"
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
      Index           =   1
      Left            =   6480
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  '����
      Caption         =   "NAme"
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
      Index           =   6
      Left            =   9840
      TabIndex        =   8
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblSum 
      BackStyle       =   0  '����
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
      Left            =   9000
      TabIndex        =   7
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label lblPTribe 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
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
      Left            =   4560
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblPSum 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
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
      Left            =   4560
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblPRank 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
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
      Left            =   4560
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblPName 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
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
      BackStyle       =   1  '�������� ����
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
Private Ŭ�� As Integer

Private Sub CmdGo_Click()
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
Next
���� = 0
����Ƚ�� = 0

Money = 5000
TeamName = TxtName

FrmMain.Show
Unload Me
End Sub

Private Sub CmdSel_Click()
MyName(������) = �̸�(���(Ŭ��))
MyTribe(������) = ����(���(Ŭ��))
MyAt(������) = ���ݷ�(���(Ŭ��))
MyR(������) = ����(���(Ŭ��))
MySt(������) = ����(���(Ŭ��))
MyAm(������) = ����(���(Ŭ��))
MyDe(������) = �����(���(Ŭ��))
MyPa(������) = ����(���(Ŭ��))
MySe(������) = ����(���(Ŭ��))
MyCo(������) = ��Ʈ��(���(Ŭ��))
MyYear(������) = OYear(���(Ŭ��))
MyRank(������) = ��ũ(���(Ŭ��))
MyTeam(������) = Team(���(Ŭ��))
MySkill(������) = Skill(���(Ŭ��))
PlayNumber(������) = Oee
Randomize Oee
Call LoadImage(ImgCon(������), MyName(������), MyYear(������))
Number = Number + val(���ݷ�(���(Ŭ��))) + val(����(���(Ŭ��))) + val(����(���(Ŭ��))) + val(����(���(Ŭ��))) + val(�����(���(Ŭ��))) + val(����(���(Ŭ��))) + val(����(���(Ŭ��))) + val(��Ʈ��(���(Ŭ��)))
Call lblNameAlter(lblName(������), 1, val(������))
DoEvents
If lblName(������).ForeColor = RGB(255, 255, 255) Then
    lblName(������).ForeColor = RGB(0, 0, 0)
End If


������ = val(������) + 1
If val(������) >= 7 Then
    CmdGo.Visible = True
    CmdSel.Visible = False
    lblSum = "�ɷ�ġ ��� : " & Int(Number / 6)
    ������ = 6
Else
    TimOee.Enabled = False
    TimRan.Enabled = True
End If
End Sub

Private Sub Form_Load()
Number = 0
Oee = Int((800 * Rnd) + 1)
������ = 1
If ��÷��� = 1 Then
    For i = 1 To 3
        Do Until (��ũ(Oee) = "Unique") And (����(Oee) = 1)
            Oee = Int((800 * Rnd) + 1)
        Loop
        ���(i) = Oee
        Oee = Int((800 * Rnd) + 1)
    Next
ElseIf ��÷��� = 2 Then
    For i = 1 To 3
        Do Until (��ũ(Oee) = "Unique") And (����(Oee) = 1)
            Oee = Int((800 * Rnd) + 1)
        Loop
        ���(i) = Oee
        Oee = Int((800 * Rnd) + 1)
    Next
ElseIf ��÷��� = 3 Then
    For i = 1 To 3
        Do Until ��ũ(Oee) = "Special" And ����(Oee) = 1
            Oee = Int((800 * Rnd) + 1)
        Loop
        ���(i) = Oee
        Oee = Int((800 * Rnd) + 1)
    Next
ElseIf ��÷��� = 4 Then
    For i = 1 To 3
        Do Until ��ũ(Oee) = "Rare" And ����(Oee) = 1
            Oee = Int((800 * Rnd) + 1)
        Loop
        ���(i) = Oee
        Oee = Int((800 * Rnd) + 1)
    Next
ElseIf ��÷��� = 5 Then
    For i = 1 To 3
        Do Until ��ũ(Oee) = "Rare" And ����(Oee) = 1
            Oee = Int((800 * Rnd) + 1)
        Loop
        ���(i) = Oee
        Oee = Int((800 * Rnd) + 1)
    Next
Else
    For i = 1 To 3
        Do Until ��ũ(Oee) = "Special" And ����(Oee) = 1
            Oee = Int((800 * Rnd) + 1)
        Loop
        ���(i) = Oee
        Oee = Int((800 * Rnd) + 1)
    Next
End If

For i = 1 To 3
    Call LoadImage(ImgChoice(i), �̸�(���(i)), OYear(���(i)))
Next

TimOee.Enabled = True
End Sub

Private Sub ImgChoice_Click(Index As Integer)
Shape1.BackColor = &HC0C0C1
Shape1.BackColor = &HC0C0C0
Ŭ�� = Index
DoEvents
Call MakeLineCom(Me, ���(Index), 4320, 3840)
Call LoadImage(Image1, �̸�(���(Index)), OYear(���(Index)))
lblPName = OYear(���(Index)) & �̸�(���(Index))
lblPRank = ��ũ(���(Index))
Call lblTribeAlter(lblPTribe, val(����(���(Index))))
lblPTribe = "���� : " & Left(Right(lblPTribe, 2), 1)
lblPSum = "�ɷ�ġ �հ� : " & val(���ݷ�(���(Index))) + val(����(���(Index))) + val(����(���(Index))) + val(����(���(Index))) + val(�����(���(Index))) + val(����(���(Index))) + val(����(���(Index))) + val(��Ʈ��(���(Index)))
End Sub

Private Sub TimOee_Timer()
Oee = Int((800 * Rnd) + 1)
End Sub

Private Sub TimRan_Timer()
TimRan.Enabled = False

If ������ = 2 Then
    If ��÷��� = 1 Then
        For i = 1 To 3
            Do Until (��ũ(Oee) = "Rare") And (����(Oee) = 2)
                Oee = Int((800 * Rnd) + 1)
            Loop
            ���(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    ElseIf ��÷��� = 2 Then
        For i = 1 To 3
            Do Until (��ũ(Oee) = "Special") And (����(Oee) = 2)
                Oee = Int((800 * Rnd) + 1)
            Loop
            ���(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    ElseIf ��÷��� = 3 Then
        For i = 1 To 3
            Do Until ��ũ(Oee) = "Unique" And ����(Oee) = 2
                Oee = Int((800 * Rnd) + 1)
            Loop
            ���(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    ElseIf ��÷��� = 4 Then
        For i = 1 To 3
            Do Until ��ũ(Oee) = "Unique" And ����(Oee) = 2
                Oee = Int((800 * Rnd) + 1)
            Loop
            ���(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    ElseIf ��÷��� = 5 Then
        For i = 1 To 3
            Do Until ��ũ(Oee) = "Special" And ����(Oee) = 2
                Oee = Int((800 * Rnd) + 1)
            Loop
            ���(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    Else
        For i = 1 To 3
            Do Until ��ũ(Oee) = "Rare" And ����(Oee) = 2
                Oee = Int((800 * Rnd) + 1)
            Loop
            ���(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    End If
ElseIf ������ = 3 Then
    If ��÷��� = 1 Then
        For i = 1 To 3
            Do Until (��ũ(Oee) = "Special") And (����(Oee) = 3)
                Oee = Int((800 * Rnd) + 1)
            Loop
            ���(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    ElseIf ��÷��� = 2 Then
        For i = 1 To 3
            Do Until (��ũ(Oee) = "Rare") And (����(Oee) = 3)
                Oee = Int((800 * Rnd) + 1)
            Loop
            ���(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    ElseIf ��÷��� = 3 Then
        For i = 1 To 3
            Do Until ��ũ(Oee) = "Rare" And ����(Oee) = 3
                Oee = Int((800 * Rnd) + 1)
            Loop
            ���(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    ElseIf ��÷��� = 4 Then
        For i = 1 To 3
            Do Until ��ũ(Oee) = "Special" And ����(Oee) = 3
                Oee = Int((800 * Rnd) + 1)
            Loop
            ���(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    ElseIf ��÷��� = 5 Then
        For i = 1 To 3
            Do Until ��ũ(Oee) = "Unique" And ����(Oee) = 3
                Oee = Int((800 * Rnd) + 1)
            Loop
            ���(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    Else
        For i = 1 To 3
            Do Until ��ũ(Oee) = "Unique" And ����(Oee) = 3
                Oee = Int((800 * Rnd) + 1)
            Loop
            ���(i) = Oee
            Oee = Int((800 * Rnd) + 1)
        Next
    End If
ElseIf ������ >= 4 Then
    For i = 1 To 3
        Do Until (��ũ(Oee) = "Normal") And (����(Oee) = val(������) - 3)
            Oee = Int((800 * Rnd) + 1)
        Loop
        ���(i) = Oee
        Oee = Int((800 * Rnd) + 1)
    Next
End If

For i = 1 To 3
    Call LoadImage(ImgChoice(i), �̸�(���(i)), OYear(���(i)))
Next

End Sub
