VERSION 5.00
Begin VB.Form FrmHighShop 
   BackColor       =   &H00FFFFFF&
   Caption         =   "��޻���"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Form19.frx":0000
   LinkTopic       =   "Form19"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'ȭ�� ���
   Begin CSO.jcbutton jcbutton1 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
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
      Caption         =   "�̱�"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin CSO.jcbutton jcbutton2 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
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
      Caption         =   "�̱�"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Label Label7 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "���� �Ͻðڽ��ϱ�?"
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
      TabIndex        =   7
      Top             =   2520
      Width           =   4695
   End
   Begin VB.Label Label6 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "[100000Cro]"
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
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   4695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   4  '���-��
      X1              =   0
      X2              =   4680
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label5 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "[Unique ~ Legend]"
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
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Width           =   4695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   4  '���-��
      X1              =   0
      X2              =   4680
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label4 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "���� �Ͻðڽ��ϱ�??"
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
      TabIndex        =   4
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "[50000Cro]"
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
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "[Normal ~ Legend]"
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
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "��� ������ ���Ű��� ȯ���մϴ�. �ΰ����� �޴��� �����մϴ�. ������ ��ðڽ��ϱ�?"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "FrmHighShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
If Mode = "Hell" Then
    Label2 = "[Normal ~ Champion]"
    Label5 = "[Unique ~ Champion]"
End If
End Sub

Private Sub jcbutton1_Click()
Randomize Oee
Oee = Int((801 * Rnd) + 0)
 If val(Money) >= 100000 Then
  If Mode = "Hard" Then
   Do Until (��ũ(����NPC) = "Unique") Or (��ũ(����NPC) = "Elite") Or (��ũ(����NPC) = "Legend")
    ����NPC = Int((800 * Rnd) + 1)
   Loop
  Else
   Do Until (��ũ(����NPC) = "Unique") Or (��ũ(����NPC) = "Elite") Or (��ũ(����NPC) = "Legend") Or (��ũ(����NPC) = "Secret") Or (��ũ(����NPC) = "Champion")
    ����NPC = Int((801 * Rnd) + 0)
   Loop
  End If
   ���Ű��� = "Yes"
   Money = val(Money) - 100000
 Else
  MsgBox "���� �����մϴ�. 10������ ��Ƽ� ���ʽÿ�."
 End If
 
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
 Unload FrmShop
 Unload Me
End If
End Sub

Private Sub jcbutton2_Click()
Randomize Oee
Oee = Int((801 * Rnd) + 0)
 If val(Money) >= 50000 Then
  If Mode = "Hard" Then
   Do Until (��ũ(����NPC) = "Normal") Or (��ũ(����NPC) = "Special") Or (��ũ(����NPC) = "Rare") Or (��ũ(����NPC) = "Unique") Or (��ũ(����NPC) = "Elite") Or (��ũ(����NPC) = "Legend")
    ����NPC = Int((800 * Rnd) + 1)
   Loop
  Else
   Do Until (��ũ(����NPC) = "Normal") Or (��ũ(����NPC) = "Special") Or (��ũ(����NPC) = "Rare") Or (��ũ(����NPC) = "Unique") Or (��ũ(����NPC) = "Elite") Or (��ũ(����NPC) = "Legend") Or (��ũ(����NPC) = "Secret") Or (��ũ(����NPC) = "Chapmion")
    ����NPC = Int((801 * Rnd) + 0)
   Loop
  End If
   ���Ű��� = "Yes"
   Money = val(Money) - 50000
 Else
  MsgBox "���� �����մϴ�. 5������ ��Ƽ� ���ʽÿ�."
 End If
 
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
 Unload FrmShop
 Unload Me
End If
End Sub
