VERSION 5.00
Begin VB.Form FrmShopConf 
   Caption         =   "Confirm"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   Icon            =   "Form20.frx":0000
   LinkTopic       =   "Form20"
   ScaleHeight     =   6615
   ScaleWidth      =   6240
   StartUpPosition =   2  'ȭ�� ���
   Begin CSO.jcbutton jcbutton2 
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   6240
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   255
      Caption         =   "No"
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin CSO.jcbutton jcbutton1 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   6240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16711680
      Caption         =   "Yes"
      ForeColor       =   16777215
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin VB.Label Label8 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "<����� ������ Ȯ���մϱ�?>"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   5880
      Width           =   6255
   End
   Begin VB.Label lblAllPrice 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   5400
      Width           =   6255
   End
   Begin VB.Label Label6 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�� ������ �ݾ�,"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   4920
      Width           =   6255
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  '�������� ����
      Height          =   2175
      Left            =   0
      Top             =   4440
      Width           =   6255
   End
   Begin VB.Label lbl���� 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3720
      Width           =   6255
   End
   Begin VB.Label Label4 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   6255
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   6255
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   6255
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '�������� ����
      Height          =   2895
      Left            =   0
      Top             =   1560
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "������ ī������ ������ ������ Ȯ���մϴ�."
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '�������� ����
      Height          =   1575
      Left            =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "FrmShopConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
If val(����) = 1 Then
 Label2 = "<Zerg Card Pack>"
 Label3 = "[Normal]"
 Label4 = "[3000Cro]"
 lblAllPrice = "[3000Cro]"
ElseIf val(����) = 2 Then
 Label2 = "<Zerg Card Pack ver.A>"
 Label3 = "[Normal ~ Rare]"
 Label4 = "[7000Cro]"
 lblAllPrice = "[7000Cro]"
ElseIf val(����) = 3 Then
 Label2 = "<Zerg Card Pack ver.S"
 Label3 = "[Normal ~ Elite]"
 Label4 = "[10000Cro]"
 lblAllPrice = "[10000Cro]"
ElseIf val(����) = 4 Then
 Label2 = "<Terran Card Pack>"
 Label3 = "[Normal]"
 Label4 = "[3000Cro]"
 lblAllPrice = "[3000Cro]"
ElseIf val(����) = 5 Then
 Label2 = "<Terran Card Pack ver.A>"
 Label3 = "[Normal ~ Rare]"
 Label4 = "[7000Cro]"
 lblAllPrice = "[7000Cro]"
ElseIf val(����) = 6 Then
 Label2 = "<Terran Card Pack ver.S>"
 Label3 = "[Normal ~ Elite]"
 Label4 = "[10000Cro]"
 lblAllPrice = "[10000Cro]"
ElseIf val(����) = 7 Then
 Label2 = "<Protoss Card Pack>"
 Label3 = "[Normal]"
 Label4 = "[3000Cro]"
 lblAllPrice = "[3000Cro]"
ElseIf val(����) = 8 Then
 Label2 = "<Protoss Card Pack ver.A>"
 Label3 = "[Normal ~ Rare]"
 Label4 = "[7000Cro]"
 lblAllPrice = "[7000Cro]"
ElseIf val(����) = 9 Then
 Label2 = "<Protoss Card Pack ver.S>"
 Label3 = "[Normal ~ Elite]"
 Label4 = "[10000Cro]"
 lblAllPrice = "[10000Cro]"
End If
lbl���� = "<���� : 1>"
End Sub

Private Sub jcbutton1_Click()
����NPC = Int((714 * Rnd) + 1)
If val(����) = 1 Then
 If val(Money) >= 3000 Then
  Do Until (��ũ(����NPC) = "Normal") And ����(����NPC) = 2
    ����NPC = Int((800 * Rnd) + 1)
  Loop
   ���Ű��� = "Yes"
   Money = val(Money) - 3000
 Else
  MsgBox "���� �����մϴ� �Ф�"
 End If
ElseIf val(����) = 2 Then
 If val(Money) >= 7000 Then
  Do Until (��ũ(����NPC) = "Normal" Or ��ũ(����NPC) = "Special" Or ��ũ(����NPC) = "Rare") And ����(����NPC) = 2
    ����NPC = Int((800 * Rnd) + 1)
  Loop
   ���Ű��� = "Yes"
   Money = val(Money) - 7000
 Else
  MsgBox "���� �����մϴ� �Ф�"
 End If
ElseIf val(����) = 3 Then
 If val(Money) >= 10000 Then
  If Mode <> "Normal" Then
   Do Until (��ũ(����NPC) = "Normal" Or ��ũ(����NPC) = "Special" Or ��ũ(����NPC) = "Rare" Or ��ũ(����NPC) = "Unique" Or ��ũ(����NPC) = "Elite") And (����(����NPC) = 2)
    ����NPC = Int((723 * Rnd) + 1)
   Loop
  Else
   Do Until (��ũ(����NPC) = "Normal" Or ��ũ(����NPC) = "Special" Or ��ũ(����NPC) = "Rare" Or ��ũ(����NPC) = "Unique") And (����(����NPC) = 2)
    ����NPC = Int((723 * Rnd) + 1)
   Loop
  End If
   ���Ű��� = "Yes"
   Money = val(Money) - 10000
 Else
  MsgBox "���� �����մϴ� �Ф�"
 End If
ElseIf val(����) = 4 Then
 If val(Money) >= 3000 Then
  Do Until (��ũ(����NPC) = "Normal") And ����(����NPC) = 1
    ����NPC = Int((800 * Rnd) + 1)
  Loop
   ���Ű��� = "Yes"
   Money = val(Money) - 3000
 Else
  MsgBox "���� �����մϴ� �Ф�"
 End If
ElseIf val(����) = 5 Then
 If val(Money) >= 7000 Then
  Do Until (��ũ(����NPC) = "Normal" Or ��ũ(����NPC) = "Special" Or ��ũ(����NPC) = "Rare") And ����(����NPC) = 1
    ����NPC = Int((800 * Rnd) + 1)
  Loop
   ���Ű��� = "Yes"
   Money = val(Money) - 7000
 Else
  MsgBox "���� �����մϴ� �Ф�"
 End If
ElseIf val(����) = 6 Then
 If val(Money) >= 10000 Then
  If Mode = "Normal" Then
   Do Until (��ũ(����NPC) = "Normal" Or ��ũ(����NPC) = "Special" Or ��ũ(����NPC) = "Rare" Or ��ũ(����NPC) = "Unique") And (����(����NPC) = 1)
    ����NPC = Int((723 * Rnd) + 1)
   Loop
  Else
   Do Until (��ũ(����NPC) = "Normal" Or ��ũ(����NPC) = "Special" Or ��ũ(����NPC) = "Rare" Or ��ũ(����NPC) = "Unique" Or ��ũ(����NPC) = "Elite") And (����(����NPC) = 1)
    ����NPC = Int((723 * Rnd) + 1)
   Loop
  End If
   ���Ű��� = "Yes"
   Money = val(Money) - 10000
 Else
  MsgBox "���� �����մϴ� �Ф�"
 End If
ElseIf val(����) = 7 Then
 If val(Money) >= 3000 Then
  Do Until (��ũ(����NPC) = "Normal") And ����(����NPC) = 3
    ����NPC = Int((800 * Rnd) + 1)
  Loop
   ���Ű��� = "Yes"
   Money = val(Money) - 3000
 Else
  MsgBox "���� �����մϴ� �Ф�"
 End If
ElseIf val(����) = 8 Then
 If val(Money) >= 7000 Then
  Do Until (��ũ(����NPC) = "Normal" Or ��ũ(����NPC) = "Special" Or ��ũ(����NPC) = "Rare") And ����(����NPC) = 3
    ����NPC = Int((800 * Rnd) + 1)
  Loop
   ���Ű��� = "Yes"
   Money = val(Money) - 7000
 Else
  MsgBox "���� �����մϴ� �Ф�"
 End If
ElseIf val(����) = 9 Then
 If val(Money) >= 10000 Then
  If Mode = "Normal" Then
   Do Until (��ũ(����NPC) = "Normal" Or ��ũ(����NPC) = "Special" Or ��ũ(����NPC) = "Rare" Or ��ũ(����NPC) = "Unique") And (����(����NPC) = 3)
    ����NPC = Int((723 * Rnd) + 1)
   Loop
  Else
   Do Until (��ũ(����NPC) = "Normal" Or ��ũ(����NPC) = "Special" Or ��ũ(����NPC) = "Rare" Or ��ũ(����NPC) = "Unique" Or ��ũ(����NPC) = "Elite") And (����(����NPC) = 3)
    ����NPC = Int((723 * Rnd) + 1)
   Loop
  End If
   ���Ű��� = "Yes"
   Money = val(Money) - 10000
 Else
  MsgBox "���� �����մϴ� �Ф�"
 End If
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


 Call Save
 FrmMain.Timer12.Enabled = True
 Unload FrmShop
 Unload Me
End If
End Sub

Private Sub jcbutton2_Click()
Unload FrmShop
Unload Me
End Sub

