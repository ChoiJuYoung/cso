VERSION 5.00
Begin VB.Form FrmResult 
   BackColor       =   &H00000000&
   Caption         =   "Winner"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
   Icon            =   "Form15.frx":0000
   LinkTopic       =   "Form15"
   ScaleHeight     =   4800
   ScaleWidth      =   9135
   StartUpPosition =   2  'ȭ�� ���
   Begin CSO.jcbutton jcbutton1 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   4440
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   661
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
      BackColor       =   0
      Caption         =   "Go"
      ForeColor       =   16777215
      ForeColorHover  =   65535
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin VB.Label lblStats 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   3480
      Width           =   9135
   End
   Begin VB.Label lblRank 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   3240
      Width           =   9135
   End
   Begin VB.Label lblMoney 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3960
      Width           =   9135
   End
   Begin VB.Label lblLose 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "0"
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
      Left            =   5160
      TabIndex        =   8
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblWin 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblTeam2 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "<Team : KT>"
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
      Left            =   5040
      TabIndex        =   6
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Label lblName2 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Loser : �̻�ȣ"
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
      Left            =   5040
      TabIndex        =   5
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Image Img2 
      Height          =   1455
      Left            =   6360
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "VS"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00000000&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00FFFFFF&
      Height          =   3135
      Left            =   5040
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label lblExp 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Exp +  <0%>"
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
      Left            =   0
      TabIndex        =   2
      Top             =   4200
      Width           =   9135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      Top             =   3840
      Width           =   9135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      Top             =   3120
      Width           =   9135
   End
   Begin VB.Label lblTeam 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "<Team : SKT>"
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
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Label lblName 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Winner : �̿�ȣ"
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
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Image Img 
      Height          =   1500
      Left            =   1320
      Top             =   240
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00FFFFFF&
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "FrmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

If Winer = "��" Then
 If Len(Dir(App.Path & "\img\����\[" & Mid(MyYear(����), 2, 2) & "]" & MyName(����) & ".gif")) <> 0 Then
  Img = LoadPicture(App.Path & "\img\����\[" & Mid(MyYear(����), 2, 2) & "]" & MyName(����) & ".gif")
 Else
  Img = LoadPicture(App.Path & "\img\����\" & MyName(����) & ".gif")
 End If
 If Len(Dir(App.Path & "\img\����\[" & Mid(OYear(Oee), 2, 2) & "]" & �̸�(Oee) & ".gif")) <> 0 Then
  Img2 = LoadPicture(App.Path & "\img\����\[" & Mid(OYear(Oee), 2, 2) & "]" & �̸�(Oee) & ".gif")
 Else
  Img2 = LoadPicture(App.Path & "\img\����\" & �̸�(Oee) & ".gif")
 End If
 lblWin = MW2
 lblLose = OW2
 lblName = "Winer : " & MyYear(����) & MyName(����)
 lblTeam = "<Owner : " & TeamName & ">"
 lblName2 = "Loser : " & OYear(Oee) & �̸�(Oee)
 lblTeam2 = "<Owner : Computer>"
 lblRank = "Rank : " & MyRank(����) & " Vs " & ��ũ(Oee)
 lblStats = "Stats : " & AA & " Vs " & AAO
 lblMoney = "Money + " & val((Int(val(val(RAAO) / 1000) + 1) * 15) / 2) & "Cro, " & Money & "Cro"
 lblExp = "Exp + " & val(RAAO) / 1000 + 1 & " <" & Int(val(MyExp(����)) * 100 / val(MyMExp(����))) & "%>"
ElseIf Winer = "���" Then
 lblName = "Winer : " & OYear(Oee) & �̸�(Oee)
 lblTeam = "<Owner : Computer>"
 lblName2 = "Loser : " & MyYear(����) & MyName(����)
 lblTeam2 = "<Owner : " & TeamName & ">"
 lblWin = OW2
 lblLose = MW2
 lblRank = "Rank : " & ��ũ(Oee) & "Vs" & MyRank(����)
 lblStats = "Stats : " & AAO & " Vs " & AA
 lblMoney = "Money = " & Money & "Cro"
 If Len(Dir(App.Path & "\img\����\[" & Mid(OYear(Oee), 2, 2) & "]" & �̸�(Oee) & ".gif")) <> 0 Then
  Img = LoadPicture(App.Path & "\img\����\[" & Mid(OYear(Oee), 2, 2) & "]" & �̸�(Oee) & ".gif")
 Else
  Img = LoadPicture(App.Path & "\img\����\" & �̸�(Oee) & ".gif")
 End If
 If Len(Dir(App.Path & "\img\����\[" & Mid(MyYear(����), 2, 2) & "]" & MyName(����) & ".gif")) <> 0 Then
  Img2 = LoadPicture(App.Path & "\img\����\[" & Mid(MyYear(����), 2, 2) & "]" & MyName(����) & ".gif")
 Else
  Img2 = LoadPicture(App.Path & "\img\����\" & MyName(����) & ".gif")
 End If
 lblExp = "Exp - " & val(Int(val(RAA) / 1500) + 1) & " <" & Int(val(MyExp(����)) * 100 / val(MyMExp(����))) & "%>"
End If
End Sub

Private Sub jcbutton1_Click()
If Turn = "OSL" Then
    Turn = "PL"
    SetN = 0
    FrmMain.Visible = True
    FrmMain.Timer2.Enabled = True
    FrmMain.Timer3.Enabled = True
    Unload Me
    MW = 0
    OW = 0
    MW2 = 0
    OW2 = 0
ElseIf Turn = "PL" Then
    If PLEnd = "True" Then
        Turn = "OSL"
        SetN = 0
        MW = 0
        OW = 0
        MW2 = 0
        OW2 = 0
        FrmMain.Visible = True
        FrmMain.Timer2.Enabled = True
        FrmMain.Timer3.Enabled = True
        PLEnd = "False"
        Unload Frm_BatInfo
        Unload Me
    ElseIf PLEnd = "False" Then
        ����Set = ����Set + 1
        Frm_BatInfo.CmdGo.Caption = ����Set & "Set : " & MyYear(SetL(����Set)) & MyName(SetL(����Set)) & " vs " & OYear(SetR(����Set)) & �̸�(SetR(����Set)) & "[" & MapName(MapL(����Set)) & "]"
        Frm_BatInfo.Visible = True
        Unload Me
    End If
End If
End Sub
