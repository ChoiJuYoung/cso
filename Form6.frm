VERSION 5.00
Begin VB.Form FrmSearch 
   BackColor       =   &H00000000&
   Caption         =   "�ɷ�ġ ����"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6735
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   5565
   ScaleWidth      =   6735
   StartUpPosition =   2  'ȭ�� ���
   Begin CSO.jcbutton CmdDetail 
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
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
      BackColor       =   8421504
      Caption         =   "���δɷ�ġ"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1920
      Top             =   4200
   End
   Begin VB.Label Label8 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label8"
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
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label7 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      Caption         =   "Label5"
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
      TabIndex        =   7
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      Caption         =   "Label4"
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
      TabIndex        =   6
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      Caption         =   "Label3"
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
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
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
      TabIndex        =   4
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblRaV 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      Caption         =   "A"
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
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label lblTri 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      Caption         =   "(R)"
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
      Left            =   3960
      TabIndex        =   1
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblName 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      Caption         =   "<11>�̿�ȣ[Ex]"
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
      Left            =   2520
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Image ImgPlayer 
      Height          =   1500
      Left            =   2520
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "FrmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdDetail_Click()
FrmSearchH.Show
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
Label1 = "vsA : " & A�¸�(Sear) & " �� " & A�й�(Sear) & " ��"
Label2 = "vsT : " & T�¸�(Sear) & " �� " & T�й�(Sear) & " ��"
Label3 = "vsZ : " & Z�¸�(Sear) & " �� " & Z�й�(Sear) & " ��"
Label4 = "vsP : " & P�¸�(Sear) & " �� " & P�й�(Sear) & " ��"
Label5 = "Team : " & Team(Sear)
Label6 = "��� : " & ���(Sear) & " �ؿ�� : " & �ؿ��(Sear)
Label7 = "Rank : " & ��ũ(Sear)
End Sub

Private Sub Timer1_Timer()
Dim XX, YY As Long


If Len(Dir(App.Path & "\img\����\[" & Mid(OYear(Sear), 2, 2) & "]" & �̸�(Sear) & ".gif")) <> 0 Then
 ImgPlayer.Picture = LoadPicture(App.Path & "\img\����\[" & Mid(OYear(Sear), 2, 2) & "]" & �̸�(Sear) & ".gif")
Else
 ImgPlayer = LoadPicture(App.Path & "\img\����\" & �̸�(Sear) & ".gif")
End If

Dim L1 As String
Dim L2 As String
Dim L3 As String
Dim L4 As String
Dim L5 As String
Dim L6 As String
Dim L7 As String
Dim L8 As String

L1 = val(���ݷ�(Sear))
L2 = val(����(Sear))
L3 = val(����(Sear))
L4 = val(����(Sear))
L5 = val(�����(Sear))
L6 = val(����(Sear))
L7 = val(����(Sear))
L8 = val(��Ʈ��(Sear))

XX = val(FrmSearch.Width) / 2
YY = 3480

Line (XX + 1100, YY)-(XX + 550 * Sqr(2), YY - 550 * Sqr(2)), RGB(255, 255, 255)
Line (XX + 550 * Sqr(2), YY - 550 * Sqr(2))-(XX, YY - 1100), RGB(255, 255, 255)
Line (XX, YY - 1100)-(XX - 550 * Sqr(2), YY - 550 * Sqr(2)), RGB(255, 255, 255)
Line (XX - 550 * Sqr(2), YY - 550 * Sqr(2))-(XX - 1100, YY), RGB(255, 255, 255)
Line (XX - 1100, YY)-(XX - 550 * Sqr(2), YY + 550 * Sqr(2)), RGB(255, 255, 255)
Line (XX - 550 * Sqr(2), YY + 550 * Sqr(2))-(XX, YY + 1100), RGB(255, 255, 255)
Line (XX, YY + 1100)-(XX + 550 * Sqr(2), YY + 550 * Sqr(2)), RGB(255, 255, 255)
Line (XX + 550 * Sqr(2), YY + 550 * Sqr(2))-(XX + 1100, YY), RGB(255, 255, 255)

Line (XX + L1, YY)-(XX + L2 * Sqr(2) / 2, YY + L2 * Sqr(2) / 2), RGB(255, 0, 0)
Line (XX + L2 * Sqr(2) / 2, YY + L2 * Sqr(2) / 2)-(XX, YY + L3), RGB(255, 0, 0)
Line (XX, YY + L3)-(XX - L4 * Sqr(2) / 2, YY + L4 * Sqr(2) / 2), RGB(255, 0, 0)
Line (XX - L4 * Sqr(2) / 2, YY + L4 * Sqr(2) / 2)-(XX - L5, YY), RGB(255, 0, 0)
Line (XX - L5, YY)-(XX - L6 * Sqr(2) / 2, YY - L6 * Sqr(2) / 2), RGB(255, 0, 0)
Line (XX - L6 * Sqr(2) / 2, YY - L6 * Sqr(2) / 2)-(XX, YY - L7), RGB(255, 0, 0)
Line (XX, YY - L7)-(XX + L8 * Sqr(2) / 2, YY - L8 * Sqr(2) / 2), RGB(255, 0, 0)
Line (XX + L8 * Sqr(2) / 2, YY - L8 * Sqr(2) / 2)-(XX + L1, YY), RGB(255, 0, 0)

If ����(Sear) = 1 Then
 lblTri = "T"
ElseIf ����(Sear) = 2 Then
 lblTri = "Z"
ElseIf ����(Sear) = 3 Then
 lblTri = "P"
End If
lblName = OYear(Sear) & �̸�(Sear)


AllPlus = val(���ݷ�(Sear)) + val(����(Sear)) + val(����(Sear)) + val(����(Sear)) + val(�����(Sear)) + val(����(Sear)) + val(����(Sear)) + val(��Ʈ��(Sear))
If val(AllPlus) < 4500 Then
lblRaV.Caption = "Rank : F"
lblRaV.ForeColor = &H4B4B4B
ElseIf val(AllPlus) >= 4500 And val(AllPlus) < 4700 Then
lblRaV.Caption = "Rank : E"
lblRaV.ForeColor = &HB0B0B0
ElseIf val(AllPlus) >= 4700 And val(AllPlus) < 4800 Then
lblRaV.Caption = "Rank : D-"
lblRaV.ForeColor = &HFF3232
ElseIf val(AllPlus) >= 4800 And val(AllPlus) < 4900 Then
lblRaV.Caption = "Rank : D"
lblRaV.ForeColor = &HFF3232
ElseIf val(AllPlus) >= 4900 And val(AllPlus) < 5000 Then
lblRaV.Caption = "Rank : D+"
lblRaV.ForeColor = &HFF3232
ElseIf val(AllPlus) >= 5000 And val(AllPlus) < 5100 Then
lblRaV.Caption = "Rank : C-"
lblRaV.ForeColor = &HFF00&
ElseIf val(AllPlus) >= 5100 And val(AllPlus) < 5200 Then
lblRaV.Caption = "Rank : C"
lblRaV.ForeColor = &HFF00&
ElseIf val(AllPlus) >= 5200 And val(AllPlus) < 5400 Then
lblRaV.Caption = "Rank : C+"
lblRaV.ForeColor = &HFF00&
ElseIf val(AllPlus) >= 5400 And val(AllPlus) < 5600 Then
lblRaV.Caption = "Rank : B-"
lblRaV.ForeColor = &HFFFD&
ElseIf val(AllPlus) >= 5600 And val(AllPlus) < 5800 Then
lblRaV.Caption = "Rank : B"
lblRaV.ForeColor = &HFFFD&
ElseIf val(AllPlus) >= 5800 And val(AllPlus) < 6000 Then
lblRaV.Caption = "Rank : B+"
lblRaV.ForeColor = &HFFFD&
ElseIf val(AllPlus) >= 6000 And val(AllPlus) < 6200 Then
lblRaV.Caption = "Rank : A-"
lblRaV.ForeColor = &H6663FF
ElseIf val(AllPlus) >= 6200 And val(AllPlus) < 6400 Then
lblRaV.Caption = "Rank : A"
lblRaV.ForeColor = &H6663FF
ElseIf val(AllPlus) >= 6400 And val(AllPlus) < 6600 Then
lblRaV.Caption = "Rank : A+"
lblRaV.ForeColor = &H6663FF
ElseIf val(AllPlus) >= 6600 And val(AllPlus) < 6800 Then
lblRaV.Caption = "Rank : S"
lblRaV.ForeColor = &H9600FF
ElseIf val(AllPlus) >= 6800 And val(AllPlus) < 7000 Then
lblRaV.Caption = "Rank : SS"
lblRaV.ForeColor = &H9600FF
ElseIf val(AllPlus) >= 7000 Then
lblRaV.Caption = "Rank : SSS"
lblRaV.ForeColor = &H9600FF
End If
Label8 = "Stats : " & AllPlus
End Sub
