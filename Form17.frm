VERSION 5.00
Begin VB.Form FrmSingPick 
   BackColor       =   &H00000000&
   Caption         =   "��������"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8235
   Icon            =   "Form17.frx":0000
   LinkTopic       =   "Form17"
   ScaleHeight     =   4410
   ScaleWidth      =   8235
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.ComboBox CmbML 
      Height          =   300
      ItemData        =   "Form17.frx":628A
      Left            =   5040
      List            =   "Form17.frx":629D
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   25
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7080
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  '���
      BackColor       =   &H00FFFFFF&
      Height          =   1500
      Left            =   3480
      Style           =   1  '�׷���
      TabIndex        =   11
      Top             =   2520
      Width           =   1500
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  '���
      BackColor       =   &H00FFFFFF&
      Height          =   1500
      Left            =   1800
      Style           =   1  '�׷���
      TabIndex        =   10
      Top             =   2520
      Width           =   1500
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  '���
      BackColor       =   &H00FFFFFF&
      Height          =   1500
      Left            =   120
      Style           =   1  '�׷���
      TabIndex        =   9
      Top             =   2520
      Width           =   1500
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  '���
      BackColor       =   &H00FFFFFF&
      Height          =   1500
      Left            =   3480
      Style           =   1  '�׷���
      TabIndex        =   8
      Top             =   600
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  '���
      BackColor       =   &H00FFFFFF&
      Height          =   1500
      Left            =   1800
      Style           =   1  '�׷���
      TabIndex        =   7
      Top             =   600
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  '���
      BackColor       =   &H00FFFFFF&
      Height          =   1500
      Left            =   120
      Style           =   1  '�׷���
      TabIndex        =   6
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '����
      Caption         =   "0 : 0"
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
      Left            =   4080
      TabIndex        =   24
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
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
      Left            =   5880
      TabIndex        =   23
      Top             =   3720
      Width           =   1575
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
      Left            =   5880
      TabIndex        =   22
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      Caption         =   "T vs Z = 60 : 40"
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
      Left            =   5880
      TabIndex        =   21
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�� �¾��� ����"
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
      Left            =   5880
      TabIndex        =   20
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Image ImgMa 
      Height          =   1500
      Left            =   5880
      Picture         =   "Form17.frx":62BF
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '����
      Caption         =   "Vs �Ｚ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "<11>"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   17
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "<11>"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   16
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "<11>"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "<11>"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   14
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "<11>"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   13
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "<11>"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   5
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   4
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   3
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "�̿�ȣ[Ex]"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "FrmSingPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
���� = 1
MN = MyName(����)
Year = MyYear(����)
AT = MyAt(����)
R = MyR(����)
St = MySt(����)
Am = MyAm(����)
De = MyDe(����)
Pa = MyPa(����)
SE = MySe(����)
Co = MyCo(����)
MT = MyTribe(����)

Text1_Change
End Sub

Private Sub Command2_Click()
���� = 2
MN = MyName(����)
Year = MyYear(����)
AT = MyAt(����)
R = MyR(����)
St = MySt(����)
Am = MyAm(����)
De = MyDe(����)
Pa = MyPa(����)
SE = MySe(����)
Co = MyCo(����)
MT = MyTribe(����)

Text1_Change
End Sub

Private Sub Command3_Click()
���� = 3
MN = MyName(����)
Year = MyYear(����)
AT = MyAt(����)
R = MyR(����)
St = MySt(����)
Am = MyAm(����)
De = MyDe(����)
Pa = MyPa(����)
SE = MySe(����)
Co = MyCo(����)
MT = MyTribe(����)

Text1_Change
End Sub

Private Sub Command4_Click()
���� = 4
MN = MyName(����)
Year = MyYear(����)
AT = MyAt(����)
R = MyR(����)
St = MySt(����)
Am = MyAm(����)
De = MyDe(����)
Pa = MyPa(����)
SE = MySe(����)
MT = MyTribe(����)
Co = MyCo(����)

Text1_Change
End Sub

Private Sub Command5_Click()
���� = 5
MN = MyName(����)
Year = MyYear(����)
AT = MyAt(����)
R = MyR(����)
St = MySt(����)
Am = MyAm(����)
De = MyDe(����)
Pa = MyPa(����)
SE = MySe(����)
MT = MyTribe(����)
Co = MyCo(����)

Text1_Change
End Sub

Private Sub Command6_Click()
���� = 6
MN = MyName(����)
Year = MyYear(����)
AT = MyAt(����)
R = MyR(����)
St = MySt(����)
Am = MyAm(����)
De = MyDe(����)
Pa = MyPa(����)
SE = MySe(����)
MT = MyTribe(����)
Co = MyCo(����)

Text1_Change
End Sub

Private Sub Form_Load()
If Turn = "PL" Then
 Label9 = MW & " : " & OW
 If val(MW) > val(OW) Then
  Label9.ForeColor = RGB(0, 255, 0)
 ElseIf val(MW) < val(OW) Then
  Label9.ForeColor = RGB(255, 0, 0)
 ElseIf val(MW) = val(OW) Then
  Label9.ForeColor = RGB(255, 255, 0)
 End If
Else
 Label9.Visible = False
End If

Label4 = MapName(Map)
If Len(Dir(App.Path & "\img\��\" & MapName(Map) & ".gif")) <> 0 Then
 ImgMa = LoadPicture(App.Path & "\img\��\" & MapName(Map) & ".gif")
Else
 ImgMa = Nothing
End If

Label5 = "T vs Z = " & TZT(Map) & " : " & TZZ(Map)
Label6 = "Z vs P = " & ZPZ(Map) & " : " & ZPP(Map)
Label7 = "P vs T = " & PTP(Map) & " : " & PTT(Map)

If Len(Dir(App.Path & "\img\����\[" & Mid(MyYear(1), 2, 2) & "]" & MyName(1) & ".gif")) <> 0 Then
 Command1.Picture = LoadPicture(App.Path & "\img\����\[" & Mid(MyYear(1), 2, 2) & "]" & MyName(1) & ".gif")
Else
 Command1.Picture = LoadPicture(App.Path & "\img\����\" & MyName(1) & ".gif")
End If

If Len(Dir(App.Path & "\img\����\[" & Mid(MyYear(2), 2, 2) & "]" & MyName(2) & ".gif")) <> 0 Then
 Command2.Picture = LoadPicture(App.Path & "\img\����\[" & Mid(MyYear(2), 2, 2) & "]" & MyName(2) & ".gif")
Else
 Command2.Picture = LoadPicture(App.Path & "\img\����\" & MyName(2) & ".gif")
End If

If Len(Dir(App.Path & "\img\����\[" & Mid(MyYear(3), 2, 2) & "]" & MyName(3) & ".gif")) <> 0 Then
 Command3.Picture = LoadPicture(App.Path & "\img\����\[" & Mid(MyYear(3), 2, 2) & "]" & MyName(3) & ".gif")
Else
 Command3.Picture = LoadPicture(App.Path & "\img\����\" & MyName(3) & ".gif")
End If

If Len(Dir(App.Path & "\img\����\[" & Mid(MyYear(4), 2, 2) & "]" & MyName(4) & ".gif")) <> 0 Then
 Command4.Picture = LoadPicture(App.Path & "\img\����\[" & Mid(MyYear(4), 2, 2) & "]" & MyName(4) & ".gif")
Else
 Command4.Picture = LoadPicture(App.Path & "\img\����\" & MyName(4) & ".gif")
End If

If Len(Dir(App.Path & "\img\����\[" & Mid(MyYear(5), 2, 2) & "]" & MyName(5) & ".gif")) <> 0 Then
 Command5.Picture = LoadPicture(App.Path & "\img\����\[" & Mid(MyYear(5), 2, 2) & "]" & MyName(5) & ".gif")
Else
 Command5.Picture = LoadPicture(App.Path & "\img\����\" & MyName(5) & ".gif")
End If

If Len(Dir(App.Path & "\img\����\[" & Mid(MyYear(6), 2, 2) & "]" & MyName(6) & ".gif")) <> 0 Then
 Command6.Picture = LoadPicture(App.Path & "\img\����\[" & Mid(MyYear(6), 2, 2) & "]" & MyName(6) & ".gif")
Else
 Command6.Picture = LoadPicture(App.Path & "\img\����\" & MyName(6) & ".gif")
End If

If PL������(1) = False Then
 Command1.Enabled = False
End If

If PL������(2) = False Then
 Command2.Enabled = False
End If

If PL������(3) = False Then
 Command3.Enabled = False
End If

If PL������(4) = False Then
 Command4.Enabled = False
End If

If PL������(5) = False Then
 Command5.Enabled = False
End If

If PL������(6) = False Then
 Command6.Enabled = False
End If

Label2(0) = MyYear(1)
Label2(1) = MyYear(2)
Label2(2) = MyYear(3)
Label2(3) = MyYear(4)
Label2(4) = MyYear(5)
Label2(5) = MyYear(6)
Label1(0) = MyName(1)
Label1(1) = MyName(2)
Label1(2) = MyName(3)
Label1(3) = MyName(4)
Label1(4) = MyName(5)
Label1(5) = MyName(6)


If MyRank(1) = "Normal" Then
 Label2(0).ForeColor = RGB(255, 255, 255)
ElseIf MyRank(1) = "Special" Then
 Label2(0).ForeColor = RGB(0, 255, 0)
ElseIf MyRank(1) = "Rare" Then
 Label2(0).ForeColor = &HFF80FF
ElseIf MyRank(1) = "Unique" Then
 Label2(0).ForeColor = &HFF8080
ElseIf MyRank(1) = "Elite" Then
 Label2(0).ForeColor = &H800080
ElseIf MyRank(1) = "Legend" Then
 Label2(0).ForeColor = &H80FF&
ElseIf MyRank(1) = "Secret" Then
 Label2(0).ForeColor = &HFFC0C0
ElseIf MyRank(1) = "Champion" Then
 Label2(0).ForeColor = RGB(255, 0, 0)
End If

If MyRank(2) = "Normal" Then
 Label2(1).ForeColor = RGB(255, 255, 255)
ElseIf MyRank(2) = "Special" Then
 Label2(1).ForeColor = RGB(0, 255, 0)
ElseIf MyRank(2) = "Rare" Then
 Label2(1).ForeColor = &HFF80FF
ElseIf MyRank(2) = "Unique" Then
 Label2(1).ForeColor = &HFF8080
ElseIf MyRank(2) = "Elite" Then
 Label2(1).ForeColor = &H800080
ElseIf MyRank(2) = "Legend" Then
 Label2(1).ForeColor = &H80FF&
ElseIf MyRank(2) = "Secret" Then
 Label2(1).ForeColor = &HFFC0C0
ElseIf MyRank(2) = "Champion" Then
 Label2(1).ForeColor = RGB(255, 0, 0)
End If

If MyRank(3) = "Normal" Then
 Label2(2).ForeColor = RGB(255, 255, 255)
ElseIf MyRank(3) = "Special" Then
 Label2(2).ForeColor = RGB(0, 255, 0)
ElseIf MyRank(3) = "Rare" Then
 Label2(2).ForeColor = &HFF80FF
ElseIf MyRank(3) = "Unique" Then
 Label2(2).ForeColor = &HFF8080
ElseIf MyRank(3) = "Elite" Then
 Label2(2).ForeColor = &H800080
ElseIf MyRank(3) = "Legend" Then
 Label2(2).ForeColor = &H80FF&
ElseIf MyRank(3) = "Secret" Then
 Label2(2).ForeColor = &HFFC0C0
ElseIf MyRank(3) = "Champion" Then
 Label2(2).ForeColor = RGB(255, 0, 0)
End If

If MyRank(4) = "Normal" Then
 Label2(3).ForeColor = RGB(255, 255, 255)
ElseIf MyRank(4) = "Special" Then
 Label2(3).ForeColor = RGB(0, 255, 0)
ElseIf MyRank(4) = "Rare" Then
 Label2(3).ForeColor = &HFF80FF
ElseIf MyRank(4) = "Unique" Then
 Label2(3).ForeColor = &HFF8080
ElseIf MyRank(4) = "Elite" Then
 Label2(3).ForeColor = &H800080
ElseIf MyRank(4) = "Legend" Then
 Label2(3).ForeColor = &H80FF&
ElseIf MyRank(4) = "Secret" Then
 Label2(3).ForeColor = &HFFC0C0
ElseIf MyRank(4) = "Champion" Then
 Label2(3).ForeColor = RGB(255, 0, 0)
End If

If MyRank(5) = "Normal" Then
 Label2(4).ForeColor = RGB(255, 255, 255)
ElseIf MyRank(5) = "Special" Then
 Label2(4).ForeColor = RGB(0, 255, 0)
ElseIf MyRank(5) = "Rare" Then
 Label2(4).ForeColor = &HFF80FF
ElseIf MyRank(5) = "Unique" Then
 Label2(4).ForeColor = &HFF8080
ElseIf MyRank(5) = "Elite" Then
 Label2(4).ForeColor = &H800080
ElseIf MyRank(5) = "Legend" Then
 Label2(4).ForeColor = &H80FF&
ElseIf MyRank(5) = "Secret" Then
 Label2(4).ForeColor = &HFFC0C0
ElseIf MyRank(5) = "Champion" Then
 Label2(4).ForeColor = RGB(255, 0, 0)
End If

If MyRank(6) = "Normal" Then
 Label2(5).ForeColor = RGB(255, 255, 255)
ElseIf MyRank(6) = "Special" Then
 Label2(5).ForeColor = RGB(0, 266, 0)
ElseIf MyRank(6) = "Rare" Then
 Label2(5).ForeColor = &HFF80FF
ElseIf MyRank(6) = "Unique" Then
 Label2(5).ForeColor = &HFF8080
ElseIf MyRank(6) = "Elite" Then
 Label2(5).ForeColor = &H800080
ElseIf MyRank(6) = "Legend" Then
 Label2(5).ForeColor = &H80FF&
ElseIf MyRank(6) = "Secret" Then
 Label2(5).ForeColor = &HFFC0C0
ElseIf MyRank(6) = "Champion" Then
 Label2(5).ForeColor = RGB(255, 0, 0)
End If

Oee = Int((801 * Rnd) + 0)


If Turn = "OSL" Then
  Label3 = "StarLeague"
Else
 MsgBox "���� �� ����"
End If
End Sub

Private Sub Text1_Change()

If Turn = "OSL" Then
 If MyNW(����) = "CB16" Then
 SetA = 1
  Do Until ��ũ(Oee) = "Normal" Or ��ũ(Oee) = "Special"
   Oee = Int((801 * Rnd) + 0)
  Loop
 ElseIf MyNW(����) = "CB8" Then
 SetA = 1
  Do Until ��ũ(Oee) = "Normal" Or ��ũ(Oee) = "Special"
   Oee = Int((801 * Rnd) + 0)
  Loop
 ElseIf MyNW(����) = "CB4" Then
 SetA = 1
  Do Until ��ũ(Oee) = "Normal" Or ��ũ(Oee) = "Special"
   Oee = Int((801 * Rnd) + 0)
  Loop
 ElseIf MyNW(����) = "CBFin" Then
 SetA = 1
  Do Until ��ũ(Oee) = "Normal" Or ��ũ(Oee) = "Special"
   Oee = Int((801 * Rnd) + 0)
  Loop
 ElseIf MyNW(����) = "CA1" Then
 SetA = 1
  Do Until ��ũ(Oee) = "Rare" Or ��ũ(Oee) = "Unique"
   Oee = Int((801 * Rnd) + 0)
  Loop
 ElseIf MyNW(����) = "CA2" Then
 SetA = 1
  Do Until ��ũ(Oee) = "Unique" Or ��ũ(Oee) = "Elite"
   Oee = Int((801 * Rnd) + 0)
  Loop
 ElseIf MyNW(����) = "CA3" Then
 SetA = 1
  Do Until ��ũ(Oee) = "Elite"
   Oee = Int((801 * Rnd) + 0)
  Loop
 ElseIf MyNW(����) = "CS32" Then
 SetA = 2
  Do Until ��ũ(Oee) = "Elite" Or ��ũ(Oee) = "Legend" Or ��ũ(Oee) = "Secret" Or ��ũ(Oee) = "Champion"
   Oee = Int((801 * Rnd) + 0)
  Loop
 ElseIf MyNW(����) = "CS16" Then
 SetA = 2
  Do Until ��ũ(Oee) = "Legend" Or ��ũ(Oee) = "Secret" Or ��ũ(Oee) = "Champion"
   Oee = Int((801 * Rnd) + 0)
  Loop
 ElseIf MyNW(����) = "CS8" Then
 SetA = 3
  Do Until ��ũ(Oee) = "Legend" Or ��ũ(Oee) = "Secret" Or ��ũ(Oee) = "Champion"
   Oee = Int((801 * Rnd) + 0)
  Loop
 ElseIf MyNW(����) = "CS4" Then
 SetA = 3
  Do Until ��ũ(Oee) = "Secret" Or ��ũ(Oee) = "Champion"
   Oee = Int((801 * Rnd) + 0)
  Loop
 ElseIf MyNW(����) = "CSFin" Then
 SetA = 4
  Do Until ��ũ(Oee) = "Secret" Or ��ũ(Oee) = "Champion"
   Oee = Int((801 * Rnd) + 0)
  Loop
 ElseIf MyNW(����) = "UpADo" Then
 SetA = 3
  Do Until ��ũ(Oee) = "Rare" Or ��ũ(Oee) = "Unique" Or ��ũ(Oee) = "Elite"
   Oee = Int((801 * Rnd) + 0)
  Loop
 End If
Else
 MsgBox "���� �� ����"
End If

FrmLoading.Show
Unload Me
End Sub
