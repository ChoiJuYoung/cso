VERSION 5.00
Begin VB.Form FrmSum 
   BackColor       =   &H00000000&
   Caption         =   "ī�� �ռ�"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   Icon            =   "Form23.frx":0000
   LinkTopic       =   "Form23"
   ScaleHeight     =   4185
   ScaleWidth      =   10830
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton Command18 
      Caption         =   "Command18"
      Height          =   255
      Left            =   9240
      TabIndex        =   18
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Command17"
      Height          =   255
      Left            =   9240
      TabIndex        =   17
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Command16"
      Height          =   255
      Left            =   9240
      TabIndex        =   16
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Command15"
      Height          =   255
      Left            =   9240
      TabIndex        =   15
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Command14"
      Height          =   255
      Left            =   7920
      TabIndex        =   14
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Command13"
      Height          =   255
      Left            =   7920
      TabIndex        =   13
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   255
      Left            =   7920
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Command11"
      Height          =   255
      Left            =   7920
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   255
      Left            =   7920
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   $"Form23.frx":628A
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2040
      Width           =   10215
   End
   Begin VB.Image Image3 
      Height          =   1500
      Left            =   4560
      Picture         =   "Form23.frx":6321
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "       <Card1>                                                 <Card2>"
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
      Left            =   2760
      TabIndex        =   0
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Image Image2 
      Height          =   1500
      Left            =   6360
      Top             =   240
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   2760
      Top             =   240
      Width           =   1500
   End
End
Attribute VB_Name = "FrmSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo ABC:
�ռ�1 = 1
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�1) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command10_Click()
On Error GoTo ABC:
�ռ�2 = 1
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command11_Click()
On Error GoTo ABC:
�ռ�2 = 2
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command12_Click()
On Error GoTo ABC:
�ռ�2 = 3
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command13_Click()
On Error GoTo ABC:
�ռ�2 = 4
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command14_Click()
On Error GoTo ABC:
�ռ�2 = 5
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command15_Click()
On Error GoTo ABC:
�ռ�2 = 6
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command16_Click()
On Error GoTo ABC:
�ռ�2 = 7
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command17_Click()
On Error GoTo ABC:
�ռ�2 = 8
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command18_Click()
On Error GoTo ABC:
�ռ�2 = 9
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�2), 2, 2) & "]" & SubName(�ռ�2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command2_Click()
On Error GoTo ABC:
�ռ�1 = 2
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�1) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command3_Click()
On Error GoTo ABC:
�ռ�1 = 3
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�1) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command4_Click()
On Error GoTo ABC:
�ռ�1 = 4
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�1) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command5_Click()
On Error GoTo ABC:
�ռ�1 = 5
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�1) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command6_Click()
On Error GoTo ABC:
�ռ�1 = 6
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�1) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command7_Click()
On Error GoTo ABC:
�ռ�1 = 7
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�1) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command8_Click()
On Error GoTo ABC:
�ռ�1 = 8
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�1) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command9_Click()
On Error GoTo ABC:
�ռ�1 = 9
If Len(Dir(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\����\[" & Mid(SubYear(�ռ�1), 2, 2) & "]" & SubName(�ռ�1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\����\" & SubName(�ռ�1) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Form_Load()
Command1.Caption = SubYear(1) & SubName(1)
Command10.Caption = SubYear(1) & SubName(1)
Command2.Caption = SubYear(2) & SubName(2)
Command11.Caption = SubYear(2) & SubName(2)
Command3.Caption = SubYear(3) & SubName(3)
Command12.Caption = SubYear(3) & SubName(3)
Command4.Caption = SubYear(4) & SubName(4)
Command13.Caption = SubYear(4) & SubName(4)
Command5.Caption = SubYear(5) & SubName(5)
Command14.Caption = SubYear(5) & SubName(5)
Command6.Caption = SubYear(6) & SubName(6)
Command15.Caption = SubYear(6) & SubName(6)
Command7.Caption = SubYear(7) & SubName(7)
Command16.Caption = SubYear(7) & SubName(7)
Command8.Caption = SubYear(8) & SubName(8)
Command17.Caption = SubYear(8) & SubName(8)
Command9.Caption = SubYear(9) & SubName(9)
Command18.Caption = SubYear(9) & SubName(9)

End Sub

Private Sub Image3_Click()
If (SubName(�ռ�1) = SubName(�ռ�2)) And (SubYear(�ռ�1) = SubYear(�ռ�2)) And (�ռ�1 <> �ռ�2) And (SubRank(�ռ�1) = SubRank(�ռ�2)) Then
 
If SubRank(�ռ�1) = "Normal" Then
 SubRank(�ռ�1) = "Special"
 SubAt(�ռ�1) = val(SubAt(�ռ�1)) + 3
 SubR(�ռ�1) = val(SubR(�ռ�1)) + 3
 SubSt(�ռ�1) = val(SubSt(�ռ�1)) + 3
 SubAm(�ռ�1) = val(SubAm(�ռ�1)) + 3
 SubDe(�ռ�1) = val(SubDe(�ռ�1)) + 3
 SubPa(�ռ�1) = val(SubPa(�ռ�1)) + 3
 SubSe(�ռ�1) = val(SubSe(�ռ�1)) + 3
 SubCo(�ռ�1) = val(SubCo(�ռ�1)) + 3
 
ElseIf SubRank(�ռ�1) = "Special" Then
 SubRank(�ռ�1) = "Rare"
 SubAt(�ռ�1) = val(SubAt(�ռ�1)) + 4
 SubR(�ռ�1) = val(SubR(�ռ�1)) + 4
 SubSt(�ռ�1) = val(SubSt(�ռ�1)) + 4
 SubAm(�ռ�1) = val(SubAm(�ռ�1)) + 4
 SubDe(�ռ�1) = val(SubDe(�ռ�1)) + 4
 SubPa(�ռ�1) = val(SubPa(�ռ�1)) + 4
 SubSe(�ռ�1) = val(SubSe(�ռ�1)) + 4
 SubCo(�ռ�1) = val(SubCo(�ռ�1)) + 4
 
ElseIf SubRank(�ռ�1) = "Rare" Then
 SubRank(�ռ�1) = "Unique"
 SubAt(�ռ�1) = val(SubAt(�ռ�1)) + 5
 SubR(�ռ�1) = val(SubR(�ռ�1)) + 5
 SubSt(�ռ�1) = val(SubSt(�ռ�1)) + 5
 SubAm(�ռ�1) = val(SubAm(�ռ�1)) + 5
 SubDe(�ռ�1) = val(SubDe(�ռ�1)) + 5
 SubPa(�ռ�1) = val(SubPa(�ռ�1)) + 5
 SubSe(�ռ�1) = val(SubSe(�ռ�1)) + 5
 SubCo(�ռ�1) = val(SubCo(�ռ�1)) + 5
 
ElseIf SubRank(�ռ�1) = "Unique" Then
 SubRank(�ռ�1) = "Elite"
 SubAt(�ռ�1) = val(SubAt(�ռ�1)) + 6
 SubR(�ռ�1) = val(SubR(�ռ�1)) + 6
 SubSt(�ռ�1) = val(SubSt(�ռ�1)) + 6
 SubAm(�ռ�1) = val(SubAm(�ռ�1)) + 6
 SubDe(�ռ�1) = val(SubDe(�ռ�1)) + 6
 SubPa(�ռ�1) = val(SubPa(�ռ�1)) + 6
 SubSe(�ռ�1) = val(SubSe(�ռ�1)) + 6
 SubCo(�ռ�1) = val(SubCo(�ռ�1)) + 6
 
ElseIf SubRank(�ռ�1) = "Elite" Then
 SubRank(�ռ�1) = "Legend"
 SubAt(�ռ�1) = val(SubAt(�ռ�1)) + 7
 SubR(�ռ�1) = val(SubR(�ռ�1)) + 7
 SubSt(�ռ�1) = val(SubSt(�ռ�1)) + 7
 SubAm(�ռ�1) = val(SubAm(�ռ�1)) + 7
 SubDe(�ռ�1) = val(SubDe(�ռ�1)) + 7
 SubPa(�ռ�1) = val(SubPa(�ռ�1)) + 7
 SubSe(�ռ�1) = val(SubSe(�ռ�1)) + 7
 SubCo(�ռ�1) = val(SubCo(�ռ�1)) + 7

ElseIf SubRank(�ռ�1) = "Legend" Then
 SubAt(�ռ�1) = val(SubAt(�ռ�1)) + 15
 SubR(�ռ�1) = val(SubR(�ռ�1)) + 15
 SubSt(�ռ�1) = val(SubSt(�ռ�1)) + 15
 SubAm(�ռ�1) = val(SubAm(�ռ�1)) + 15
 SubDe(�ռ�1) = val(SubDe(�ռ�1)) + 15
 SubPa(�ռ�1) = val(SubPa(�ռ�1)) + 15
 SubSe(�ռ�1) = val(SubSe(�ռ�1)) + 15
 SubCo(�ռ�1) = val(SubCo(�ռ�1)) + 15
 SubRank(�ռ�1) = "Secret"
If SubName(�ռ�1) = "�ӿ�ȯ" Then
    SubSkill(�ռ�1) = 1
ElseIf SubName(�ռ�1) = "�̿�ȣ" Then
    SubSkill(�ռ�1) = 2
ElseIf SubName(�ռ�1) = "ȫ��ȣ" Then
    SubSkill(�ռ�1) = 3
ElseIf SubName(�ռ�1) = "������" Then
    SubSkill(�ռ�1) = 4
ElseIf SubName(�ռ�1) = "������" Then
    SubSkill(�ռ�1) = 5
ElseIf SubName(�ռ�1) = "������" Then
    SubSkill(�ռ�1) = 6
ElseIf SubName(�ռ�1) = "�ֿ���" Then
    SubSkill(�ռ�1) = 7
ElseIf SubName(�ռ�1) = "���ÿ�" Then
    SubSkill(�ռ�1) = 8
ElseIf SubName(�ռ�1) = "�ۺ���" Then
    SubSkill(�ռ�1) = 9
ElseIf SubName(�ռ�1) = "�豸��" Then
    SubSkill(�ռ�1) = 10
ElseIf SubName(�ռ�1) = "�㿵��" Then
    SubSkill(�ռ�1) = 11
ElseIf SubName(�ռ�1) = "�����" Then
    SubSkill(�ռ�1) = 12
ElseIf SubName(�ռ�1) = "������" Then
    SubSkill(�ռ�1) = 13
ElseIf SubName(�ռ�1) = "������" Then
    SubSkill(�ռ�1) = 14
ElseIf SubName(�ռ�1) = "�ڼ���" Then
    SubSkill(�ռ�1) = 15
ElseIf SubName(�ռ�1) = "������" Then
    SubSkill(�ռ�1) = 16
ElseIf SubName(�ռ�1) = "���ؿ�" Then
    SubSkill(�ռ�1) = 17
ElseIf SubName(�ռ�1) = "������" Then
    SubSkill(�ռ�1) = 18
ElseIf SubName(�ռ�1) = "����ȣ" Then
    SubSkill(�ռ�1) = 19
ElseIf SubName(�ռ�1) = "������" Then
    SubSkill(�ռ�1) = 20
ElseIf SubName(�ռ�1) = "������" Then
    SubSkill(�ռ�1) = 21
ElseIf SubName(�ռ�1) = "������" Then
    SubSkill(�ռ�1) = 22
ElseIf SubName(�ռ�1) = "���¾�" Then
    SubSkill(�ռ�1) = 23
ElseIf SubName(�ռ�1) = "������" Then
    SubSkill(�ռ�1) = 24
ElseIf SubName(�ռ�1) = "����ȯ1" Then
    SubSkill(�ռ�1) = 25
ElseIf SubName(�ռ�1) = "����ȣ" Then
    SubSkill(�ռ�1) = 26
ElseIf SubName(�ռ�1) = "����" Then
    SubSkill(�ռ�1) = 27
ElseIf SubName(�ռ�1) = "���ö" Then
    SubSkill(�ռ�1) = 28
ElseIf SubName(�ռ�1) = "�̼���" Then
    SubSkill(�ռ�1) = 29
ElseIf SubName(�ռ�1) = "����" Then
    SubSkill(�ռ�1) = 30
Else
    SubSkill(�ռ�1) = Int((8 * Rnd) + 31)
End If

ElseIf SubRank(�ռ�1) = "Secret" Then
 SubAt(�ռ�1) = val(SubAt(�ռ�1)) + 20
 SubR(�ռ�1) = val(SubR(�ռ�1)) + 20
 SubSt(�ռ�1) = val(SubSt(�ռ�1)) + 20
 SubAm(�ռ�1) = val(SubAm(�ռ�1)) + 20
 SubDe(�ռ�1) = val(SubDe(�ռ�1)) + 20
 SubPa(�ռ�1) = val(SubPa(�ռ�1)) + 20
 SubSe(�ռ�1) = val(SubSe(�ռ�1)) + 20
 SubCo(�ռ�1) = val(SubCo(�ռ�1)) + 20
 Dim è�Ǿ� As Integer
 è�Ǿ� = Int((200 * Rnd) + 1)
 If val(è�Ǿ�) <> 1 Then
  SubRank(�ռ�1) = "Secret"
 Else
  SubRank(�ռ�1) = "Champion"
  MsgBox "Congratulations!" & SubYear(�ռ�1) & SubName(�ռ�1) & "'s Rank. Champion.!"
 End If



ElseIf SubRank(�ռ�2) = "Champion" Then
 SubAt(�ռ�1) = val(SubAt(�ռ�1)) + 25
 SubR(�ռ�1) = val(SubR(�ռ�1)) + 25
 SubSt(�ռ�1) = val(SubSt(�ռ�1)) + 25
 SubAm(�ռ�1) = val(SubAm(�ռ�1)) + 25
 SubDe(�ռ�1) = val(SubDe(�ռ�1)) + 25
 SubPa(�ռ�1) = val(SubPa(�ռ�1)) + 25
 SubSe(�ռ�1) = val(SubSe(�ռ�1)) + 25
 SubCo(�ռ�1) = val(SubCo(�ռ�1)) + 25

End If

������ = val(������) - 1
SubName(�ռ�2) = ""
SubTribe(�ռ�2) = ""
SubAt(�ռ�2) = ""
SubR(�ռ�2) = ""
SubSt(�ռ�2) = ""
SubAm(�ռ�2) = ""
SubDe(�ռ�2) = ""
SubPa(�ռ�2) = ""
SubSe(�ռ�2) = ""
SubCo(�ռ�2) = ""
SubYear(�ռ�2) = ""
SubRank(�ռ�2) = ""
SubAW(�ռ�2) = ""
SubAL(�ռ�2) = ""
SubTW(�ռ�2) = ""
SubTL(�ռ�2) = ""
SubZW(�ռ�2) = ""
SubZL(�ռ�2) = ""
SubPW(�ռ�2) = ""
SubPL(�ռ�2) = ""
SubTeam(�ռ�2) = ""
SubLev(�ռ�2) = ""
SubExp(�ռ�2) = ""
SubMExp(�ռ�2) = ""
SubPoint(�ռ�2) = ""
SubVic(�ռ�2) = ""
SubSeVic(�ռ�2) = ""
SubNum(�ռ�2) = 0
SubNW(�ռ�2) = ""
SubSkill(�ռ�2) = 0

Do Until �ռ�2 = val(������) - 5
 SubName(�ռ�2) = SubName(�ռ�2 + 1)
 SubTribe(�ռ�2) = SubTribe(�ռ�2 + 1)
 SubAt(�ռ�2) = SubAt(�ռ�2 + 1)
 SubR(�ռ�2) = SubR(�ռ�2 + 1)
 SubSt(�ռ�2) = SubSt(�ռ�2 + 1)
 SubAm(�ռ�2) = SubAm(�ռ�2 + 1)
 SubDe(�ռ�2) = SubDe(�ռ�2 + 1)
 SubPa(�ռ�2) = SubPa(�ռ�2 + 1)
 SubSe(�ռ�2) = SubSe(�ռ�2 + 1)
 SubCo(�ռ�2) = SubCo(�ռ�2 + 1)
 SubYear(�ռ�2) = SubYear(�ռ�2 + 1)
 SubRank(�ռ�2) = SubRank(�ռ�2 + 1)
 SubAW(�ռ�2) = SubAW(�ռ�2 + 1)
 SubAL(�ռ�2) = SubAL(�ռ�2 + 1)
 SubTW(�ռ�2) = SubTW(�ռ�2 + 1)
 SubTL(�ռ�2) = SubTL(�ռ�2 + 1)
 SubZW(�ռ�2) = SubZW(�ռ�2 + 1)
 SubZL(�ռ�2) = SubZL(�ռ�2 + 1)
 SubPW(�ռ�2) = SubPW(�ռ�2 + 1)
 SubPL(�ռ�2) = SubPL(�ռ�2 + 1)
 SubTeam(�ռ�2) = SubTeam(�ռ�2 + 1)
 SubLev(�ռ�2) = SubLev(�ռ�2 + 1)
 SubExp(�ռ�2) = SubExp(�ռ�2 + 1)
 SubMExp(�ռ�2) = SubMExp(�ռ�2 + 1)
 SubPoint(�ռ�2) = SubPoint(�ռ�2 + 1)
 SubVic(�ռ�2) = SubVic(�ռ�2 + 1)
 SubSeVic(�ռ�2) = SubSeVic(�ռ�2 + 1)
 SubNum(�ռ�2) = SubNum(�ռ�2 + 1)
 SubNW(�ռ�2) = SubNW(�ռ�2 + 1)
 SubSkill(�ռ�2) = SubSkill(�ռ�2 + 1)
 �ռ�2 = val(�ռ�2) + 1
Loop

SubName(�ռ�2) = ""
SubTribe(�ռ�2) = ""
SubAt(�ռ�2) = ""
SubR(�ռ�2) = ""
SubSt(�ռ�2) = ""
SubAm(�ռ�2) = ""
SubDe(�ռ�2) = ""
SubPa(�ռ�2) = ""
SubSe(�ռ�2) = ""
SubCo(�ռ�2) = ""
SubYear(�ռ�2) = ""
SubRank(�ռ�2) = ""
SubAW(�ռ�2) = ""
SubAL(�ռ�2) = ""
SubTW(�ռ�2) = ""
SubTL(�ռ�2) = ""
SubZW(�ռ�2) = ""
SubZL(�ռ�2) = ""
SubPW(�ռ�2) = ""
SubPL(�ռ�2) = ""
SubTeam(�ռ�2) = ""
SubLev(�ռ�2) = ""
SubExp(�ռ�2) = ""
SubMExp(�ռ�2) = ""
SubPoint(�ռ�2) = ""
SubVic(�ռ�2) = ""
SubSeVic(�ռ�2) = ""
SubNum(�ռ�2) = 0
SubNW(�ռ�2) = ""
SubSkill(�ռ�2) = 0

FrmMain.Timer14.Enabled = True
FrmMain.Timer12.Enabled = True
Command1.Caption = SubYear(1) & SubName(1)
Command10.Caption = SubYear(1) & SubName(1)
Command2.Caption = SubYear(2) & SubName(2)
Command11.Caption = SubYear(2) & SubName(2)
Command3.Caption = SubYear(3) & SubName(3)
Command12.Caption = SubYear(3) & SubName(3)
Command4.Caption = SubYear(4) & SubName(4)
Command13.Caption = SubYear(4) & SubName(4)
Command5.Caption = SubYear(5) & SubName(5)
Command14.Caption = SubYear(5) & SubName(5)
Command6.Caption = SubYear(6) & SubName(6)
Command15.Caption = SubYear(6) & SubName(6)
Command7.Caption = SubYear(7) & SubName(7)
Command16.Caption = SubYear(7) & SubName(7)
Command8.Caption = SubYear(8) & SubName(8)
Command17.Caption = SubYear(8) & SubName(8)
Command9.Caption = SubYear(9) & SubName(9)
Command18.Caption = SubYear(9) & SubName(9)
Command1_Click
Command10_Click
Else
 MsgBox "���� �̸�,�⵵,��ũ�� �ٸ� ������ �����մϴ�."
End If
End Sub
