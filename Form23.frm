VERSION 5.00
Begin VB.Form FrmSum 
   BackColor       =   &H00000000&
   Caption         =   "카드 합성"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   Icon            =   "Form23.frx":0000
   LinkTopic       =   "Form23"
   ScaleHeight     =   4185
   ScaleWidth      =   10830
   StartUpPosition =   2  '화면 가운데
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
         Name            =   "돋움"
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
합성1 = 1
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(합성1) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command10_Click()
On Error GoTo ABC:
합성2 = 1
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\선수\" & SubName(합성2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command11_Click()
On Error GoTo ABC:
합성2 = 2
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\선수\" & SubName(합성2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command12_Click()
On Error GoTo ABC:
합성2 = 3
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\선수\" & SubName(합성2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command13_Click()
On Error GoTo ABC:
합성2 = 4
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\선수\" & SubName(합성2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command14_Click()
On Error GoTo ABC:
합성2 = 5
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\선수\" & SubName(합성2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command15_Click()
On Error GoTo ABC:
합성2 = 6
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\선수\" & SubName(합성2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command16_Click()
On Error GoTo ABC:
합성2 = 7
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\선수\" & SubName(합성2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command17_Click()
On Error GoTo ABC:
합성2 = 8
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\선수\" & SubName(합성2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command18_Click()
On Error GoTo ABC:
합성2 = 9
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")) <> 0 Then
 Image2 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성2), 2, 2) & "]" & SubName(합성2) & ".gif")
Else
 Image2 = LoadPicture(App.Path & "\img\선수\" & SubName(합성2) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command2_Click()
On Error GoTo ABC:
합성1 = 2
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(합성1) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command3_Click()
On Error GoTo ABC:
합성1 = 3
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(합성1) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command4_Click()
On Error GoTo ABC:
합성1 = 4
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(합성1) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command5_Click()
On Error GoTo ABC:
합성1 = 5
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(합성1) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command6_Click()
On Error GoTo ABC:
합성1 = 6
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(합성1) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command7_Click()
On Error GoTo ABC:
합성1 = 7
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(합성1) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command8_Click()
On Error GoTo ABC:
합성1 = 8
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(합성1) & ".gif")
End If
Exit Sub
ABC:
Exit Sub
End Sub

Private Sub Command9_Click()
On Error GoTo ABC:
합성1 = 9
If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")) <> 0 Then
 Image1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(합성1), 2, 2) & "]" & SubName(합성1) & ".gif")
Else
 Image1 = LoadPicture(App.Path & "\img\선수\" & SubName(합성1) & ".gif")
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
If (SubName(합성1) = SubName(합성2)) And (SubYear(합성1) = SubYear(합성2)) And (합성1 <> 합성2) And (SubRank(합성1) = SubRank(합성2)) Then
 
If SubRank(합성1) = "Normal" Then
 SubRank(합성1) = "Special"
 SubAt(합성1) = val(SubAt(합성1)) + 3
 SubR(합성1) = val(SubR(합성1)) + 3
 SubSt(합성1) = val(SubSt(합성1)) + 3
 SubAm(합성1) = val(SubAm(합성1)) + 3
 SubDe(합성1) = val(SubDe(합성1)) + 3
 SubPa(합성1) = val(SubPa(합성1)) + 3
 SubSe(합성1) = val(SubSe(합성1)) + 3
 SubCo(합성1) = val(SubCo(합성1)) + 3
 
ElseIf SubRank(합성1) = "Special" Then
 SubRank(합성1) = "Rare"
 SubAt(합성1) = val(SubAt(합성1)) + 4
 SubR(합성1) = val(SubR(합성1)) + 4
 SubSt(합성1) = val(SubSt(합성1)) + 4
 SubAm(합성1) = val(SubAm(합성1)) + 4
 SubDe(합성1) = val(SubDe(합성1)) + 4
 SubPa(합성1) = val(SubPa(합성1)) + 4
 SubSe(합성1) = val(SubSe(합성1)) + 4
 SubCo(합성1) = val(SubCo(합성1)) + 4
 
ElseIf SubRank(합성1) = "Rare" Then
 SubRank(합성1) = "Unique"
 SubAt(합성1) = val(SubAt(합성1)) + 5
 SubR(합성1) = val(SubR(합성1)) + 5
 SubSt(합성1) = val(SubSt(합성1)) + 5
 SubAm(합성1) = val(SubAm(합성1)) + 5
 SubDe(합성1) = val(SubDe(합성1)) + 5
 SubPa(합성1) = val(SubPa(합성1)) + 5
 SubSe(합성1) = val(SubSe(합성1)) + 5
 SubCo(합성1) = val(SubCo(합성1)) + 5
 
ElseIf SubRank(합성1) = "Unique" Then
 SubRank(합성1) = "Elite"
 SubAt(합성1) = val(SubAt(합성1)) + 6
 SubR(합성1) = val(SubR(합성1)) + 6
 SubSt(합성1) = val(SubSt(합성1)) + 6
 SubAm(합성1) = val(SubAm(합성1)) + 6
 SubDe(합성1) = val(SubDe(합성1)) + 6
 SubPa(합성1) = val(SubPa(합성1)) + 6
 SubSe(합성1) = val(SubSe(합성1)) + 6
 SubCo(합성1) = val(SubCo(합성1)) + 6
 
ElseIf SubRank(합성1) = "Elite" Then
 SubRank(합성1) = "Legend"
 SubAt(합성1) = val(SubAt(합성1)) + 7
 SubR(합성1) = val(SubR(합성1)) + 7
 SubSt(합성1) = val(SubSt(합성1)) + 7
 SubAm(합성1) = val(SubAm(합성1)) + 7
 SubDe(합성1) = val(SubDe(합성1)) + 7
 SubPa(합성1) = val(SubPa(합성1)) + 7
 SubSe(합성1) = val(SubSe(합성1)) + 7
 SubCo(합성1) = val(SubCo(합성1)) + 7

ElseIf SubRank(합성1) = "Legend" Then
 SubAt(합성1) = val(SubAt(합성1)) + 15
 SubR(합성1) = val(SubR(합성1)) + 15
 SubSt(합성1) = val(SubSt(합성1)) + 15
 SubAm(합성1) = val(SubAm(합성1)) + 15
 SubDe(합성1) = val(SubDe(합성1)) + 15
 SubPa(합성1) = val(SubPa(합성1)) + 15
 SubSe(합성1) = val(SubSe(합성1)) + 15
 SubCo(합성1) = val(SubCo(합성1)) + 15
 SubRank(합성1) = "Secret"
If SubName(합성1) = "임요환" Then
    SubSkill(합성1) = 1
ElseIf SubName(합성1) = "이영호" Then
    SubSkill(합성1) = 2
ElseIf SubName(합성1) = "홍진호" Then
    SubSkill(합성1) = 3
ElseIf SubName(합성1) = "박정석" Then
    SubSkill(합성1) = 4
ElseIf SubName(합성1) = "이윤열" Then
    SubSkill(합성1) = 5
ElseIf SubName(합성1) = "마재윤" Then
    SubSkill(합성1) = 6
ElseIf SubName(합성1) = "최연성" Then
    SubSkill(합성1) = 7
ElseIf SubName(합성1) = "김택용" Then
    SubSkill(합성1) = 8
ElseIf SubName(합성1) = "송병구" Then
    SubSkill(합성1) = 9
ElseIf SubName(합성1) = "김구현" Then
    SubSkill(합성1) = 10
ElseIf SubName(합성1) = "허영무" Then
    SubSkill(합성1) = 11
ElseIf SubName(합성1) = "도재욱" Then
    SubSkill(합성1) = 12
ElseIf SubName(합성1) = "윤용태" Then
    SubSkill(합성1) = 13
ElseIf SubName(합성1) = "정명훈" Then
    SubSkill(합성1) = 14
ElseIf SubName(합성1) = "박성준" Then
    SubSkill(합성1) = 15
ElseIf SubName(합성1) = "이제동" Then
    SubSkill(합성1) = 16
ElseIf SubName(합성1) = "김준영" Then
    SubSkill(합성1) = 17
ElseIf SubName(합성1) = "오영종" Then
    SubSkill(합성1) = 18
ElseIf SubName(합성1) = "조용호" Then
    SubSkill(합성1) = 19
ElseIf SubName(합성1) = "서지수" Then
    SubSkill(합성1) = 20
ElseIf SubName(합성1) = "진영수" Then
    SubSkill(합성1) = 21
ElseIf SubName(합성1) = "김정우" Then
    SubSkill(합성1) = 22
ElseIf SubName(합성1) = "전태양" Then
    SubSkill(합성1) = 23
ElseIf SubName(합성1) = "서지훈" Then
    SubSkill(합성1) = 24
ElseIf SubName(합성1) = "김윤환1" Then
    SubSkill(합성1) = 25
ElseIf SubName(합성1) = "이재호" Then
    SubSkill(합성1) = 26
ElseIf SubName(합성1) = "김명운" Then
    SubSkill(합성1) = 27
ElseIf SubName(합성1) = "김민철" Then
    SubSkill(합성1) = 28
ElseIf SubName(합성1) = "이성은" Then
    SubSkill(합성1) = 29
ElseIf SubName(합성1) = "강민" Then
    SubSkill(합성1) = 30
Else
    SubSkill(합성1) = Int((8 * Rnd) + 31)
End If

ElseIf SubRank(합성1) = "Secret" Then
 SubAt(합성1) = val(SubAt(합성1)) + 20
 SubR(합성1) = val(SubR(합성1)) + 20
 SubSt(합성1) = val(SubSt(합성1)) + 20
 SubAm(합성1) = val(SubAm(합성1)) + 20
 SubDe(합성1) = val(SubDe(합성1)) + 20
 SubPa(합성1) = val(SubPa(합성1)) + 20
 SubSe(합성1) = val(SubSe(합성1)) + 20
 SubCo(합성1) = val(SubCo(합성1)) + 20
 Dim 챔피언 As Integer
 챔피언 = Int((200 * Rnd) + 1)
 If val(챔피언) <> 1 Then
  SubRank(합성1) = "Secret"
 Else
  SubRank(합성1) = "Champion"
  MsgBox "Congratulations!" & SubYear(합성1) & SubName(합성1) & "'s Rank. Champion.!"
 End If



ElseIf SubRank(합성2) = "Champion" Then
 SubAt(합성1) = val(SubAt(합성1)) + 25
 SubR(합성1) = val(SubR(합성1)) + 25
 SubSt(합성1) = val(SubSt(합성1)) + 25
 SubAm(합성1) = val(SubAm(합성1)) + 25
 SubDe(합성1) = val(SubDe(합성1)) + 25
 SubPa(합성1) = val(SubPa(합성1)) + 25
 SubSe(합성1) = val(SubSe(합성1)) + 25
 SubCo(합성1) = val(SubCo(합성1)) + 25

End If

선수수 = val(선수수) - 1
SubName(합성2) = ""
SubTribe(합성2) = ""
SubAt(합성2) = ""
SubR(합성2) = ""
SubSt(합성2) = ""
SubAm(합성2) = ""
SubDe(합성2) = ""
SubPa(합성2) = ""
SubSe(합성2) = ""
SubCo(합성2) = ""
SubYear(합성2) = ""
SubRank(합성2) = ""
SubAW(합성2) = ""
SubAL(합성2) = ""
SubTW(합성2) = ""
SubTL(합성2) = ""
SubZW(합성2) = ""
SubZL(합성2) = ""
SubPW(합성2) = ""
SubPL(합성2) = ""
SubTeam(합성2) = ""
SubLev(합성2) = ""
SubExp(합성2) = ""
SubMExp(합성2) = ""
SubPoint(합성2) = ""
SubVic(합성2) = ""
SubSeVic(합성2) = ""
SubNum(합성2) = 0
SubNW(합성2) = ""
SubSkill(합성2) = 0

Do Until 합성2 = val(선수수) - 5
 SubName(합성2) = SubName(합성2 + 1)
 SubTribe(합성2) = SubTribe(합성2 + 1)
 SubAt(합성2) = SubAt(합성2 + 1)
 SubR(합성2) = SubR(합성2 + 1)
 SubSt(합성2) = SubSt(합성2 + 1)
 SubAm(합성2) = SubAm(합성2 + 1)
 SubDe(합성2) = SubDe(합성2 + 1)
 SubPa(합성2) = SubPa(합성2 + 1)
 SubSe(합성2) = SubSe(합성2 + 1)
 SubCo(합성2) = SubCo(합성2 + 1)
 SubYear(합성2) = SubYear(합성2 + 1)
 SubRank(합성2) = SubRank(합성2 + 1)
 SubAW(합성2) = SubAW(합성2 + 1)
 SubAL(합성2) = SubAL(합성2 + 1)
 SubTW(합성2) = SubTW(합성2 + 1)
 SubTL(합성2) = SubTL(합성2 + 1)
 SubZW(합성2) = SubZW(합성2 + 1)
 SubZL(합성2) = SubZL(합성2 + 1)
 SubPW(합성2) = SubPW(합성2 + 1)
 SubPL(합성2) = SubPL(합성2 + 1)
 SubTeam(합성2) = SubTeam(합성2 + 1)
 SubLev(합성2) = SubLev(합성2 + 1)
 SubExp(합성2) = SubExp(합성2 + 1)
 SubMExp(합성2) = SubMExp(합성2 + 1)
 SubPoint(합성2) = SubPoint(합성2 + 1)
 SubVic(합성2) = SubVic(합성2 + 1)
 SubSeVic(합성2) = SubSeVic(합성2 + 1)
 SubNum(합성2) = SubNum(합성2 + 1)
 SubNW(합성2) = SubNW(합성2 + 1)
 SubSkill(합성2) = SubSkill(합성2 + 1)
 합성2 = val(합성2) + 1
Loop

SubName(합성2) = ""
SubTribe(합성2) = ""
SubAt(합성2) = ""
SubR(합성2) = ""
SubSt(합성2) = ""
SubAm(합성2) = ""
SubDe(합성2) = ""
SubPa(합성2) = ""
SubSe(합성2) = ""
SubCo(합성2) = ""
SubYear(합성2) = ""
SubRank(합성2) = ""
SubAW(합성2) = ""
SubAL(합성2) = ""
SubTW(합성2) = ""
SubTL(합성2) = ""
SubZW(합성2) = ""
SubZL(합성2) = ""
SubPW(합성2) = ""
SubPL(합성2) = ""
SubTeam(합성2) = ""
SubLev(합성2) = ""
SubExp(합성2) = ""
SubMExp(합성2) = ""
SubPoint(합성2) = ""
SubVic(합성2) = ""
SubSeVic(합성2) = ""
SubNum(합성2) = 0
SubNW(합성2) = ""
SubSkill(합성2) = 0

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
 MsgBox "같은 이름,년도,랭크의 다른 선수만 가능합니다."
End If
End Sub
