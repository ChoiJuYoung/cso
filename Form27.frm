VERSION 5.00
Begin VB.Form FrmCopyRight 
   Caption         =   "CopyRight"
   ClientHeight    =   10650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "Form27.frx":0000
   LinkTopic       =   "Form27"
   ScaleHeight     =   10650
   ScaleWidth      =   15240
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11760
      Top             =   8880
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "제작 : 최주영(hajuu96123@naver.com)"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   0
      Top             =   9360
      Width           =   8415
   End
   Begin VB.Image Image1 
      Height          =   10995
      Left            =   0
      Picture         =   "Form27.frx":628A
      Stretch         =   -1  'True
      Top             =   -360
      Width           =   15300
   End
End
Attribute VB_Name = "FrmCopyRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Copy = 0
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
If Label1.Visible = "False" Then
 Label1.Visible = True
Else
 Label1.Visible = False
End If

If (val(Copy)) = 0 Or (val(Copy) = 1) Then
 Label1 = "제작 : 최주영(hajuu96123@naver.com)"
Copy = val(Copy) + 1
ElseIf (val(Copy)) = 2 Or (val(Copy) = 3) Then
 Label1 = "원작 : 권우진(Crow)"
Copy = val(Copy) + 1
ElseIf (val(Copy) = 4) Or (val(Copy) = 5) Then
 Label1 = "사진 : (제옹신, 오늘은, 권우진)"
Copy = val(Copy) + 1
ElseIf (val(Copy) = 6) Or (val(Copy) = 7) Then
 Label1 = "오프닝 : 플투군"
Copy = val(Copy) + 1
ElseIf (val(Copy) = 8) Then
 Label1 = "FPTeam. 동훈(Smallestman@naver.com)"
 Copy = val(Copy) + 1
ElseIf (val(Copy) = 9) Then
 Copy = 0
End If
End Sub
