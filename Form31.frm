VERSION 5.00
Begin VB.Form Form31 
   BackColor       =   &H00000000&
   Caption         =   "진행창"
   ClientHeight    =   10275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   LinkTopic       =   "Form31"
   ScaleHeight     =   10275
   ScaleWidth      =   11850
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Tim_Proc 
      Interval        =   500
      Left            =   120
      Top             =   120
   End
   Begin CSO.ProgressBar Prg_Map 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2640
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      Value           =   55
      Theme           =   5
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextForeColor   =   16777215
      Text            =   "T vs Z = 55 : 45"
      TextEffectColor =   0
      TextEffect      =   4
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   4
      Text            =   "Form31.frx":0000
      Top             =   3240
      Width           =   6735
   End
   Begin VB.Label lblRScore 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   27.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8520
      TabIndex        =   8
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblLScore 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   27.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2520
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblMName 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "네오 글래디에이터"
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
      Left            =   4800
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   1500
      Left            =   5040
      Picture         =   "Form31.frx":0006
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label lblNameR 
      BackStyle       =   0  '투명
      Caption         =   "<11>크로우[Ex]"
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
      Left            =   9480
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblNameL 
      BackStyle       =   0  '투명
      Caption         =   "Label1"
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
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblTribeR 
      BackStyle       =   0  '투명
      Caption         =   "(P)"
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
      Left            =   11040
      TabIndex        =   1
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblTribeL 
      BackStyle       =   0  '투명
      Caption         =   "(P)"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   9480
      Picture         =   "Form31.frx":2BD7
      Top             =   720
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  '투명하지 않음
      Height          =   7095
      Left            =   2520
      Top             =   3120
      Width           =   6975
   End
   Begin VB.Image Img_Left 
      Height          =   1500
      Left            =   840
      Picture         =   "Form31.frx":4431
      Top             =   720
      Width           =   1500
   End
End
Attribute VB_Name = "Form31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Unload Me
'종족,이름,스코어 기입 부분
If MyTribe(선택) = 1 Then
    lblTribeL = "(T)"
ElseIf MyTribe(선택) = 2 Then
    lblTribeL = "(Z)"
End If

If 종족(Oee) = 1 Then
    lblTribeR = "(T)"
ElseIf 종족(Oee) = 2 Then
    lblTribeR = "(Z)"
End If

lblNameL = MyYear(선택) & MyName(선택)
lblNameR = OYear(Oee) & 이름(Oee)

lblLScore = MW
lblRScore = OW
'종족,이름,스코어 기입 부분 종료
'Call Strategy(val(선택), val(Oee))
'↑전략 결정
End Sub

Private Sub Tim_Proc_Timer()
Call Fight(val(선택), val(Oee))
End Sub
