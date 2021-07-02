VERSION 5.00
Begin VB.Form FrmPLCheat 
   Caption         =   "Form30"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   1920
   LinkTopic       =   "Form30"
   ScaleHeight     =   1980
   ScaleWidth      =   1920
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   960
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   960
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   960
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "위치"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "pl패"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "pl승"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "pl경기수"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FrmPLCheat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
PL경기수 = Text1
PL승 = Text2
PL패 = Text3
PL진행 = Text4
End Sub

Private Sub Form_Load()
Text1 = PL경기수
Text2 = PL승
Text3 = PL패
Text4 = PL진행
End Sub
