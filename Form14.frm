VERSION 5.00
Begin VB.Form FrmGameInfo 
   Caption         =   "Á¤º¸Ã¢"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form14"
   ScaleHeight     =   5775
   ScaleWidth      =   9255
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin CSO.jcbutton jcbutton1 
      Height          =   375
      Left            =   3120
      TabIndex        =   18
      Top             =   3600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Go"
      CaptionEffects  =   0
   End
   Begin VB.Label lblO¿¬½Â 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   25
      Top             =   5400
      Width           =   3015
   End
   Begin VB.Label lblM¿¬½Â 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label lblMapTri 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   23
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label lblOTeam 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   22
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label lblMTeam 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label lblOrank 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   20
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label lblMrank 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label lblSAO 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "<¾øÀ½>"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   17
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label lblSA 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "<¾øÀ½>"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label Label5 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "-Special Ability-"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   15
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "-Special Ability-"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label lblOTT 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "[10½Â 0ÆÐ]"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   13
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label lblMTT 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "[10½Â 0ÆÐ]"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Label lblMT 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "[Vs T]"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   11
      Top             =   4920
      Width           =   3015
   End
   Begin VB.Label lblOT 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "[Vs Z]"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "<»ó´ë Á¾Á·Àü>"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Label lblOSt 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Stats : 6500"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label lblOR 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Rank : C-"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   7
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label lblMSt 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Stats : 6500"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label lblMR 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Rank : C-"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "<Card Rank>"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Åõ¸íÇÏÁö ¾ÊÀ½
      Height          =   1095
      Index           =   5
      Left            =   6240
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Åõ¸íÇÏÁö ¾ÊÀ½
      Height          =   1095
      Index           =   4
      Left            =   3120
      Top             =   4680
      Width           =   3135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Åõ¸íÇÏÁö ¾ÊÀ½
      Height          =   1095
      Index           =   3
      Left            =   0
      Top             =   4680
      Width           =   3135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Åõ¸íÇÏÁö ¾ÊÀ½
      Height          =   735
      Index           =   2
      Left            =   6240
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Åõ¸íÇÏÁö ¾ÊÀ½
      Height          =   735
      Index           =   1
      Left            =   3120
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Åõ¸íÇÏÁö ¾ÊÀ½
      Height          =   735
      Index           =   0
      Left            =   0
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label lblMapName 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "[³×¿À¹®±Û·¹ÀÌºê]"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "<Map Order>"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   3135
   End
   Begin VB.Image ImgMa 
      Height          =   1500
      Left            =   3960
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label lblON 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "ÀÌ¿µÈ£"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label lblMN 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "ÀÓÅÂ±Ô"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Image ImgOp 
      Height          =   1500
      Left            =   6960
      Top             =   360
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Åõ¸íÇÏÁö ¾ÊÀ½
      BorderColor     =   &H00000000&
      Height          =   4000
      Index           =   2
      Left            =   6240
      Top             =   0
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Åõ¸íÇÏÁö ¾ÊÀ½
      BorderColor     =   &H00000000&
      Height          =   4000
      Index           =   1
      Left            =   3120
      Top             =   0
      Width           =   3135
   End
   Begin VB.Image ImgMe 
      Height          =   1500
      Left            =   720
      Top             =   360
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Åõ¸íÇÏÁö ¾ÊÀ½
      BorderColor     =   &H00000000&
      Height          =   4005
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "FrmGameInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

If val(MyTribe(¼±ÅÃ)) = 1 Then
 If val(Á¾Á·(Oee)) = 1 Then
  lblMapTri = "T vs T = 50 : 50"
 ElseIf val(Á¾Á·(Oee)) = 2 Then
  lblMapTri = "T vs Z = " & TZT(Map) & " : " & TZZ(Map)
 ElseIf val(Á¾Á·(Oee)) = 3 Then
  lblMapTri = "T vs P = " & PTT(Map) & " : " & PTP(Map)
 End If
ElseIf val(MyTribe(¼±ÅÃ)) = 2 Then
 If val(Á¾Á·(Oee)) = 1 Then
  lblMapTri = "Z vs T = " & TZZ(Map) & " : " & TZT(Map)
 ElseIf val(Á¾Á·(Oee)) = 2 Then
  lblMapTri = "Z vs Z = 50 : 50"
 ElseIf val(Á¾Á·(Oee)) = 3 Then
  lblMapTri = "Z vs P = " & ZPZ(Map) & " : " & ZPP(Map)
 End If
ElseIf val(MyTribe(¼±ÅÃ)) = 3 Then
 If val(Á¾Á·(Oee)) = 1 Then
  lblMapTri = "P vs T = " & PTP(Map) & " : " & PTT(Map)
 ElseIf val(Á¾Á·(Oee)) = 2 Then
  lblMapTri = "P vs Z = " & ZPP(Map) & " : " & ZPZ(Map)
 ElseIf val(Á¾Á·(Oee)) = 3 Then
  lblMapTri = "P vs P = 50 : 50"
 End If
End If


Dim ¸¶ÀÌ·©Å© As String, »ó´ë·©Å© As String
lblMN = MyYear(¼±ÅÃ) & " " & MyName(¼±ÅÃ)
lblON = OYear(Oee) & " " & ÀÌ¸§(Oee)
lblMTeam = MyTeam(¼±ÅÃ)
lblOTeam = Team(Oee)

If MySkill(¼±ÅÃ) = 1 Then
    lblSA = "üÕð¤"
ElseIf MySkill(¼±ÅÃ) = 2 Then
    lblSA = "õÌðûÜ²Ðï"
ElseIf MySkill(¼±ÅÃ) = 3 Then
    lblSA = "øìù¦"
ElseIf MySkill(¼±ÅÃ) = 4 Then
    lblSA = "çÈê©"
ElseIf MySkill(¼±ÅÃ) = 5 Then
    lblSA = "ô¸î¦"
ElseIf MySkill(¼±ÅÃ) = 6 Then
    lblSA = "Maestro"
ElseIf MySkill(¼±ÅÃ) = 7 Then
    lblSA = "ÎÖÚª"
ElseIf MySkill(¼±ÅÃ) = 8 Then
    lblSA = "Þ«×£"
ElseIf MySkill(¼±ÅÃ) = 9 Then
    lblSA = "Íð×£"
ElseIf MySkill(¼±ÅÃ) = 10 Then
    lblSA = "îå×£"
ElseIf MySkill(¼±ÅÃ) = 11 Then
    lblSA = "ê£×£"
ElseIf MySkill(¼±ÅÃ) = 12 Then
    lblSA = "ÎÖ×£"
ElseIf MySkill(¼±ÅÃ) = 13 Then
    lblSA = "Òâ×£"
ElseIf MySkill(¼±ÅÃ) = 14 Then
    lblSA = "ÏÐÜâ"
ElseIf MySkill(¼±ÅÃ) = 15 Then
    lblSA = "÷ããê"
ElseIf MySkill(¼±ÅÃ) = 16 Then
    lblSA = "øìÏÖ"
ElseIf MySkill(¼±ÅÃ) = 17 Then
    lblSA = "ÓÞìÑ"
ElseIf MySkill(¼±ÅÃ) = 18 Then
    lblSA = "ÞÝãê"
ElseIf MySkill(¼±ÅÃ) = 19 Then
    lblSA = "ÙÌÔÛ"
ElseIf MySkill(¼±ÅÃ) = 20 Then
    lblSA = "åüð¤"
ElseIf MySkill(¼±ÅÃ) = 21 Then
    lblSA = "rEd sNipeR"
ElseIf MySkill(¼±ÅÃ) = 22 Then
    lblSA = "ÝÕÞÝðè"
ElseIf MySkill(¼±ÅÃ) = 23 Then
    lblSA = "Sun"
ElseIf MySkill(¼±ÅÃ) = 24 Then
    lblSA = "pErfecT tErraN"
ElseIf MySkill(¼±ÅÃ) = 25 Then
    lblSA = "Brain"
ElseIf MySkill(¼±ÅÃ) = 26 Then
    lblSA = "zErg sPeicaL kILLeR"
ElseIf MySkill(¼±ÅÃ) = 27 Then
    lblSA = "¾î¸°¿ÕÀÚ"
ElseIf MySkill(¼±ÅÃ) = 28 Then
    lblSA = "ôÑÛú"
ElseIf MySkill(¼±ÅÃ) = 29 Then
    lblSA = "ýÙê£íþ"
ElseIf MySkill(¼±ÅÃ) = 30 Then
    lblSA = "ÙÓßÌÊ«"
End If

If Skill(Oee) = 1 Then
    lblSAO = "üÕð¤"
ElseIf Skill(Oee) = 2 Then
    lblSAO = "õÌðûÜ²Ðï"
ElseIf Skill(Oee) = 3 Then
    lblSAO = "øìù¦"
ElseIf Skill(Oee) = 4 Then
    lblSAO = "çÈê©"
ElseIf Skill(Oee) = 5 Then
    lblSAO = "ô¸î¦"
ElseIf Skill(Oee) = 6 Then
    lblSAO = "Maestro"
ElseIf Skill(Oee) = 7 Then
    lblSAO = "ÎÖÚª"
ElseIf Skill(Oee) = 8 Then
    lblSAO = "Þ«×£"
ElseIf Skill(Oee) = 9 Then
    lblSAO = "Íð×£"
ElseIf Skill(Oee) = 10 Then
    lblSAO = "îå×£"
ElseIf Skill(Oee) = 11 Then
    lblSAO = "ê£×£"
ElseIf Skill(Oee) = 12 Then
    lblSAO = "ÎÖ×£"
ElseIf Skill(Oee) = 13 Then
    lblSAO = "Òâ×£"
ElseIf Skill(Oee) = 14 Then
    lblSAO = "ÏÐÜâ"
ElseIf Skill(Oee) = 15 Then
    lblSAO = "÷ããê"
ElseIf Skill(Oee) = 16 Then
    lblSAO = "øìÏÖ"
ElseIf Skill(Oee) = 17 Then
    lblSAO = "ÓÞìÑ"
ElseIf Skill(Oee) = 18 Then
    lblSAO = "ÞÝãê"
ElseIf Skill(Oee) = 19 Then
    lblSAO = "ÙÌÔÛ"
ElseIf Skill(Oee) = 20 Then
    lblSAO = "åüð¤"
ElseIf Skill(Oee) = 21 Then
    lblSAO = "rEd sNipeR"
ElseIf Skill(Oee) = 22 Then
    lblSAO = "ÝÕÞÝðè"
ElseIf Skill(Oee) = 23 Then
    lblSAO = "Sun"
ElseIf Skill(Oee) = 24 Then
    lblSAO = "pErfecT tErraN"
ElseIf Skill(Oee) = 25 Then
    lblSAO = "Brain"
ElseIf Skill(Oee) = 26 Then
    lblSAO = "zErg sPeicaL kILLeR"
ElseIf Skill(Oee) = 27 Then
    lblSAO = "¾î¸°¿ÕÀÚ"
ElseIf Skill(Oee) = 28 Then
    lblSAO = "ôÑÛú"
ElseIf Skill(Oee) = 29 Then
    lblSAO = "ýÙê£íþ"
ElseIf Skill(Oee) = 30 Then
    lblSAO = "ÙÓßÌÊ«"
End If


If Á¾Á·(Oee) = 1 Then
 lblOT = "[Vs T]"
 lblMTT = "[" & MyTW(¼±ÅÃ) & "½Â " & MyTL(¼±ÅÃ) & "ÆÐ]"
 If MyT¿¬(¼±ÅÃ) = "W" Then
  lblM¿¬½Â = "[" & MyT¿¬½Â(¼±ÅÃ) & "¿¬½ÂÁß" & "]"
  lblM¿¬½Â.ForeColor = RGB(0, 255, 255)
 ElseIf MyT¿¬(¼±ÅÃ) = "L" Then
  lblM¿¬½Â = "[" & MyT¿¬½Â(¼±ÅÃ) & "¿¬ÆÐÁß" & "]"
  lblM¿¬½Â.ForeColor = RGB(255, 0, 0)
 End If
ElseIf Á¾Á·(Oee) = 2 Then
 lblOT = "[Vs Z]"
 lblMTT = "[" & MyZW(¼±ÅÃ) & "½Â " & MyZL(¼±ÅÃ) & "ÆÐ]"
 If MyZ¿¬(¼±ÅÃ) = "W" Then
  lblM¿¬½Â = "[" & MyZ¿¬½Â(¼±ÅÃ) & "¿¬½ÂÁß" & "]"
  lblM¿¬½Â.ForeColor = RGB(0, 255, 255)
 ElseIf MyZ¿¬(¼±ÅÃ) = "L" Then
  lblM¿¬½Â = "[" & MyZ¿¬½Â(¼±ÅÃ) & "¿¬ÆÐÁß" & "]"
  lblM¿¬½Â.ForeColor = RGB(255, 0, 0)
 End If
ElseIf Á¾Á·(Oee) = 3 Then
 lblOT = "[Vs P]"
 lblMTT = "[" & MyPW(¼±ÅÃ) & "½Â " & MyPL(¼±ÅÃ) & "ÆÐ]"
 If MyP¿¬(¼±ÅÃ) = "W" Then
  lblM¿¬½Â = "[" & MyP¿¬½Â(¼±ÅÃ) & "¿¬½ÂÁß" & "]"
  lblM¿¬½Â.ForeColor = RGB(0, 255, 255)
 ElseIf MyP¿¬(¼±ÅÃ) = "L" Then
  lblM¿¬½Â = "[" & MyP¿¬½Â(¼±ÅÃ) & "¿¬ÆÐÁß" & "]"
  lblM¿¬½Â.ForeColor = RGB(255, 0, 0)
 End If
End If

If MyTribe(¼±ÅÃ) = 1 Then
 lblMT = "[Vs T]"
 lblOTT = "[" & T½Â¸®(Oee) & "½Â " & TÆÐ¹è(Oee) & "ÆÐ]"
 If T¿¬(Oee) = "W" Then
  lblO¿¬½Â = "[" & T¿¬½Â(Oee) & "¿¬½ÂÁß" & "]"
  lblO¿¬½Â.ForeColor = RGB(0, 255, 255)
 ElseIf T¿¬(Oee) = "L" Then
  lblO¿¬½Â = "[" & T¿¬½Â(Oee) & "¿¬ÆÐÁß" & "]"
  lblO¿¬½Â.ForeColor = RGB(255, 0, 0)
 End If
ElseIf MyTribe(¼±ÅÃ) = 2 Then
 lblMT = "[Vs Z]"
 lblOTT = "[" & Z½Â¸®(Oee) & "½Â " & ZÆÐ¹è(Oee) & "ÆÐ]"
 If Z¿¬(Oee) = "W" Then
  lblO¿¬½Â = "[" & Z¿¬½Â(Oee) & "¿¬½ÂÁß" & "]"
  lblO¿¬½Â.ForeColor = RGB(0, 255, 255)
 ElseIf Z¿¬(Oee) = "L" Then
  lblO¿¬½Â = "[" & Z¿¬½Â(Oee) & "¿¬ÆÐÁß" & "]"
  lblO¿¬½Â.ForeColor = RGB(255, 0, 0)
 End If
ElseIf MyTribe(¼±ÅÃ) = 3 Then
 lblMT = "[Vs P]"
 lblOTT = "[" & P½Â¸®(Oee) & "½Â " & PÆÐ¹è(Oee) & "ÆÐ]"
 If P¿¬(Oee) = "W" Then
  lblO¿¬½Â = "[" & P¿¬½Â(Oee) & "¿¬½ÂÁß" & "]"
  lblO¿¬½Â.ForeColor = RGB(0, 255, 255)
 ElseIf P¿¬(Oee) = "L" Then
  lblO¿¬½Â = "[" & P¿¬½Â(Oee) & "¿¬ÆÐÁß" & "]"
  lblO¿¬½Â.ForeColor = RGB(255, 0, 0)
 End If
End If

If Len(Dir(App.Path & "\img\¼±¼ö\[" & Mid(MyYear(¼±ÅÃ), 2, 2) & "]" & MN & ".gif")) <> 0 Then
 ImgMe = LoadPicture(App.Path & "\img\¼±¼ö\[" & Mid(MyYear(¼±ÅÃ), 2, 2) & "]" & MN & ".gif")
Else
 ImgMe = LoadPicture(App.Path & "\img\¼±¼ö\" & MyName(¼±ÅÃ) & ".gif")
End If

If Len(Dir(App.Path & "\img\¼±¼ö\[" & Mid(OYear(Oee), 2, 2) & "]" & ÀÌ¸§(Oee) & ".gif")) <> 0 Then
 ImgOp = LoadPicture(App.Path & "\img\¼±¼ö\[" & Mid(OYear(Oee), 2, 2) & "]" & ÀÌ¸§(Oee) & ".gif")
Else
 ImgOp = LoadPicture(App.Path & "\img\¼±¼ö\" & ÀÌ¸§(Oee) & ".gif")
End If

If Len(Dir(App.Path & "\img\¸Ê\" & MapName(Map) & ".gif")) <> 0 Then
 ImgMa = LoadPicture(App.Path & "\img\¸Ê\" & MapName(Map) & ".gif")
Else
 ImgMa = Nothing
End If


AA = val(AT) + val(R) + val(St) + val(Am) + val(De) + val(Pa) + val(SE) + val(Co)
If val(AA) < 4500 Then
¸¶ÀÌ·©Å© = "F"
'&H4B4B4B
ElseIf val(AA) >= 4500 And val(AA) < 4700 Then
¸¶ÀÌ·©Å© = "E"
'&HB0B0B0
ElseIf val(AA) >= 4700 And val(AA) < 4800 Then
¸¶ÀÌ·©Å© = "D-"
'&HFF3232
ElseIf val(AA) >= 4800 And val(AA) < 4900 Then
¸¶ÀÌ·©Å© = "D"
'&HFF3232
ElseIf val(AA) >= 4900 And val(AA) < 5000 Then
¸¶ÀÌ·©Å© = "D+"
'&HFF3232
ElseIf val(AA) >= 5000 And val(AA) < 5100 Then
¸¶ÀÌ·©Å© = "C-"
'&HFF00&
ElseIf val(AA) >= 5100 And val(AA) < 5200 Then
¸¶ÀÌ·©Å© = "C"
'&HFF00&
ElseIf val(AA) >= 5200 And val(AA) < 5400 Then
¸¶ÀÌ·©Å© = "C+"
'&HFF00&
ElseIf val(AA) >= 5400 And val(AA) < 5600 Then
¸¶ÀÌ·©Å© = "B-"
'&HFFFD&
ElseIf val(AA) >= 5600 And val(AA) < 5800 Then
¸¶ÀÌ·©Å© = "B"
'&HFFFD&
ElseIf val(AA) >= 5800 And val(AA) < 6000 Then
¸¶ÀÌ·©Å© = "B+"
'&HFFFD&
ElseIf val(AA) >= 6000 And val(AA) < 6200 Then
¸¶ÀÌ·©Å© = "A-"
'&H6663FF
ElseIf val(AA) >= 6200 And val(AA) < 6400 Then
¸¶ÀÌ·©Å© = "A"
'&H6663FF
ElseIf val(AA) >= 6400 And val(AA) < 6600 Then
¸¶ÀÌ·©Å© = "A+"
'&H6663FF
ElseIf val(AA) >= 6600 And val(AA) < 6800 Then
¸¶ÀÌ·©Å© = "S"
ElseIf val(AA) >= 6800 And val(AA) < 7000 Then
¸¶ÀÌ·©Å© = "SS"
ElseIf val(AA) >= 7000 Then
¸¶ÀÌ·©Å© = "SSS"
End If

AAO = val(°ø°Ý·Â(Oee)) + val(°ßÁ¦(Oee)) + val(Àü·«(Oee)) + val(¹°·®(Oee)) + val(¼öºñ·Â(Oee)) + val(Á¤Âû(Oee)) + val(¼¾½º(Oee)) + val(ÄÁÆ®·Ñ(Oee))
If val(AAO) < 4500 Then
»ó´ë·©Å© = "F"
'&H4B4B4B
ElseIf val(AAO) >= 4500 And val(AAO) < 4700 Then
»ó´ë·©Å© = "E"
'&HB0B0B0
ElseIf val(AAO) >= 4700 And val(AAO) < 4800 Then
»ó´ë·©Å© = "D-"
'&HFF3232
ElseIf val(AAO) >= 4800 And val(AAO) < 4900 Then
»ó´ë·©Å© = "D"
'&HFF3232
ElseIf val(AAO) >= 4900 And val(AAO) < 5000 Then
»ó´ë·©Å© = "D+"
'&HFF3232
ElseIf val(AAO) >= 5000 And val(AAO) < 5100 Then
»ó´ë·©Å© = "C-"
'&HFF00&
ElseIf val(AAO) >= 5100 And val(AAO) < 5200 Then
»ó´ë·©Å© = "C"
'&HFF00&
ElseIf val(AAO) >= 5200 And val(AAO) < 5400 Then
»ó´ë·©Å© = "C+"
'&HFF00&
ElseIf val(AAO) >= 5400 And val(AAO) < 5600 Then
»ó´ë·©Å© = "B-"
'&HFFFD&
ElseIf val(AAO) >= 5600 And val(AAO) < 5800 Then
»ó´ë·©Å© = "B"
'&HFFFD&
ElseIf val(AAO) >= 5800 And val(AAO) < 6000 Then
»ó´ë·©Å© = "B+"
'&HFFFD&
ElseIf val(AAO) >= 6000 And val(AAO) < 6200 Then
»ó´ë·©Å© = "A-"
'&H6663FF
ElseIf val(AAO) >= 6200 And val(AAO) < 6400 Then
»ó´ë·©Å© = "A"
'&H6663FF
ElseIf val(AAO) >= 6400 And val(AAO) < 6600 Then
»ó´ë·©Å© = "A+"
'&H6663FF
ElseIf val(AAO) >= 6600 And val(AAO) < 6800 Then
»ó´ë·©Å© = "S"
ElseIf val(AAO) >= 6800 And val(AAO) < 7000 Then
»ó´ë·©Å© = "SS"
ElseIf val(AAO) >= 7000 Then
»ó´ë·©Å© = "SSS"
End If

lblMrank = MyRank(¼±ÅÃ)
lblOrank = ·©Å©(Oee)
If MyRank(¼±ÅÃ) = "Normal" Then
 lblMrank.ForeColor = RGB(0, 0, 0)
ElseIf MyRank(¼±ÅÃ) = "Special" Then
 lblMrank.ForeColor = RGB(0, 255, 0)
ElseIf MyRank(¼±ÅÃ) = "Rare" Then
 lblMrank.ForeColor = &HFF80FF
ElseIf MyRank(¼±ÅÃ) = "Unique" Then
 lblMrank.ForeColor = &HFF8080
ElseIf MyRank(¼±ÅÃ) = "Elite" Then
 lblMrank.ForeColor = &H800080
ElseIf MyRank(¼±ÅÃ) = "Legend" Then
 lblMrank.ForeColor = &H80FF&
ElseIf MyRank(¼±ÅÃ) = "Secret" Then
 lblMrank.ForeColor = &HFFC0C0
ElseIf MyRank(¼±ÅÃ) = "Champion" Then
 lblMrank.ForeColor = RGB(255, 0, 0)
End If

If ·©Å©(Oee) = "Normal" Then
 lblOrank.ForeColor = RGB(0, 0, 0)
ElseIf ·©Å©(Oee) = "Special" Then
 lblOrank.ForeColor = RGB(0, 255, 0)
ElseIf ·©Å©(Oee) = "Rare" Then
 lblOrank.ForeColor = &HFF80FF
ElseIf ·©Å©(Oee) = "Unique" Then
 lblOrank.ForeColor = &HFF8080
ElseIf ·©Å©(Oee) = "Elite" Then
 lblOrank.ForeColor = &H800080
ElseIf ·©Å©(Oee) = "Legend" Then
 lblOrank.ForeColor = &H80FF&
ElseIf ·©Å©(Oee) = "Secret" Then
 lblOrank.ForeColor = &HFFC0C0
ElseIf ·©Å©(Oee) = "Champion" Then
 lblOrank.ForeColor = RGB(255, 0, 0)
End If



lblMapName = MapName(Map)
lblMR = "Rank : " & ¸¶ÀÌ·©Å©
lblOR = "Rank : " & »ó´ë·©Å©
lblMSt = "Stats : " & AA
lblOSt = "Stats : " & AAO
End Sub

Private Sub jcbutton1_Click()
OStyle = Int((5 * Rnd) + 1)
If val(OStyle) = 1 Then
 OStyle = "°ø°ÝÇü"
ElseIf val(OStyle) = 2 Then
 OStyle = "¼öºñÇü"
ElseIf val(OStyle) = 3 Then
 OStyle = "°ßÁ¦Çü"
ElseIf val(OStyle) = 4 Then
 OStyle = "¿î¿µÇü"
ElseIf val(OStyle) = 5 Then
 OStyle = "³ë¸ÖÇü"
End If
FrmPickSt.Visible = True
Unload Me
End Sub
