VERSION 5.00
Begin VB.Form FrmStat 
   BackColor       =   &H00000000&
   Caption         =   "�������"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4875
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form11"
   ScaleHeight     =   2490
   ScaleWidth      =   4875
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.TextBox Text1 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3360
      TabIndex        =   26
      Text            =   "1"
      Top             =   1560
      Width           =   615
   End
   Begin CSO.jcbutton jcbutton8 
      Height          =   375
      Left            =   4080
      TabIndex        =   23
      Top             =   1080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Up!"
      CaptionEffects  =   0
   End
   Begin CSO.jcbutton jcbutton7 
      Height          =   375
      Left            =   4080
      TabIndex        =   22
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Up!"
      CaptionEffects  =   0
   End
   Begin CSO.jcbutton jcbutton6 
      Height          =   375
      Left            =   4080
      TabIndex        =   21
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Up!"
      CaptionEffects  =   0
   End
   Begin CSO.jcbutton jcbutton5 
      Height          =   375
      Left            =   1680
      TabIndex        =   20
      Top             =   2040
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Up!"
      CaptionEffects  =   0
   End
   Begin CSO.jcbutton jcbutton4 
      Height          =   375
      Left            =   1680
      TabIndex        =   19
      Top             =   1560
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Up!"
      CaptionEffects  =   0
   End
   Begin CSO.jcbutton jcbutton3 
      Height          =   375
      Left            =   1680
      TabIndex        =   18
      Top             =   1080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Up!"
      CaptionEffects  =   0
   End
   Begin CSO.jcbutton jcbutton2 
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Up!"
      CaptionEffects  =   0
   End
   Begin CSO.jcbutton jcbutton1 
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Up!"
      CaptionEffects  =   0
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "������ :"
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
      TabIndex        =   27
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      Caption         =   "����Ʈ :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   25
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label15 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   24
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "���ݷ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "��Ʈ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblAT 
      BackColor       =   &H00000000&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblR 
      BackColor       =   &H00000000&
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblST 
      BackColor       =   &H00000000&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblAm 
      BackColor       =   &H00000000&
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblDe 
      BackColor       =   &H00000000&
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lblPa 
      BackColor       =   &H00000000&
      Caption         =   "Label14"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblSe 
      BackColor       =   &H00000000&
      Caption         =   "Label15"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblCo 
      BackColor       =   &H00000000&
      Caption         =   "Label16"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "FrmStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
lblAt = MyAt(����)
lblR = MyR(����)
lblSt = MySt(����)
lblAm = MyAm(����)
lblDe = MyDe(����)
lblPa = MyPa(����)
lblSe = MySe(����)
lblCo = MyCo(����)
Label15 = MyPoint(����)
End Sub

Private Sub jcbutton1_Click()
If val(MyPoint(����)) >= val(Text1) Then
 MyAt(����) = val(MyAt(����)) + val(Text1)
 MyPoint(����) = val(MyPoint(����)) - val(Text1)
 lblAt = MyAt(����)
Label15 = MyPoint(����)
End If
End Sub

Private Sub jcbutton2_Click()
If val(MyPoint(����)) >= val(Text1) Then
 MyR(����) = val(MyR(����)) + val(Text1)
 MyPoint(����) = val(MyPoint(����)) - val(Text1)
 lblR = MyR(����)
Label15 = MyPoint(����)
End If
End Sub

Private Sub jcbutton3_Click()
If val(MyPoint(����)) >= val(Text1) Then
 MySt(����) = val(MySt(����)) + val(Text1)
 MyPoint(����) = val(MyPoint(����)) - val(Text1)
 lblSt = MySt(����)
Label15 = MyPoint(����)
End If
End Sub

Private Sub jcbutton4_Click()
If val(MyPoint(����)) >= val(Text1) Then
 MyAm(����) = val(MyAm(����)) + val(Text1)
 MyPoint(����) = val(MyPoint(����)) - val(Text1)
 lblAm = MyAm(����)
Label15 = MyPoint(����)
End If
End Sub

Private Sub jcbutton5_Click()
If val(MyPoint(����)) >= val(Text1) Then
 MyDe(����) = val(MyDe(����)) + val(Text1)
 MyPoint(����) = val(MyPoint(����)) - val(Text1)
 lblDe = MyDe(����)
Label15 = MyPoint(����)
End If
End Sub

Private Sub jcbutton6_Click()
If val(MyPoint(����)) >= val(Text1) Then
 MyPa(����) = val(MyPa(����)) + val(Text1)
 MyPoint(����) = val(MyPoint(����)) - val(Text1)
 lblPa = MyPa(����)
Label15 = MyPoint(����)
End If
End Sub

Private Sub jcbutton7_Click()
If val(MyPoint(����)) >= val(Text1) Then
 MySe(����) = val(MySe(����)) + val(Text1)
 MyPoint(����) = val(MyPoint(����)) - val(Text1)
 lblSe = MySe(����)
Label15 = MyPoint(����)
End If
End Sub

Private Sub jcbutton8_Click()
If val(MyPoint(����)) >= val(Text1) Then
 MyCo(����) = val(MyCo(����)) + val(Text1)
 MyPoint(����) = val(MyPoint(����)) - val(Text1)
 lblCo = MyCo(����)
Label15 = MyPoint(����)
End If
End Sub
