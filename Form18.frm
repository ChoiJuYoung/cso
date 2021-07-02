VERSION 5.00
Begin VB.Form FrmShop 
   BackColor       =   &H00808080&
   Caption         =   "Shop"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   Icon            =   "Form18.frx":0000
   LinkTopic       =   "Form18"
   ScaleHeight     =   9255
   ScaleWidth      =   10095
   StartUpPosition =   2  '화면 가운데
   Begin CSO.jcbutton jcbutton1 
      Height          =   375
      Left            =   3120
      TabIndex        =   30
      Top             =   1320
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   255
      Caption         =   "고급상점가기"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.Label Label6 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[10000 Cro]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   6720
      TabIndex        =   29
      Top             =   8880
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[Normal ~ Elite]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   6720
      TabIndex        =   28
      Top             =   8520
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "<Protoss Card Pack ver.S>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   6720
      TabIndex        =   27
      Top             =   8280
      Width           =   3375
   End
   Begin VB.Image Protoss3 
      Height          =   1125
      Index           =   8
      Left            =   7320
      Picture         =   "Form18.frx":628A
      Top             =   7080
      Width           =   2250
   End
   Begin VB.Label Label6 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[7000 Cro]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   6720
      TabIndex        =   26
      Top             =   6480
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[Normal ~ Rare]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   6720
      TabIndex        =   25
      Top             =   6120
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "<Protoss Card Pack ver.A>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   6720
      TabIndex        =   24
      Top             =   5880
      Width           =   3375
   End
   Begin VB.Image Protoss2 
      Height          =   1500
      Index           =   7
      Left            =   7320
      Picture         =   "Form18.frx":76E4
      Top             =   4320
      Width           =   2250
   End
   Begin VB.Label Label6 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[3000 Cro]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   6720
      TabIndex        =   23
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[Normal]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   6720
      TabIndex        =   22
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "<Protoss Card Pack>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   6720
      TabIndex        =   21
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Image Protoss1 
      Height          =   1275
      Index           =   6
      Left            =   7800
      Picture         =   "Form18.frx":A51A
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label Label6 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[10000 Cro]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   20
      Top             =   8880
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[Normal ~ Elite]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   19
      Top             =   8520
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "<Terran Card Pack ver.S>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   18
      Top             =   8280
      Width           =   3375
   End
   Begin VB.Image Terran3 
      Height          =   1125
      Index           =   5
      Left            =   3960
      Picture         =   "Form18.frx":B1C9
      Top             =   7080
      Width           =   2250
   End
   Begin VB.Label Label6 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[7000 Cro]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   17
      Top             =   6480
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[Normal ~ Rare]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   16
      Top             =   6120
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "<Terran Card Pack ver.A>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   15
      Top             =   5880
      Width           =   3375
   End
   Begin VB.Image Terran2 
      Height          =   1500
      Index           =   4
      Left            =   3960
      Picture         =   "Form18.frx":C799
      Top             =   4320
      Width           =   2250
   End
   Begin VB.Label Label6 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[3000 Cro]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   14
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[Normal]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   13
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "<Terran Card Pack>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   12
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Image Terran1 
      Height          =   1275
      Index           =   3
      Left            =   4440
      Picture         =   "Form18.frx":F11F
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label Label6 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[3000 Cro]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   11
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[Normal]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "<Zerg Card Pack>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Image Zerg1 
      Height          =   1275
      Index           =   2
      Left            =   1080
      Picture         =   "Form18.frx":1010A
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label Label6 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[7000 Cro]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   6480
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[Normal ~ Rare]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   6120
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "<Zerg Card Pack ver.A>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   5880
      Width           =   3375
   End
   Begin VB.Image Zerg2 
      Height          =   1500
      Index           =   1
      Left            =   600
      Picture         =   "Form18.frx":10EF6
      Top             =   4320
      Width           =   2250
   End
   Begin VB.Shape Shape2 
      Height          =   2535
      Index           =   8
      Left            =   6720
      Top             =   6720
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      Height          =   2535
      Index           =   7
      Left            =   3360
      Top             =   6720
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      Height          =   2535
      Index           =   6
      Left            =   6720
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      Height          =   2535
      Index           =   5
      Left            =   6720
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      Height          =   2535
      Index           =   4
      Left            =   3360
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      Height          =   2535
      Index           =   3
      Left            =   3360
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      Height          =   2535
      Index           =   2
      Left            =   0
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      Height          =   2535
      Index           =   1
      Left            =   0
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label6 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[10000 Cro]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   8880
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "[Normal ~ Elite]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   8520
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "<Zerg Card Pack ver.S>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   8280
      Width           =   3375
   End
   Begin VB.Image Zerg3 
      Height          =   1125
      Index           =   0
      Left            =   600
      Picture         =   "Form18.frx":12F44
      Top             =   7080
      Width           =   2250
   End
   Begin VB.Shape Shape2 
      Height          =   2535
      Index           =   0
      Left            =   0
      Top             =   6720
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "<금액에 따른 카드는 1장이 나옵니다.>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   10095
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "<게임에 사용하는 선수 카드를 구매하실 수 있습니다.>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   10095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "<Card Shop>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   10095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      Height          =   1695
      Left            =   0
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "FrmShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If Mode = "Normal" Then
    Label5(0) = "[Normal ~ Unique]"
    Label5(5) = "[Normal ~ Unique]"
    Label5(8) = "[Normal ~ Unique]"
    jcbutton1.Enabled = False
End If
End Sub

Private Sub jcbutton1_Click()
FrmHighShop.Show
End Sub

Private Sub Protoss1_Click(Index As Integer)
구매 = 7
FrmShopConf.Show
End Sub

Private Sub Protoss2_Click(Index As Integer)
구매 = 8
FrmShopConf.Show
End Sub

Private Sub Protoss3_Click(Index As Integer)
구매 = 9
FrmShopConf.Show
End Sub

Private Sub Terran1_Click(Index As Integer)
구매 = 4
FrmShopConf.Show
End Sub

Private Sub Terran2_Click(Index As Integer)
구매 = 5
FrmShopConf.Show
End Sub

Private Sub Terran3_Click(Index As Integer)
구매 = 6
FrmShopConf.Show
End Sub

Private Sub Zerg1_Click(Index As Integer)
구매 = 1
FrmShopConf.Show
End Sub

Private Sub Zerg2_Click(Index As Integer)
구매 = 2
FrmShopConf.Show
End Sub

Private Sub Zerg3_Click(Index As Integer)
구매 = 3
FrmShopConf.Show
End Sub
