VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H009A9A9A&
   Caption         =   "Information"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12990
   Icon            =   "Form16.frx":0000
   LinkTopic       =   "Form16"
   ScaleHeight     =   7920
   ScaleWidth      =   12990
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer15 
      Interval        =   100
      Left            =   11760
      Top             =   8160
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   240
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   8280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   11280
      Top             =   8760
   End
   Begin CSO.jcbutton CmdDelete 
      Height          =   495
      Left            =   1920
      TabIndex        =   40
      Top             =   6840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "카드 판매"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin VB.Timer Timer13 
      Interval        =   10
      Left            =   11280
      Top             =   8160
   End
   Begin CSO.jcbutton CmdSetting 
      Height          =   495
      Left            =   120
      TabIndex        =   39
      Top             =   6840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "카드 세팅"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   10800
      Top             =   8760
   End
   Begin CSO.jcbutton Cmd합성 
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   6240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "카드 합성"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin CSO.jcbutton CmdShop 
      Height          =   495
      Left            =   1920
      TabIndex        =   18
      Top             =   6240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "상점"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin VB.Timer Timer11 
      Interval        =   10
      Left            =   10800
      Top             =   8160
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8400
      Top             =   8160
   End
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   8880
      Top             =   8160
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   9360
      Top             =   8160
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   9840
      Top             =   8160
   End
   Begin VB.Timer Timer5 
      Interval        =   100
      Left            =   10320
      Top             =   8160
   End
   Begin VB.Timer Timer6 
      Interval        =   10
      Left            =   8400
      Top             =   8760
   End
   Begin VB.Timer Timer7 
      Interval        =   100
      Left            =   8880
      Top             =   8760
   End
   Begin VB.Timer Timer8 
      Interval        =   200
      Left            =   9360
      Top             =   8760
   End
   Begin VB.Timer Timer9 
      Interval        =   100
      Left            =   9840
      Top             =   8760
   End
   Begin VB.Timer Timer10 
      Interval        =   10
      Left            =   10320
      Top             =   8760
   End
   Begin CSO.jcbutton CmdMa 
      Height          =   495
      Left            =   4320
      TabIndex        =   13
      Top             =   5640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "메인으로"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin CSO.jcbutton CmdSear 
      Height          =   495
      Left            =   1920
      TabIndex        =   14
      Top             =   5640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "선수검색"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin CSO.jcbutton CmdGo 
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   5640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "진행"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin CSO.jcbutton CmdSa 
      Height          =   495
      Left            =   4320
      TabIndex        =   16
      Top             =   6240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "저장하기"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin VB.Label lblNews2 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Label4"
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
      Left            =   6480
      TabIndex        =   50
      Top             =   120
      Width           =   6495
   End
   Begin VB.Label lblNews1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Label4"
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
      TabIndex        =   49
      Top             =   120
      Width           =   6495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   1  '투명하지 않음
      Height          =   495
      Left            =   6480
      Top             =   0
      Width           =   6495
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label lblDeck 
      BackStyle       =   0  '투명
      Caption         =   "Label4"
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
      Left            =   3720
      TabIndex        =   48
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Label lbl하향 
      BackStyle       =   0  '투명
      Caption         =   "하향횟수 : 5000"
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
      Left            =   3720
      TabIndex        =   47
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label lbl쿠폰 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "쿠폰 : 100000"
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
      Left            =   2280
      TabIndex        =   46
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label lblMode 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Label4"
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
      Left            =   1440
      TabIndex        =   45
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "Label3"
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
      Left            =   3720
      TabIndex        =   43
      Top             =   7200
      Width           =   2655
   End
   Begin VB.Label lblPL 
      BackStyle       =   0  '투명
      Caption         =   "Label3"
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
      Left            =   3720
      TabIndex        =   42
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label lblTurn 
      BackStyle       =   0  '투명
      Caption         =   "Turn : OSL"
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
      Left            =   3720
      TabIndex        =   41
      Top             =   7680
      Width           =   2775
   End
   Begin VB.Image ImgS6 
      Height          =   1500
      Left            =   11160
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Image ImgS2 
      Height          =   1500
      Left            =   9000
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label lblSN9 
      BackStyle       =   0  '투명
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
      Left            =   11640
      TabIndex        =   38
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label lblSy9 
      BackStyle       =   0  '투명
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
      Left            =   11160
      TabIndex        =   37
      Top             =   7440
      Width           =   495
   End
   Begin VB.Label lblSN8 
      BackStyle       =   0  '투명
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
      Left            =   9480
      TabIndex        =   36
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label lblSy8 
      BackStyle       =   0  '투명
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
      Left            =   9000
      TabIndex        =   35
      Top             =   7440
      Width           =   495
   End
   Begin VB.Label lblSN7 
      BackStyle       =   0  '투명
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
      Left            =   7320
      TabIndex        =   34
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label lblSy7 
      BackStyle       =   0  '투명
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
      Left            =   6840
      TabIndex        =   33
      Top             =   7440
      Width           =   495
   End
   Begin VB.Label lblSN6 
      BackStyle       =   0  '투명
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
      Left            =   11640
      TabIndex        =   32
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label lblSy6 
      BackStyle       =   0  '투명
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
      Left            =   11160
      TabIndex        =   31
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label lblSN5 
      BackStyle       =   0  '투명
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
      Left            =   9480
      TabIndex        =   30
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label lblSy5 
      BackStyle       =   0  '투명
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
      Left            =   9000
      TabIndex        =   29
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label lblSN4 
      BackStyle       =   0  '투명
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
      Left            =   7320
      TabIndex        =   28
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label lblSy4 
      BackStyle       =   0  '투명
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
      Left            =   6840
      TabIndex        =   27
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label lblSN3 
      BackStyle       =   0  '투명
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
      Left            =   11640
      TabIndex        =   26
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblSy3 
      BackStyle       =   0  '투명
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
      Left            =   11160
      TabIndex        =   25
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblSN2 
      BackStyle       =   0  '투명
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
      Left            =   9480
      TabIndex        =   24
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblSy2 
      BackStyle       =   0  '투명
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
      Left            =   9000
      TabIndex        =   23
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblSN1 
      BackStyle       =   0  '투명
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
      Left            =   7320
      TabIndex        =   22
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblSy1 
      BackStyle       =   0  '투명
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
      Left            =   6840
      TabIndex        =   21
      Top             =   2640
      Width           =   495
   End
   Begin VB.Image ImgS9 
      Height          =   1500
      Left            =   11160
      Top             =   5880
      Width           =   1500
   End
   Begin VB.Image ImgS8 
      Height          =   1500
      Left            =   9000
      Top             =   5880
      Width           =   1500
   End
   Begin VB.Image ImgS7 
      Height          =   1500
      Left            =   6840
      Top             =   5880
      Width           =   1500
   End
   Begin VB.Image ImgS5 
      Height          =   1500
      Left            =   9000
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Image ImgS4 
      Height          =   1500
      Left            =   6840
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Image ImgS3 
      Height          =   1500
      Left            =   11160
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Image ImgS1 
      Height          =   1500
      Left            =   6840
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Line Line11 
      X1              =   10800
      X2              =   10800
      Y1              =   960
      Y2              =   7920
   End
   Begin VB.Line Line10 
      X1              =   8640
      X2              =   8640
      Y1              =   960
      Y2              =   7920
   End
   Begin VB.Line Line9 
      X1              =   6480
      X2              =   12960
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line8 
      X1              =   6480
      X2              =   12960
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "<Gamer Card> - Sub Card"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009A9A9A&
      Height          =   255
      Left            =   6480
      TabIndex        =   19
      Top             =   600
      Width           =   6495
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  '투명하지 않음
      Height          =   495
      Left            =   6480
      Top             =   480
      Width           =   6495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0FFFF&
      Height          =   6975
      Left            =   6480
      Top             =   960
      Width           =   6495
   End
   Begin VB.Label lblMoney 
      BackStyle       =   0  '투명
      Caption         =   "Money : 299792458Cro"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   7380
      Width           =   6255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  '투명하지 않음
      Height          =   2415
      Left            =   0
      Top             =   5520
      Width           =   6495
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   0
      Y1              =   5040
      Y2              =   480
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   6480
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line5 
      X1              =   4320
      X2              =   4320
      Y1              =   960
      Y2              =   5520
   End
   Begin VB.Line Line4 
      X1              =   2160
      X2              =   2160
      Y1              =   960
      Y2              =   5520
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   6480
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line2 
      X1              =   6480
      X2              =   6480
      Y1              =   960
      Y2              =   5520
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6480
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label lblYe6 
      BackStyle       =   0  '투명
      Caption         =   "<11>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label lblName6 
      BackStyle       =   0  '투명
      Caption         =   "이영호"
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
      Left            =   5160
      TabIndex        =   11
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label lblYe5 
      BackStyle       =   0  '투명
      Caption         =   "<11>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label lblName5 
      BackStyle       =   0  '투명
      Caption         =   "이영호"
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
      Left            =   3000
      TabIndex        =   9
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label lblYe4 
      BackStyle       =   0  '투명
      Caption         =   "<11>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label lblName4 
      BackStyle       =   0  '투명
      Caption         =   "이영호"
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
      Left            =   840
      TabIndex        =   7
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label lblYe3 
      BackStyle       =   0  '투명
      Caption         =   "<11>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblName3 
      BackStyle       =   0  '투명
      Caption         =   "이영호"
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
      Left            =   5160
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblYe2 
      BackStyle       =   0  '투명
      Caption         =   "<11>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblName2 
      BackStyle       =   0  '투명
      Caption         =   "이영호"
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
      Left            =   3000
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblYe1 
      BackStyle       =   0  '투명
      Caption         =   "<11>"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblName1 
      BackStyle       =   0  '투명
      Caption         =   "이영호"
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
      Left            =   840
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "<Gamer Card> - Main Card"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009A9A9A&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   6495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      Height          =   495
      Left            =   0
      Top             =   480
      Width           =   6495
   End
   Begin VB.Image ImgP6 
      Height          =   1500
      Left            =   4680
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Image ImgP5 
      Height          =   1500
      Left            =   2520
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Image ImgP4 
      Height          =   1500
      Left            =   360
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Image ImgP3 
      Height          =   1500
      Left            =   4680
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Image ImgP2 
      Height          =   1500
      Left            =   2520
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Image ImgP1 
      Height          =   1500
      Left            =   360
      Top             =   1080
      Width           =   1500
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function SplitString(ByVal sStr As String, Optional ByVal CutString As String = " ", Optional ByVal CutLen As Long = 1) As String
     Dim i As Long, vStr As String
     For i = 1 To Len(sStr) Step CutLen
          vStr = vStr & Mid(sStr, i, CutLen) & CutString
     Next i
     SplitString = vStr
End Function

Private Sub Cmd합성_Click()
If val(선수수) >= 7 Then
FrmSum.Show
ElseIf val(선수수) = 6 Then
MsgBox "Sub카드가 없습니다."
End If
End Sub

Private Sub CmdDelete_Click()
If val(선수수) >= 7 Then
FrmFire.Show
ElseIf val(선수수) = 6 Then
MsgBox "Sub카드가 없습니다."
End If
End Sub

Private Sub CmdSetting_Click()
If val(선수수) >= 7 Then
FrmSetting.Show
ElseIf val(선수수) = 6 Then
MsgBox "Sub카드가 없습니다."
End If
End Sub

Private Sub CmdShop_Click()
If val(선수수) >= 15 Then
 MsgBox "선수는 15명이 최대입니다."
Else
FrmCoupon.Show
End If
End Sub

Private Sub Form_Load()
lblMode = Mode
For i = 1 To 6
    PL출전자(i) = True
Next
If Mode = "Hard" Then
    lblMode.ForeColor = RGB(0, 0, 255)
ElseIf Mode = "Hell" Then
    lblMode.ForeColor = RGB(255, 0, 0)
End If

For 마이 = 1 To 6
    If MyNW(마이) = "CB64" Or MyNW(마이) = "CB32" Then
        MyNW(마이) = "CB16"
    End If
Next

Visible확인 = True
Timer12.Enabled = True
lblYe1 = MyYear(1)
lblYe2 = MyYear(2)
lblYe3 = MyYear(3)
lblYe4 = MyYear(4)
lblYe5 = MyYear(5)
lblYe6 = MyYear(6)
lblName1 = MyName(1)
lblName2 = MyName(2)
lblName3 = MyName(3)
lblName4 = MyName(4)
lblName5 = MyName(5)
lblName6 = MyName(6)

If Len(Dir(App.Path & "\img\선수\[" & Mid(MyYear(1), 2, 2) & "]" & MyName(1) & ".gif")) <> 0 Then
 ImgP1 = LoadPicture(App.Path & "\img\선수\[" & Mid(MyYear(1), 2, 2) & "]" & MyName(1) & ".gif")
Else
 ImgP1 = LoadPicture(App.Path & "\img\선수\" & MyName(1) & ".gif")
End If

If Len(Dir(App.Path & "\img\선수\[" & Mid(MyYear(2), 2, 2) & "]" & MyName(2) & ".gif")) <> 0 Then
 ImgP2 = LoadPicture(App.Path & "\img\선수\[" & Mid(MyYear(2), 2, 2) & "]" & MyName(2) & ".gif")
Else
 ImgP2 = LoadPicture(App.Path & "\img\선수\" & MyName(2) & ".gif")
End If

If Len(Dir(App.Path & "\img\선수\[" & Mid(MyYear(3), 2, 2) & "]" & MyName(3) & ".gif")) <> 0 Then
 ImgP3 = LoadPicture(App.Path & "\img\선수\[" & Mid(MyYear(3), 2, 2) & "]" & MyName(3) & ".gif")
Else
 ImgP3 = LoadPicture(App.Path & "\img\선수\" & MyName(3) & ".gif")
End If

If Len(Dir(App.Path & "\img\선수\[" & Mid(MyYear(4), 2, 2) & "]" & MyName(4) & ".gif")) <> 0 Then
 ImgP4 = LoadPicture(App.Path & "\img\선수\[" & Mid(MyYear(4), 2, 2) & "]" & MyName(4) & ".gif")
Else
 ImgP4 = LoadPicture(App.Path & "\img\선수\" & MyName(4) & ".gif")
End If

If Len(Dir(App.Path & "\img\선수\[" & Mid(MyYear(5), 2, 2) & "]" & MyName(5) & ".gif")) <> 0 Then
 ImgP5 = LoadPicture(App.Path & "\img\선수\[" & Mid(MyYear(5), 2, 2) & "]" & MyName(5) & ".gif")
Else
 ImgP5 = LoadPicture(App.Path & "\img\선수\" & MyName(5) & ".gif")
End If

If Len(Dir(App.Path & "\img\선수\[" & Mid(MyYear(6), 2, 2) & "]" & MyName(6) & ".gif")) <> 0 Then
 ImgP6 = LoadPicture(App.Path & "\img\선수\[" & Mid(MyYear(6), 2, 2) & "]" & MyName(6) & ".gif")
Else
 ImgP6 = LoadPicture(App.Path & "\img\선수\" & MyName(6) & ".gif")
End If


If MyTeam(1) = MyTeam(2) And MyTeam(2) = MyTeam(3) And MyTeam(3) = MyTeam(4) And MyTeam(4) = MyTeam(5) And MyTeam(5) = MyTeam(6) Then
    If MyTeam(1) = "삼성전자" Then
        Deck = "삼성전자"
    ElseIf MyTeam(1) = "STX" Then
        Deck = "STX"
    ElseIf MyTeam(1) = "Mystar" Then
        Deck = "Mystar"
    ElseIf MyTeam(1) = "eSTRO" Then
        Deck = "eSTRO"
    ElseIf MyTeam(1) = "공군" Then
        Deck = "공군"
    ElseIf MyTeam(1) = "8th" Then
        Deck = "8th"
    End If
ElseIf MyTeam(1) = "POS" Or MyTeam(1) = "MBC" Then
    If MyTeam(2) = "POS" Or MyTeam(2) = "MBC" Then
        If MyTeam(3) = "POS" Or MyTeam(3) = "MBC" Then
            If MyTeam(4) = "POS" Or MyTeam(4) = "MBC" Then
                If MyTeam(5) = "POS" Or MyTeam(5) = "MBC" Then
                    If MyTeam(6) = "POS" Or MyTeam(6) = "MBC" Then
                        Deck = "MBC"
                    End If
                End If
            End If
        End If
    End If
ElseIf MyTeam(1) = "GO" Or MyTeam(1) = "CJ" Then
    If MyTeam(2) = "GO" Or MyTeam(2) = "CJ" Then
        If MyTeam(3) = "GO" Or MyTeam(3) = "CJ" Then
            If MyTeam(4) = "GO" Or MyTeam(4) = "CJ" Then
                If MyTeam(5) = "GO" Or MyTeam(5) = "CJ" Then
                    If MyTeam(6) = "GO" Or MyTeam(6) = "CJ" Then
                        Deck = "CJ"
                    End If
                End If
            End If
        End If
    End If
ElseIf MyTeam(1) = "온게임넷" Or MyTeam(1) = "하이트" Then
    If MyTeam(2) = "온게임넷" Or MyTeam(2) = "하이트" Then
        If MyTeam(3) = "온게임넷" Or MyTeam(3) = "하이트" Then
            If MyTeam(4) = "온게임넷" Or MyTeam(4) = "하이트" Then
                If MyTeam(5) = "온게임넷" Or MyTeam(5) = "하이트" Then
                    If MyTeam(6) = "온게임넷" Or MyTeam(6) = "하이트" Then
                        Deck = "하이트"
                    End If
                End If
            End If
        End If
    End If
ElseIf MyTeam(1) = "르까프" Or MyTeam(1) = "화승" Or MyTeam(1) = "PLUS" Then
    If MyTeam(2) = "르까프" Or MyTeam(2) = "화승" Or MyTeam(2) = "PLUS" Then
        If MyTeam(3) = "르까프" Or MyTeam(3) = "화승" Or MyTeam(3) = "PLUS" Then
            If MyTeam(4) = "르까프" Or MyTeam(4) = "화승" Or MyTeam(4) = "PLUS" Then
                If MyTeam(5) = "르까프" Or MyTeam(5) = "화승" Or MyTeam(5) = "PLUS" Then
                    If MyTeam(6) = "르까프" Or MyTeam(6) = "화승" Or MyTeam(6) = "PLUS" Then
                        Deck = "화승"
                    End If
                End If
            End If
        End If
    End If
ElseIf MyTeam(1) = "한빛" Or MyTeam(1) = "웅진" Then
    If MyTeam(2) = "한빛" Or MyTeam(2) = "웅진" Then
        If MyTeam(3) = "한빛" Or MyTeam(3) = "웅진" Then
            If MyTeam(4) = "한빛" Or MyTeam(4) = "웅진" Then
                If MyTeam(5) = "한빛" Or MyTeam(5) = "웅진" Then
                    If MyTeam(6) = "한빛" Or MyTeam(6) = "웅진" Then
                        Deck = "웅진"
                    End If
                End If
            End If
        End If
    End If
ElseIf MyTeam(1) = "KTF" Or MyTeam(1) = "KT" Then
    If MyTeam(2) = "KTF" Or MyTeam(2) = "KT" Then
        If MyTeam(3) = "KTF" Or MyTeam(3) = "KT" Then
            If MyTeam(4) = "KTF" Or MyTeam(4) = "KT" Then
                If MyTeam(5) = "KTF" Or MyTeam(5) = "KT" Then
                    If MyTeam(6) = "KTF" Or MyTeam(6) = "KT" Then
                        Deck = "KT"
                    End If
                End If
            End If
        End If
    End If
ElseIf MyTeam(1) = "4U" Or MyTeam(1) = "SK" Or MyTeam(1) = "Orion" Or MyTeam(1) = "IS" Then
    If MyTeam(2) = "4U" Or MyTeam(2) = "SK" Or MyTeam(2) = "Orion" Or MyTeam(2) = "IS" Then
        If MyTeam(3) = "4U" Or MyTeam(3) = "SK" Or MyTeam(3) = "Orion" Or MyTeam(3) = "IS" Then
            If MyTeam(4) = "4U" Or MyTeam(4) = "SK" Or MyTeam(4) = "Orion" Or MyTeam(4) = "IS" Then
                If MyTeam(5) = "4U" Or MyTeam(5) = "SK" Or MyTeam(5) = "Orion" Or MyTeam(5) = "IS" Then
                    If MyTeam(6) = "4U" Or MyTeam(6) = "SK" Or MyTeam(6) = "Orion" Or MyTeam(6) = "IS" Then
                        Deck = "SK"
                    End If
                End If
            End If
        End If
    End If
ElseIf MyTeam(1) = "Toona" Or MyTeam(1) = "폭스" Or MyTeam(1) = "Curitel" Or MyTeam(1) = "Pantech" Then
    If MyTeam(2) = "Toona" Or MyTeam(2) = "폭스" Or MyTeam(2) = "Curitel" Or MyTeam(2) = "Pantech" Then
        If MyTeam(3) = "Toona" Or MyTeam(3) = "폭스" Or MyTeam(3) = "Curitel" Or MyTeam(3) = "Pantech" Then
            If MyTeam(4) = "Toona" Or MyTeam(4) = "폭스" Or MyTeam(4) = "Curitel" Or MyTeam(4) = "Pantech" Then
                If MyTeam(5) = "Toona" Or MyTeam(5) = "폭스" Or MyTeam(5) = "Curitel" Or MyTeam(5) = "Pantech" Then
                    If MyTeam(6) = "Toona" Or MyTeam(6) = "폭스" Or MyTeam(6) = "Curitel" Or MyTeam(6) = "Pantech" Then
                        Deck = "폭스"
                    End If
                End If
            End If
        End If
    End If
Else
    Deck = ""
End If
If Deck <> "" Then
    If MyYear(1) = MyYear(2) And MyYear(2) = MyYear(3) And MyYear(3) = MyYear(4) And MyYear(5) = MyYear(4) And MyYear(5) = MyYear(6) Then
        Deck년도 = True
    Else
        Deck년도 = False
    End If
End If

Randomize AP
Randomize Oee
Randomize Map
Randomize 상점NPC

End Sub

Private Sub CmdGo_Click()
FrmMain.Visible = False
Timer2.Enabled = False
Timer3.Enabled = False
CmdGO.Visible = False

If Turn = "OSL" Then
    FrmSingPick.Show
Else
    Frm_BatInfo.Show
End If

Dim i As Integer
For i = 1 To 6
    PL출전자(i) = True
Next
If Visible확인 = True Then
 CmdSa.Visible = False
 Visible확인 = False
End If
End Sub

Private Sub CmdMa_Click()
Shell App.Path & "\cso.exe", vbNormalFocus
End
End Sub

Private Sub CmdSa_Click()
Visible확인 = False
If 불러온이름 <> "" Then
    Call Save
    불러옴 = True
    MsgBox "완료"
    CmdSa.Visible = False
ElseIf 네임 = "" Then
 MsgBox "취소되었습니다."
End If
End Sub

Private Sub CmdSear_Click()
검색 = "Yes"
SearName = InputBox("선수 이름을 넣어주세요. 김윤환 = 김윤환1, 프녕 = 이영호1", "이름넣기", "<년도>이름   으로 기입")
Sear = 0
If SearName <> "" Then
 Do Until OYear(Sear) & 이름(Sear) = SearName
 If val(Sear) <= 800 Then
  Sear = val(Sear) + 1
 End If
 If val(Sear) = 801 Then
  MsgBox "그런 선수는 없습니다."
  검색 = "No"
  Exit Do
 End If
 Loop

If 검색 = "No" Then
Else
 FrmSearch.Show
End If

ElseIf SearName = "" Then
 MsgBox "취소되었습니다."
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub ImgP1_Click()
선택 = 1
FrmAbility.Show
End Sub

Private Sub ImgP2_Click()
선택 = 2
FrmAbility.Show
End Sub

Private Sub ImgP3_Click()
선택 = 3
FrmAbility.Show
End Sub

Private Sub ImgP4_Click()
선택 = 4
FrmAbility.Show
End Sub

Private Sub ImgP5_Click()
선택 = 5
FrmAbility.Show
End Sub

Private Sub ImgP6_Click()
선택 = 6
FrmAbility.Show
End Sub

Private Sub Label1_Click()
Dim 코드입력 As String
Dim 상점도우미보조 As Integer
코드입력 = InputBox(코드입력)
If val(코드입력) = 48772561 Then
상점도우미 = 0
상점도우미보조 = val(선수수) - 5
 If (크로우생산 = "No") Then
 SubName(상점도우미보조) = 이름(상점도우미)
 SubTeam(상점도우미보조) = Team(상점도우미)
 SubAt(상점도우미보조) = NPC공격력(상점도우미)
 SubR(상점도우미보조) = NPC견제(상점도우미)
 SubSt(상점도우미보조) = NPC전략(상점도우미)
 SubAm(상점도우미보조) = NPC물량(상점도우미)
 SubDe(상점도우미보조) = NPC수비력(상점도우미)
 SubPa(상점도우미보조) = NPC정찰(상점도우미)
 SubSe(상점도우미보조) = NPC센스(상점도우미)
 SubCo(상점도우미보조) = NPC컨트롤(상점도우미)
 SubRank(상점도우미보조) = 랭크(상점도우미)
 SubYear(상점도우미보조) = OYear(상점도우미)
 SubTribe(상점도우미보조) = 종족(상점도우미)
 SubLev(상점도우미보조) = 1
 SubExp(상점도우미보조) = 0
 SubMExp(상점도우미보조) = 50
 SubPoint(상점도우미보조) = 0
 SubNum(상점도우미보조) = val(상점도우미)
 SubAW(상점도우미보조) = 0
 SubAL(상점도우미보조) = 0
 SubTW(상점도우미보조) = 0
 SubTL(상점도우미보조) = 0
 SubZW(상점도우미보조) = 0
 SubZL(상점도우미보조) = 0
 SubPW(상점도우미보조) = 0
 SubPL(상점도우미보조) = 0
 SubVic(상점도우미보조) = 0
 SubSeVic(상점도우미보조) = 0
 
If SubRank(상점도우미보조) = "Normal" Or SubRank(상점도우미보조) = "Special" Then
 SubNW(상점도우미보조) = "CB16"
ElseIf SubRank(상점도우미보조) = "Rare" Then
 SubNW(상점도우미보조) = "CA1"
ElseIf SubRank(상점도우미보조) = "Unique" Then
 SubNW(상점도우미보조) = "CA2"
ElseIf SubRank(상점도우미보조) = "Elite" Then
 SubNW(상점도우미보조) = "CA3"
Else
 SubNW(상점도우미보조) = "CS32"
End If

 선수수 = val(선수수) + 1
 
 Timer12.Enabled = True
 크로우생산 = "Yes"
 Else
 MsgBox "크로우 선수는 한명만 가능합니다."
End If
ElseIf 코드입력 = "Administrator" Then
 Dim AdCode As String
 AdCode = InputBox("입력")
 If AdCode = "SavE" Then
  If Text1 <> "세이브" Then
   Text1 = "세이브"
  Else
   Text1 = "헤이브"
  End If
 ElseIf AdCode = "mOneY" Then
  Dim MoneyCode As String
   MoneyCode = InputBox("입력")
   Money = MoneyCode
 ElseIf AdCode = "Player" Then
 
 
  Dim PlayerCode As String
  Dim PlayerCodeNPC As String
  Dim 검색 As String
  Dim 구매가능 As String
  PlayerCode = InputBox("이름입력")

PlayerCodeNPC = 0
If PlayerCode <> "" Then
 Do Until OYear(PlayerCodeNPC) & 이름(PlayerCodeNPC) = PlayerCode
 If val(PlayerCodeNPC) <= 800 Then
  PlayerCodeNPC = val(PlayerCodeNPC) + 1
 End If
 If val(PlayerCodeNPC) = 801 Then
  MsgBox "그런 선수는 없습니다."
  검색 = "No"
  Exit Do
 End If
 Loop

If PlayerCodeNPC <= 800 Then
    If OYear(PlayerCodeNPC) & 이름(PlayerCodeNPC) = PlayerCode Then
     구매가능 = "Yes"
    End If
End If

If 검색 = "No" Then
Else
Dim PlayerCodeNpc보조 As String
PlayerCodeNpc보조 = val(선수수) - 5
If 구매가능 = "Yes" Then
 SubName(PlayerCodeNpc보조) = 이름(PlayerCodeNPC)
 SubTeam(PlayerCodeNpc보조) = Team(PlayerCodeNPC)
 SubAt(PlayerCodeNpc보조) = NPC공격력(PlayerCodeNPC)
 SubR(PlayerCodeNpc보조) = NPC견제(PlayerCodeNPC)
 SubSt(PlayerCodeNpc보조) = NPC전략(PlayerCodeNPC)
 SubAm(PlayerCodeNpc보조) = NPC물량(PlayerCodeNPC)
 SubDe(PlayerCodeNpc보조) = NPC수비력(PlayerCodeNPC)
 SubPa(PlayerCodeNpc보조) = NPC정찰(PlayerCodeNPC)
 SubSe(PlayerCodeNpc보조) = NPC센스(PlayerCodeNPC)
 SubCo(PlayerCodeNpc보조) = NPC컨트롤(PlayerCodeNPC)
 SubRank(PlayerCodeNpc보조) = 랭크(PlayerCodeNPC)
 SubYear(PlayerCodeNpc보조) = OYear(PlayerCodeNPC)
 SubTribe(PlayerCodeNpc보조) = 종족(PlayerCodeNPC)
 SubLev(PlayerCodeNpc보조) = 1
 SubExp(PlayerCodeNpc보조) = 0
 SubMExp(PlayerCodeNpc보조) = 50
 SubPoint(PlayerCodeNpc보조) = 0
 SubNum(PlayerCodeNpc보조) = val(PlayerCodeNPC)
 SubAW(PlayerCodeNpc보조) = 0
 SubAL(PlayerCodeNpc보조) = 0
 SubTW(PlayerCodeNpc보조) = 0
 SubTL(PlayerCodeNpc보조) = 0
 SubZW(PlayerCodeNpc보조) = 0
 SubZL(PlayerCodeNpc보조) = 0
 SubPW(PlayerCodeNpc보조) = 0
 SubPL(PlayerCodeNpc보조) = 0
 SubVic(PlayerCodeNpc보조) = 0
 SubSeVic(PlayerCodeNpc보조) = 0
 
If SubRank(PlayerCodeNpc보조) = "Normal" Or SubRank(PlayerCodeNpc보조) = "Special" Then
 SubNW(PlayerCodeNpc보조) = "CB16"
ElseIf SubRank(PlayerCodeNpc보조) = "Rare" Then
 SubNW(PlayerCodeNpc보조) = "CA1"
ElseIf SubRank(PlayerCodeNpc보조) = "Unique" Then
 SubNW(PlayerCodeNpc보조) = "CA2"
ElseIf SubRank(PlayerCodeNpc보조) = "Elite" Then
 SubNW(PlayerCodeNpc보조) = "CA3"
Else
 SubNW(PlayerCodeNpc보조) = "CS32"
End If

 SubSkill(PlayerCodeNpc보조) = Skill(PlayerCodeNPC)
 선수수 = val(선수수) + 1
End If
End If
ElseIf PlayerCode = "" Then
 MsgBox "취소되었습니다."
End If
 ElseIf AdCode = "OSL" Then
  Turn = "OSL"
 ElseIf AdCode = "PL" Then
  Turn = "PL"
 ElseIf AdCode = "MYNW" Then
  Dim AddMYNWcode As String
  Dim AddMYNWTurn As String
  AddMYNWcode = InputBox("입력")
  AddMYNWTurn = InputBox("선수넘버")
  MyNW(AddMYNWTurn) = AddMYNWcode
 ElseIf AdCode = "Set" Then
 
  PL경기수 = 11
  End If
Else
 MsgBox "코드오류"
End If
 FrmMain.Timer12.Enabled = True
End Sub

Private Sub Label2_Click()
Dim SourcePL As String
SourcePL = InputBox("CODE 입력")
If SourcePL = "PL조정" Then
    FrmPLCheat.Show
End If
End Sub

Private Sub lbl쿠폰_Click()
Dim CouponHe As String
CouponHe = InputBox("Code")
If CouponHe = "CodeNameMKP" Then
    쿠폰 = val(쿠폰) + 1
End If
End Sub

Private Sub lbl하향_Click()
Turn = "PL"
End Sub

Private Sub lblNews1_Click()
Frm_BatInfo.Show
End Sub

Private Sub Text1_Change()
CmdSa_Click
End Sub

Private Sub Timer1_Timer()

Randomize 행컨
Randomize 행능

End Sub

Private Sub Timer10_Timer()
Con = 100
Dim 우히히히히히히
For 우히히히히히히 = 0 To 800
 컨디션(우히히히히히히) = 100
Next 우히히히히히히

End Sub

Private Sub Timer11_Timer()
lblMoney = "Money : " & Money & "Cro"
If MyRank(1) = "Normal" Then
 lblYe1.ForeColor = RGB(0, 0, 0)
ElseIf MyRank(1) = "Special" Then
 lblYe1.ForeColor = RGB(0, 255, 0)
ElseIf MyRank(1) = "Rare" Then
 lblYe1.ForeColor = &HFF80FF
ElseIf MyRank(1) = "Unique" Then
 lblYe1.ForeColor = &HFF8080
ElseIf MyRank(1) = "Elite" Then
 lblYe1.ForeColor = &H800080
ElseIf MyRank(1) = "Legend" Then
 lblYe1.ForeColor = &H80FF&
ElseIf MyRank(1) = "Secret" Then
 lblYe1.ForeColor = &HFFC0C0
ElseIf MyRank(1) = "Champion" Then
 lblYe1.ForeColor = RGB(255, 0, 0)
End If

If MyRank(2) = "Normal" Then
 lblYe2.ForeColor = RGB(0, 0, 0)
ElseIf MyRank(2) = "Special" Then
 lblYe2.ForeColor = RGB(0, 255, 0)
ElseIf MyRank(2) = "Rare" Then
 lblYe2.ForeColor = &HFF80FF
ElseIf MyRank(2) = "Unique" Then
 lblYe2.ForeColor = &HFF8080
ElseIf MyRank(2) = "Elite" Then
 lblYe2.ForeColor = &H800080
ElseIf MyRank(2) = "Legend" Then
 lblYe2.ForeColor = &H80FF&
ElseIf MyRank(2) = "Secret" Then
 lblYe2.ForeColor = &HFFC0C0
ElseIf MyRank(2) = "Champion" Then
 lblYe2.ForeColor = RGB(255, 0, 0)
End If

If MyRank(3) = "Normal" Then
 lblYe3.ForeColor = RGB(0, 0, 0)
ElseIf MyRank(3) = "Special" Then
 lblYe3.ForeColor = RGB(0, 255, 0)
ElseIf MyRank(3) = "Rare" Then
 lblYe3.ForeColor = &HFF80FF
ElseIf MyRank(3) = "Unique" Then
 lblYe3.ForeColor = &HFF8080
ElseIf MyRank(3) = "Elite" Then
 lblYe3.ForeColor = &H800080
ElseIf MyRank(3) = "Legend" Then
 lblYe3.ForeColor = &H80FF&
ElseIf MyRank(3) = "Secret" Then
 lblYe3.ForeColor = &HFFC0C0
ElseIf MyRank(3) = "Champion" Then
 lblYe3.ForeColor = RGB(255, 0, 0)
End If

If MyRank(4) = "Normal" Then
 lblYe4.ForeColor = RGB(0, 0, 0)
ElseIf MyRank(4) = "Special" Then
 lblYe4.ForeColor = RGB(0, 255, 0)
ElseIf MyRank(4) = "Rare" Then
 lblYe4.ForeColor = &HFF80FF
ElseIf MyRank(4) = "Unique" Then
 lblYe4.ForeColor = &HFF8080
ElseIf MyRank(4) = "Elite" Then
 lblYe4.ForeColor = &H800080
ElseIf MyRank(4) = "Legend" Then
 lblYe4.ForeColor = &H80FF&
ElseIf MyRank(4) = "Secret" Then
 lblYe4.ForeColor = &HFFC0C0
ElseIf MyRank(4) = "Champion" Then
 lblYe4.ForeColor = RGB(255, 0, 0)
End If

If MyRank(5) = "Normal" Then
 lblYe5.ForeColor = RGB(0, 0, 0)
ElseIf MyRank(5) = "Special" Then
 lblYe5.ForeColor = RGB(0, 255, 0)
ElseIf MyRank(5) = "Rare" Then
 lblYe5.ForeColor = &HFF80FF
ElseIf MyRank(5) = "Unique" Then
 lblYe5.ForeColor = &HFF8080
ElseIf MyRank(5) = "Elite" Then
 lblYe5.ForeColor = &H800080
ElseIf MyRank(5) = "Legend" Then
 lblYe5.ForeColor = &H80FF&
ElseIf MyRank(5) = "Secret" Then
 lblYe5.ForeColor = &HFFC0C0
ElseIf MyRank(5) = "Champion" Then
 lblYe5.ForeColor = RGB(255, 0, 0)
End If

If MyRank(6) = "Normal" Then
 lblYe6.ForeColor = RGB(0, 0, 0)
ElseIf MyRank(6) = "Special" Then
 lblYe6.ForeColor = RGB(0, 266, 0)
ElseIf MyRank(6) = "Rare" Then
 lblYe6.ForeColor = &HFF80FF
ElseIf MyRank(6) = "Unique" Then
 lblYe6.ForeColor = &HFF8080
ElseIf MyRank(6) = "Elite" Then
 lblYe6.ForeColor = &H800080
ElseIf MyRank(6) = "Legend" Then
 lblYe6.ForeColor = &H80FF&
ElseIf MyRank(6) = "Secret" Then
 lblYe6.ForeColor = &HFFC0C0
ElseIf MyRank(6) = "Champion" Then
 lblYe6.ForeColor = RGB(255, 0, 0)
End If
End Sub

Private Sub Timer12_Timer()



If SubRank(1) = "Normal" Then
 lblSy1.ForeColor = RGB(0, 0, 0)
ElseIf SubRank(1) = "Special" Then
 lblSy1.ForeColor = RGB(0, 255, 0)
ElseIf SubRank(1) = "Rare" Then
 lblSy1.ForeColor = &HFF80FF
ElseIf SubRank(1) = "Unique" Then
 lblSy1.ForeColor = &HFF8080
ElseIf SubRank(1) = "Elite" Then
 lblSy1.ForeColor = &H800080
ElseIf SubRank(1) = "Legend" Then
 lblSy1.ForeColor = &H80FF&
ElseIf SubRank(1) = "Secret" Then
 lblSy1.ForeColor = &HFFC0C0
ElseIf SubRank(1) = "Champion" Then
 lblSy1.ForeColor = RGB(255, 0, 0)
End If

If SubRank(2) = "Normal" Then
 lblSy2.ForeColor = RGB(0, 0, 0)
ElseIf SubRank(2) = "Special" Then
 lblSy2.ForeColor = RGB(0, 255, 0)
ElseIf SubRank(2) = "Rare" Then
 lblSy2.ForeColor = &HFF80FF
ElseIf SubRank(2) = "Unique" Then
 lblSy2.ForeColor = &HFF8080
ElseIf SubRank(2) = "Elite" Then
 lblSy2.ForeColor = &H800080
ElseIf SubRank(2) = "Legend" Then
 lblSy2.ForeColor = &H80FF&
ElseIf SubRank(2) = "Secret" Then
 lblSy2.ForeColor = &HFFC0C0
ElseIf SubRank(2) = "Champion" Then
 lblSy2.ForeColor = RGB(255, 0, 0)
End If

If SubRank(3) = "Normal" Then
 lblSy3.ForeColor = RGB(0, 0, 0)
ElseIf SubRank(3) = "Special" Then
 lblSy3.ForeColor = RGB(0, 255, 0)
ElseIf SubRank(3) = "Rare" Then
 lblSy3.ForeColor = &HFF80FF
ElseIf SubRank(3) = "Unique" Then
 lblSy3.ForeColor = &HFF8080
ElseIf SubRank(3) = "Elite" Then
 lblSy3.ForeColor = &H800080
ElseIf SubRank(3) = "Legend" Then
 lblSy3.ForeColor = &H80FF&
ElseIf SubRank(3) = "Secret" Then
 lblSy3.ForeColor = &HFFC0C0
ElseIf SubRank(3) = "Champion" Then
 lblSy3.ForeColor = RGB(255, 0, 0)
End If

If SubRank(4) = "Normal" Then
 lblSy4.ForeColor = RGB(0, 0, 0)
ElseIf SubRank(4) = "Special" Then
 lblSy4.ForeColor = RGB(0, 255, 0)
ElseIf SubRank(4) = "Rare" Then
 lblSy4.ForeColor = &HFF80FF
ElseIf SubRank(4) = "Unique" Then
 lblSy4.ForeColor = &HFF8080
ElseIf SubRank(4) = "Elite" Then
 lblSy4.ForeColor = &H800080
ElseIf SubRank(4) = "Legend" Then
 lblSy4.ForeColor = &H80FF&
ElseIf SubRank(4) = "Secret" Then
 lblSy4.ForeColor = &HFFC0C0
ElseIf SubRank(4) = "Champion" Then
 lblSy4.ForeColor = RGB(255, 0, 0)
End If

If SubRank(5) = "Normal" Then
 lblSy5.ForeColor = RGB(0, 0, 0)
ElseIf SubRank(5) = "Special" Then
 lblSy5.ForeColor = RGB(0, 255, 0)
ElseIf SubRank(5) = "Rare" Then
 lblSy5.ForeColor = &HFF80FF
ElseIf SubRank(5) = "Unique" Then
 lblSy5.ForeColor = &HFF8080
ElseIf SubRank(5) = "Elite" Then
 lblSy5.ForeColor = &H800080
ElseIf SubRank(5) = "Legend" Then
 lblSy5.ForeColor = &H80FF&
ElseIf SubRank(5) = "Secret" Then
 lblSy5.ForeColor = &HFFC0C0
ElseIf SubRank(5) = "Champion" Then
 lblSy5.ForeColor = RGB(255, 0, 0)
End If

If SubRank(6) = "Normal" Then
 lblSy6.ForeColor = RGB(0, 0, 0)
ElseIf SubRank(6) = "Special" Then
 lblSy6.ForeColor = RGB(0, 266, 0)
ElseIf SubRank(6) = "Rare" Then
 lblSy6.ForeColor = &HFF80FF
ElseIf SubRank(6) = "Unique" Then
 lblSy6.ForeColor = &HFF8080
ElseIf SubRank(6) = "Elite" Then
 lblSy6.ForeColor = &H800080
ElseIf SubRank(6) = "Legend" Then
 lblSy6.ForeColor = &H80FF&
ElseIf SubRank(6) = "Secret" Then
 lblSy6.ForeColor = &HFFC0C0
ElseIf SubRank(6) = "Champion" Then
 lblSy6.ForeColor = RGB(255, 0, 0)
End If

If SubRank(7) = "Normal" Then
 lblSy7.ForeColor = RGB(0, 0, 0)
ElseIf SubRank(7) = "Special" Then
 lblSy7.ForeColor = RGB(0, 277, 0)
ElseIf SubRank(7) = "Rare" Then
 lblSy7.ForeColor = &HFF80FF
ElseIf SubRank(7) = "Unique" Then
 lblSy7.ForeColor = &HFF8080
ElseIf SubRank(7) = "Elite" Then
 lblSy7.ForeColor = &H800080
ElseIf SubRank(7) = "Legend" Then
 lblSy7.ForeColor = &H80FF&
ElseIf SubRank(7) = "Secret" Then
 lblSy7.ForeColor = &HFFC0C0
ElseIf SubRank(7) = "Champion" Then
 lblSy7.ForeColor = RGB(255, 0, 0)
End If

If SubRank(8) = "Normal" Then
 lblSy8.ForeColor = RGB(0, 0, 0)
ElseIf SubRank(8) = "Special" Then
 lblSy8.ForeColor = RGB(0, 288, 0)
ElseIf SubRank(8) = "Rare" Then
 lblSy8.ForeColor = &HFF80FF
ElseIf SubRank(8) = "Unique" Then
 lblSy8.ForeColor = &HFF8080
ElseIf SubRank(8) = "Elite" Then
 lblSy8.ForeColor = &H800080
ElseIf SubRank(8) = "Legend" Then
 lblSy8.ForeColor = &H80FF&
ElseIf SubRank(8) = "Secret" Then
 lblSy8.ForeColor = &HFFC0C0
ElseIf SubRank(8) = "Champion" Then
 lblSy8.ForeColor = RGB(255, 0, 0)
End If

If SubRank(9) = "Normal" Then
 lblSy9.ForeColor = RGB(0, 0, 0)
ElseIf SubRank(9) = "Special" Then
 lblSy9.ForeColor = RGB(0, 299, 0)
ElseIf SubRank(9) = "Rare" Then
 lblSy9.ForeColor = &HFF90FF
ElseIf SubRank(9) = "Unique" Then
 lblSy9.ForeColor = &HFF9090
ElseIf SubRank(9) = "Elite" Then
 lblSy9.ForeColor = &H900090
ElseIf SubRank(9) = "Legend" Then
 lblSy9.ForeColor = &H90FF&
ElseIf SubRank(9) = "Secret" Then
 lblSy9.ForeColor = &HFFC0C0
ElseIf SubRank(9) = "Champion" Then
 lblSy9.ForeColor = RGB(255, 0, 0)
End If

If val(선수수) = 7 Then
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")) <> 0 Then
  ImgS1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")
 Else
  ImgS1 = LoadPicture(App.Path & "\img\선수\" & SubName(1) & ".gif")
 End If
 lblSy1 = SubYear(1)
 lblSN1 = SubName(1)
ElseIf val(선수수) = 8 Then
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")) <> 0 Then
  ImgS1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")
 Else
  ImgS1 = LoadPicture(App.Path & "\img\선수\" & SubName(1) & ".gif")
 End If
 
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(2), 2, 2) & "]" & SubName(2) & ".gif")) <> 0 Then
  ImgS2 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(2), 2, 2) & "]" & SubName(2) & ".gif")
 Else
  ImgS2 = LoadPicture(App.Path & "\img\선수\" & SubName(2) & ".gif")
 End If
 lblSy1 = SubYear(1)
 lblSN1 = SubName(1)
 lblSy2 = SubYear(2)
 lblSN2 = SubName(2)
ElseIf val(선수수) = 9 Then
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")) <> 0 Then
  ImgS1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")
 Else
  ImgS1 = LoadPicture(App.Path & "\img\선수\" & SubName(1) & ".gif")
 End If
 
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(2), 2, 2) & "]" & SubName(2) & ".gif")) <> 0 Then
  ImgS2 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(2), 2, 2) & "]" & SubName(2) & ".gif")
 Else
  ImgS2 = LoadPicture(App.Path & "\img\선수\" & SubName(2) & ".gif")
 End If
 
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(3), 2, 2) & "]" & SubName(3) & ".gif")) <> 0 Then
  ImgS3 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(3), 2, 2) & "]" & SubName(3) & ".gif")
 Else
  ImgS3 = LoadPicture(App.Path & "\img\선수\" & SubName(3) & ".gif")
 End If
 lblSy1 = SubYear(1)
 lblSN1 = SubName(1)
 lblSy2 = SubYear(2)
 lblSN2 = SubName(2)
 lblSy3 = SubYear(3)
 lblSN3 = SubName(3)
ElseIf val(선수수) = 10 Then
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")) <> 0 Then
  ImgS1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")
 Else
  ImgS1 = LoadPicture(App.Path & "\img\선수\" & SubName(1) & ".gif")
 End If
 
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(2), 2, 2) & "]" & SubName(2) & ".gif")) <> 0 Then
  ImgS2 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(2), 2, 2) & "]" & SubName(2) & ".gif")
 Else
  ImgS2 = LoadPicture(App.Path & "\img\선수\" & SubName(2) & ".gif")
 End If
 
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(3), 2, 2) & "]" & SubName(3) & ".gif")) <> 0 Then
  ImgS3 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(3), 2, 2) & "]" & SubName(3) & ".gif")
 Else
  ImgS3 = LoadPicture(App.Path & "\img\선수\" & SubName(3) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(4), 2, 2) & "]" & SubName(4) & ".gif")) <> 0 Then
  ImgS4 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(4), 2, 2) & "]" & SubName(4) & ".gif")
 Else
  ImgS4 = LoadPicture(App.Path & "\img\선수\" & SubName(4) & ".gif")
 End If
 lblSy1 = SubYear(1)
 lblSN1 = SubName(1)
 lblSy2 = SubYear(2)
 lblSN2 = SubName(2)
 lblSy3 = SubYear(3)
 lblSN3 = SubName(3)
 lblSy4 = SubYear(4)
 lblSN4 = SubName(4)
ElseIf val(선수수) = 11 Then
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")) <> 0 Then
  ImgS1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")
 Else
  ImgS1 = LoadPicture(App.Path & "\img\선수\" & SubName(1) & ".gif")
 End If
 
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(2), 2, 2) & "]" & SubName(2) & ".gif")) <> 0 Then
  ImgS2 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(2), 2, 2) & "]" & SubName(2) & ".gif")
 Else
  ImgS2 = LoadPicture(App.Path & "\img\선수\" & SubName(2) & ".gif")
 End If
 
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(3), 2, 2) & "]" & SubName(3) & ".gif")) <> 0 Then
  ImgS3 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(3), 2, 2) & "]" & SubName(3) & ".gif")
 Else
  ImgS3 = LoadPicture(App.Path & "\img\선수\" & SubName(3) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(4), 2, 2) & "]" & SubName(4) & ".gif")) <> 0 Then
  ImgS4 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(4), 2, 2) & "]" & SubName(4) & ".gif")
 Else
  ImgS4 = LoadPicture(App.Path & "\img\선수\" & SubName(4) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(5), 2, 2) & "]" & SubName(5) & ".gif")) <> 0 Then
  ImgS5 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(5), 2, 2) & "]" & SubName(5) & ".gif")
 Else
  ImgS5 = LoadPicture(App.Path & "\img\선수\" & SubName(5) & ".gif")
 End If
 lblSy1 = SubYear(1)
 lblSN1 = SubName(1)
 lblSy2 = SubYear(2)
 lblSN2 = SubName(2)
 lblSy3 = SubYear(3)
 lblSN3 = SubName(3)
 lblSy4 = SubYear(4)
 lblSN4 = SubName(4)
 lblSy5 = SubYear(5)
 lblSN5 = SubName(5)
ElseIf val(선수수) = 12 Then
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")) <> 0 Then
  ImgS1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")
 Else
  ImgS1 = LoadPicture(App.Path & "\img\선수\" & SubName(1) & ".gif")
 End If
 
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(2), 2, 2) & "]" & SubName(2) & ".gif")) <> 0 Then
  ImgS2 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(2), 2, 2) & "]" & SubName(2) & ".gif")
 Else
  ImgS2 = LoadPicture(App.Path & "\img\선수\" & SubName(2) & ".gif")
 End If
 
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(3), 2, 2) & "]" & SubName(3) & ".gif")) <> 0 Then
  ImgS3 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(3), 2, 2) & "]" & SubName(3) & ".gif")
 Else
  ImgS3 = LoadPicture(App.Path & "\img\선수\" & SubName(3) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(4), 2, 2) & "]" & SubName(4) & ".gif")) <> 0 Then
  ImgS4 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(4), 2, 2) & "]" & SubName(4) & ".gif")
 Else
  ImgS4 = LoadPicture(App.Path & "\img\선수\" & SubName(4) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(5), 2, 2) & "]" & SubName(5) & ".gif")) <> 0 Then
  ImgS5 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(5), 2, 2) & "]" & SubName(5) & ".gif")
 Else
  ImgS5 = LoadPicture(App.Path & "\img\선수\" & SubName(5) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(6), 2, 2) & "]" & SubName(6) & ".gif")) <> 0 Then
  ImgS6 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(6), 2, 2) & "]" & SubName(6) & ".gif")
 Else
  ImgS6 = LoadPicture(App.Path & "\img\선수\" & SubName(6) & ".gif")
 End If
 lblSy1 = SubYear(1)
 lblSN1 = SubName(1)
 lblSy2 = SubYear(2)
 lblSN2 = SubName(2)
 lblSy3 = SubYear(3)
 lblSN3 = SubName(3)
 lblSy4 = SubYear(4)
 lblSN4 = SubName(4)
 lblSy5 = SubYear(5)
 lblSN5 = SubName(5)
 lblSy6 = SubYear(6)
 lblSN6 = SubName(6)
ElseIf val(선수수) = 13 Then
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")) <> 0 Then
  ImgS1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")
 Else
  ImgS1 = LoadPicture(App.Path & "\img\선수\" & SubName(1) & ".gif")
 End If
 
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(2), 2, 2) & "]" & SubName(2) & ".gif")) <> 0 Then
  ImgS2 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(2), 2, 2) & "]" & SubName(2) & ".gif")
 Else
  ImgS2 = LoadPicture(App.Path & "\img\선수\" & SubName(2) & ".gif")
 End If
 
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(3), 2, 2) & "]" & SubName(3) & ".gif")) <> 0 Then
  ImgS3 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(3), 2, 2) & "]" & SubName(3) & ".gif")
 Else
  ImgS3 = LoadPicture(App.Path & "\img\선수\" & SubName(3) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(4), 2, 2) & "]" & SubName(4) & ".gif")) <> 0 Then
  ImgS4 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(4), 2, 2) & "]" & SubName(4) & ".gif")
 Else
  ImgS4 = LoadPicture(App.Path & "\img\선수\" & SubName(4) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(5), 2, 2) & "]" & SubName(5) & ".gif")) <> 0 Then
  ImgS5 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(5), 2, 2) & "]" & SubName(5) & ".gif")
 Else
  ImgS5 = LoadPicture(App.Path & "\img\선수\" & SubName(5) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(6), 2, 2) & "]" & SubName(6) & ".gif")) <> 0 Then
  ImgS6 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(6), 2, 2) & "]" & SubName(6) & ".gif")
 Else
  ImgS6 = LoadPicture(App.Path & "\img\선수\" & SubName(6) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(7), 2, 2) & "]" & SubName(7) & ".gif")) <> 0 Then
  ImgS7 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(7), 2, 2) & "]" & SubName(7) & ".gif")
 Else
  ImgS7 = LoadPicture(App.Path & "\img\선수\" & SubName(7) & ".gif")
 End If
 lblSy1 = SubYear(1)
 lblSN1 = SubName(1)
 lblSy2 = SubYear(2)
 lblSN2 = SubName(2)
 lblSy3 = SubYear(3)
 lblSN3 = SubName(3)
 lblSy4 = SubYear(4)
 lblSN4 = SubName(4)
 lblSy5 = SubYear(5)
 lblSN5 = SubName(5)
 lblSy6 = SubYear(6)
 lblSN6 = SubName(6)
 lblSy7 = SubYear(7)
 lblSN7 = SubName(7)
ElseIf val(선수수) = 14 Then
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")) <> 0 Then
  ImgS1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")
 Else
  ImgS1 = LoadPicture(App.Path & "\img\선수\" & SubName(1) & ".gif")
 End If
 
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(2), 2, 2) & "]" & SubName(2) & ".gif")) <> 0 Then
  ImgS2 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(2), 2, 2) & "]" & SubName(2) & ".gif")
 Else
  ImgS2 = LoadPicture(App.Path & "\img\선수\" & SubName(2) & ".gif")
 End If
 
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(3), 2, 2) & "]" & SubName(3) & ".gif")) <> 0 Then
  ImgS3 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(3), 2, 2) & "]" & SubName(3) & ".gif")
 Else
  ImgS3 = LoadPicture(App.Path & "\img\선수\" & SubName(3) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(4), 2, 2) & "]" & SubName(4) & ".gif")) <> 0 Then
  ImgS4 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(4), 2, 2) & "]" & SubName(4) & ".gif")
 Else
  ImgS4 = LoadPicture(App.Path & "\img\선수\" & SubName(4) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(5), 2, 2) & "]" & SubName(5) & ".gif")) <> 0 Then
  ImgS5 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(5), 2, 2) & "]" & SubName(5) & ".gif")
 Else
  ImgS5 = LoadPicture(App.Path & "\img\선수\" & SubName(5) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(6), 2, 2) & "]" & SubName(6) & ".gif")) <> 0 Then
  ImgS6 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(6), 2, 2) & "]" & SubName(6) & ".gif")
 Else
  ImgS6 = LoadPicture(App.Path & "\img\선수\" & SubName(6) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(7), 2, 2) & "]" & SubName(7) & ".gif")) <> 0 Then
  ImgS7 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(7), 2, 2) & "]" & SubName(7) & ".gif")
 Else
  ImgS7 = LoadPicture(App.Path & "\img\선수\" & SubName(7) & ".gif")
 End If
 
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(8), 2, 2) & "]" & SubName(8) & ".gif")) <> 0 Then
  ImgS8 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(8), 2, 2) & "]" & SubName(8) & ".gif")
 Else
  ImgS8 = LoadPicture(App.Path & "\img\선수\" & SubName(8) & ".gif")
 End If
 lblSy1 = SubYear(1)
 lblSN1 = SubName(1)
 lblSy2 = SubYear(2)
 lblSN2 = SubName(2)
 lblSy3 = SubYear(3)
 lblSN3 = SubName(3)
 lblSy4 = SubYear(4)
 lblSN4 = SubName(4)
 lblSy5 = SubYear(5)
 lblSN5 = SubName(5)
 lblSy6 = SubYear(6)
 lblSN6 = SubName(6)
 lblSy7 = SubYear(7)
 lblSN7 = SubName(7)
 lblSy8 = SubYear(8)
 lblSN8 = SubName(8)
ElseIf val(선수수) = 15 Then
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")) <> 0 Then
  ImgS1 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(1), 2, 2) & "]" & SubName(1) & ".gif")
 Else
  ImgS1 = LoadPicture(App.Path & "\img\선수\" & SubName(1) & ".gif")
 End If
 
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(2), 2, 2) & "]" & SubName(2) & ".gif")) <> 0 Then
  ImgS2 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(2), 2, 2) & "]" & SubName(2) & ".gif")
 Else
  ImgS2 = LoadPicture(App.Path & "\img\선수\" & SubName(2) & ".gif")
 End If
 
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(3), 2, 2) & "]" & SubName(3) & ".gif")) <> 0 Then
  ImgS3 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(3), 2, 2) & "]" & SubName(3) & ".gif")
 Else
  ImgS3 = LoadPicture(App.Path & "\img\선수\" & SubName(3) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(4), 2, 2) & "]" & SubName(4) & ".gif")) <> 0 Then
  ImgS4 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(4), 2, 2) & "]" & SubName(4) & ".gif")
 Else
  ImgS4 = LoadPicture(App.Path & "\img\선수\" & SubName(4) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(5), 2, 2) & "]" & SubName(5) & ".gif")) <> 0 Then
  ImgS5 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(5), 2, 2) & "]" & SubName(5) & ".gif")
 Else
  ImgS5 = LoadPicture(App.Path & "\img\선수\" & SubName(5) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(6), 2, 2) & "]" & SubName(6) & ".gif")) <> 0 Then
  ImgS6 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(6), 2, 2) & "]" & SubName(6) & ".gif")
 Else
  ImgS6 = LoadPicture(App.Path & "\img\선수\" & SubName(6) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(7), 2, 2) & "]" & SubName(7) & ".gif")) <> 0 Then
  ImgS7 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(7), 2, 2) & "]" & SubName(7) & ".gif")
 Else
  ImgS7 = LoadPicture(App.Path & "\img\선수\" & SubName(7) & ".gif")
 End If
 
 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(8), 2, 2) & "]" & SubName(8) & ".gif")) <> 0 Then
  ImgS8 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(8), 2, 2) & "]" & SubName(8) & ".gif")
 Else
  ImgS8 = LoadPicture(App.Path & "\img\선수\" & SubName(8) & ".gif")
 End If

 If Len(Dir(App.Path & "\img\선수\[" & Mid(SubYear(9), 2, 2) & "]" & SubName(9) & ".gif")) <> 0 Then
  ImgS9 = LoadPicture(App.Path & "\img\선수\[" & Mid(SubYear(9), 2, 2) & "]" & SubName(9) & ".gif")
 Else
  ImgS9 = LoadPicture(App.Path & "\img\선수\" & SubName(9) & ".gif")
 End If



 lblSy1 = SubYear(1)
 lblSN1 = SubName(1)
 lblSy2 = SubYear(2)
 lblSN2 = SubName(2)
 lblSy3 = SubYear(3)
 lblSN3 = SubName(3)
 lblSy4 = SubYear(4)
 lblSN4 = SubName(4)
 lblSy5 = SubYear(5)
 lblSN5 = SubName(5)
 lblSy6 = SubYear(6)
 lblSN6 = SubName(6)
 lblSy7 = SubYear(7)
 lblSN7 = SubName(7)
 lblSy8 = SubYear(8)
 lblSN8 = SubName(8)
 lblSy9 = SubYear(9)
 lblSN9 = SubName(9)
End If
Timer12.Enabled = False
End Sub

Private Sub Timer13_Timer()
Timer12.Enabled = True
lblYe1 = MyYear(1)
lblYe2 = MyYear(2)
lblYe3 = MyYear(3)
lblYe4 = MyYear(4)
lblYe5 = MyYear(5)
lblYe6 = MyYear(6)
lblName1 = MyName(1)
lblName2 = MyName(2)
lblName3 = MyName(3)
lblName4 = MyName(4)
lblName5 = MyName(5)
lblName6 = MyName(6)
If Len(Dir(App.Path & "\img\선수\[" & Mid(MyYear(1), 2, 2) & "]" & MyName(1) & ".gif")) <> 0 Then
 ImgP1 = LoadPicture(App.Path & "\img\선수\[" & Mid(MyYear(1), 2, 2) & "]" & MyName(1) & ".gif")
Else
 ImgP1 = LoadPicture(App.Path & "\img\선수\" & MyName(1) & ".gif")
End If

If Len(Dir(App.Path & "\img\선수\[" & Mid(MyYear(2), 2, 2) & "]" & MyName(2) & ".gif")) <> 0 Then
 ImgP2 = LoadPicture(App.Path & "\img\선수\[" & Mid(MyYear(2), 2, 2) & "]" & MyName(2) & ".gif")
Else
 ImgP2 = LoadPicture(App.Path & "\img\선수\" & MyName(2) & ".gif")
End If

If Len(Dir(App.Path & "\img\선수\[" & Mid(MyYear(3), 2, 2) & "]" & MyName(3) & ".gif")) <> 0 Then
 ImgP3 = LoadPicture(App.Path & "\img\선수\[" & Mid(MyYear(3), 2, 2) & "]" & MyName(3) & ".gif")
Else
 ImgP3 = LoadPicture(App.Path & "\img\선수\" & MyName(3) & ".gif")
End If

If Len(Dir(App.Path & "\img\선수\[" & Mid(MyYear(4), 2, 2) & "]" & MyName(4) & ".gif")) <> 0 Then
 ImgP4 = LoadPicture(App.Path & "\img\선수\[" & Mid(MyYear(4), 2, 2) & "]" & MyName(4) & ".gif")
Else
 ImgP4 = LoadPicture(App.Path & "\img\선수\" & MyName(4) & ".gif")
End If

If Len(Dir(App.Path & "\img\선수\[" & Mid(MyYear(5), 2, 2) & "]" & MyName(5) & ".gif")) <> 0 Then
 ImgP5 = LoadPicture(App.Path & "\img\선수\[" & Mid(MyYear(5), 2, 2) & "]" & MyName(5) & ".gif")
Else
 ImgP5 = LoadPicture(App.Path & "\img\선수\" & MyName(5) & ".gif")
End If

If Len(Dir(App.Path & "\img\선수\[" & Mid(MyYear(6), 2, 2) & "]" & MyName(6) & ".gif")) <> 0 Then
 ImgP6 = LoadPicture(App.Path & "\img\선수\[" & Mid(MyYear(6), 2, 2) & "]" & MyName(6) & ".gif")
Else
 ImgP6 = LoadPicture(App.Path & "\img\선수\" & MyName(6) & ".gif")
End If
Timer13.Enabled = False
End Sub

Private Sub Timer14_Timer()
If val(선수수) = 14 Then
ImgS9.Picture = Nothing
lblSN9 = ""
lblSy9 = ""
ElseIf val(선수수) = 13 Then
ImgS8.Picture = Nothing
lblSN8 = ""
lblSy8 = ""
ElseIf val(선수수) = 12 Then
ImgS7.Picture = Nothing
lblSN7 = ""
lblSy7 = ""
ElseIf val(선수수) = 11 Then
ImgS6.Picture = Nothing
lblSN6 = ""
lblSy6 = ""
ElseIf val(선수수) = 10 Then
ImgS5.Picture = Nothing
lblSN5 = ""
lblSy5 = ""
ElseIf val(선수수) = 9 Then
ImgS4.Picture = Nothing
lblSN4 = ""
lblSy4 = ""
ElseIf val(선수수) = 8 Then
ImgS3.Picture = Nothing
lblSN3 = ""
lblSy3 = ""
ElseIf val(선수수) = 7 Then
ImgS2.Picture = Nothing
lblSN2 = ""
lblSy2 = ""
ElseIf val(선수수) = 6 Then
ImgS1.Picture = Nothing
lblSN1 = ""
lblSy1 = ""
End If
End Sub

Private Sub Timer15_Timer()
For 확인 = 0 To 800
    총스탯 = val(공격력(확인)) + val(견제(확인)) + val(전략(확인)) + val(물량(확인)) + val(수비력(확인)) + val(정찰(확인)) + val(센스(확인)) + val(컨트롤(확인))
    If val(총스탯) >= 8500 Then
        공격력(확인) = val(공격력(확인)) - 100
        견제(확인) = val(견제(확인)) - 100
        전략(확인) = val(전략(확인)) - 100
        물량(확인) = val(물량(확인)) - 100
        수비력(확인) = val(수비력(확인)) - 100
        정찰(확인) = val(정찰(확인)) - 100
        센스(확인) = val(센스(확인)) - 100
        컨트롤(확인) = val(컨트롤(확인)) - 100
        MsgBox "상대" & OYear(확인) & 이름(확인) & " 총스탯 100씩 하향"
        하향 = val(하향) + 1
        하향횟수 = val(하향횟수) + 1
        MsgBox "하향횟수 : " & 하향
        If val(하향횟수) >= 5 Then
            쿠폰 = val(쿠폰) + 1
            하향횟수 = val(하향횟수) - 5
            MsgBox "쿠폰 + 1"
        End If
    End If
Next

For 확인 = 1 To 6
    총스탯 = val(MyAt(확인)) + val(MyR(확인)) + val(MySt(확인)) + val(MyAm(확인)) + val(MyDe(확인)) + val(MyPa(확인)) + val(MySe(확인)) + val(MyCo(확인))
    If val(총스탯) >= 8500 Then
        MyAt(확인) = val(MyAt(확인)) - 100
        MyR(확인) = val(MyR(확인)) - 100
        MySt(확인) = val(MySt(확인)) - 100
        MyAm(확인) = val(MyAm(확인)) - 100
        MyDe(확인) = val(MyDe(확인)) - 100
        MyPa(확인) = val(MyPa(확인)) - 100
        MySe(확인) = val(MySe(확인)) - 100
        MyCo(확인) = val(MyCo(확인)) - 100
        MsgBox MyYear(확인) & MyName(확인) & "총스탯 100씩 하향"
        MsgBox "쿠폰 + 1"
        쿠폰 = val(쿠폰) + 1
    End If
Next
End Sub

Private Sub Timer2_Timer()

행컨회 = Int((3 * Rnd) + 3)
행능 = Int((3 * Rnd) + 1)
행컨훈 = Int((3 * Rnd) + 0)

Oee = Int((801 * Rnd) + 0)
Oee = Int((801 * Rnd) + 0)

For i = 1 To 7
    Randomize MapL(i)
    MapL(i) = Int((12 * Rnd) + 1)
Next

 상점NPC = Int((800 * Rnd) + 1)

CmdGO.Visible = True
End Sub

Private Sub Timer3_Timer()
If PL경기수 >= 12 Then
    PL경기수 = 0
End If
  '****맵 뽑기
  Randomize Map
 Map = Int((12 * Rnd) + 1)
Randomize Map
Randomize 팬미팅
Randomize 로또

팬미팅 = Int((5000 * Rnd) + 500)
로또 = Int((90000 * Rnd) + 10000)

End Sub


Private Sub Timer4_Timer()
RandomAbility = Int((3 * Rnd) + 1) - 2
AP = val((101 * Rnd) + 0)
AP = val((101 * Rnd) + 0)
lblMoney = "Money : " & Money & "Cro"
lblTurn = "Turn : " & Turn
lblPL = "프로리그 : " & PL진행 & " " & PL승 & "승 " & PL패 & "패"
Label3 = "프로리그 우승 : " & PL우승 & "준우승 : " & PL준우승
lbl쿠폰 = "쿠폰 : " & 쿠폰
lbl하향 = "하향횟수 : " & 하향

If Deck <> "" Then
    If Deck년도 = False Then
        lblDeck = Deck & " DECK"
    Else
        lblDeck = MyYear(1) & Deck & " DECK"
    End If
Else
    lblDeck = "No DECK"
End If
End Sub


Private Sub Timer6_Timer()
Dim 마이 As Integer
For 마이 = 1 To 6
MyMExp(마이) = val(MyLev(마이)) * 50
If val(MyExp(마이)) <= 0 Then
 MyExp(마이) = 0
End If
Next 마이

For 마이 = 1 To 6
    If val(MyExp(마이)) >= val(MyMExp(마이)) Then
        MsgBox MyName(마이) & ", 레벨업!"
        MyExp(마이) = val(MyExp(마이)) - val(MyMExp(마이))
        MyLev(마이) = val(MyLev(마이)) + 1
        MyPoint(마이) = val(MyPoint(마이)) + 10
        MyMExp(마이) = val(MyLev(마이)) * 50
    End If
Next

End Sub


Private Sub Timer9_Timer()
잇힝 = Int((100 * Rnd) + 1) - 50
If val(Money) <= 0 Then
 Money = 0
End If
End Sub


'If 불러옴 = True Then
'    Call Save
'Else
'End If
'frmmain.Timer12.Enabled = True
