VERSION 5.00
Begin VB.Form Form28 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "Cafe. Kaiknight Checking Page."
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4845
   Icon            =   "Form28.frx":0000
   LinkTopic       =   "Form28"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   4845
   StartUpPosition =   2  '화면 가운데
   Begin VB.ListBox List1 
      Height          =   960
      Left            =   1320
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin CSO.jcbutton jcbutton2 
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Sign In"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin CSO.jcbutton jcbutton1 
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      Caption         =   "Start CSO."
      ForeColor       =   255
      ForeColorHover  =   33023
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.TextBox TxtPW 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Naver_PassWord"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox TxtID 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Text            =   "CSO_ID"
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "CSO_PassWord"
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
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "CSO_ID"
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim Regi1 As String, Regi2 As String
Regi1 = GetSetting("CSO", "CSO", "CSOID")
Regi2 = GetSetting("CSO", "CSO", "CSOPW")
TxtID = Regi1
TxtPW = Regi2

dB.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & _
                          "SERVER=localhost;DATABASE=SourceDB;" & _
                          "UID=root;PWD=chl123; OPTION=16427;" & _
                          "STMT= set names euckr"
dB.ConnectionTimeout = 30
dB.Mode = adModeReadWrite
dB.Open

Dim sql As String
sql = "select * from input" '테이블 입력
Dim rs
Set rs = New ADODB.Recordset '레코드 셋
rs.Open sql, dB
List1.Clear
Do While Not rs.EOF '목록의 끝까지 불러오기
    List1.AddItem (rs(1) & "," & rs(2) & vbCrLf)
    rs.MoveNext '데이터를 계속 밑으로 내려가며 참조
Loop ' 조건 루프문




End Sub

Private Sub jcbutton1_Click()
On Error GoTo error:
i = 0
Do
    If Split(List1.List(i), ",")(0) = TxtID Then
        If Left(Split(List1.List(i), ",")(1), Len(Split(List1.List(i), ",")(1)) - 2) = TxtPW Then
            Form1.Show
            SaveSetting "CSO", "CSO", "CSOID", TxtID
            SaveSetting "CSO", "CSO", "CSOPW", TxtPW
            Unload Me
            Exit Do
        Else
            MsgBox "PassWord 오류입니다."
            Exit Do
        End If
    Else
        i = i + 1
    End If
Loop

Exit Sub
error:
MsgBox "존재하지 않는 ID입니다."
End Sub

Private Sub jcbutton2_Click()
Form29.Show
Unload Me
End Sub

Private Sub TxtID_Click()
TxtID = vbNullString
End Sub

Private Sub TxtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    jcbutton1_Click
End If
End Sub

Private Sub TxtPW_Click()
TxtPW = vbNullString
End Sub

Private Sub TxtPW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    jcbutton1_Click
End If
End Sub
