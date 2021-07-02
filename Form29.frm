VERSION 5.00
Begin VB.Form Form29 
   BorderStyle     =   1  '단일 고정
   Caption         =   "CSO종료"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3765
   Icon            =   "Form29.frx":0000
   LinkTopic       =   "Form29"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3765
   StartUpPosition =   2  '화면 가운데
   Begin VB.ListBox List1 
      Height          =   1320
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin CSO.jcbutton jcbutton1 
      Height          =   1215
      Left            =   2040
      TabIndex        =   2
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2143
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "JiNie는 착하고 잘생기고 매너있고 훈훈하고 멋있다!! 를 인정합니다."
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.TextBox TxtPWD 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Text            =   "PassWord (최대 50글자)"
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox TxtID 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "ID (최대8글자)"
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dbsi As New ADODB.Connection
Public sql As String
'맨위에 선언 해줍니다 db는 데이터 연결 변수고 sql 은 쿼리문 작성할때 사용할 변수입니다.

Private Sub TxtID_Change()
If Len(TxtID) >= 15 Then
    TxtID = Left(TxtID, 15)
End If
End Sub

Private Sub TxtID_Click()
TxtID = ""
End Sub

Private Sub TxtPWD_Change()
If Len(TxtPWD) >= 20 Then
    TxtPWD = Left(TxtPWD, 20)
End If
End Sub

Private Sub TxtPWD_Click()
TxtPWD = ""
End Sub


Private Sub jcbutton1_Click()
Call sqlload(TxtID)
End Sub

Private Sub Form_Load()
dbsi.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & _
                          "SERVER=localhost;DATABASE=SourceDB;" & _
                          "UID=root;PWD=chl123; OPTION=16427;" & _
                          "STMT= set names euckr"
dbsi.ConnectionTimeout = 30
dbsi.Mode = adModeReadWrite
dbsi.Open
End Sub

 

'부가설명:
'*SERVER=localhost 는 localhost의 아이피 사용
'DATABASE=SourceDB 는 스키마의 이름
'UID=root 는 쿼리 아이디 이름
'PWD=root 는 쿼리 비밀번호 이름
'여기서 db.ConnectionString은 mysql 의 연결이고 ConnectionTimeout은 타임아웃 시간을 설정하는겁니다
'Mode는 데이터 사용 방식을 설정 하는것이고
'Open은 이제 mysql 를 사용한다는 뜻입니다.
'자이렇게 하면 mysql 연결은 완료입니다




Private Function sqlload(IDCheck As String) '쿼리로딩
sql = "select * from input" '테이블 입력
Dim rs
Set rs = New ADODB.Recordset '레코드 셋
rs.Open sql, dbsi
List1.Clear
Do While Not rs.EOF '목록의 끝까지 불러오기
    If IDCheck = rs(1) Then
        MsgBox "이미 동일한 아이디가 존재합니다."
        Exit Function
    Else
        List1.AddItem (rs(1) & vbCrLf)
        rs.MoveNext '데이터를 계속 밑으로 내려가며 참조
    End If
Loop ' 조건 루프문
sql = "INSERT INTO `input` VALUES (" & (List1.ListCount + 1) & ",'" & TxtID.Text & "','" & TxtPWD.Text & "',15);"
dbsi.Execute sql

MsgBox "생성이 완료되었습니다."
End Function
