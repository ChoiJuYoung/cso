VERSION 5.00
Begin VB.Form Form29 
   BorderStyle     =   1  '���� ����
   Caption         =   "CSO����"
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
   StartUpPosition =   2  'ȭ�� ���
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
      Caption         =   "JiNie�� ���ϰ� �߻���� �ų��ְ� �����ϰ� ���ִ�!! �� �����մϴ�."
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.TextBox TxtPWD 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Text            =   "PassWord (�ִ� 50����)"
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox TxtID 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "ID (�ִ�8����)"
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
'������ ���� ���ݴϴ� db�� ������ ���� ������ sql �� ������ �ۼ��Ҷ� ����� �����Դϴ�.

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

 

'�ΰ�����:
'*SERVER=localhost �� localhost�� ������ ���
'DATABASE=SourceDB �� ��Ű���� �̸�
'UID=root �� ���� ���̵� �̸�
'PWD=root �� ���� ��й�ȣ �̸�
'���⼭ db.ConnectionString�� mysql �� �����̰� ConnectionTimeout�� Ÿ�Ӿƿ� �ð��� �����ϴ°̴ϴ�
'Mode�� ������ ��� ����� ���� �ϴ°��̰�
'Open�� ���� mysql �� ����Ѵٴ� ���Դϴ�.
'���̷��� �ϸ� mysql ������ �Ϸ��Դϴ�




Private Function sqlload(IDCheck As String) '�����ε�
sql = "select * from input" '���̺� �Է�
Dim rs
Set rs = New ADODB.Recordset '���ڵ� ��
rs.Open sql, dbsi
List1.Clear
Do While Not rs.EOF '����� ������ �ҷ�����
    If IDCheck = rs(1) Then
        MsgBox "�̹� ������ ���̵� �����մϴ�."
        Exit Function
    Else
        List1.AddItem (rs(1) & vbCrLf)
        rs.MoveNext '�����͸� ��� ������ �������� ����
    End If
Loop ' ���� ������
sql = "INSERT INTO `input` VALUES (" & (List1.ListCount + 1) & ",'" & TxtID.Text & "','" & TxtPWD.Text & "',15);"
dbsi.Execute sql

MsgBox "������ �Ϸ�Ǿ����ϴ�."
End Function
