Attribute VB_Name = "Mod_Funcs"
Option Explicit
Public Function RandomOee(Hexa As Integer, PLHexaBer As Integer)
Hexa = Int((801 * Rnd) + 0)
If PLHexaBer = 0 Then
    Do Until Team(Hexa) = "삼성전자"
        Hexa = Int((801 * Rnd) + 0)
    Loop
ElseIf PLHexaBer = 1 Then
    Do Until Team(Hexa) = "eSTRO" Or Team(Hexa) = "공군"
        Hexa = Int((801 * Rnd) + 0)
    Loop
ElseIf PLHexaBer = 2 Then
    Do Until (Team(Hexa) = "MBC") Or (Team(Hexa) = "POS")
        Hexa = Int((801 * Rnd) + 0)
    Loop
ElseIf PLHexaBer = 3 Then
    Do Until (Team(Hexa) = "CJ") Or (Team(Hexa) = "GO")
        Hexa = Int((801 * Rnd) + 0)
    Loop
ElseIf PLHexaBer = 4 Then
    Do Until (Team(Hexa) = "온게임넷") Or (Team(Hexa) = "하이트")
        Hexa = Int((801 * Rnd) + 0)
    Loop
ElseIf PLHexaBer = 5 Then
    Do Until (Team(Hexa) = "STX")
        Hexa = Int((801 * Rnd) + 0)
    Loop
ElseIf PLHexaBer = 6 Then
    Do Until (Team(Hexa) = "르까프") Or (Team(Hexa) = "화승") Or (Team(Hexa) = "PLUS")
        Hexa = Int((801 * Rnd) + 0)
    Loop
ElseIf PLHexaBer = 7 Then
    Do Until Team(Hexa) = "Mystar" Or Team(Hexa) = "8th"
        Hexa = Int((801 * Rnd) + 0)
    Loop
ElseIf PLHexaBer = 8 Then
    Do Until (Team(Hexa) = "웅진") Or (Team(Hexa) = "한빛")
        Hexa = Int((801 * Rnd) + 0)
    Loop
ElseIf PLHexaBer = 9 Then
    Do Until (Team(Hexa) = "SK") Or (Team(Hexa) = "Orion") Or (Team(Hexa) = "IS") Or (Team(Hexa) = "4U")
        Hexa = Int((801 * Rnd) + 0)
    Loop
ElseIf PLHexaBer = 10 Then
    Do Until (Team(Hexa) = "KT") Or (Team(Hexa) = "KTF")
        Hexa = Int((801 * Rnd) + 0)
    Loop
ElseIf PLHexaBer = 11 Then
    Do Until (Team(Hexa) = "폭스") Or (Team(Hexa) = "Toona") Or (Team(Hexa) = "Pantech") Or (Team(Hexa) = "Curitel")
        Hexa = Int((801 * Rnd) + 0)
    Loop
ElseIf PLHexaBer = 12 Then
    Do Until (Team(Hexa) = "Mystar")
        Hexa = Int((801 * Rnd) + 0)
    Loop
End If
End Function
Public Function MakeLine(DrawForm As Form, Num As Integer, X As Integer, Y As Integer)
DrawForm.BackColor = RGB(0, 0, 0)
DrawForm.BackColor = RGB(0, 0, 1)
DrawForm.Line (X + 1100, Y)-(X + 550 * Sqr(2), Y + 550 * Sqr(2)), RGB(255, 255, 255)
DrawForm.Line (X + 550 * Sqr(2), Y + 550 * Sqr(2))-(X, Y + 1100), RGB(255, 255, 255)
DrawForm.Line (X, Y + 1100)-(X - 550 * Sqr(2), Y + 550 * Sqr(2)), RGB(255, 255, 255)
DrawForm.Line (X - 550 * Sqr(2), Y + 550 * Sqr(2))-(X - 1100, Y), RGB(255, 255, 255)
DrawForm.Line (X - 1100, Y)-(X - 550 * Sqr(2), Y - 550 * Sqr(2)), RGB(255, 255, 255)
DrawForm.Line (X - 550 * Sqr(2), Y - 550 * Sqr(2))-(X, Y - 1100), RGB(255, 255, 255)
DrawForm.Line (X, Y - 1100)-(X + 550 * Sqr(2), Y - 550 * Sqr(2)), RGB(255, 255, 255)
DrawForm.Line (X + 550 * Sqr(2), Y - 550 * Sqr(2))-(X + 1100, Y), RGB(255, 255, 255)

DrawForm.Line (X + MyAt(Num), Y)-(X + (MyR(Num) / 2) * Sqr(2), Y + (MyR(Num) / 2) * Sqr(2)), RGB(255, 0, 0)
DrawForm.Line (X + (MyR(Num) / 2) * Sqr(2), Y + (MyR(Num) / 2) * Sqr(2))-(X, Y + MySt(Num)), RGB(255, 0, 0)
DrawForm.Line (X, Y + MySt(Num))-(X - (MyAm(Num) / 2) * Sqr(2), Y + (MyAm(Num) / 2) * Sqr(2)), RGB(255, 0, 0)
DrawForm.Line (X - (MyAm(Num) / 2) * Sqr(2), Y + (MyAm(Num) / 2) * Sqr(2))-(X - MyDe(Num), Y), RGB(255, 0, 0)
DrawForm.Line (X - MyDe(Num), Y)-(X - (MyPa(Num) / 2) * Sqr(2), Y - (MyPa(Num) / 2) * Sqr(2)), RGB(255, 0, 0)
DrawForm.Line (X - (MyPa(Num) / 2) * Sqr(2), Y - (MyPa(Num) / 2) * Sqr(2))-(X, Y - MySe(Num)), RGB(255, 0, 0)
DrawForm.Line (X, Y - MySe(Num))-(X + (MyCo(Num) / 2) * Sqr(2), Y - (MyCo(Num) / 2) * Sqr(2)), RGB(255, 0, 0)
DrawForm.Line (X + (MyCo(Num) / 2) * Sqr(2), Y - (MyCo(Num) / 2) * Sqr(2))-(X + MyAt(Num), Y), RGB(255, 0, 0)
End Function

Public Function MakeLineCom(DrawForm As Form, Num As Integer, X As Integer, Y As Integer)
DrawForm.Line (X + 1100, Y)-(X + 550 * Sqr(2), Y + 550 * Sqr(2)), RGB(255, 255, 255)
DrawForm.Line (X + 550 * Sqr(2), Y + 550 * Sqr(2))-(X, Y + 1100), RGB(255, 255, 255)
DrawForm.Line (X, Y + 1100)-(X - 550 * Sqr(2), Y + 550 * Sqr(2)), RGB(255, 255, 255)
DrawForm.Line (X - 550 * Sqr(2), Y + 550 * Sqr(2))-(X - 1100, Y), RGB(255, 255, 255)
DrawForm.Line (X - 1100, Y)-(X - 550 * Sqr(2), Y - 550 * Sqr(2)), RGB(255, 255, 255)
DrawForm.Line (X - 550 * Sqr(2), Y - 550 * Sqr(2))-(X, Y - 1100), RGB(255, 255, 255)
DrawForm.Line (X, Y - 1100)-(X + 550 * Sqr(2), Y - 550 * Sqr(2)), RGB(255, 255, 255)
DrawForm.Line (X + 550 * Sqr(2), Y - 550 * Sqr(2))-(X + 1100, Y), RGB(255, 255, 255)

DrawForm.Line (X + 공격력(Num), Y)-(X + (견제(Num) / 2) * Sqr(2), Y + (견제(Num) / 2) * Sqr(2)), RGB(255, 0, 0)
DrawForm.Line (X + (견제(Num) / 2) * Sqr(2), Y + (견제(Num) / 2) * Sqr(2))-(X, Y + 전략(Num)), RGB(255, 0, 0)
DrawForm.Line (X, Y + 전략(Num))-(X - (물량(Num) / 2) * Sqr(2), Y + (물량(Num) / 2) * Sqr(2)), RGB(255, 0, 0)
DrawForm.Line (X - (물량(Num) / 2) * Sqr(2), Y + (물량(Num) / 2) * Sqr(2))-(X - 수비력(Num), Y), RGB(255, 0, 0)
DrawForm.Line (X - 수비력(Num), Y)-(X - (정찰(Num) / 2) * Sqr(2), Y - (정찰(Num) / 2) * Sqr(2)), RGB(255, 0, 0)
DrawForm.Line (X - (정찰(Num) / 2) * Sqr(2), Y - (정찰(Num) / 2) * Sqr(2))-(X, Y - 센스(Num)), RGB(255, 0, 0)
DrawForm.Line (X, Y - 센스(Num))-(X + (컨트롤(Num) / 2) * Sqr(2), Y - (컨트롤(Num) / 2) * Sqr(2)), RGB(255, 0, 0)
DrawForm.Line (X + (컨트롤(Num) / 2) * Sqr(2), Y - (컨트롤(Num) / 2) * Sqr(2))-(X + 공격력(Num), Y), RGB(255, 0, 0)
End Function

Public Function LoadImage(Pics As Image, Hexa As String, Hexa2 As String)
'Hexa = 이름, Hexa2 = 년도
If Len(Dir(App.Path & "\img\선수\" & "[" & Mid(Hexa2, 2, 2) & "]" & Hexa & ".gif")) <> 0 Then
    Pics = LoadPicture(App.Path & "\img\선수\" & "[" & Mid(Hexa2, 2, 2) & "]" & Hexa & ".gif")
ElseIf Len(Dir(App.Path & "\img\선수\" & Hexa & ".gif")) <> 0 Then
    Pics = LoadPicture(App.Path & "\img\선수\" & Hexa & ".gif")
ElseIf Len(Dir(App.Path & "\img\선수\아마추어.gif")) <> 0 Then
    Pics = LoadPicture(App.Path & "\img\선수\아마추어.gif")
Else
    Pics = Nothing
End If
End Function

Public Function LoadMapImg(Pics As Image, Hexa As String)
If Len(Dir(App.Path & "\img\맵\" & Hexa & ".gif")) <> 0 Then
    Pics = LoadPicture(App.Path & "\img\맵\" & Hexa & ".gif")
Else
    Pics = Nothing
End If
End Function
Public Function lblTribeAlter(Labelgi As Label, Num As Integer)
If Num = 1 Then
    Labelgi = "(T)"
ElseIf Num = 2 Then
    Labelgi = "(Z)"
Else
    Labelgi = "(P)"
End If
End Function

Public Function lblNameAlter(Labelgi As Label, Num As Integer, Num2 As Integer)
If Num = 1 Then
    Labelgi = MyYear(Num2) & MyName(Num2)
    If MyRank(Num2) = "Normal" Then
        Labelgi.ForeColor = RGB(255, 255, 255)
    ElseIf MyRank(Num2) = "Special" Then
        Labelgi.ForeColor = RGB(0, 255, 0)
    ElseIf MyRank(Num2) = "Rare" Then
        Labelgi.ForeColor = &HFF80FF
    ElseIf MyRank(Num2) = "Unique" Then
        Labelgi.ForeColor = &HFF8080
    ElseIf MyRank(Num2) = "Elite" Then
        Labelgi.ForeColor = &H800080
    ElseIf MyRank(Num2) = "Legend" Then
        Labelgi.ForeColor = &H80FF&
    ElseIf MyRank(Num2) = "Secret" Then
        Labelgi.ForeColor = &HFFC0C0
    ElseIf MyRank(Num2) = "Champion" Then
        Labelgi.ForeColor = RGB(255, 0, 0)
    End If
Else
    Labelgi = OYear(Num2) & 이름(Num2)
    If 랭크(Num2) = "Normal" Then
        Labelgi.ForeColor = RGB(255, 255, 255)
    ElseIf 랭크(Num2) = "Special" Then
        Labelgi.ForeColor = RGB(0, 255, 0)
    ElseIf 랭크(Num2) = "Rare" Then
        Labelgi.ForeColor = &HFF80FF
    ElseIf 랭크(Num2) = "Unique" Then
        Labelgi.ForeColor = &HFF8080
    ElseIf 랭크(Num2) = "Elite" Then
        Labelgi.ForeColor = &H800080
    ElseIf 랭크(Num2) = "Legend" Then
        Labelgi.ForeColor = &H80FF&
    ElseIf 랭크(Num2) = "Secret" Then
        Labelgi.ForeColor = &HFFC0C0
    ElseIf 랭크(Num2) = "Champion" Then
        Labelgi.ForeColor = RGB(255, 0, 0)
    End If
End If
End Function
