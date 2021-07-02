Attribute VB_Name = "Mod_Save"
Option Explicit

Public Function Save()
Dim 돌려 As Integer
Open App.Path & "\save\" & 불러온이름 & ".JiNie" For Output As #1
Dim 저장 As Integer
For Oee = 0 To 800
 Print #1, Skill(Oee)
Next

 Print #1, 쿠폰
 Print #1, 하향
 Print #1, 하향횟수

 
For Oee = 0 To 800
 Print #1, T연승(Oee)
 Print #1, Z연승(Oee)
 Print #1, P연승(Oee)
 Print #1, A연승(Oee)
 Print #1, T연(Oee)
 Print #1, Z연(Oee)
 Print #1, P연(Oee)
 Print #1, A연(Oee)
 Print #1, Buff(Oee)
Next

For 저장 = 1 To 6
 Print #1, MyT연승(저장)
 Print #1, MyZ연승(저장)
 Print #1, MyP연승(저장)
 Print #1, MyA연승(저장)
 Print #1, MyT연(저장)
 Print #1, MyZ연(저장)
 Print #1, MyP연(저장)
 Print #1, MyA연(저장)
 Print #1, MyBuff(저장)
 Print #1, MyCode(저장)
Next

For 저장 = 1 To 9
 Print #1, SubT연승(저장)
 Print #1, SubZ연승(저장)
 Print #1, SubP연승(저장)
 Print #1, SubA연승(저장)
 Print #1, SubT연(저장)
 Print #1, SubZ연(저장)
 Print #1, SubP연(저장)
 Print #1, SubA연(저장)
 Print #1, SubBuff(저장)
 Print #1, SubCode(저장)
Next

For 저장 = 1 To 6
 Print #1, MySkill(저장)
Next

For 돌려 = 1 To 9
 Print #1, SubSkill(돌려)
Next
 Print #1, PL넘버
 Print #1, 선수수
 Print #1, Money
 Print #1, 크로우생산
For Oee = 0 To 800
 Print #1, 공격력(Oee)
 Print #1, NPC공격력(Oee)
Next Oee
For 저장 = 1 To 6
 Print #1, MyAt(저장)
Next 저장

 Print #1, PL승
 Print #1, PL패
 Print #1, PL경기수
 Print #1, PL진행
For Oee = 0 To 800
 Print #1, 견제(Oee)
 Print #1, NPC견제(Oee)
Next Oee
For 저장 = 1 To 6
 Print #1, MyR(저장)
Next 저장
 
For Oee = 0 To 800
 Print #1, 전략(Oee)
 Print #1, NPC전략(Oee)
Next Oee
For 저장 = 1 To 6
 Print #1, MySt(저장)
Next 저장
 
For Oee = 0 To 800
 Print #1, 물량(Oee)
 Print #1, NPC물량(Oee)
Next Oee
For 저장 = 1 To 6
 Print #1, MyAm(저장)
Next 저장
 
For Oee = 0 To 800
 Print #1, 수비력(Oee)
 Print #1, NPC수비력(Oee)
Next Oee
For 저장 = 1 To 6
 Print #1, MyDe(저장)
Next 저장
 
For Oee = 0 To 800
 Print #1, 정찰(Oee)
 Print #1, NPC정찰(Oee)
Next Oee
For 저장 = 1 To 6
 Print #1, MyPa(저장)
Next 저장
 
 
For Oee = 0 To 800
 Print #1, 센스(Oee)
 Print #1, NPC센스(Oee)
Next Oee
For 저장 = 1 To 6
 Print #1, MySe(저장)
Next 저장
 
For Oee = 0 To 800
 Print #1, 컨트롤(Oee)
 Print #1, NPC컨트롤(Oee)
Next Oee
For 저장 = 1 To 6
 Print #1, MyCo(저장)
Next 저장
 
For Oee = 0 To 800
 Print #1, 종족(Oee)
Next Oee
For 저장 = 1 To 6
 Print #1, MyTribe(저장)
Next 저장

For Oee = 0 To 800
 Print #1, A승리(Oee)
 Print #1, A패배(Oee)
 Print #1, T승리(Oee)
 Print #1, T패배(Oee)
 Print #1, Z승리(Oee)
 Print #1, Z패배(Oee)
 Print #1, P승리(Oee)
 Print #1, P패배(Oee)
Next Oee
For 돌려 = 1 To 9
Print #1, SubNum(돌려)
Next 돌려
For 저장 = 1 To 6
 Print #1, MyAW(저장)
 Print #1, MyAL(저장)
 Print #1, MyTW(저장)
 Print #1, MyTL(저장)
 Print #1, MyZW(저장)
 Print #1, MyZL(저장)
 Print #1, MyPW(저장)
 Print #1, MyPL(저장)
Next 저장

For Oee = 0 To 800
 Print #1, 우승(Oee)
 Print #1, 준우승(Oee)
 Print #1, 랭크(Oee)
 Print #1, OYear(Oee)
Next Oee
For 저장 = 1 To 6
 Print #1, MyVic(저장)
 Print #1, MySeVic(저장)
 Print #1, MyExp(저장)
 Print #1, MyMExp(저장)
 Print #1, MyLev(저장)
 Print #1, MyRank(저장)
 Print #1, MyYear(저장)
 Print #1, MyPoint(저장)
Next 저장
 Print #1, Mode
 Print #1, Turn
 Print #1, TeamName
 Print #1, val(Money) / 2
 Print #1, PlayNumber(1)
 Print #1, PlayNumber(2)
 Print #1, PlayNumber(3)
 Print #1, PlayNumber(4)
 Print #1, PlayNumber(5)
 Print #1, PlayNumber(6)
 
For Oee = 0 To 800
 Print #1, Team(Oee)
Next Oee
For 저장 = 1 To 6
 Print #1, MyTeam(저장)
Next 저장
For 저장 = 1 To 6
 Print #1, MyNW(저장)
Next
For 저장 = 1 To 9
 Print #1, SubNW(저장)
Next
For 돌려 = 1 To 12
 Print #1, MapName(돌려)
 Print #1, 러쉬거리(돌려)
 Print #1, 자원(돌려)
 Print #1, 복잡도(돌려)
 Print #1, TZT(돌려)
 Print #1, TZZ(돌려)
 Print #1, ZPZ(돌려)
 Print #1, ZPP(돌려)
 Print #1, PTP(돌려)
 Print #1, PTT(돌려)
Next
For 돌려 = 1 To 9
 Print #1, SubTeam(돌려)
 Print #1, SubAt(돌려)
 Print #1, SubR(돌려)
 Print #1, SubSt(돌려)
 Print #1, SubAm(돌려)
 Print #1, SubDe(돌려)
 Print #1, SubPa(돌려)
 Print #1, SubSe(돌려)
 Print #1, SubCo(돌려)
 Print #1, SubLev(돌려)
 Print #1, SubExp(돌려)
 Print #1, SubMExp(돌려)
 Print #1, SubAW(돌려)
 Print #1, SubAL(돌려)
 Print #1, SubTW(돌려)
 Print #1, SubTL(돌려)
 Print #1, SubZW(돌려)
 Print #1, SubZL(돌려)
 Print #1, SubPW(돌려)
 Print #1, SubPL(돌려)
 Print #1, SubRank(돌려)
 Print #1, SubYear(돌려)
 Print #1, SubTribe(돌려)
 Print #1, SubPoint(돌려)
 Print #1, SubVic(돌려)
 Print #1, SubSeVic(돌려)
 Print #1, SubNum(돌려)
Next 돌려
 Print #1, PLEnd
 Print #1, PL우승
 Print #1, PL준우승
Close #1
End Function
