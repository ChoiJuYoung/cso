Attribute VB_Name = "Mod_Save"
Option Explicit

Public Function Save()
Dim ���� As Integer
Open App.Path & "\save\" & �ҷ����̸� & ".JiNie" For Output As #1
Dim ���� As Integer
For Oee = 0 To 800
 Print #1, Skill(Oee)
Next

 Print #1, ����
 Print #1, ����
 Print #1, ����Ƚ��

 
For Oee = 0 To 800
 Print #1, T����(Oee)
 Print #1, Z����(Oee)
 Print #1, P����(Oee)
 Print #1, A����(Oee)
 Print #1, T��(Oee)
 Print #1, Z��(Oee)
 Print #1, P��(Oee)
 Print #1, A��(Oee)
 Print #1, Buff(Oee)
Next

For ���� = 1 To 6
 Print #1, MyT����(����)
 Print #1, MyZ����(����)
 Print #1, MyP����(����)
 Print #1, MyA����(����)
 Print #1, MyT��(����)
 Print #1, MyZ��(����)
 Print #1, MyP��(����)
 Print #1, MyA��(����)
 Print #1, MyBuff(����)
 Print #1, MyCode(����)
Next

For ���� = 1 To 9
 Print #1, SubT����(����)
 Print #1, SubZ����(����)
 Print #1, SubP����(����)
 Print #1, SubA����(����)
 Print #1, SubT��(����)
 Print #1, SubZ��(����)
 Print #1, SubP��(����)
 Print #1, SubA��(����)
 Print #1, SubBuff(����)
 Print #1, SubCode(����)
Next

For ���� = 1 To 6
 Print #1, MySkill(����)
Next

For ���� = 1 To 9
 Print #1, SubSkill(����)
Next
 Print #1, PL�ѹ�
 Print #1, ������
 Print #1, Money
 Print #1, ũ�ο����
For Oee = 0 To 800
 Print #1, ���ݷ�(Oee)
 Print #1, NPC���ݷ�(Oee)
Next Oee
For ���� = 1 To 6
 Print #1, MyAt(����)
Next ����

 Print #1, PL��
 Print #1, PL��
 Print #1, PL����
 Print #1, PL����
For Oee = 0 To 800
 Print #1, ����(Oee)
 Print #1, NPC����(Oee)
Next Oee
For ���� = 1 To 6
 Print #1, MyR(����)
Next ����
 
For Oee = 0 To 800
 Print #1, ����(Oee)
 Print #1, NPC����(Oee)
Next Oee
For ���� = 1 To 6
 Print #1, MySt(����)
Next ����
 
For Oee = 0 To 800
 Print #1, ����(Oee)
 Print #1, NPC����(Oee)
Next Oee
For ���� = 1 To 6
 Print #1, MyAm(����)
Next ����
 
For Oee = 0 To 800
 Print #1, �����(Oee)
 Print #1, NPC�����(Oee)
Next Oee
For ���� = 1 To 6
 Print #1, MyDe(����)
Next ����
 
For Oee = 0 To 800
 Print #1, ����(Oee)
 Print #1, NPC����(Oee)
Next Oee
For ���� = 1 To 6
 Print #1, MyPa(����)
Next ����
 
 
For Oee = 0 To 800
 Print #1, ����(Oee)
 Print #1, NPC����(Oee)
Next Oee
For ���� = 1 To 6
 Print #1, MySe(����)
Next ����
 
For Oee = 0 To 800
 Print #1, ��Ʈ��(Oee)
 Print #1, NPC��Ʈ��(Oee)
Next Oee
For ���� = 1 To 6
 Print #1, MyCo(����)
Next ����
 
For Oee = 0 To 800
 Print #1, ����(Oee)
Next Oee
For ���� = 1 To 6
 Print #1, MyTribe(����)
Next ����

For Oee = 0 To 800
 Print #1, A�¸�(Oee)
 Print #1, A�й�(Oee)
 Print #1, T�¸�(Oee)
 Print #1, T�й�(Oee)
 Print #1, Z�¸�(Oee)
 Print #1, Z�й�(Oee)
 Print #1, P�¸�(Oee)
 Print #1, P�й�(Oee)
Next Oee
For ���� = 1 To 9
Print #1, SubNum(����)
Next ����
For ���� = 1 To 6
 Print #1, MyAW(����)
 Print #1, MyAL(����)
 Print #1, MyTW(����)
 Print #1, MyTL(����)
 Print #1, MyZW(����)
 Print #1, MyZL(����)
 Print #1, MyPW(����)
 Print #1, MyPL(����)
Next ����

For Oee = 0 To 800
 Print #1, ���(Oee)
 Print #1, �ؿ��(Oee)
 Print #1, ��ũ(Oee)
 Print #1, OYear(Oee)
Next Oee
For ���� = 1 To 6
 Print #1, MyVic(����)
 Print #1, MySeVic(����)
 Print #1, MyExp(����)
 Print #1, MyMExp(����)
 Print #1, MyLev(����)
 Print #1, MyRank(����)
 Print #1, MyYear(����)
 Print #1, MyPoint(����)
Next ����
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
For ���� = 1 To 6
 Print #1, MyTeam(����)
Next ����
For ���� = 1 To 6
 Print #1, MyNW(����)
Next
For ���� = 1 To 9
 Print #1, SubNW(����)
Next
For ���� = 1 To 12
 Print #1, MapName(����)
 Print #1, �����Ÿ�(����)
 Print #1, �ڿ�(����)
 Print #1, ���⵵(����)
 Print #1, TZT(����)
 Print #1, TZZ(����)
 Print #1, ZPZ(����)
 Print #1, ZPP(����)
 Print #1, PTP(����)
 Print #1, PTT(����)
Next
For ���� = 1 To 9
 Print #1, SubTeam(����)
 Print #1, SubAt(����)
 Print #1, SubR(����)
 Print #1, SubSt(����)
 Print #1, SubAm(����)
 Print #1, SubDe(����)
 Print #1, SubPa(����)
 Print #1, SubSe(����)
 Print #1, SubCo(����)
 Print #1, SubLev(����)
 Print #1, SubExp(����)
 Print #1, SubMExp(����)
 Print #1, SubAW(����)
 Print #1, SubAL(����)
 Print #1, SubTW(����)
 Print #1, SubTL(����)
 Print #1, SubZW(����)
 Print #1, SubZL(����)
 Print #1, SubPW(����)
 Print #1, SubPL(����)
 Print #1, SubRank(����)
 Print #1, SubYear(����)
 Print #1, SubTribe(����)
 Print #1, SubPoint(����)
 Print #1, SubVic(����)
 Print #1, SubSeVic(����)
 Print #1, SubNum(����)
Next ����
 Print #1, PLEnd
 Print #1, PL���
 Print #1, PL�ؿ��
Close #1
End Function
