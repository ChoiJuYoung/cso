Attribute VB_Name = "Mod_변수"
Public 통계(1 To 8, 1 To 800) As Integer, 통계도움(1 To 4) As String
Public 총합(1 To 800) As Integer, 이름통계(1 To 800) As String, 년도통계(1 To 800) As String, 랭크통계(1 To 800) As String

Public SetL(1 To 7) As Integer, SetR(1 To 7) As Integer 'Frm_BatInfo 에서 양쪽 엔트리
Public SetNum As Integer, MapL(1 To 7) As Integer '프로리그 현재 세트, 맵
Public Num As Integer
Public SetSel(1 To 6) As Boolean '프로리그 이미 출전했는지 여부
Public SelLNum As Integer, 완료여부 As Boolean, 진행Set As Integer
Public 명단(1 To 3) As Integer, 추첨경우 As Integer, Number As Integer

Public 확인 As Long, 총스탯 As Long, 하향 As String, 하향횟수 As String, Deck As String, Deck년도 As Boolean
Public L1 As Long, L2 As Long, L3 As Long, L4 As Long
Public L5 As Long, L6 As Long, L7 As Long, L8 As Long

Public r1 As Long, r2 As Long, r3 As Long, R4 As Long
Public R5 As Long, R6 As Long, R7 As Long, R8 As Long

Public W1 As Long, W2 As Long, W3 As Long, W4 As Long
Public W5 As Long, W8 As Long, W6 As Long, W7 As Long

Public 갠리그 As Long, Mode As String
Public 돌려 As Integer
Public OYear(0 To 800) As String, Year As String, Map As Integer, AT As String, R As String, St As String, Am As String, De As String, Pa As String, SE As String, Co As String, Con As String, ATO As String, RO As String, StO As String, AmO As String, DeO As String, PaO As String, SeO As String, CoO As String, ConO As String, OpN As String, MN As String, MT As String, OT As String
Public AA As Long, AAO As Long, RAA As String, RAAO As String, AW As String, AL As String, TW As String, TL As String, ZW As String, ZL As String, PW As String, PL As String, SubNW(1 To 9) As String, MyNW(1 To 6) As String, Turn As String, Winer As String
Public MP As Long, OP As Long, AP, SetA As Long, SetN As Long, MW As Long, OW As Long, RUD, ResV, M As Long
Public Oee As Long
Public 저장용 As String
Public ABC As Long, i As Long
Public MyCode(1 To 6) As String, SubCode(1 To 9) As String
Public BGM As String

Public MW2 As Long, OW2 As Long
Public 히힛 As Integer
Public Version As String

Public 히히히 As Long
Public 우히힛 As Long
Public AAA As String
Public A As String, AN As String, B As String, BN As String, c As String, CN As String, a1 As String, a2 As String, A3 As String, b1 As String, b2 As String, b3 As String, C1 As String, C2 As String, C3 As String
Public RAT As String, rR As String, RSt As String, RAm As String, RDe As String, RPa As String, RSe As String, RCo As String, RATO As String, RRO As String, RStO As String, RAmO As String, RDeO As String, RPaO As String, RSeO As String, RCoO As String
Public Victory, SemiVictory As String
Public 행동력 As String
Public 행컨회, 행능, 행컨훈 As Long
Public Choice As String
Public Money As String

Public TeamName As String

Public 공격력(0 To 800) As String, 견제(0 To 800) As String, 전략(0 To 800) As String, 물량(0 To 800) As String, 수비력(0 To 800) As String, 정찰(0 To 800) As String
Public 종족(0 To 800) As String, 센스(0 To 800) As String, 컨트롤(0 To 800) As String, 컨디션(0 To 800) As String, 우승(0 To 800) As String, 준우승(0 To 800) As String, 이름(0 To 800) As String
Public R공격력(0 To 800) As String, R견제(0 To 800) As String, R전략(0 To 800) As String, R물량(0 To 800) As String, R수비력(0 To 800) As String, R정찰(0 To 800) As String, R센스(0 To 800) As String, R컨트롤(0 To 800) As String
Public Buff(0 To 800) As String, MyBuff(0 To 6) As String, SubBuff(0 To 9) As String

'연승,연패
Public T연승(0 To 800) As String, Z연승(0 To 800) As String, P연승(0 To 800) As String, A연승(0 To 800) As String
Public MyT연승(1 To 6) As String, MyZ연승(1 To 6) As String, MyP연승(1 To 6) As String, MyA연승(1 To 6) As String
Public SubT연승(1 To 9) As String, SubZ연승(1 To 9) As String, SubP연승(1 To 9) As String, SubA연승(1 To 9) As String
Public T연(0 To 800) As String, Z연(0 To 800) As String, P연(0 To 800) As String, A연(0 To 800) As String
Public MyT연(1 To 6) As String, MyZ연(1 To 6) As String, MyP연(1 To 6) As String, MyA연(1 To 6) As String
Public SubT연(1 To 9) As String, SubZ연(1 To 9) As String, SubP연(1 To 9) As String, SubA연(1 To 9) As String
Public StatPlusFin As Long


Public X, Y As Long
Public 쿠폰 As String, 팬미팅 As Long, 로또 As Long
Public SearNa(0 To 800) As String, SearName As String
Public Sear As Long

Public RandomAbility As Long

Public RandomPM As Long

Public AllPlus As Long

Public PlusMinus As Long

Public 로딩 As String, 팁 As Integer

Public 이히 As Long
Public 선택량 As Long

Public 경험치 As String, M경험치 As String, 포인트 As String
Public 레벨 As String
Public 합성1 As Integer, 합성2 As Integer
Public 올천 As String

Public Style, PSty As String

Public 잇힝 As Integer

Public A승리(0 To 800) As String, A패배(0 To 800) As String
Public T승리(0 To 800) As String, T패배(0 To 800) As String
Public Z승리(0 To 800) As String, Z패배(0 To 800) As String
Public P승리(0 To 800) As String, P패배(0 To 800) As String


Public MySelect As Long
Public MyName(1 To 6) As String, MyTribe(1 To 6) As String, MyAt(1 To 6) As String
Public MyR(1 To 6) As String, MySt(1 To 6) As String, MyAm(1 To 6) As String, MyCo(1 To 6) As String
Public MyDe(1 To 6) As String, MyPa(1 To 6) As String, MySe(1 To 6) As String
Public MyAW(1 To 6) As String, MyAL(1 To 6) As String, MyTW(1 To 6) As String
Public MyTL(1 To 6) As String, MyZW(1 To 6) As String, MyZL(1 To 6) As String
Public MyPW(1 To 6) As String, MyPL(1 To 6) As String, MyVic(1 To 6) As String
Public MySeVic(1 To 6) As String, MyLev(1 To 6) As String, MyExp(1 To 6) As String, MyMExp(1 To 6) As String
Public MyPoint(1 To 6) As String
Public MyYear(1 To 6) As String

Public 선택 As Integer
Public PlayNumber(1 To 6) As String

Public 랭크(0 To 800) As String, MyRank(1 To 6) As String
Public Team(0 To 800) As String, MyTeam(1 To 6) As String
Public 코스트(0 To 800) As String, MyCost(1 To 6) As String

Public 확인용1 As String
Public 상점NPC As Long, 구매가능 As String
Public 구매 As String, 구매수량 As String
Public 선수수 As String
Public 크로우생산 As String

Public SubName(1 To 9) As String, SubTribe(1 To 9) As String, SubAt(1 To 9) As String
Public SubR(1 To 9) As String, SubSt(1 To 9) As String, SubAm(1 To 9) As String, SubCo(1 To 9) As String
Public SubDe(1 To 9) As String, SubPa(1 To 9) As String, SubSe(1 To 9) As String
Public SubAW(1 To 9) As String, SubAL(1 To 9) As String, SubTW(1 To 9) As String
Public SubTL(1 To 9) As String, SubZW(1 To 9) As String, SubZL(1 To 9) As String
Public SubPW(1 To 9) As String, SubPL(1 To 9) As String, SubVic(1 To 9) As String
Public SubSeVic(1 To 9) As String, SubLev(1 To 9) As String, SubExp(1 To 9) As String, SubMExp(1 To 9) As String
Public SubPoint(1 To 9) As String, SubRank(1 To 9) As String
Public SubYear(1 To 9) As String, SubTeam(1 To 9) As String
Public SubCost(1 To 9) As String, SubChange As Integer


Public 교체Name As String, 교체Tribe As String, 교체At As String, 교체R As String, 교체St As String
Public 교체Am As String, 교체De As String, 교체Pa As String, 교체Se As String, 교체Co As String
Public 교체AW As String, 교체AL As String, 교체TW As String, 교체TL As String
Public 교체ZW As String, 교체ZL As String, 교체PW As String, 교체PL As String, 교체Code As String
Public 교체Team As String, 교체Rank As String, 교체Year As String, 교체Skill As String
Public 교체Exp As String, 교체MExp As String, 교체Point As String, 교체Lev As String
Public 교체Vic As String, 교체SeVic As String, 교체Num As Long, 교체NW As String
Public 교체T연승 As String, 교체Z연승 As String, 교체P연승 As String, 교체A연승 As String
Public 교체T연 As String, 교체Z연 As String, 교체P연 As String, 교체A연 As String

Public SubNum(1 To 9) As String
Public My랭크량 As Integer, O랭크량 As Integer
Public PL출전자(1 To 6) As Boolean
Public PLEnd As String
Public 상점도우미 As String
Public Copy As Integer


'form 24 변수
Public 줄수 As Integer, 알림NPC As Integer, My자원 As Long, My병력 As Long
Public My멀티 As Long, O자원 As Long, O병력 As Long, O멀티 As Long, OStyle As String
Public 경과시간 As Long, My피해 As Long, O피해 As Long, 병력차이 As Long
Public My컨병 As Long, O컨병 As Long

'프로리그 부가
Public PL경기수 As String, PL승 As String, PL패 As String, PL진행 As String
Public 검색 As String, PL넘버 As String
Public PL우승 As String, PL준우승 As String

Public Skill(0 To 800) As String, MySkill(1 To 6) As String, SubSkill(1 To 9) As String

Public 동훈이 As String, 돈량 As String

Public Visible확인 As Boolean
Public NPC공격력(0 To 800) As String, NPC견제(0 To 800) As String, NPC전략(0 To 800) As String
Public NPC물량(0 To 800) As String, NPC수비력(0 To 800) As String, NPC정찰(0 To 800) As String
Public NPC센스(0 To 800) As String, NPC컨트롤(0 To 800) As String
Public CR As Long
Public 불러온이름 As String, 불러옴 As Boolean
Public Helper As Integer


Public MapName(1 To 12) As String, 러쉬거리(1 To 12) As String
Public 자원(1 To 12) As String
Public 복잡도(1 To 12) As String, TZT(1 To 12) As String, TZZ(1 To 12) As String
Public ZPZ(1 To 12) As String, ZPP(1 To 12) As String, PTP(1 To 12) As String, PTT(1 To 12) As String

Public M우세 As Long, O우세 As Long
