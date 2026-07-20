/* =====================================================================
   ⚙️ 동기화 설정 — 여기 한 곳에만 배포한 구글 앱스 스크립트 웹앱 URL을 넣으면
   가족 모두 어떤 기기/브라우저로 열어도 별도 입력 없이 동기화 버튼이 바로 동작합니다.
   (비워두면 대시보드 화면에서 각자 URL을 입력해서 저장하는 기존 방식으로 동작)
===================================================================== */
const DEFAULT_WEBHOOK_URL = 'https://script.google.com/macros/s/AKfycbxkanqaM7kc88Xf4S9awiof_Q9N99JTJrj5-SRdPJxsplfO962AKMpTRa8fScVTXoS5/exec'; // 예: 'https://script.google.com/macros/s/AKfycb.../exec'

/* =====================================================================
   SEED DATA — 2026_유럽여행_일정_ING.xlsx '상세일정표' 시트 기준
===================================================================== */
const SEED_DAYS = [
{date:'2026-07-29', label:'7/29(수)', city:'프라하', country:'체코', stay:'Best Western Moran', checkin:true, blocks:[
  {period:'오후', tag:'이동', title:'인천→프라하 도착 · 호텔 체크인', time:'11:05–19:30', place:'프라하 국제공항 → Best Western Moran', tip:'입국심사·수하물 수령 18시 전후(러시아워, 대중교통 추천). 공항 정류장 100번 버스→종점 Zličín역→지하철 B선(노란색) 환승→Karlovo náměstí역 하차, 도보 2~3분(약 50분). 90분권 약 50 CZK.', map:'Best Western Premier Kings Court Prague'},
  {period:'저녁', tag:'식사', title:'★ 저녁식사', time:'20:00', place:'프라하참조 풀에서 선택', tip:'', map:''},
  {period:'밤', tag:'관광', title:'프라하 구시가지 투어(천문시계·구시청사)', time:'', place:'구시가지', tip:'천문시계탑 정시 퍼포먼스. 전망대 9~21시, 300/200 CZK(+100 CZK EV).', map:'Prague Astronomical Clock'},
  {period:'밤', tag:'관광', title:'카를로보 나메스티 광장 산책', time:'', place:'카를로보 나메스티 광장', tip:'프라하 투척사건 무대인 신시청사 옆 대형 광장. 산책하기 좋음(도보 3분).', map:'Karlovo náměstí Prague'},
  {period:'밤', tag:'관광', title:'바츨라프 광장 & 국립박물관', time:'', place:'바츨라프 광장', tip:'프라하 민주화의 상징. 쇼핑몰·레스토랑·환전소 밀집(도보 15~20분).', map:'Wenceslas Square Prague'},
]},
{date:'2026-07-30', label:'7/30(목)', city:'프라하', country:'체코', stay:'Best Western Moran', blocks:[
  {period:'아침', tag:'식사', title:'★ 조식', time:'07:30', place:'숙소', tip:'', map:''},
  {period:'오전', tag:'관광', title:'프라하 성 투어', time:'09:00–13:00', place:'프라하 성', tip:'트램 이용, 성 종탑 전망대에서 시내 조망.', map:'Prague Castle'},
  {period:'오후', tag:'식사', title:'★ 점심식사', time:'', place:'', tip:'', map:''},
  {period:'오후', tag:'관광', title:'카렐교 & 말라스트라나 산책', time:'14:00–17:00', place:'카렐교', tip:'도보 이동, 보트투어도 가능.', map:'Charles Bridge Prague'},
  {period:'저녁', tag:'식사', title:'★ 저녁식사 & 블타바 강변 야경 산책', time:'19:00–21:00', place:'블타바 강변(바플라브카)', tip:'도보 산책, 강변 레스토랑 사전예약 추천.', map:''},
]},
{date:'2026-07-31', label:'7/31(금)', city:'체스키크룸로프', country:'체코', stay:'Best Western Moran(당일투어)', blocks:[
  {period:'아침', tag:'식사', title:'★ 조식', time:'07:00', place:'숙소', tip:'', map:''},
  {period:'오전', tag:'관광', title:'체스키 크룸로프 당일투어', time:'09:00–18:00', place:'체스키 크룸로프', tip:'버스 약 2시간30분 이동. 마이리얼트립 예약 완료.', map:'Český Krumlov'},
  {period:'오후', tag:'식사', title:'★ 점심 · 자유시간', time:'', place:'마을중심부 / 강가', tip:'개별 방문. 보헤미안 전통 립(BBQ), 굴라시, 꼴레뇨, 에겐베르크 양조장 맥주 추천.', map:''},
  {period:'저녁', tag:'식사', title:'★ 저녁식사(예약필수)', time:'19:00', place:'투어가이드 추천 식당', tip:'스비치코바(크림소스 소고기), 흘레비치키(오픈 샌드위치 타파스) 추천.', map:''},
]},
{date:'2026-08-01', label:'8/1(토)', city:'드레스덴', country:'독일', stay:'Best Western Moran(당일투어)', blocks:[
  {period:'아침', tag:'식사', title:'조식', time:'07:00', place:'숙소', tip:'', map:''},
  {period:'오전', tag:'관광', title:'드레스덴 & 바스테이 투어', time:'08:00–19:00', place:'드레스덴 / 바스테이', tip:'전용차량 1.5~2시간 이동. 예약 완료(바스테이 관광 포함).', map:'Bastei Germany'},
  {period:'오후', tag:'식사', title:'★ 점심 · 슈니첼&학센', time:'', place:'Augustiner an der Frauenkirche', tip:'프라우엔 교회 바로 옆, 노이마르크트 광장 근처. "짜지 않게 잘한다"는 후기 다수.', map:'Augustiner an der Frauenkirche Dresden'},
  {period:'오후', tag:'식사', title:'바스테이 간식', time:'', place:'바스테이', tip:'스낵/샌드위치 위주 간단 간식.', map:''},
  {period:'저녁', tag:'식사', title:'★ 저녁식사', time:'', place:'', tip:'', map:''},
]},
{date:'2026-08-02', label:'8/2(일)', city:'뮌헨', country:'독일', stay:'Hilton Munchen City', checkin:true, blocks:[
  {period:'아침', tag:'식사', title:'★ 조식 · 짐싸기 · 체크아웃', time:'07:00–08:40', place:'Best Western Moran', tip:'', map:''},
  {period:'오전', tag:'이동', title:'플로렌츠 버스터미널 이동', time:'09:00', place:'UAN Florenc', tip:'볼트/우버 약 10분(150~250 CZK). 지하철 B·C선 환승역과 직결.', map:'UAN Florenc bus terminal Prague'},
  {period:'오전', tag:'이동', title:'프라하 → 뮌헨 버스 이동 · 점심', time:'09:30–14:15', place:'RegioJet', tip:'노란색 버스, 와이파이·핫음료·콘센트·화장실 완비. 예약 완료. 간단 식사 챙기기.', map:'RegioJet'},
  {period:'오후', tag:'관광', title:'마리엔 광장 · 시내 중심가 투어', time:'16:00–18:00', place:'뮌헨 시내', tip:'도보/U-Bahn 이용.', map:'Marienplatz Munich'},
  {period:'저녁', tag:'식사', title:'★ 저녁식사', time:'19:00', place:'뮌헨', tip:'', map:''},
]},
{date:'2026-08-03', label:'8/3(월)', city:'퓌센', country:'독일', stay:'Hilton Munchen City', blocks:[
  {period:'아침', tag:'식사', title:'★ 조식', time:'07:00', place:'숙소', tip:'', map:''},
  {period:'오전', tag:'관광', title:'퓌센 · 노이슈반슈타인 성 투어', time:'09:00–18:00', place:'노이슈반슈타인 성(13:55 입장예약)', tip:'차량/기차 이동. 1일투어 예약 완료.', map:'Neuschwanstein Castle'},
  {period:'오후', tag:'식사', title:'★ 점심식사', time:'13:00', place:'퓌센', tip:'성 올라가는 길목 노점의 콰르크벨헨(치즈 도넛) 간식 추천.', map:''},
  {period:'저녁', tag:'식사', title:'★ 저녁식사', time:'19:00', place:'퓌센/뮌헨', tip:'', map:''},
]},
{date:'2026-08-04', label:'8/4(화)', city:'뮌헨', country:'독일', stay:'Hilton Munchen City', blocks:[
  {period:'아침', tag:'식사', title:'★ 조식', time:'08:00', place:'숙소', tip:'', map:''},
  {period:'오전', tag:'관광', title:'BMW 박물관 & 벨트 투어', time:'11:00–12:00', place:'BMW Welt', tip:'예약 완료. U-Bahn 이용.', map:'BMW Welt Munich'},
  {period:'오후', tag:'식사', title:'★ 점심식사', time:'12:30', place:'BMW Welt 인근', tip:'벨트 내 레스토랑: 바바리에 바이 케퍼 / 쿠퍼스 가든 / M1 비스트로.', map:''},
  {period:'오후', tag:'관광', title:'올림픽 공원 및 주변 산책', time:'14:00–17:00', place:'올림피아파크', tip:'올림피아 암(비어가든) 또는 안 티엔(베트남 쌀국수) 추천.', map:'Olympiapark Munich'},
  {period:'저녁', tag:'식사', title:'★ 저녁식사', time:'19:00', place:'뮌헨', tip:'', map:''},
]},
{date:'2026-08-05', label:'8/5(수)', city:'린다우', country:'독일', stay:'Air BnB Langenweg 33, Lindau', checkin:true, blocks:[
  {period:'아침', tag:'식사', title:'★ 조식 · 짐싸기 · 체크아웃', time:'07:30', place:'Hilton Munchen City', tip:'', map:''},
  {period:'오전', tag:'이동', title:'뮌헨 → 린다우 이동', time:'09:00–12:00', place:'렌터카(Europcar)', tip:'예약 완료. 에어비앤비 체크인.', map:'Lindau Island'},
  {period:'오후', tag:'식사', title:'★ 점심식사', time:'13:00', place:'린다우', tip:'', map:''},
  {period:'오후', tag:'관광', title:'린다우 섬 항구 · 구시가지 산책', time:'14:00–17:00', place:'린다우 섬', tip:'사자상·등대 포토스팟. 전망 맛집 Restaurant Reutemann(항구/사자상 조망 테라스).', map:'Lindau Harbour lion statue'},
  {period:'저녁', tag:'식사', title:'★ 저녁식사', time:'19:00', place:'린다우', tip:'', map:''},
  {period:'밤', tag:'관광', title:'브레겐츠 페스티벌 · 라 트라비아타', time:'21:00–23:00', place:'Seebühne(보덴호수 위 무대)', tip:'공연 예약 완료.', map:'Bregenzer Festspiele Seebühne'},
]},
{date:'2026-08-06', label:'8/6(목)', city:'브레겐츠', country:'오스트리아', stay:'Air BnB Langenweg 33, Lindau', blocks:[
  {period:'아침', tag:'식사', title:'★ 조식', time:'08:00', place:'숙소', tip:'', map:''},
  {period:'오전', tag:'관광', title:'린다우 항구 · 선셋크루즈 예약확인', time:'10:00–12:00', place:'린다우 항구', tip:'일몰 맛집 EIL.GUT.HALLE(항구 옆 야외 좌석).', map:'Lindau Harbour'},
  {period:'오후', tag:'식사', title:'★ 점심식사', time:'12:30', place:'린다우/브레겐츠', tip:'', map:''},
  {period:'오후', tag:'관광', title:'팬더산(Pfänder) 케이블카 전망', time:'14:00–16:30', place:'브레겐츠', tip:'추천 Berghaus Pfänder(호수 전체 조망 테라스 식사).', map:'Pfänder cable car Bregenz'},
  {period:'저녁', tag:'식사', title:'★ 저녁식사', time:'19:00', place:'린다우', tip:'', map:''},
]},
{date:'2026-08-07', label:'8/7(금)', city:'취리히', country:'스위스', stay:'Air BnB Langenweg 33, Lindau', blocks:[
  {period:'아침', tag:'식사', title:'★ 조식 or 브런치', time:'09:00', place:'숙소', tip:'', map:''},
  {period:'오후', tag:'관광', title:'린트 홈 오브 초콜릿 박물관', time:'12:00–14:00', place:'취리히 킬히베르크', tip:'사전예약 필수, 예약 완료.', map:'Lindt Home of Chocolate'},
  {period:'오후', tag:'관광', title:'취리히 시내 투어', time:'14:30–17:00', place:'취리히', tip:'스위스 비네트(고속도로 통행권) 구매 필요.', map:'Zurich old town'},
  {period:'저녁', tag:'식사', title:'★ 저녁식사', time:'19:00', place:'Valentin Wine Dining', tip:'분위기 좋은 와인 다이닝 추천.', map:'Valentin Wine Dining Zurich'},
]},
{date:'2026-08-08', label:'8/8(토)', city:'파두츠·아펜첼', country:'리히텐슈타인·스위스', stay:'Air BnB Langenweg 33, Lindau', blocks:[
  {period:'아침', tag:'식사', title:'★ 조식', time:'08:30', place:'숙소', tip:'', map:''},
  {period:'오전', tag:'관광', title:'리히텐슈타인 파두츠 투어', time:'10:30–12:30', place:'파두츠', tip:'파두츠 성 외관 관람(내부 불가). 파두츠→아펜첼 42km, 약 40분.', map:'Vaduz Castle'},
  {period:'오후', tag:'식사', title:'★ 점심식사', time:'12:30', place:'파두츠', tip:'기념일 추천 Restaurant Bayerischer Hof(호수 전망 테라스).', map:''},
  {period:'오후', tag:'관광', title:'아펜첼 동화마을', time:'13:00–15:30', place:'아펜첼', tip:'치즈 공방 견학·시식, 전통맥주(알펜비터), 에벤알프 케이블카.', map:'Appenzell village'},
  {period:'저녁', tag:'식사', title:'★ 저녁식사', time:'19:00', place:'아펜첼러 치즈공장(Schaukäserei) 인근', tip:'주차: Parkplatz Brauereiplatz. 전통 치즈 곁들인 슈니첼/뢰스티 추천. 아펜첼→린다우 54km, 약 1시간.', map:''},
]},
{date:'2026-08-09', label:'8/9(일)', city:'프라하', country:'체코', stay:'Hilton Prague Atrium', checkin:true, blocks:[
  {period:'아침', tag:'식사', title:'★ 조식 · 체크아웃', time:'08:00', place:'린다우 숙소', tip:'렌터카 반납.', map:''},
  {period:'오전', tag:'이동', title:'린다우 → 뮌헨 이동', time:'09:00–12:00', place:'메어스부르크 경유', tip:'', map:''},
  {period:'오후', tag:'식사', title:'★ 점심식사', time:'12:30', place:'이동 중', tip:'', map:''},
  {period:'오후', tag:'이동', title:'뮌헨 경유 프라하 복귀', time:'14:00–19:30', place:'RegioJet', tip:'예약 완료. 뚜껑 있는 음료·샌드위치·견과류·바나나 등 간식 챙기기.', map:'RegioJet'},
  {period:'저녁', tag:'식사', title:'★ 저녁식사 & 프라하 체크인', time:'20:00', place:'Hilton Prague Atrium', tip:'', map:''},
]},
{date:'2026-08-10', label:'8/10(월)', city:'프라하', country:'체코', stay:'Hilton Prague Atrium', blocks:[
  {period:'아침', tag:'식사', title:'조식', time:'08:00', place:'숙소', tip:'', map:''},
  {period:'오전', tag:'관광', title:'자유일정 · 기념품 쇼핑', time:'09:00–15:00', place:'하벨시장 / 팔라디움 백화점', tip:'마지막 프라하 자유시간. 하단 쇼핑 리스트 참고.', map:'Havelské tržiště market Prague'},
  {period:'저녁', tag:'식사', title:'★ 마지막 만찬', time:'19:00', place:'그란 피에로 등', tip:'프라하참조 풀에서 특별한 곳으로 예약 추천.', map:''},
]},
{date:'2026-08-11', label:'8/11(화)', city:'프라하 → 인천', country:'귀국일', stay:'Hilton Prague Atrium(체크아웃)', blocks:[
  {period:'아침', tag:'식사', title:'조식', time:'08:00', place:'숙소', tip:'', map:''},
  {period:'오후', tag:'이동', title:'호텔 → 프라하 공항', time:'16:00', place:'우버/볼트', tip:'약 25~35분(500~700 CZK).', map:'Václav Havel Airport Prague'},
  {period:'오후', tag:'이동', title:'공항 도착', time:'17:00', place:'프라하 공항', tip:'', map:''},
  {period:'저녁', tag:'이동', title:'프라하 → 인천 출발', time:'19:05', place:'항공편', tip:'다음날(+1일) 13:20 인천 2터미널 도착 예정.', map:''},
]},
];

const SEED_POOL = [
 {city:'프라하', title:"Pork's Vodičkova", desc:'시그니처 꼴레뇨(479 CZK)+양파수프+감자전. 돼지기름으로 튀긴 감자전이 별미.', map:"Pork's Vodičkova Prague"},
 {city:'프라하', title:'Gran Fierro', desc:'프라하 최고 인기 아르헨티나 스테이크 전문점. 훌륭한 와인 리스트.', map:'Gran Fierro Prague'},
 {city:'프라하', title:'Ginger & Fred', desc:'댄싱하우스 7층 파인다이닝. 블타바강·프라하성 야경 통유리 뷰.', map:'Ginger and Fred restaurant Prague'},
 {city:'프라하', title:"L'Osteria Národní", desc:'대형 화덕피자 반반(Half&Half), 생면 파스타.', map:"L'Osteria Národní Prague"},
 {city:'프라하', title:"Sad Man's Tongue", desc:'빈티지 미국식 인테리어의 수제버거 맛집. 현금 결제만 가능.', map:"Sad Man's Tongue Prague"},
 {city:'프라하', title:'Ekant Indian Restaurant', desc:'카를로보 나메스티 광장 옆 인도 요리. 매운맛 조절 가능.', map:'Ekant Indian Restaurant Prague'},
 {city:'프라하', title:'Lobkowicz Palace Restaurant & Café', desc:'프라하성에서 시내 전경 보며 차 한잔.', map:'Lobkowicz Palace Restaurant Prague'},
 {city:'프라하', title:'Our Lady of Exile', desc:'스트라호프 수도원 전망대 근처, 수도원 맥주로 유명.', map:'Strahov Monastery viewpoint Prague'},
 {city:'프라하', title:'Năm - Viet Kitchen', desc:'느끼할 때 생각나는 쌀국수 한 그릇.', map:'Nam Viet Kitchen Prague'},
 {city:'프라하', title:'Vyšehradské hradby', desc:'프라하 일몰 명소, 성벽에서 보는 노을.', map:'Vyšehrad walls Prague'},
 {city:'프라하', title:'아르테 비앙카 빵집 & 카페', desc:'커피와 빵이 맛있는 베이커리 카페.', map:'Arte Bianca bakery Prague'},
 {city:'체스키크룸로프', title:'구시가지 강가 식당가', desc:'보헤미안 전통 립, 굴라시, 꼴레뇨, 에겐베르크 로컬 맥주.', map:'Český Krumlov restaurants near river'},
 {city:'드레스덴', title:'Augustiner an der Frauenkirche', desc:'슈바인학센·슈니첼, 짜지 않다는 후기 다수.', map:'Augustiner an der Frauenkirche Dresden'},
 {city:'드레스덴', title:'Pulverturm', desc:'옛 화약탑 지하 개조 레스토랑, 한국어 메뉴판 제공.', map:'Pulverturm Dresden'},
 {city:'드레스덴', title:'Gänsedieb', desc:'구시가지 맛집거리, 신선한 샐러드 곁들인 슈니첼.', map:'Gänsedieb Dresden'},
 {city:'뮌헨', title:'올림피아 암 (Olympia Alm)', desc:'뮌헨 최고지대 비어가든. 소시지구이·슈니첼·생맥주.', map:'Olympia Alm Munich'},
 {city:'뮌헨', title:'안 티엔 (Anh Tien)', desc:'BMW박물관 근처 인기 베트남 식당, 가성비 런치세트.', map:'Anh Tien restaurant Munich'},
 {city:'린다우', title:'Restaurant Reutemann', desc:'린다우 항구·사자상 조망 테라스.', map:'Restaurant Reutemann Lindau'},
 {city:'린다우', title:'EIL.GUT.HALLE', desc:'항구 옆 일몰 감상하기 좋은 야외 좌석.', map:'EIL.GUT.HALLE Lindau'},
 {city:'브레겐츠', title:'Berghaus Pfänder', desc:'팬더산 정상, 호수 전체 조망 테라스 식사.', map:'Berghaus Pfänder Bregenz'},
 {city:'취리히', title:'Valentin Wine Dining', desc:'분위기 좋은 와인 다이닝.', map:'Valentin Wine Dining Zurich'},
 {city:'파두츠·아펜첼', title:'Restaurant Bayerischer Hof', desc:'고급스러운 호수 전망 테라스, 기념일 추천.', map:'Restaurant Bayerischer Hof Vaduz'},
 {city:'파두츠·아펜첼', title:'Appenzeller Schaukäserei', desc:'전통 치즈 공방, 제조 과정 견학 및 시식.', map:'Appenzeller Schaukäserei'},
];

const SEED_SHOPPING = [
 {country:'체코', city:'프라하', items:['마누팍투라(Manufaktura) 맥주 샴푸·핸드크림','보타니쿠스(Botanicus) 여왕의 팩 장미오일·수제비누','아포테카(Apoteka) 3분 마스크팩','마를렌카(Marlenka) 꿀 케이크','코로나다(Kolonada) 국민 웨하스','벌꿀주(Medovina)','베헤롭카(Becherovka) 허브 리큐르','체코 와인(보헤미아/모라비아)','마리오네트 나무인형','보헤미안 글라스/크리스탈','필스너 우르켈 캔맥주(마트, 1유로 미만)'], loc:'추천 쇼핑지: 올드타운, 하벨시장, 프라하성 황금소로, 팔라디움 백화점'},
 {country:'체코', city:'체스키 크룸로프', items:['전통 보헤미안 진저브레드(Český Perník)','수제 꿀 비누·꿀 화장품·벌꿀주(메도비나)','에겐베르크(Eggenberg) 로컬 양조장 맥주']},
 {country:'독일', city:'드레스덴 & 바스테이', items:['마이센(Meissen) 국립 도자기','에르츠게비르게(Erzgebirge) 목공예품','드레스덴 슈톨렌("Dresdner Stollen" 인증마크 확인)','푼츠 몰케라이(Pfunds Molkerei) 유제품','바스테이 허브 리큐르(Kräuterlikör)','로컬 천연꿀 & 사과와인(Apfelwein)']},
 {country:'독일', city:'뮌헨', items:['드럭스토어(dm/Rossmann): 바레아(Balea) 앰플·마스크팩','아요나(Ajona) 치약','발포 비타민','알페신(Alpecin) 카페인 샴푸','쌍둥이칼(Zwilling)','독일 전통 맥주잔','하리보(Haribo) 젤리','리터 스포트(Ritter Sport) 초콜릿']},
 {country:'독일', city:'린다우', items:['드럭스토어(dm): 발레아 앰플·립밤, 카밀(Kamill) 핸드크림','린트(Lindt) 초콜릿','독일 꿀','하리보(Haribo) 젤리','과일 브랜디(슈냅스)','지역 와인']},
 {country:'독일', city:'퓌센 · 노이슈반슈타인', items:['독일 뻐꾸기시계(택스리펀 10~12%)','알고이(Allgäu) 천연 치즈 & 벌꿀','콰르크벨헨(Quarkbällchen) 튀김도넛','엠오이칼(Em-eukal) 허브 목캔디']},
 {country:'독일', city:'BMW 박물관 & 벨트', items:['BMW 로고 티셔츠/후드티','서모 텀블러·보틀','우산(Puma 협업)','50.01유로 이상 구매 시 텍스리펀 서류 요청']},
];

const SEED_CHECKLIST = [
 {cat:'Basic Item', items:['여권','항공권','국제학생증','국제운전면허증','여권사본','여행자 보험','로밍']},
 {cat:'Reservation', items:['렌터카','버스(프라하<->뮌헨)','드레스덴 투어','체스키크룸로프 투어','노이반슈타인성','린트초콜릿','BMW 박물관','브레겐트 오페라공연']},
 {cat:'Payment Item', items:['트래블 월릿','유로','예비 카드']},
 {cat:'Electric Item', items:['충전기','충전단자','멀티탭','보조배터리']},
 {cat:'Sanitary', items:['세면도구','로션','렌즈','고무줄','손소독제','손소독티슈','인공눈물','립밤']},
 {cat:'Sun Protect', items:['모자','선크림','선글라스','팔토시, 목토시']},
 {cat:'Theft Protect', items:['소형열쇠','옷핀','복대','휴대폰 분실방지','슬링백']},
 {cat:'Photo Item', items:['폴라로이드 카메라']},
 {cat:'Clothes', items:['속옷','양말','바람막이','바지','티셔츠','손수건','후드','신발']},
 {cat:'Medicine', items:['종합감기약','지사제','소화제','해열진통제','밴드','연고']},
 {cat:'etc', items:['지퍼백/비닐봉지','동전지갑','필기도구','안경']},
 {cat:'Korean food', items:['라면','누릉지','생수','과자']},
];

/* =====================================================================
   STATE
===================================================================== */
const STATE_KEY = 'euro2026_state_v3';
function uid(p){ return p+'_'+Math.random().toString(36).slice(2,9); }
function esc(s){ return (s===null||s===undefined?'':String(s)).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function loadJSON(key, fallback){ try{ const v = JSON.parse(localStorage.getItem(key)); return v==null? fallback : v; }catch(e){ return fallback; } }
function saveJSON(key, val){ localStorage.setItem(key, JSON.stringify(val)); }

function defaultState(){
  return {
    days: SEED_DAYS.map((d,di)=>({ id:'day'+di, date:d.date, label:d.label, city:d.city, country:d.country, stay:d.stay, checkin:!!d.checkin,
      blocks: d.blocks.map(b=>({ id:uid('blk'), period:b.period, tag:b.tag, title:b.title, time:b.time, place:b.place, tip:b.tip, map:b.map, estCost:null, actCost:null, memo:'' })) })),
    pool: SEED_POOL.map(p=>({ id:uid('pool'), city:p.city, title:p.title, desc:p.desc, map:p.map })),
    shopping: SEED_SHOPPING,
    checklist: SEED_CHECKLIST.map(c=>({ id:uid('cat'), cat:c.cat, items:c.items.map(label=>({ id:uid('item'), label, done:false })) })),
    memos: [],
    shopChecked: {},
    meta: { webhookUrl:'', lastSync:null }
  };
}
let state = loadJSON(STATE_KEY, null) || defaultState();
if(!state.shopping) state.shopping = SEED_SHOPPING;
if(!state.memos) state.memos = [];
if(!state.shopChecked) state.shopChecked = {};
if(!state.meta) state.meta = {webhookUrl:'', lastSync:null};
if(!state.meta.webhookUrl && DEFAULT_WEBHOOK_URL) state.meta.webhookUrl = DEFAULT_WEBHOOK_URL;
function persist(){ saveJSON(STATE_KEY, state); }

let activeDayId = null;

/* =====================================================================
   NAV
===================================================================== */
const titles = {dashboard:'종합 대시보드', timeline:'인터랙티브 타임라인', checklist:'준비물 체크리스트', memo:'메모 & 쇼핑 가이드'};
function switchTab(name){
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.getElementById('tab-'+name).classList.add('active');
  document.querySelectorAll('.navbtn').forEach(b=>b.classList.toggle('active', b.dataset.tab===name));
  document.getElementById('topTitle').textContent = titles[name];
  if(name==='timeline') renderTimeline();
  if(name==='dashboard') renderDashboard();
}

/* =====================================================================
   DASHBOARD
===================================================================== */
function daysBetween(a,b){ return Math.round((b-a)/86400000); }

function renderDashboard(){
  const today = new Date(); today.setHours(0,0,0,0);
  const start = new Date(state.days[0].date);
  const end = new Date(state.days[state.days.length-1].date);
  const totalDays = daysBetween(start,end)+1;
  let elapsed = daysBetween(start, today);
  let dday = daysBetween(today, start);

  const ddayText = document.getElementById('ddayText');
  const ddaySub = document.getElementById('ddaySub');
  const progressBar = document.getElementById('progressBar');
  const progressLabel = document.getElementById('progressLabel');
  let currentDay = null;

  if(dday > 0){
    ddayText.textContent = 'D-'+dday;
    ddaySub.textContent = '2026.07.29(수) 출발까지';
    progressBar.style.width = '0%'; progressLabel.textContent = '0 / '+totalDays+' 일차';
  } else if(dday === 0){
    ddayText.textContent = 'D-DAY'; ddaySub.textContent = '오늘 출발! 좋은 여행 되세요 ✈️';
    currentDay = state.days[0];
    progressBar.style.width = (100/totalDays)+'%'; progressLabel.textContent = '1 / '+totalDays+' 일차';
  } else if(elapsed >= 0 && elapsed < totalDays){
    const dayNum = elapsed+1;
    ddayText.textContent = dayNum+'일차'; ddaySub.textContent = '여정이 진행 중입니다';
    currentDay = state.days[elapsed];
    progressBar.style.width = Math.round(dayNum/totalDays*100)+'%'; progressLabel.textContent = dayNum+' / '+totalDays+' 일차';
  } else {
    ddayText.textContent = 'END'; ddaySub.textContent = '2026년 여정이 마무리되었습니다';
    currentDay = state.days[state.days.length-1];
    progressBar.style.width = '100%'; progressLabel.textContent = totalDays+' / '+totalDays+' 일차';
  }

  const stayDay = currentDay || state.days[0];
  document.getElementById('stayCard').innerHTML = `
    <div style="display:flex; gap:12px; align-items:center;">
      <div style="width:44px; height:44px; border-radius:11px; background:var(--paper-2); display:flex; align-items:center; justify-content:center; flex-shrink:0;">
        <svg width="22" height="22" fill="none" stroke="var(--gold-deep)" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M3 21h18M5 21V7l8-4v18M13 21V11l6 4v6M9 9v.01M9 12v.01M9 15v.01"/></svg>
      </div>
      <div style="flex:1;">
        <div style="font-weight:700; font-size:15px;">${esc(stayDay.stay)}</div>
        <div style="font-size:12px; color:var(--muted); margin-top:2px;">${esc(stayDay.city)}, ${esc(stayDay.country)} · ${esc(stayDay.label)}</div>
      </div>
      <span class="badge ${stayDay.checkin?'badge-pending':'badge-done'}">${stayDay.checkin? '체크인 예정' : '투숙 중'}</span>
    </div>`;

  renderRouteLine(elapsed);
  renderBudget();
  renderSyncStatus();
}

function renderBudget(){
  let est=0, act=0;
  state.days.forEach(d=>d.blocks.forEach(b=>{ est += Number(b.estCost)||0; act += Number(b.actCost)||0; }));
  document.getElementById('budgetEst').textContent = est.toLocaleString()+'원';
  document.getElementById('budgetAct').textContent = act.toLocaleString()+'원';
}

function renderRouteLine(elapsed){
  const svg = document.getElementById('routeSvg');
  const cities = [];
  state.days.forEach((d,i)=>{ if(cities.length===0 || cities[cities.length-1].city!==d.city) cities.push({city:d.city, idx:i}); });
  const stepW = 92, padX = 24, y = 54;
  const w = padX*2 + stepW*(cities.length-1) + 40;
  svg.setAttribute('width', w);
  svg.setAttribute('viewBox', `0 0 ${w} 108`);
  let html = `<line x1="${padX}" y1="${y}" x2="${padX+stepW*(cities.length-1)}" y2="${y}" stroke="#DED5C0" stroke-width="4"/>`;
  cities.forEach((c,i)=>{
    const x = padX + i*stepW;
    const isNow = elapsed!==null && elapsed>=0 && (c.idx<=elapsed && (i+1>=cities.length || cities[i+1].idx>elapsed));
    let fill = '#FBF8F1', stroke = '#C99A3E', r = 7;
    if(elapsed!==null && elapsed>=0){
      if(c.idx < elapsed){ fill = '#6C8A6F'; stroke='#4B6650'; }
      if(isNow){ fill = '#C99A3E'; stroke='#9C7325'; r=9; }
    }
    html += `<circle cx="${x}" cy="${y}" r="${r}" fill="${fill}" stroke="${stroke}" stroke-width="2.5"/>`;
    html += `<text x="${x}" y="${y-16}" font-family="IBM Plex Mono, monospace" font-size="10" fill="#6B7280" text-anchor="middle">${esc(state.days[c.idx].label.replace(/\(.\)/,''))}</text>`;
    html += `<text x="${x}" y="${y+24}" font-family="Inter, sans-serif" font-size="11.5" font-weight="700" fill="${isNow?'#9C7325':'#1E2430'}" text-anchor="middle">${esc(c.city)}</text>`;
  });
  svg.innerHTML = html;
}

/* =====================================================================
   TIMELINE
===================================================================== */
function renderDayChips(){
  const wrap = document.getElementById('dayChips');
  const today = new Date(); today.setHours(0,0,0,0);
  wrap.innerHTML = state.days.map(d=>{
    const isToday = new Date(d.date).getTime()===today.getTime();
    return `<div class="daychip ${d.id===activeDayId?'active':''} ${isToday?'today':''}" onclick="selectDay('${d.id}')">
      <div class="d">${esc(d.label)}</div><div class="c">${esc(d.city)}</div>
    </div>`;
  }).join('');
}
function selectDay(id){ activeDayId = id; renderDayChips(); renderTimelineBody(); renderPool(); }
function renderTimeline(){
  if(!activeDayId){
    const today = new Date(); today.setHours(0,0,0,0);
    const found = state.days.find(d=>new Date(d.date).getTime()>=today.getTime());
    activeDayId = (found || state.days[0]).id;
  }
  renderDayChips(); renderTimelineBody(); renderPool();
}
function tagIcon(tag){ return {'이동':'🚌','관광':'📍','식사':'🍽️','숙소':'🏨'}[tag] || '•'; }

function renderTimelineBody(){
  const day = state.days.find(d=>d.id===activeDayId);
  const body = document.getElementById('timelineBody');
  body.innerHTML = `<div class="card" style="margin-bottom:12px;">
    <div style="display:flex; justify-content:space-between; align-items:center;">
      <div>
        <div class="font-display" style="font-size:17px; font-weight:700;">${esc(day.city)} <span style="font-weight:400; font-size:13px; color:var(--muted);">· ${esc(day.country)}</span></div>
        <div style="font-size:12px; color:var(--muted); margin-top:2px;">${esc(day.label)} · ${esc(day.stay)}</div>
      </div>
      <span style="font-size:10px; font-family:'IBM Plex Mono',monospace; color:var(--muted); text-align:right;">⠿ 눌러서<br>순서 변경</span>
    </div>
  </div>
  <div id="sortableContainer"></div>
  <button class="btn-primary" style="width:100%; margin-top:8px;" onclick="openBlockModal('${day.id}', null)">＋ 새 일정 카드 추가</button>`;

  const container = document.getElementById('sortableContainer');
  let currentPeriod = null, inner = '';
  day.blocks.forEach(b=>{
    if(b.period !== currentPeriod){ inner += `<div class="tl-period">${esc(b.period)}</div>`; currentPeriod = b.period; }
    const costRow = (b.estCost||b.actCost) ? `<div style="display:flex; gap:8px; margin-top:8px;">${b.estCost?`<span class="badge badge-pending">예상 ${Number(b.estCost).toLocaleString()}원</span>`:''}${b.actCost?`<span class="badge badge-done">지출 ${Number(b.actCost).toLocaleString()}원</span>`:''}</div>` : '';
    inner += `<div class="tl-card tag-${esc(b.tag)}" data-id="${b.id}" onclick="openBlockModal('${day.id}','${b.id}')">
      <div class="tl-top">
        <div>
          <span class="drag-handle" onclick="event.stopPropagation();">⠿</span><span class="tl-time">${esc(b.time)}</span>
          <div class="tl-title">${tagIcon(b.tag)} ${esc(b.title)}</div>
          ${b.place?`<div class="tl-place">${esc(b.place)}</div>`:''}
        </div>
        <div style="display:flex; flex-direction:column; align-items:flex-end; gap:6px;">
          <span class="tl-tag-pill">${esc(b.tag)}</span>
          ${b.map?`<div class="tl-map-btn" onclick="event.stopPropagation(); openMap('${b.map.replace(/'/g,"\\'")}')"><svg fill="none" viewBox="0 0 24 24" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M9 20l-5.447-2.724A1 1 0 013 16.382V5.618a1 1 0 011.447-.894L9 7m0 13l6-3m-6 3V7m6 10l4.553 2.276A1 1 0 0021 18.382V7.618a1 1 0 00-.553-.894L15 4m0 13V4m0 0L9 7"/></svg></div>`:''}
        </div>
      </div>
      ${b.tip?`<div class="tl-tip">💡 ${esc(b.tip)}</div>`:''}
      ${b.memo?`<div class="tl-tip" style="background:#EFEAF7;">📝 ${esc(b.memo)}</div>`:''}
      ${costRow}
    </div>`;
  });
  container.innerHTML = inner;

  Sortable.create(container, {
    animation:180, handle:'.drag-handle', ghostClass:'sortable-ghost', dragClass:'sortable-drag',
    onEnd:function(){
      const newIds = Array.from(container.querySelectorAll('.tl-card')).map(el=>el.dataset.id);
      const map = {}; day.blocks.forEach(b=>map[b.id]=b);
      day.blocks = newIds.map(id=>map[id]);
      persist();
      renderTimelineBody();
    }
  });
}

function renderPool(){
  const day = state.days.find(d=>d.id===activeDayId);
  const zone = document.getElementById('poolZone');
  const spots = state.pool.filter(p=>p.city===day.city);
  let html = spots.map(s=>`
    <div class="pool-card">
      <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px;">
        <div><b>${esc(s.title)}</b>${esc(s.desc)}</div>
        <div style="display:flex; gap:8px; flex-shrink:0;">
          <span class="icon-btn" onclick="openPoolModal('${s.id}')">✎</span>
          <span class="icon-btn danger" onclick="deletePoolSpot('${s.id}')">🗑</span>
        </div>
      </div>
      <div style="margin-top:7px; display:flex; gap:14px; align-items:center;">
        ${s.map?`<span onclick="openMap('${s.map.replace(/'/g,"\\'")}')" style="font-size:11px; color:var(--brick); font-weight:700; cursor:pointer;">📍 지도 보기</span>`:''}
        <span onclick="addPoolToTimeline('${s.id}')" style="font-size:11px; color:var(--moss); font-weight:700; cursor:pointer;">＋ 오늘 일정에 추가</span>
      </div>
    </div>`).join('');
  if(spots.length===0) html = `<div style="font-size:12.5px; color:var(--muted); text-align:center; padding:10px;">이 도시에 등록된 추천 스팟이 없습니다.</div>`;
  html += `<button class="btn-ghost" style="width:100%; margin-top:4px;" onclick="openPoolModal(null)">＋ 새 스팟 추가</button>`;
  zone.innerHTML = html;
}

function openMap(query){ window.open('https://www.google.com/maps/search/?api=1&query='+encodeURIComponent(query), '_blank'); }

/* --- Block create/edit modal --- */
function openBlockModal(dayId, blockId){
  const day = state.days.find(d=>d.id===dayId);
  const block = blockId ? day.blocks.find(b=>b.id===blockId) : {period: day.blocks.length? day.blocks[day.blocks.length-1].period : '오전', tag:'관광', title:'', time:'', place:'', tip:'', map:'', estCost:null, actCost:null, memo:''};
  const periods = ['아침','오전','오후','저녁','밤','추가 일정'];
  const tags = ['이동','관광','식사','숙소'];
  const html = `
    <div class="font-display" style="font-size:17px; font-weight:700; margin-bottom:6px;">${blockId?'일정 카드 수정':'새 일정 카드 추가'}</div>
    <div class="formGrid">
      <label>시간대<select id="fm_period">${periods.map(p=>`<option ${p===block.period?'selected':''}>${p}</option>`).join('')}</select></label>
      <label>태그<select id="fm_tag">${tags.map(t=>`<option ${t===block.tag?'selected':''}>${t}</option>`).join('')}</select></label>
    </div>
    <label>제목<input id="fm_title" value="${esc(block.title)}" placeholder="예: 프라하 성 투어"></label>
    <div class="formGrid">
      <label>시간<input id="fm_time" value="${esc(block.time)}" placeholder="09:00–13:00"></label>
      <label>장소<input id="fm_place" value="${esc(block.place)}" placeholder="장소명"></label>
    </div>
    <label>팁/설명<textarea id="fm_tip">${esc(block.tip)}</textarea></label>
    <label>구글맵 검색어<input id="fm_map" value="${esc(block.map)}" placeholder="영문 장소명 권장"></label>
    <div class="formGrid">
      <label>예상비용(원)<input id="fm_est" type="number" value="${block.estCost??''}" placeholder="0"></label>
      <label>실제 사용비용(원)<input id="fm_act" type="number" value="${block.actCost??''}" placeholder="0"></label>
    </div>
    <label>현지 메모<textarea id="fm_memo">${esc(block.memo||'')}</textarea></label>
    <div style="display:flex; gap:8px; margin-top:14px;">
      <button class="btn-primary" style="flex:1;" onclick="saveBlockModal('${dayId}', ${blockId?`'${blockId}'`:null})">저장</button>
      ${blockId? `<button class="btn-ghost" onclick="demoteToPool('${dayId}','${blockId}')">풀로 내리기</button>` : ''}
      ${blockId? `<button class="btn-danger" onclick="deleteBlock('${dayId}','${blockId}')">삭제</button>` : ''}
    </div>
    <button class="btn-cancel" onclick="closeFormModal()">취소</button>`;
  openFormModal(html);
}
function saveBlockModal(dayId, blockId){
  const day = state.days.find(d=>d.id===dayId);
  const val = id=>document.getElementById(id).value;
  const title = val('fm_title').trim();
  if(!title){ alert('제목을 입력해주세요.'); return; }
  const data = { period: val('fm_period'), tag: val('fm_tag'), title, time: val('fm_time').trim(), place: val('fm_place').trim(), tip: val('fm_tip').trim(), map: val('fm_map').trim(), estCost: val('fm_est')? Number(val('fm_est')) : null, actCost: val('fm_act')? Number(val('fm_act')) : null, memo: val('fm_memo').trim() };
  if(blockId){
    const idx = day.blocks.findIndex(b=>b.id===blockId);
    day.blocks[idx] = {...day.blocks[idx], ...data};
  } else {
    day.blocks.push({id:uid('blk'), ...data});
  }
  persist(); closeFormModal(); renderTimelineBody(); renderDashboard();
}
function deleteBlock(dayId, blockId){
  if(!confirm('이 일정 카드를 삭제할까요?')) return;
  const day = state.days.find(d=>d.id===dayId);
  day.blocks = day.blocks.filter(b=>b.id!==blockId);
  persist(); closeFormModal(); renderTimelineBody(); renderDashboard();
}
function demoteToPool(dayId, blockId){
  const day = state.days.find(d=>d.id===dayId);
  const block = day.blocks.find(b=>b.id===blockId);
  if(!confirm('"'+block.title+'" 카드를 일정에서 빼고 추천 스팟 풀로 옮길까요?')) return;
  state.pool.push({id:uid('pool'), city:day.city, title:block.title.replace(/^📌\s*/,'').replace(/^★\s*/,''), desc: block.tip||block.place||'', map: block.map||''});
  day.blocks = day.blocks.filter(b=>b.id!==blockId);
  persist(); closeFormModal(); renderTimelineBody(); renderPool();
}

/* --- Pool create/edit modal --- */
function openPoolModal(poolId){
  const day = state.days.find(d=>d.id===activeDayId);
  const spot = poolId ? state.pool.find(p=>p.id===poolId) : {city: day.city, title:'', desc:'', map:''};
  const html = `
    <div class="font-display" style="font-size:17px; font-weight:700; margin-bottom:6px;">${poolId?'추천 스팟 수정':'새 추천 스팟 추가'}</div>
    <label>도시<input id="pm_city" value="${esc(spot.city)}" placeholder="예: 프라하"></label>
    <label>이름<input id="pm_title" value="${esc(spot.title)}" placeholder="가게/장소 이름"></label>
    <label>설명<textarea id="pm_desc">${esc(spot.desc)}</textarea></label>
    <label>구글맵 검색어<input id="pm_map" value="${esc(spot.map)}" placeholder="영문 장소명 권장"></label>
    <div style="display:flex; gap:8px; margin-top:14px;">
      <button class="btn-primary" style="flex:1;" onclick="savePoolModal(${poolId?`'${poolId}'`:null})">저장</button>
      ${poolId? `<button class="btn-danger" onclick="deletePoolSpot('${poolId}')">삭제</button>` : ''}
    </div>
    <button class="btn-cancel" onclick="closeFormModal()">취소</button>`;
  openFormModal(html);
}
function savePoolModal(poolId){
  const val = id=>document.getElementById(id).value.trim();
  const title = val('pm_title');
  if(!title){ alert('이름을 입력해주세요.'); return; }
  const data = {city: val('pm_city')||'미지정', title, desc: val('pm_desc'), map: val('pm_map')};
  if(poolId){
    const idx = state.pool.findIndex(p=>p.id===poolId);
    state.pool[idx] = {...state.pool[idx], ...data};
  } else {
    state.pool.push({id:uid('pool'), ...data});
  }
  persist(); closeFormModal(); renderPool();
}
function deletePoolSpot(poolId){
  if(!confirm('이 추천 스팟을 삭제할까요?')) return;
  state.pool = state.pool.filter(p=>p.id!==poolId);
  persist(); closeFormModal(); renderPool();
}
function addPoolToTimeline(poolId){
  const spot = state.pool.find(p=>p.id===poolId);
  const day = state.days.find(d=>d.id===activeDayId);
  day.blocks.push({id:uid('blk'), period:'추가 일정', tag:'관광', title:'📌 '+spot.title, time:'', place:spot.desc, tip:'추천 스팟 풀에서 추가됨', map:spot.map, estCost:null, actCost:null, memo:''});
  persist(); renderTimelineBody();
}

/* =====================================================================
   CHECKLIST
===================================================================== */
function renderChecklist(){
  const body = document.getElementById('checklistBody');
  let totalItems=0, totalDone=0, html='';
  state.checklist.forEach(cat=>{
    const done = cat.items.filter(i=>i.done).length;
    totalItems += cat.items.length; totalDone += done;
    html += `<div class="card chk-cat">
      <div class="chk-cat-head">
        <h3>${esc(cat.cat)}</h3>
        <div style="display:flex; align-items:center; gap:10px;">
          <span class="chk-progress">${done}/${cat.items.length}</span>
          <span class="icon-btn" onclick="editCatName('${cat.id}')">✎</span>
          <span class="icon-btn danger" onclick="deleteCategory('${cat.id}')">🗑</span>
        </div>
      </div>
      ${cat.items.map(item=>`
        <div class="chk-item">
          <div class="chk-box ${item.done?'on':''}" onclick="toggleChk('${cat.id}','${item.id}')">
            <svg viewBox="0 0 24 24" fill="none" stroke-width="3"><path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7"/></svg>
          </div>
          <div class="chk-label ${item.done?'done':''}">${esc(item.label)}</div>
          <span class="icon-btn" onclick="editChkItem('${cat.id}','${item.id}')">✎</span>
          <span class="icon-btn danger" onclick="deleteChkItem('${cat.id}','${item.id}')">×</span>
        </div>`).join('')}
      <div style="display:flex; gap:6px; margin-top:9px;">
        <input id="newitem-${cat.id}" placeholder="새 항목 추가" style="flex:1; border:1px solid var(--line); border-radius:8px; padding:7px 9px; font-size:12.5px;" onkeydown="if(event.key==='Enter')addChecklistItem('${cat.id}')">
        <button class="btn-ghost" onclick="addChecklistItem('${cat.id}')">추가</button>
      </div>
    </div>`;
  });
  html += `<div class="card">
    <div style="display:flex; gap:6px;">
      <input id="newCatInput" placeholder="새 카테고리 이름" style="flex:1; border:1px solid var(--line); border-radius:8px; padding:9px 10px; font-size:13px;" onkeydown="if(event.key==='Enter')addCategory()">
      <button class="btn-primary" onclick="addCategory()">카테고리 추가</button>
    </div>
  </div>`;
  body.innerHTML = html;
  const pct = totalItems? Math.round(totalDone/totalItems*100) : 0;
  document.getElementById('chkOverallPct').textContent = pct+'%';
  document.getElementById('chkOverallBar').style.width = pct+'%';
}
function toggleChk(catId, itemId){
  const cat = state.checklist.find(c=>c.id===catId);
  const item = cat.items.find(i=>i.id===itemId);
  item.done = !item.done; persist(); renderChecklist();
}
function editChkItem(catId, itemId){
  const cat = state.checklist.find(c=>c.id===catId);
  const item = cat.items.find(i=>i.id===itemId);
  const val = prompt('항목 이름 수정', item.label);
  if(val===null || !val.trim()) return;
  item.label = val.trim(); persist(); renderChecklist();
}
function deleteChkItem(catId, itemId){
  if(!confirm('이 항목을 삭제할까요?')) return;
  const cat = state.checklist.find(c=>c.id===catId);
  cat.items = cat.items.filter(i=>i.id!==itemId); persist(); renderChecklist();
}
function addChecklistItem(catId){
  const input = document.getElementById('newitem-'+catId);
  const val = input.value.trim(); if(!val) return;
  const cat = state.checklist.find(c=>c.id===catId);
  cat.items.push({id:uid('item'), label:val, done:false});
  input.value=''; persist(); renderChecklist();
}
function editCatName(catId){
  const cat = state.checklist.find(c=>c.id===catId);
  const val = prompt('카테고리 이름 수정', cat.cat);
  if(val===null || !val.trim()) return;
  cat.cat = val.trim(); persist(); renderChecklist();
}
function deleteCategory(catId){
  if(!confirm('이 카테고리 전체를 삭제할까요?')) return;
  state.checklist = state.checklist.filter(c=>c.id!==catId); persist(); renderChecklist();
}
function addCategory(){
  const input = document.getElementById('newCatInput');
  const val = input.value.trim(); if(!val) return;
  state.checklist.push({id:uid('cat'), cat:val, items:[]});
  input.value=''; persist(); renderChecklist();
}

/* =====================================================================
   MEMO & SHOPPING
===================================================================== */
function addMemo(){
  const input = document.getElementById('memoInput');
  const val = input.value.trim(); if(!val) return;
  state.memos.unshift({text:val, ts:new Date().toLocaleString('ko-KR',{month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit'})});
  persist(); input.value=''; renderMemo();
}
function deleteMemo(i){ state.memos.splice(i,1); persist(); renderMemo(); }
function renderMemo(){
  const list = document.getElementById('memoList');
  if(state.memos.length===0){ list.innerHTML = `<div style="font-size:12.5px; color:var(--muted); text-align:center; padding:14px 0;">아직 저장된 메모가 없습니다.</div>`; return; }
  list.innerHTML = state.memos.map((m,i)=>`<div class="memo-item"><span class="memo-del" onclick="deleteMemo(${i})">×</span><div class="ts">${esc(m.ts)}</div>${esc(m.text)}</div>`).join('');
}
function renderShopping(){
  const body = document.getElementById('shoppingBody');
  let html = '';
  state.shopping.forEach((grp, gi)=>{
    html += `<div class="shop-city"><span class="shop-country-tag">${esc(grp.country)}</span>${esc(grp.city)}</div>`;
    grp.items.forEach((item, ii)=>{
      const key = gi+'-'+ii;
      const isOn = !!state.shopChecked[key];
      html += `<div class="chk-item">
        <div class="chk-box ${isOn?'on':''}" onclick="toggleShop('${key}')"><svg viewBox="0 0 24 24" fill="none" stroke-width="3"><path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7"/></svg></div>
        <div class="chk-label ${isOn?'done':''}" style="font-size:13px;">${esc(item)}</div>
      </div>`;
    });
    if(grp.loc) html += `<div style="font-size:11.5px; color:var(--muted); margin-top:6px; margin-bottom:4px;">📍 ${esc(grp.loc)}</div>`;
  });
  body.innerHTML = html;
}
function toggleShop(key){ state.shopChecked[key] = !state.shopChecked[key]; persist(); renderShopping(); }

/* =====================================================================
   SYNC (Google Apps Script Web App)
===================================================================== */
function saveWebhookUrl(){
  state.meta.webhookUrl = document.getElementById('webhookInput').value.trim();
  persist(); renderSyncStatus(); alert('저장되었습니다.');
}
function renderSyncStatus(){
  const el = document.getElementById('syncStatus');
  if(el) el.textContent = state.meta.lastSync ? ('마지막 동기화: '+new Date(state.meta.lastSync).toLocaleString('ko-KR')) : '아직 동기화한 적 없음';
  const input = document.getElementById('webhookInput');
  if(input && document.activeElement!==input) input.value = state.meta.webhookUrl||'';
}
function syncUpload(evt){
  if(!state.meta.webhookUrl){ alert('먼저 웹앱 URL을 입력하고 저장해주세요.'); return; }
  const btn = evt.target; const orig = btn.textContent; btn.disabled = true; btn.textContent = '업로드 중…';
  fetch(state.meta.webhookUrl, {method:'POST', headers:{'Content-Type':'text/plain;charset=utf-8'}, body:JSON.stringify(state)})
    .then(r=>r.json())
    .then(()=>{ state.meta.lastSync = new Date().toISOString(); persist(); renderSyncStatus(); alert('업로드 완료! 다른 가족 기기에서 같은 URL로 "지금 불러오기"를 누르면 받아볼 수 있어요.'); })
    .catch(err=>{ alert('업로드 실패: '+err.message+'\n웹앱 URL과 배포 설정(액세스 권한: 전체 허용)을 확인해주세요.'); })
    .finally(()=>{ btn.disabled=false; btn.textContent=orig; });
}
function syncDownload(){
  if(!state.meta.webhookUrl){ alert('먼저 웹앱 URL을 입력하고 저장해주세요.'); return; }
  fetch(state.meta.webhookUrl).then(r=>r.json()).then(data=>{
    if(!data || !data.days){ alert('서버에 저장된 데이터가 없습니다. 먼저 "지금 업로드"를 눌러주세요.'); return; }
    if(confirm('불러온 데이터로 이 기기 내용을 덮어씁니다. 계속할까요?')){
      const url = state.meta.webhookUrl;
      state = data;
      state.meta = state.meta || {};
      state.meta.webhookUrl = url;
      state.meta.lastSync = new Date().toISOString();
      persist(); location.reload();
    }
  }).catch(err=>alert('불러오기 실패: '+err.message));
}
function openSyncHelp(){
  openModal('구글시트 동기화 설정 방법',
`1) 구글 드라이브에서 새 스프레드시트를 만듭니다.
2) 상단 메뉴 확장 프로그램 > Apps Script 클릭.
3) 함께 제공된 GoogleAppsScript_Code.gs 내용 전체를 붙여넣기.
4) 우측 상단 배포 > 새 배포 선택.
   - 유형: 웹 앱
   - 실행 계정: 나
   - 액세스 권한: 전체 허용
5) 배포 후 나오는 웹 앱 URL을 복사해서 위 입력창에 붙여넣고 "URL 저장".
6) "지금 업로드"로 현재 데이터를 올리고, 다른 가족 기기에서는 같은 URL을 입력한 뒤 "지금 불러오기"로 받아오면 됩니다.

※ 자동 실시간 동기화가 아니라 수동 업로드/불러오기 방식입니다. 현지에서 일정을 바꾼 뒤에는 "지금 업로드"를 눌러주는 것을 잊지 마세요.`);
}

/* =====================================================================
   MODAL (info)
===================================================================== */
function openModal(title, body){
  document.getElementById('modalTitle').textContent = title;
  document.getElementById('modalBody').textContent = body;
  document.getElementById('modalBg').style.display = 'flex';
}
function closeModal(){ document.getElementById('modalBg').style.display = 'none'; }
function openFormModal(html){
  document.getElementById('formModalContent').innerHTML = html;
  document.getElementById('formModalBg').style.display = 'flex';
}
function closeFormModal(){ document.getElementById('formModalBg').style.display = 'none'; }

/* =====================================================================
   INIT
===================================================================== */
renderDashboard();
renderChecklist();
renderMemo();
renderShopping();
