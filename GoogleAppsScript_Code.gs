/**
 * EuroTrip 2026 대시보드 - 동기화 백엔드
 *
 * 사용법
 * 1) 구글 드라이브에서 새 스프레드시트를 만듭니다. (내용은 비워둬도 됩니다)
 * 2) 상단 메뉴 [확장 프로그램] > [Apps Script] 클릭
 * 3) 기본으로 열려 있는 코드를 모두 지우고, 이 파일 내용 전체를 붙여넣기
 * 4) 우측 상단 [배포] > [새 배포] 클릭
 *      - 유형 선택(톱니바퀴): 웹 앱
 *      - 설명: 아무거나(예: euro2026 sync)
 *      - 실행 계정: 나
 *      - 액세스 권한이 있는 사용자: 전체 허용(anyone)
 * 5) [배포] 클릭 후 나오는 "웹 앱 URL"을 복사
 * 6) EuroTrip2026_Dashboard.html 의 대시보드 탭 > "가족 동기화" 카드에
 *    복사한 URL을 붙여넣고 "URL 저장" 클릭
 * 7) "지금 업로드"로 현재 기기의 데이터를 서버에 저장하고,
 *    다른 가족 구성원은 같은 URL을 입력한 뒤 "지금 불러오기"로 받아오면 됩니다.
 *
 * 참고: 자동 실시간 동기화가 아니라, 버튼을 눌러 주고받는 수동 동기화 방식입니다.
 * PropertiesService에 JSON 문자열 전체를 저장하는 간단한 key-value 저장 방식이라
 * 별도의 시트 구조 설정 없이 바로 동작합니다.
 */

const STORAGE_KEY = 'euro2026_trip_data';

function doGet(e) {
  const saved = PropertiesService.getScriptProperties().getProperty(STORAGE_KEY);
  const payload = saved || '{}';
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = e && e.postData ? e.postData.contents : '{}';
    // 유효한 JSON인지 확인 후 저장 (깨진 데이터로 덮어쓰는 것을 방지)
    JSON.parse(body);
    PropertiesService.getScriptProperties().setProperty(STORAGE_KEY, body);
    const result = { status: 'ok', savedAt: new Date().toISOString() };
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    const result = { status: 'error', message: String(err) };
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  }
}
