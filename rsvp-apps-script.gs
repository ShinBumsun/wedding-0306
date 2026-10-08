/**
 * 청첩장 RSVP → 구글 시트 저장용 Apps Script
 *
 * 1. 구글 시트 새로 만들기 (sheets.new)
 * 2. 메뉴 [확장 프로그램] → [Apps Script] → 기존 코드 지우고 이 파일 내용 붙여넣기 → 저장
 * 3. 오른쪽 위 [배포] → [새 배포] → 유형 선택(톱니바퀴) [웹 앱]
 *      - 다음 사용자로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 *    → [배포] → 권한 승인 → 나오는 "웹 앱 URL"(https://script.google.com/macros/s/.../exec) 복사
 * 4. index.html 의 CONFIG.rsvpUrl 에 붙여넣기
 */
function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['제출시각', '구분', '참석여부', '식사', '성함', '동반인원', '연락처', '메시지']);
    sheet.setFrozenRows(1);
  }
  const p = e.parameter;
  sheet.appendRow([new Date(), p.side, p.attend, p.meal, p.name, p.companions, "'" + (p.phone || ''), p.message]);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
