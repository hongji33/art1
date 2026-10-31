// ==========================================
// 뱅크시와 고정관념 깨기 - 투표 서버 (Google Apps Script)
// 이 코드는 투표 결과 스프레드시트에 붙어 있는 Apps Script에 넣습니다.
// ==========================================

// 1. 설정 정보 ------------------------------------------
// 입장 비밀번호 (index.html 의 반 비밀번호와 반드시 같아야 합니다)
const CLASS_PASSWORDS = {
  "3-1": "1111",
  "3-2": "2222",
  "3-3": "3333",
  "3-4": "4444",
  "3-5": "5555",
  "3-6": "6666",
  "3-7": "7777",
  "3-8": "8888",
  "3-9": "9999",
  "3-10": "1010"
};

// 투표 데이터가 저장될 시트 이름 (스프레드시트 하단 탭 이름과 일치해야 합니다)
const SHEET_NAME = "시트1";

// ★ 시트 주소 확인용: 위쪽 함수 선택 칸에서 '시트주소보기'를 고르고 [실행] → [실행 로그]에 주소가 나옵니다.
function 시트주소보기() {
  const url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  Logger.log("투표 결과 시트 주소: " + url);
  return url;
}

// '3-2'가 날짜(3월 2일)로 바뀌어 저장된 칸도 다시 '3-2'로 읽어 줍니다.
function normalizeClass_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
    return (value.getMonth() + 1) + "-" + value.getDate();
  }
  return String(value).trim();
}

// 2. GET 요청 처리 (웹앱에서 데이터를 가져갈 때) ------------------
function doGet(e) {
  const action = e.parameter.action;
  const classId = e.parameter.classId;
  const password = e.parameter.password;

  // A. 학급 로그인 비밀번호 인증 검증
  if (action === "verify") {
    const expectedPassword = CLASS_PASSWORDS[classId];
    if (expectedPassword && password === expectedPassword) {
      return createJsonResponse({ status: "success", message: "인증 성공" });
    } else {
      return createJsonResponse({ status: "fail", message: "비밀번호 불일치" });
    }
  }

  // B. 실시간 투표 결과 데이터 반환 (모든 학급 데이터 통합)
  if (action === "getVotes") {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return createJsonResponse({});
    }

    const data = sheet.getDataRange().getValues();
    const result = {}; // { "3-1": { "q1": { "옵션A": 5 } }, ... }

    // 첫 번째 행(헤더: 일시, 학급, 질문ID, 선택안)을 제외하고 데이터 순회
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[1] || !row[2] || !row[3]) continue;
      const rowClass = normalizeClass_(row[1]); // B열: 학급 (예: 3-1)
      const rowTopic = String(row[2]);           // C열: 질문 ID (예: q1)
      const rowOption = String(row[3]);          // D열: 선택한 항목

      if (!result[rowClass]) result[rowClass] = {};
      if (!result[rowClass][rowTopic]) result[rowClass][rowTopic] = {};
      result[rowClass][rowTopic][rowOption] = (result[rowClass][rowTopic][rowOption] || 0) + 1;
    }

    return createJsonResponse(result);
  }

  return createJsonResponse({ status: "error", message: "올바르지 않은 명령입니다." });
}

// 3. POST 요청 처리 (웹앱에서 데이터를 저장/삭제할 때) -------------
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    const classId = postData.classId;
    const password = postData.password;

    // 패스워드 2차 보안 확인
    if (!CLASS_PASSWORDS[classId] || CLASS_PASSWORDS[classId] !== password) {
      return createJsonResponse({ status: "fail", message: "인증 실패" });
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return createJsonResponse({ status: "fail", message: "시트를 찾을 수 없음" });
    }

    // A. 새로운 투표 데이터 추가 저장
    if (action === "addVote") {
      const topicId = postData.topicId;
      const selectedOption = postData.selectedOption;
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        // 학급 칸(B열)을 '일반 텍스트'로 지정해 3-2가 날짜로 바뀌지 않게 함
        sheet.getRange("B:B").setNumberFormat("@");
        sheet.appendRow([new Date(), String(classId), topicId, selectedOption]);
      } finally {
        lock.releaseLock();
      }
      return createJsonResponse({ status: "success" });
    }

    // B. 특정 학급 투표 초기화 (리셋 버튼)
    if (action === "resetVotes") {
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        const data = sheet.getDataRange().getValues();
        // 아래에서부터 거꾸로 지워야 줄 번호가 꼬이지 않습니다.
        for (let i = data.length - 1; i >= 1; i--) {
          if (normalizeClass_(data[i][1]) === classId) {
            sheet.deleteRow(i + 1);
          }
        }
      } finally {
        lock.releaseLock();
      }
      return createJsonResponse({ status: "success" });
    }

    return createJsonResponse({ status: "error", message: "올바르지 않은 명령입니다." });
  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}

// JSON 응답 생성 공통 함수
function createJsonResponse(object) {
  return ContentService.createTextOutput(JSON.stringify(object))
                       .setMimeType(ContentService.MimeType.JSON);
}
