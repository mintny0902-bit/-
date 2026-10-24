/**
 * Google Apps Script (GAS) 연동 모듈 (submit.js)
 * 
 * [개인정보 완전 익명 보장 원칙]
 * 학생의 이름, 번호, 생년월일, IP 등 일체의 개인정보는 수집하거나 전송하지 않습니다.
 * 전송 데이터: [제출 시각, 총점, 총 문항 수, 각 문항별 정오답 여부]
 */

const GAS_CONFIG = {
  /**
   * 구글 스프레드시트와 연동된 구글 앱스 스크립트 웹 앱 URL을 아래에 입력하세요.
   * 비워두거나 기본값일 경우 시뮬레이션 모드로 안전하게 작동하여 오류 없이 결과를 보여줍니다.
   * 예: "https://script.google.com/macros/s/AKfycbxAbCdEfGhIjKlMnOpQrStUvWxYz/exec"
   */
  webAppUrl: "YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE",
  
  // 요청 제한 시간 (밀리초)
  timeoutMs: 8000
};

/**
 * 퀴즈 결과 익명 통계 전송 함수
 * @param {Object} quizResult
 * @param {number} quizResult.score - 맞힌 문항 수 (예: 8)
 * @param {number} quizResult.total - 총 문항 수 (예: 10)
 * @param {Array<{qId: number, topic: string, selectedIndex: number, correctIndex: number, isCorrect: boolean}>} quizResult.details - 문항별 정오답
 * @returns {Promise<{success: boolean, message: string, simulated: boolean, timestamp: string}>}
 */
async function submitQuizResult(quizResult) {
  const now = new Date();
  const timestamp = now.getFullYear() + "-" +
    String(now.getMonth() + 1).padStart(2, "0") + "-" +
    String(now.getDate()).padStart(2, "0") + " " +
    String(now.getHours()).padStart(2, "0") + ":" +
    String(now.getMinutes()).padStart(2, "0") + ":" +
    String(now.getSeconds()).padStart(2, "0");

  // 완전 익명 통계 페이로드 구성 (개인정보 없음)
  const payload = {
    timestamp: timestamp,
    score: quizResult.score,
    total: quizResult.total,
    percentage: Math.round((quizResult.score / quizResult.total) * 100),
    answers: quizResult.details.map((item, idx) => ({
      questionNum: idx + 1,
      topic: item.topic,
      isCorrect: item.isCorrect ? "O" : "X",
      selectedChoice: item.selectedIndex + 1,
      correctChoice: item.correctIndex + 1
    }))
  };

  const targetUrl = (GAS_CONFIG.webAppUrl || "").trim();

  // 1. GAS URL이 설정되지 않은 경우 (데모 및 시뮬레이션 모드)
  if (!targetUrl || targetUrl.includes("YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE") || !targetUrl.startsWith("http")) {
    console.info("ℹ️ [GAS 연동 안내] GAS 웹 앱 URL이 설정되지 않아 브라우저 시뮬레이션 모드로 결과를 처리합니다.");
    console.table(payload.answers);
    
    // 자연스러운 전송 UX를 위한 0.8초 딜레이
    await new Promise(resolve => setTimeout(resolve, 800));
    
    return {
      success: true,
      simulated: true,
      timestamp: timestamp,
      message: "시뮬레이션 모드로 전송되었습니다. (선생님용 GAS URL을 설정하면 구글 시트에 자동 저장됩니다.)"
    };
  }

  // 2. 실제 GAS 웹 앱으로 전송 (CORS 이슈를 방지하기 위해 mode: 'no-cors' 및 text/plain 사용)
  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), GAS_CONFIG.timeoutMs);

    await fetch(targetUrl, {
      method: "POST",
      mode: "no-cors",
      cache: "no-cache",
      headers: {
        "Content-Type": "text/plain;charset=utf-8"
      },
      body: JSON.stringify(payload),
      signal: controller.signal
    });

    clearTimeout(timeoutId);

    return {
      success: true,
      simulated: false,
      timestamp: timestamp,
      message: "결과가 구글 스프레드시트에 안전하게 익명으로 저장되었습니다."
    };
  } catch (error) {
    console.error("❌ 구글 앱스 스크립트 전송 중 오류 발생:", error);
    return {
      success: false,
      simulated: false,
      timestamp: timestamp,
      message: "결과 전송 중 네트워크 오류가 발생했으나, 퀴즈 결과 확인에는 지장이 없습니다."
    };
  }
}

// 브라우저 전역 등록
if (typeof window !== "undefined") {
  window.submitQuizResult = submitQuizResult;
  window.GAS_CONFIG = GAS_CONFIG;
}

/**
 * ============================================================================
 * [선생님 및 관리자용 안내: 구글 앱스 스크립트(GAS) 백엔드 코드]
 * ============================================================================
 * 1. 구글 드라이브에서 새 '구글 스프레드시트'를 생성합니다.
 * 2. 상단 메뉴 [확장 프로그램] -> [Apps Script]를 클릭합니다.
 * 3. 기존 코드를 모두 지우고 아래의 Code.gs 내용을 복사하여 붙여넣고 저장(Ctrl+S)합니다.
 * 
 * --- 아래 코드를 Apps Script에 붙여넣으세요 ---
 * 
 * function doPost(e) {
 *   try {
 *     var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 *     
 *     // 첫 행(헤더)이 없으면 자동 생성
 *     if (sheet.getLastRow() === 0) {
 *       var headers = ["제출 시각", "총점", "총 문항수", "백분율(%)"];
 *       for (var i = 1; i <= 10; i++) {
 *         headers.push("문항" + i + "(정오답)");
 *       }
 *       sheet.appendRow(headers);
 *       sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#EEF2FF");
 *     }
 *     
 *     var data = JSON.parse(e.postData.contents);
 *     var row = [
 *       data.timestamp,
 *       data.score,
 *       data.total,
 *       data.percentage + "%"
 *     ];
 *     
 *     // 문항별 정오답 (O 또는 X)
 *     if (data.answers && Array.isArray(data.answers)) {
 *       data.answers.forEach(function(ans) {
 *         row.push(ans.isCorrect);
 *       });
 *     }
 *     
 *     sheet.appendRow(row);
 *     
 *     return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
 *       .setMimeType(ContentService.MimeType.JSON);
 *   } catch (err) {
 *     return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
 *       .setMimeType(ContentService.MimeType.JSON);
 *   }
 * }
 * 
 * 4. 오른쪽 상단 [배포] -> [새 배포] 클릭
 * 5. 유형 선택: [웹 앱]
 *    - 설명: 디지털 시민성 퀴즈 결과 수집
 *    - 다음 사용자로 실행: [나]
 *    - 액세스 권한이 있는 사용자: [모든 사용자(Anyone)]  <-- 필수!
 * 6. [배포] 버튼 클릭 후 생성된 '웹 앱 URL'을 복사하여 위 GAS_CONFIG.webAppUrl 에 넣으세요.
 * ============================================================================
 */
