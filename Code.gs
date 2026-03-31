// Google Apps Script - 웹앱 API 서버 역할
// 이 파일을 Google Apps Script 에디터에 붙여넣기

const SHEET_ID = '여기에_구글시트_ID_입력';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}

// 슬롯 목록 + 잔여석 계산
function getSlots() {
  const ss = getSpreadsheet();
  const slotSheet = ss.getSheetByName('슬롯');
  const regSheet = ss.getSheetByName('신청');

  const slotData = slotSheet.getDataRange().getValues();
  const regData = regSheet.getDataRange().getValues();

  const slots = [];
  // 슬롯 시트: 슬롯ID, 장소, 날짜, 시간, 최대인원, 오픈채팅링크
  for (let i = 1; i < slotData.length; i++) {
    const row = slotData[i];
    const slotId = row[0];
    const count = regData.filter(r => r[5] === slotId).length; // 5번 컬럼(0-indexed) = 선택한 슬롯ID
    slots.push({
      id: slotId,
      location: row[1],
      date: row[2],
      time: row[3],
      maxPeople: row[4],
      chatLink: row[5],
      currentCount: count,
      remaining: row[4] - count,
      isFull: count >= row[4]
    });
  }
  return slots;
}

// 신청 등록
function submitRegistration(data) {
  const ss = getSpreadsheet();
  const regSheet = ss.getSheetByName('신청');

  // 다시 한번 잔여석 확인 (동시 신청 방지)
  const slots = getSlots();
  const selectedSlot = slots.find(s => s.id === data.slotId);
  if (!selectedSlot || selectedSlot.isFull) {
    return { success: false, message: '선택한 슬롯이 마감되었습니다.' };
  }

  regSheet.appendRow([
    new Date(),         // 신청일시
    data.name,          // 이름
    data.contact,       // 연락처
    data.os,            // 노트북 OS
    data.experience,    // 코딩 경험
    data.slotId,        // 선택한 슬롯
    data.serviceType,   // 만들고 싶은 서비스
    data.blocker || '', // 막히는 부분
    data.memo || ''     // 기타 요청
  ]);

  return {
    success: true,
    chatLink: selectedSlot.chatLink,
    slotInfo: `${selectedSlot.date} ${selectedSlot.time} - ${selectedSlot.location}`
  };
}

// HTTP GET 요청 처리 (슬롯 조회)
function doGet(e) {
  const result = getSlots();
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// HTTP POST 요청 처리 (신청 등록)
function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const result = submitRegistration(data);
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
