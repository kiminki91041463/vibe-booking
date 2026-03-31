// Google Apps Script - 웹앱 API 서버 역할
// 이 파일을 Google Apps Script 에디터에 붙여넣기

const SHEET_ID = '1dk07mZjdnky0HvV1PFVUFSddhNedlWE--RRbDaJ-QLo';

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
    const count = regData.filter((r) => r[5] === slotId).length; // 5번 컬럼(0-indexed) = 선택한 슬롯ID
    // 날짜 문자열 "4/8 (수)" → "2026-04-08" 형식으로 변환
    const dateStr = String(row[2]);
    const match = dateStr.match(/(\d+)\/(\d+)/);
    const dateKey = match
      ? `2026-${match[1].padStart(2, '0')}-${match[2].padStart(2, '0')}`
      : '';

    slots.push({
      id: slotId,
      location: row[1],
      dateKey: dateKey,
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
  const selectedSlot = slots.find((s) => s.id === data.slotId);
  if (!selectedSlot || selectedSlot.isFull) {
    return { success: false, message: '선택한 슬롯이 마감되었습니다.' };
  }

  regSheet.appendRow([
    new Date(), // 신청일시
    data.name, // 이름
    data.contact, // 연락처
    data.os, // 노트북 OS
    data.experience, // 코딩 경험
    data.slotId, // 선택한 슬롯
    data.serviceType, // 만들고 싶은 서비스
    data.blocker || '', // 막히는 부분
    data.memo || '' // 기타 요청
  ]);

  return {
    success: true,
    chatLink: selectedSlot.chatLink,
    slotInfo: `${selectedSlot.date} ${selectedSlot.time} - ${selectedSlot.location}`
  };
}

// HTTP GET 요청 처리 (슬롯 조회 + 신청 등록)
function doGet(e) {
  const action = e.parameter.action;

  if (action === 'submit') {
    const data = {
      name: e.parameter.name,
      contact: e.parameter.contact,
      os: e.parameter.os,
      experience: e.parameter.experience,
      serviceType: e.parameter.serviceType,
      blocker: e.parameter.blocker || '',
      slotId: e.parameter.slotId,
      memo: e.parameter.memo || ''
    };
    const result = submitRegistration(data);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(
      ContentService.MimeType.JSON
    );
  }

  // 기본: 슬롯 조회
  const result = getSlots();
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// HTTP POST 요청 처리 (폴백)
function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const result = submitRegistration(data);
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(
    ContentService.MimeType.JSON
  );
}
