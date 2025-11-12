/**
 * 명함 신청 관리 시스템 - Google Apps Script 메인 코드
 * 
 * 주요 기능:
 * - 신청 데이터 처리
 * - 이메일 알림 발송
 * - 상태 변경 관리
 * - 파일 업로드 처리
 */

// 설정 변수
const CONFIG = {
  SHEET_NAME: '신청목록',
  ADMIN_EMAIL: 'bysong@law-lin.com',
  ADMIN_PASSWORD: 'admin123', // 실제 사용 시 변경 필요
  DRIVE_FOLDER_NAME: '명함 초안 파일'
};

/**
 * 웹 앱 진입점 - 사용자 신청 폼
 */
function doGet(e) {
  var page = e.parameter.page || 'form';
  
  if (page === 'form') {
    return HtmlService.createTemplateFromFile('UserForm')
        .evaluate()
        .setTitle('명함 신청')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === 'check') {
    return HtmlService.createTemplateFromFile('CheckStatus')
        .evaluate()
        .setTitle('신청 내역 조회')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === 'review') {
    var applicationId = e.parameter.id;
    if (!applicationId) {
      return HtmlService.createHtmlOutput('잘못된 접근입니다.');
    }
    var template = HtmlService.createTemplateFromFile('ReviewDraft');
    template.applicationId = applicationId;
    return template.evaluate()
        .setTitle('명함 초안 확인')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === 'admin') {
    return HtmlService.createTemplateFromFile('AdminPage')
        .evaluate()
        .setTitle('관리자 페이지')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  return HtmlService.createHtmlOutput('페이지를 찾을 수 없습니다.');
}

/**
 * 신청 데이터 제출 처리
 */
function submitApplication(data) {
  try {
    var sheet = getSheet();
    var lastRow = sheet.getLastRow();
    
    // 신청ID 생성 (LN-YYYY-NNNN 형식)
    var year = new Date().getFullYear();
    var sequence = getNextSequence();
    var applicationId = 'LN-' + year + '-' + String(sequence).padStart(4, '0');
    
    // 현재 날짜
    var today = new Date();
    var dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    
    // 데이터 행 구성
    var rowData = [
      applicationId,                    // A: 신청ID
      '신청',                           // B: Status
      data.email,                       // C: 신청자계정
      data.name,                        // D: 등록자
      data.quantity,                    // E: 통
      data.isSame,                      // F: 기존동일여부
      data.isLawyer,                    // G: 변호사여부
      data.lawyerName || '',            // H: 담당변호사
      data.notes || '',                 // I: 비고
      dateStr,                          // J: 등록일
      '',                               // K: 첨부파일
      '',                               // L: 처리자
      ''                                // M: 처리일자
    ];
    
    // 시트에 데이터 추가
    sheet.appendRow(rowData);
    
    // 이메일 알림 발송
    sendNewApplicationEmails(applicationId, data);
    
    return {
      success: true,
      applicationId: applicationId,
      message: '신청이 완료되었습니다.'
    };
  } catch (error) {
    Logger.log('Error in submitApplication: ' + error.toString());
    return {
      success: false,
      message: '신청 처리 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 신청 내역 조회
 */
function checkApplicationStatus(email, applicationId) {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    
    // 헤더 행 제외하고 검색
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === applicationId && data[i][2] === email) {
        return {
          success: true,
          status: data[i][1],
          applicationId: data[i][0],
          registeredDate: data[i][9],
          message: "귀하의 명함은 현재 '" + data[i][1] + "' 상태입니다."
        };
      }
    }
    
    return {
      success: false,
      message: '일치하는 신청 내역을 찾을 수 없습니다.'
    };
  } catch (error) {
    Logger.log('Error in checkApplicationStatus: ' + error.toString());
    return {
      success: false,
      message: '조회 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 초안 확인 페이지 데이터 조회
 */
function getApplicationForReview(applicationId) {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === applicationId) {
        return {
          success: true,
          applicationId: data[i][0],
          status: data[i][1],
          email: data[i][2],
          name: data[i][3],
          fileUrl: data[i][10],
          canApprove: data[i][1] === '초안전달' || data[i][1] === '수정요청'
        };
      }
    }
    
    return {
      success: false,
      message: '신청 내역을 찾을 수 없습니다.'
    };
  } catch (error) {
    Logger.log('Error in getApplicationForReview: ' + error.toString());
    return {
      success: false,
      message: '조회 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 제작 승인 처리
 */
function approveProduction(applicationId) {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    // 신청 내역 찾기
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === applicationId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: '신청 내역을 찾을 수 없습니다.' };
    }
    
    // 상태 변경
    sheet.getRange(rowIndex, 2).setValue('제작');
    
    // 관리자에게 알림 메일 발송
    var applicationData = data[rowIndex - 1];
    sendProductionApprovalEmail(applicationId, applicationData);
    
    return {
      success: true,
      message: '제작 승인이 완료되었습니다.'
    };
  } catch (error) {
    Logger.log('Error in approveProduction: ' + error.toString());
    return {
      success: false,
      message: '처리 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 수정 요청 처리
 */
function requestModification(applicationId, reason) {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    // 신청 내역 찾기
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === applicationId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: '신청 내역을 찾을 수 없습니다.' };
    }
    
    // 상태 변경
    sheet.getRange(rowIndex, 2).setValue('수정요청');
    
    // 비고에 수정 사유 추가
    var currentNotes = sheet.getRange(rowIndex, 9).getValue();
    var newNotes = currentNotes + '\n[수정요청] ' + reason;
    sheet.getRange(rowIndex, 9).setValue(newNotes);
    
    // 관리자에게 알림 메일 발송
    var applicationData = data[rowIndex - 1];
    sendModificationRequestEmail(applicationId, applicationData, reason);
    
    return {
      success: true,
      message: '수정 요청이 완료되었습니다.'
    };
  } catch (error) {
    Logger.log('Error in requestModification: ' + error.toString());
    return {
      success: false,
      message: '처리 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 관리자 로그인 인증
 */
function adminLogin(password) {
  return password === CONFIG.ADMIN_PASSWORD;
}

/**
 * 관리자 - 모든 신청 목록 조회
 */
function getAdminApplicationList(statusFilter) {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var result = [];
    
    // 헤더 제외하고 데이터 변환
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // 상태 필터 적용
      if (statusFilter && statusFilter !== '전체' && row[1] !== statusFilter) {
        continue;
      }
      
      result.push({
        applicationId: row[0],
        status: row[1],
        email: row[2],
        name: row[3],
        quantity: row[4],
        isSame: row[5],
        isLawyer: row[6],
        lawyerName: row[7],
        notes: row[8],
        registeredDate: row[9],
        fileUrl: row[10],
        processor: row[11],
        processedDate: row[12]
      });
    }
    
    return {
      success: true,
      data: result
    };
  } catch (error) {
    Logger.log('Error in getAdminApplicationList: ' + error.toString());
    return {
      success: false,
      message: '목록 조회 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 관리자 - 초안 파일 첨부 및 상태 변경
 */
function adminAttachDraft(applicationId, fileData, fileName, mimeType) {
  try {
    // Base64 데이터를 Blob으로 변환
    var base64Data = fileData.split(',')[1] || fileData; // data:image/png;base64, 부분 제거
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    
    // Google Drive에 파일 업로드
    var folder = getOrCreateDriveFolder();
    var file = folder.createFile(blob);
    file.setName(fileName || '초안_' + applicationId);
    
    // 공유 설정 (링크로 접근 가능하도록)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = file.getUrl();
    
    // 시트에 파일 링크 저장
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === applicationId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: '신청 내역을 찾을 수 없습니다.' };
    }
    
    // 파일 링크 저장
    sheet.getRange(rowIndex, 11).setValue(fileUrl);
    
    return {
      success: true,
      fileUrl: fileUrl,
      message: '파일이 업로드되었습니다.'
    };
  } catch (error) {
    Logger.log('Error in adminAttachDraft: ' + error.toString());
    return {
      success: false,
      message: '파일 업로드 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 관리자 - 초안 전달 처리
 */
function adminSendDraft(applicationId) {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    var applicationData = null;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === applicationId) {
        rowIndex = i + 1;
        applicationData = data[i];
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: '신청 내역을 찾을 수 없습니다.' };
    }
    
    // 파일 링크 확인
    var fileUrl = applicationData[10];
    if (!fileUrl) {
      return { success: false, message: '초안 파일이 첨부되지 않았습니다.' };
    }
    
    // 상태를 '초안전달'로 변경
    sheet.getRange(rowIndex, 2).setValue('초안전달');
    
    // 신청자에게 초안 확인 요청 이메일 발송
    sendDraftReviewEmail(applicationId, applicationData[2], applicationData[3], fileUrl);
    
    return {
      success: true,
      message: '초안이 전달되었습니다.'
    };
  } catch (error) {
    Logger.log('Error in adminSendDraft: ' + error.toString());
    return {
      success: false,
      message: '처리 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 관리자 - 상태 변경
 */
function adminChangeStatus(applicationId, newStatus) {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === applicationId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: '신청 내역을 찾을 수 없습니다.' };
    }
    
    // 상태 변경
    sheet.getRange(rowIndex, 2).setValue(newStatus);
    
    // 처리자 및 처리일자 기록
    var today = new Date();
    var dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(rowIndex, 12).setValue(Session.getActiveUser().getEmail());
    sheet.getRange(rowIndex, 13).setValue(dateStr);
    
    return {
      success: true,
      message: '상태가 변경되었습니다.'
    };
  } catch (error) {
    Logger.log('Error in adminChangeStatus: ' + error.toString());
    return {
      success: false,
      message: '처리 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

// ==================== 유틸리티 함수 ====================

/**
 * 시트 가져오기
 */
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    // 시트가 없으면 생성
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    // 헤더 설정
    var headers = [
      '신청ID', 'Status', '신청자계정', '등록자', '통', 
      '기존동일여부', '변호사여부', '담당변호사', '비고', 
      '등록일', '첨부파일', '처리자', '처리일자'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  
  return sheet;
}

/**
 * 다음 시퀀스 번호 가져오기
 */
function getNextSequence() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return 1001;
  }
  
  // 마지막 신청ID에서 시퀀스 추출
  var lastId = sheet.getRange(lastRow, 1).getValue();
  if (lastId && lastId.toString().match(/LN-\d{4}-(\d+)/)) {
    var match = lastId.toString().match(/LN-\d{4}-(\d+)/);
    return parseInt(match[1]) + 1;
  }
  
  return 1001;
}

/**
 * Drive 폴더 가져오기 또는 생성
 */
function getOrCreateDriveFolder() {
  var folders = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(CONFIG.DRIVE_FOLDER_NAME);
  }
}

// ==================== 이메일 발송 함수 ====================

/**
 * 새 신청 알림 이메일 발송
 */
function sendNewApplicationEmails(applicationId, data) {
  // 관리자에게 알림
  var adminSubject = '[명함 신청] 새 명함 신청이 접수되었습니다';
  var adminBody = '새로운 명함 신청이 접수되었습니다.\n\n' +
                  '신청ID: ' + applicationId + '\n' +
                  '신청자: ' + data.name + '\n' +
                  '이메일: ' + data.email + '\n' +
                  '신청 통 수: ' + data.quantity + '\n' +
                  '기존 동일 여부: ' + data.isSame + '\n' +
                  '변호사 여부: ' + data.isLawyer + '\n' +
                  (data.lawyerName ? '담당 변호사: ' + data.lawyerName + '\n' : '') +
                  (data.notes ? '비고: ' + data.notes + '\n' : '');
  
  MailApp.sendEmail({
    to: CONFIG.ADMIN_EMAIL,
    subject: adminSubject,
    body: adminBody
  });
  
  // 신청자에게 확인 이메일
  var userSubject = '[명함 신청] 신청이 접수되었습니다';
  var userBody = data.name + '님,\n\n' +
                 '명함 신청이 정상적으로 접수되었습니다.\n\n' +
                 '신청ID: ' + applicationId + '\n' +
                 '신청 상태는 아래 링크에서 확인하실 수 있습니다.\n\n' +
                 '신청 내역 조회: ' + getWebAppUrl() + '?page=check';
  
  MailApp.sendEmail({
    to: data.email,
    subject: userSubject,
    body: userBody
  });
}

/**
 * 초안 확인 요청 이메일 발송
 */
function sendDraftReviewEmail(applicationId, email, name, fileUrl) {
  var reviewUrl = getWebAppUrl() + '?page=review&id=' + applicationId;
  
  var subject = '[명함 신청] 명함 초안 확인 요청';
  var body = name + '님,\n\n' +
             '명함 초안이 준비되었습니다. 아래 링크에서 확인해주세요.\n\n' +
             '초안 확인: ' + reviewUrl + '\n\n' +
             '초안 파일: ' + fileUrl + '\n\n' +
             '초안을 확인하신 후 제작 승인 또는 수정 요청을 선택해주세요.';
  
  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: body
  });
}

/**
 * 제작 승인 알림 이메일 발송
 */
function sendProductionApprovalEmail(applicationId, applicationData) {
  var subject = '[명함 신청] 제작 승인 알림';
  var body = '명함 제작이 승인되었습니다.\n\n' +
             '신청ID: ' + applicationId + '\n' +
             '신청자: ' + applicationData[3] + '\n' +
             '이메일: ' + applicationData[2] + '\n' +
             '제작을 진행해주세요.';
  
  MailApp.sendEmail({
    to: CONFIG.ADMIN_EMAIL,
    subject: subject,
    body: body
  });
}

/**
 * 수정 요청 알림 이메일 발송
 */
function sendModificationRequestEmail(applicationId, applicationData, reason) {
  var subject = '[명함 신청] 수정 요청 알림';
  var body = '명함 초안에 대한 수정 요청이 접수되었습니다.\n\n' +
             '신청ID: ' + applicationId + '\n' +
             '신청자: ' + applicationData[3] + '\n' +
             '이메일: ' + applicationData[2] + '\n' +
             '수정 사유: ' + reason + '\n\n' +
             '초안을 수정하여 다시 전달해주세요.';
  
  MailApp.sendEmail({
    to: CONFIG.ADMIN_EMAIL,
    subject: subject,
    body: body
  });
}

/**
 * 웹 앱 URL 가져오기
 */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * HTML 파일 포함 (스타일/스크립트 공유용)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

