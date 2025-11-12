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
  EXCEL_FILE_NAME: '명함 신청 관리 DB.xlsx', // OneDrive Excel 파일 이름
  WORKSHEET_NAME: '신청목록', // 워크시트 이름
  ADMIN_EMAIL: 'bysong@law-lin.com',
  ADMIN_PASSWORD: 'admin123', // 실제 사용 시 변경 필요
  ONEDRIVE_FOLDER_PATH: '/명함 초안 파일' // OneDrive 폴더 경로
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
    var fileId = getExcelFileId(CONFIG.EXCEL_FILE_NAME);
    
    // 신청ID 생성 (LN-YYYY-NNNN 형식)
    var year = new Date().getFullYear();
    var sequence = getNextSequence(fileId);
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
    
    // Excel 파일에 데이터 추가
    addExcelRow(fileId, CONFIG.WORKSHEET_NAME, '신청목록테이블', rowData);
    
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
    var fileId = getExcelFileId(CONFIG.EXCEL_FILE_NAME);
    var data = readExcelRange(fileId, CONFIG.WORKSHEET_NAME, 'A:M');
    
    // 헤더 행 제외하고 검색
    for (var i = 1; i < data.length; i++) {
      if (data[i] && data[i][0] === applicationId && data[i][2] === email) {
        return {
          success: true,
          status: data[i][1] || '',
          applicationId: data[i][0],
          registeredDate: data[i][9] || '',
          message: "귀하의 명함은 현재 '" + (data[i][1] || '') + "' 상태입니다."
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
    var fileId = getExcelFileId(CONFIG.EXCEL_FILE_NAME);
    var data = readExcelRange(fileId, CONFIG.WORKSHEET_NAME, 'A:M');
    
    for (var i = 1; i < data.length; i++) {
      if (data[i] && data[i][0] === applicationId) {
        return {
          success: true,
          applicationId: data[i][0],
          status: data[i][1] || '',
          email: data[i][2] || '',
          name: data[i][3] || '',
          fileUrl: data[i][10] || '',
          canApprove: (data[i][1] === '초안전달' || data[i][1] === '수정요청')
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
    var fileId = getExcelFileId(CONFIG.EXCEL_FILE_NAME);
    var data = readExcelRange(fileId, CONFIG.WORKSHEET_NAME, 'A:M');
    var rowIndex = -1;
    
    // 신청 내역 찾기
    for (var i = 1; i < data.length; i++) {
      if (data[i] && data[i][0] === applicationId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: '신청 내역을 찾을 수 없습니다.' };
    }
    
    // 상태 변경 (B열 = 인덱스 1)
    var cellAddress = 'B' + rowIndex;
    updateExcelCell(fileId, CONFIG.WORKSHEET_NAME, cellAddress, '제작');
    
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
    var fileId = getExcelFileId(CONFIG.EXCEL_FILE_NAME);
    var data = readExcelRange(fileId, CONFIG.WORKSHEET_NAME, 'A:M');
    var rowIndex = -1;
    
    // 신청 내역 찾기
    for (var i = 1; i < data.length; i++) {
      if (data[i] && data[i][0] === applicationId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: '신청 내역을 찾을 수 없습니다.' };
    }
    
    // 상태 변경 (B열)
    var statusCell = 'B' + rowIndex;
    updateExcelCell(fileId, CONFIG.WORKSHEET_NAME, statusCell, '수정요청');
    
    // 비고에 수정 사유 추가 (I열 = 인덱스 8)
    var currentNotes = data[rowIndex - 1][8] || '';
    var newNotes = currentNotes + '\n[수정요청] ' + reason;
    var notesCell = 'I' + rowIndex;
    updateExcelCell(fileId, CONFIG.WORKSHEET_NAME, notesCell, newNotes);
    
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
    var fileId = getExcelFileId(CONFIG.EXCEL_FILE_NAME);
    var data = readExcelRange(fileId, CONFIG.WORKSHEET_NAME, 'A:M');
    var result = [];
    
    // 헤더 제외하고 데이터 변환
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row || !row[0]) continue; // 빈 행 건너뛰기
      
      // 상태 필터 적용
      if (statusFilter && statusFilter !== '전체' && row[1] !== statusFilter) {
        continue;
      }
      
      result.push({
        applicationId: row[0] || '',
        status: row[1] || '',
        email: row[2] || '',
        name: row[3] || '',
        quantity: row[4] || '',
        isSame: row[5] || '',
        isLawyer: row[6] || '',
        lawyerName: row[7] || '',
        notes: row[8] || '',
        registeredDate: row[9] || '',
        fileUrl: row[10] || '',
        processor: row[11] || '',
        processedDate: row[12] || ''
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
    
    // OneDrive에 파일 업로드
    var uploadFileName = fileName || '초안_' + applicationId;
    var uploadedFile = uploadFileToOneDrive(blob, uploadFileName, CONFIG.ONEDRIVE_FOLDER_PATH);
    var fileUrl = uploadedFile.shareUrl || uploadedFile.webUrl || '';
    
    // Excel 파일에 파일 링크 저장
    var fileId = getExcelFileId(CONFIG.EXCEL_FILE_NAME);
    var data = readExcelRange(fileId, CONFIG.WORKSHEET_NAME, 'A:M');
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i] && data[i][0] === applicationId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: '신청 내역을 찾을 수 없습니다.' };
    }
    
    // 파일 링크 저장 (K열 = 인덱스 10)
    var fileUrlCell = 'K' + rowIndex;
    updateExcelCell(fileId, CONFIG.WORKSHEET_NAME, fileUrlCell, fileUrl);
    
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
    var fileId = getExcelFileId(CONFIG.EXCEL_FILE_NAME);
    var data = readExcelRange(fileId, CONFIG.WORKSHEET_NAME, 'A:M');
    var rowIndex = -1;
    var applicationData = null;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i] && data[i][0] === applicationId) {
        rowIndex = i + 1;
        applicationData = data[i];
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: '신청 내역을 찾을 수 없습니다.' };
    }
    
    // 파일 링크 확인
    var fileUrl = applicationData[10] || '';
    if (!fileUrl) {
      return { success: false, message: '초안 파일이 첨부되지 않았습니다.' };
    }
    
    // 상태를 '초안전달'로 변경 (B열)
    var statusCell = 'B' + rowIndex;
    updateExcelCell(fileId, CONFIG.WORKSHEET_NAME, statusCell, '초안전달');
    
    // 신청자에게 초안 확인 요청 이메일 발송
    sendDraftReviewEmail(applicationId, applicationData[2] || '', applicationData[3] || '', fileUrl);
    
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
    var fileId = getExcelFileId(CONFIG.EXCEL_FILE_NAME);
    var data = readExcelRange(fileId, CONFIG.WORKSHEET_NAME, 'A:M');
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i] && data[i][0] === applicationId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: '신청 내역을 찾을 수 없습니다.' };
    }
    
    // 상태 변경 (B열)
    var statusCell = 'B' + rowIndex;
    updateExcelCell(fileId, CONFIG.WORKSHEET_NAME, statusCell, newStatus);
    
    // 처리자 및 처리일자 기록
    var today = new Date();
    var dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    var processorCell = 'L' + rowIndex;
    var processedDateCell = 'M' + rowIndex;
    updateExcelCell(fileId, CONFIG.WORKSHEET_NAME, processorCell, Session.getActiveUser().getEmail());
    updateExcelCell(fileId, CONFIG.WORKSHEET_NAME, processedDateCell, dateStr);
    
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
 * 다음 시퀀스 번호 가져오기
 */
function getNextSequence(fileId) {
  try {
    var data = readExcelRange(fileId, CONFIG.WORKSHEET_NAME, 'A:A');
    
    if (data.length <= 1) {
      return 1001;
    }
    
    // 마지막 신청ID에서 시퀀스 추출
    var lastId = data[data.length - 1][0];
    if (lastId && lastId.toString().match(/LN-\d{4}-(\d+)/)) {
      var match = lastId.toString().match(/LN-\d{4}-(\d+)/);
      return parseInt(match[1]) + 1;
    }
    
    return 1001;
  } catch (error) {
    Logger.log('Error in getNextSequence: ' + error.toString());
    return 1001;
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

