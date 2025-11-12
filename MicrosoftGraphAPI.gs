/**
 * Microsoft Graph API 통신 모듈
 * OneDrive Excel 파일 접근을 위한 함수들
 */

// Microsoft Graph API 설정
const MS_GRAPH_CONFIG = {
  CLIENT_ID: '5b165094-6f98-4818-9667-f81e651ce7d0',
  TENANT_ID: '20c53626-8c84-4cb2-b049-ae1f603445a6',
  CLIENT_SECRET: '', // PropertiesService에서 가져오거나 직접 설정
  GRAPH_API_ENDPOINT: 'https://graph.microsoft.com/v1.0',
  TOKEN_CACHE_KEY: 'ms_graph_token',
  TOKEN_EXPIRY_KEY: 'ms_graph_token_expiry'
};

/**
 * Microsoft Graph API 액세스 토큰 획득
 * Client Credentials Flow 사용 (서버 간 인증)
 */
function getMicrosoftGraphToken() {
  try {
    // 캐시된 토큰 확인
    var properties = PropertiesService.getScriptProperties();
    var cachedToken = properties.getProperty(MS_GRAPH_CONFIG.TOKEN_CACHE_KEY);
    var tokenExpiry = properties.getProperty(MS_GRAPH_CONFIG.TOKEN_EXPIRY_KEY);
    
    if (cachedToken && tokenExpiry && new Date().getTime() < parseInt(tokenExpiry)) {
      return cachedToken;
    }
    
    // Client Secret 가져오기 (PropertiesService 또는 직접 설정)
    var clientSecret = properties.getProperty('MS_CLIENT_SECRET') || MS_GRAPH_CONFIG.CLIENT_SECRET;
    if (!clientSecret) {
      throw new Error('Microsoft Client Secret이 설정되지 않았습니다. PropertiesService에 MS_CLIENT_SECRET을 설정하세요.');
    }
    
    // 토큰 요청
    var tokenUrl = 'https://login.microsoftonline.com/' + MS_GRAPH_CONFIG.TENANT_ID + '/oauth2/v2.0/token';
    
    var payload = {
      'client_id': MS_GRAPH_CONFIG.CLIENT_ID,
      'client_secret': clientSecret,
      'scope': 'https://graph.microsoft.com/.default',
      'grant_type': 'client_credentials'
    };
    
    var options = {
      'method': 'post',
      'contentType': 'application/x-www-form-urlencoded',
      'payload': Object.keys(payload).map(function(key) {
        return encodeURIComponent(key) + '=' + encodeURIComponent(payload[key]);
      }).join('&')
    };
    
    var response = UrlFetchApp.fetch(tokenUrl, options);
    var result = JSON.parse(response.getContentText());
    
    if (result.error) {
      throw new Error('토큰 획득 실패: ' + result.error_description);
    }
    
    // 토큰 캐시 (만료 시간 50분 전에 갱신)
    var expiryTime = new Date().getTime() + (result.expires_in - 600) * 1000;
    properties.setProperty(MS_GRAPH_CONFIG.TOKEN_CACHE_KEY, result.access_token);
    properties.setProperty(MS_GRAPH_CONFIG.TOKEN_EXPIRY_KEY, expiryTime.toString());
    
    return result.access_token;
  } catch (error) {
    Logger.log('Error in getMicrosoftGraphToken: ' + error.toString());
    throw error;
  }
}

/**
 * OneDrive에서 Excel 파일 찾기
 * @param {string} fileName - 파일 이름 (예: '명함 신청 관리 DB.xlsx')
 * @return {object} 파일 정보 (id, name, webUrl 등)
 */
function findOneDriveExcelFile(fileName) {
  try {
    var token = getMicrosoftGraphToken();
    var url = MS_GRAPH_CONFIG.GRAPH_API_ENDPOINT + '/me/drive/root/search(q=\'' + encodeURIComponent(fileName) + '\')';
    
    var options = {
      'method': 'get',
      'headers': {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      }
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());
    
    if (result.value && result.value.length > 0) {
      // Excel 파일 찾기
      for (var i = 0; i < result.value.length; i++) {
        var file = result.value[i];
        if (file.name === fileName && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
          return file;
        }
      }
    }
    
    return null;
  } catch (error) {
    Logger.log('Error in findOneDriveExcelFile: ' + error.toString());
    throw error;
  }
}

/**
 * Excel 파일의 워크시트에서 데이터 읽기
 * @param {string} fileId - OneDrive 파일 ID
 * @param {string} worksheetName - 워크시트 이름 (예: '신청목록')
 * @param {string} range - 범위 (예: 'A1:M100')
 * @return {array} 2차원 배열 데이터
 */
function readExcelRange(fileId, worksheetName, range) {
  try {
    var token = getMicrosoftGraphToken();
    var url = MS_GRAPH_CONFIG.GRAPH_API_ENDPOINT + 
              '/me/drive/items/' + fileId + 
              '/workbook/worksheets/' + encodeURIComponent(worksheetName) + 
              '/range(address=\'' + range + '\')';
    
    var options = {
      'method': 'get',
      'headers': {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      }
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());
    
    if (result.values) {
      return result.values;
    }
    
    return [];
  } catch (error) {
    Logger.log('Error in readExcelRange: ' + error.toString());
    throw error;
  }
}

/**
 * Excel 파일의 워크시트에 데이터 쓰기
 * @param {string} fileId - OneDrive 파일 ID
 * @param {string} worksheetName - 워크시트 이름
 * @param {string} range - 범위 (예: 'A2')
 * @param {array} values - 2차원 배열 데이터
 */
function writeExcelRange(fileId, worksheetName, range, values) {
  try {
    var token = getMicrosoftGraphToken();
    var url = MS_GRAPH_CONFIG.GRAPH_API_ENDPOINT + 
              '/me/drive/items/' + fileId + 
              '/workbook/worksheets/' + encodeURIComponent(worksheetName) + 
              '/range(address=\'' + range + '\')';
    
    var payload = {
      'values': values
    };
    
    var options = {
      'method': 'patch',
      'headers': {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(payload)
    };
    
    var response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) {
      throw new Error('Excel 쓰기 실패: ' + response.getContentText());
    }
    
    return true;
  } catch (error) {
    Logger.log('Error in writeExcelRange: ' + error.toString());
    throw error;
  }
}

/**
 * Excel 파일에 새 행 추가 (테이블 사용)
 * @param {string} fileId - OneDrive 파일 ID
 * @param {string} worksheetName - 워크시트 이름
 * @param {string} tableName - 테이블 이름 (없으면 자동 생성)
 * @param {array} rowData - 행 데이터 배열
 */
function addExcelRow(fileId, worksheetName, tableName, rowData) {
  try {
    var token = getMicrosoftGraphToken();
    
    // 먼저 테이블이 있는지 확인하고 없으면 생성
    var tableId = getOrCreateTable(fileId, worksheetName, tableName);
    
    // 테이블에 행 추가
    var url = MS_GRAPH_CONFIG.GRAPH_API_ENDPOINT + 
              '/me/drive/items/' + fileId + 
              '/workbook/worksheets/' + encodeURIComponent(worksheetName) + 
              '/tables/' + tableId + '/rows/add';
    
    var payload = {
      'values': [rowData],
      'index': null
    };
    
    var options = {
      'method': 'post',
      'headers': {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(payload)
    };
    
    var response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 201) {
      throw new Error('행 추가 실패: ' + response.getContentText());
    }
    
    return JSON.parse(response.getContentText());
  } catch (error) {
    Logger.log('Error in addExcelRow: ' + error.toString());
    // 테이블이 없으면 일반 범위에 추가
    return addExcelRowToRange(fileId, worksheetName, rowData);
  }
}

/**
 * 범위에 직접 행 추가 (테이블이 없는 경우)
 */
function addExcelRowToRange(fileId, worksheetName, rowData) {
  try {
    // 먼저 마지막 행 찾기
    var data = readExcelRange(fileId, worksheetName, 'A:M');
    var lastRow = data.length;
    var nextRow = lastRow + 1;
    
    // 새 행에 데이터 쓰기
    var range = 'A' + nextRow + ':M' + nextRow;
    writeExcelRange(fileId, worksheetName, range, [rowData]);
    
    return { success: true };
  } catch (error) {
    Logger.log('Error in addExcelRowToRange: ' + error.toString());
    throw error;
  }
}

/**
 * 테이블 가져오기 또는 생성
 */
function getOrCreateTable(fileId, worksheetName, tableName) {
  try {
    var token = getMicrosoftGraphToken();
    
    // 기존 테이블 목록 가져오기
    var url = MS_GRAPH_CONFIG.GRAPH_API_ENDPOINT + 
              '/me/drive/items/' + fileId + 
              '/workbook/worksheets/' + encodeURIComponent(worksheetName) + 
              '/tables';
    
    var options = {
      'method': 'get',
      'headers': {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      }
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());
    
    // 기존 테이블 찾기
    if (result.value) {
      for (var i = 0; i < result.value.length; i++) {
        if (result.value[i].name === tableName) {
          return result.value[i].id;
        }
      }
    }
    
    // 테이블이 없으면 생성 (헤더 행이 있다고 가정)
    var createUrl = MS_GRAPH_CONFIG.GRAPH_API_ENDPOINT + 
                    '/me/drive/items/' + fileId + 
                    '/workbook/worksheets/' + encodeURIComponent(worksheetName) + 
                    '/tables/add';
    
    var payload = {
      'address': 'A1:M1', // 헤더 범위
      'hasHeaders': true
    };
    
    var createOptions = {
      'method': 'post',
      'headers': {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(payload)
    };
    
    var createResponse = UrlFetchApp.fetch(createUrl, createOptions);
    var createResult = JSON.parse(createResponse.getContentText());
    
    // 테이블 이름 변경
    if (tableName && createResult.id) {
      var renameUrl = MS_GRAPH_CONFIG.GRAPH_API_ENDPOINT + 
                      '/me/drive/items/' + fileId + 
                      '/workbook/worksheets/' + encodeURIComponent(worksheetName) + 
                      '/tables/' + createResult.id;
      
      var renamePayload = { 'name': tableName };
      var renameOptions = {
        'method': 'patch',
        'headers': {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'application/json'
        },
        'payload': JSON.stringify(renamePayload)
      };
      
      UrlFetchApp.fetch(renameUrl, renameOptions);
    }
    
    return createResult.id;
  } catch (error) {
    Logger.log('Error in getOrCreateTable: ' + error.toString());
    throw error;
  }
}

/**
 * Excel 파일의 특정 셀 업데이트
 * @param {string} fileId - OneDrive 파일 ID
 * @param {string} worksheetName - 워크시트 이름
 * @param {string} cellAddress - 셀 주소 (예: 'B2')
 * @param {string} value - 값
 */
function updateExcelCell(fileId, worksheetName, cellAddress, value) {
  try {
    writeExcelRange(fileId, worksheetName, cellAddress, [[value]]);
    return true;
  } catch (error) {
    Logger.log('Error in updateExcelCell: ' + error.toString());
    throw error;
  }
}

/**
 * OneDrive에 파일 업로드
 * @param {Blob} fileBlob - 업로드할 파일 Blob
 * @param {string} fileName - 파일 이름
 * @param {string} folderPath - 폴더 경로 (예: '/명함 초안 파일')
 * @return {object} 업로드된 파일 정보 (webUrl 포함)
 */
function uploadFileToOneDrive(fileBlob, fileName, folderPath) {
  try {
    var token = getMicrosoftGraphToken();
    
    // 폴더 경로 처리
    var uploadPath = folderPath ? folderPath + '/' + fileName : '/' + fileName;
    
    var url = MS_GRAPH_CONFIG.GRAPH_API_ENDPOINT + 
              '/me/drive/root:' + encodeURIComponent(uploadPath) + ':/content';
    
    var options = {
      'method': 'put',
      'headers': {
        'Authorization': 'Bearer ' + token,
        'Content-Type': fileBlob.getContentType()
      },
      'payload': fileBlob.getBytes()
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());
    
    // 공유 링크 생성
    if (result.id) {
      var shareUrl = MS_GRAPH_CONFIG.GRAPH_API_ENDPOINT + 
                     '/me/drive/items/' + result.id + 
                     '/createLink';
      
      var sharePayload = {
        'type': 'view',
        'scope': 'anonymous'
      };
      
      var shareOptions = {
        'method': 'post',
        'headers': {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'application/json'
        },
        'payload': JSON.stringify(sharePayload)
      };
      
      var shareResponse = UrlFetchApp.fetch(shareUrl, shareOptions);
      var shareResult = JSON.parse(shareResponse.getContentText());
      
      result.shareUrl = shareResult.link.webUrl;
    }
    
    return result;
  } catch (error) {
    Logger.log('Error in uploadFileToOneDrive: ' + error.toString());
    throw error;
  }
}

/**
 * Excel 파일 정보 캐시 (성능 향상)
 */
var excelFileCache = {
  fileId: null,
  fileName: null,
  lastChecked: null
};

/**
 * Excel 파일 ID 가져오기 (캐시 사용)
 */
function getExcelFileId(fileName) {
  var cacheTimeout = 3600000; // 1시간
  
  if (excelFileCache.fileName === fileName && 
      excelFileCache.lastChecked && 
      (new Date().getTime() - excelFileCache.lastChecked) < cacheTimeout) {
    return excelFileCache.fileId;
  }
  
  var file = findOneDriveExcelFile(fileName);
  if (file) {
    excelFileCache.fileId = file.id;
    excelFileCache.fileName = fileName;
    excelFileCache.lastChecked = new Date().getTime();
    return file.id;
  }
  
  throw new Error('Excel 파일을 찾을 수 없습니다: ' + fileName);
}

