# OneDrive 설정 가이드

이 문서는 Microsoft 365 OneDrive를 데이터베이스로 사용하기 위한 설정 가이드를 제공합니다.

## 사전 준비사항

1. Microsoft 365 계정 (OneDrive 접근 권한)
2. Azure Portal 접근 권한
3. Azure AD 앱 등록 (이미 완료된 것으로 보임)

## 1단계: Azure AD 앱 등록 확인 및 Client Secret 생성

### 1.1 Azure Portal 접속

1. [Azure Portal](https://portal.azure.com) 접속
2. **Azure Active Directory** > **앱 등록** 클릭
3. Client ID `5b165094-6f98-4818-9667-f81e651ce7d0` 검색

### 1.2 Client Secret 생성

1. 앱 등록 페이지에서 **인증서 및 암호** 클릭
2. **새 클라이언트 암호** 클릭
3. 설명 입력 (예: "Google Apps Script용")
4. 만료 기간 선택 (권장: 24개월)
5. **추가** 클릭
6. **값** 필드의 값을 복사 (한 번만 표시됨!)
7. `access_certi` 파일의 `MS_CLIENT_SECRET=` 뒤에 붙여넣기

### 1.3 API 권한 설정

1. **API 사용 권한** 클릭
2. **권한 추가** 클릭
3. **Microsoft Graph** 선택
4. **애플리케이션 권한** 선택
5. 다음 권한 추가:
   - `Files.ReadWrite.All` (OneDrive 파일 읽기/쓰기)
   - `Sites.ReadWrite.All` (SharePoint 사이트 접근, 필요한 경우)
6. **권한 추가** 클릭
7. **관리자 동의 부여** 클릭 (테넌트 관리자 권한 필요)

## 2단계: OneDrive Excel 파일 생성

### 2.1 Excel 파일 생성

1. [OneDrive](https://onedrive.live.com) 접속
2. **새로 만들기** > **Excel 통합 문서** 클릭
3. 파일 이름을 "명함 신청 관리 DB.xlsx"로 변경
4. 첫 번째 워크시트 이름을 "신청목록"으로 변경

### 2.2 헤더 행 설정

A1부터 M1까지 다음 헤더를 입력:

```
A1: 신청ID
B1: Status
C1: 신청자계정
D1: 등록자
E1: 통
F1: 기존동일여부
G1: 변호사여부
H1: 담당변호사
I1: 비고
J1: 등록일
K1: 첨부파일
L1: 처리자
M1: 처리일자
```

3. 첫 번째 행을 선택하고 **홈** 탭 > **굵게** 적용

### 2.3 OneDrive 폴더 생성 (초안 파일용)

1. OneDrive에서 **새로 만들기** > **폴더** 클릭
2. 폴더 이름을 "명함 초안 파일"로 설정
3. 폴더 생성 완료

## 3단계: Google Apps Script 설정

### 3.1 MicrosoftGraphAPI.gs 파일 추가

1. Google Apps Script 편집기에서 **파일 > 새로 만들기 > 스크립트** 클릭
2. 파일 이름을 "MicrosoftGraphAPI"로 변경
3. 제공된 `MicrosoftGraphAPI.gs` 내용을 복사하여 붙여넣기

### 3.2 Client Secret 설정

Google Apps Script에서 Client Secret을 안전하게 저장:

**방법 1: PropertiesService 사용 (권장)**

1. Apps Script 편집기에서 다음 코드를 실행 (한 번만):

```javascript
function setClientSecret() {
  var properties = PropertiesService.getScriptProperties();
  // access_certi 파일의 MS_CLIENT_SECRET 값을 여기에 입력
  properties.setProperty('MS_CLIENT_SECRET', '여기에_Client_Secret_값_입력');
  Logger.log('Client Secret이 설정되었습니다.');
}
```

2. **실행** > `setClientSecret` 선택 > **실행** 클릭
3. 권한 승인
4. 실행 후 `setClientSecret` 함수는 삭제하거나 주석 처리

**방법 2: 직접 설정 (덜 안전)**

`MicrosoftGraphAPI.gs` 파일의 `MS_GRAPH_CONFIG` 객체에서 직접 설정:

```javascript
const MS_GRAPH_CONFIG = {
  CLIENT_ID: '5b165094-6f98-4818-9667-f81e651ce7d0',
  TENANT_ID: '20c53626-8c84-4cb2-b049-ae1f603445a6',
  CLIENT_SECRET: '여기에_Client_Secret_값_입력', // 직접 설정
  // ...
};
```

⚠️ **주의**: 이 방법은 코드에 비밀번호가 노출되므로 권장하지 않습니다.

### 3.3 Code.gs 설정 확인

`Code.gs` 파일의 `CONFIG` 객체를 확인:

```javascript
const CONFIG = {
  EXCEL_FILE_NAME: '명함 신청 관리 DB.xlsx', // OneDrive Excel 파일 이름
  WORKSHEET_NAME: '신청목록', // 워크시트 이름
  ADMIN_EMAIL: 'bysong@law-lin.com',
  ADMIN_PASSWORD: 'admin123', // 반드시 변경!
  ONEDRIVE_FOLDER_PATH: '/명함 초안 파일' // OneDrive 폴더 경로
};
```

## 4단계: 테스트

### 4.1 토큰 획득 테스트

1. Apps Script 편집기에서 다음 함수 실행:

```javascript
function testToken() {
  try {
    var token = getMicrosoftGraphToken();
    Logger.log('토큰 획득 성공!');
    Logger.log('토큰 (처음 20자): ' + token.substring(0, 20) + '...');
    return '성공';
  } catch (error) {
    Logger.log('오류: ' + error.toString());
    return '실패: ' + error.toString();
  }
}
```

2. **실행** > `testToken` 선택 > **실행** 클릭
3. 로그 확인 (보기 > 로그)

### 4.2 Excel 파일 찾기 테스트

```javascript
function testFindFile() {
  try {
    var file = findOneDriveExcelFile('명함 신청 관리 DB.xlsx');
    if (file) {
      Logger.log('파일 찾기 성공!');
      Logger.log('파일 ID: ' + file.id);
      Logger.log('파일 이름: ' + file.name);
      return '성공';
    } else {
      Logger.log('파일을 찾을 수 없습니다.');
      return '파일 없음';
    }
  } catch (error) {
    Logger.log('오류: ' + error.toString());
    return '실패: ' + error.toString();
  }
}
```

### 4.3 데이터 읽기 테스트

```javascript
function testReadData() {
  try {
    var fileId = getExcelFileId('명함 신청 관리 DB.xlsx');
    var data = readExcelRange(fileId, '신청목록', 'A1:M1');
    Logger.log('데이터 읽기 성공!');
    Logger.log('헤더: ' + JSON.stringify(data[0]));
    return '성공';
  } catch (error) {
    Logger.log('오류: ' + error.toString());
    return '실패: ' + error.toString();
  }
}
```

## 5단계: 문제 해결

### 토큰 획득 실패

**증상**: "토큰 획득 실패" 오류

**해결 방법**:
1. Client Secret이 올바르게 설정되었는지 확인
2. Azure Portal에서 Client Secret이 만료되지 않았는지 확인
3. API 권한이 올바르게 설정되었는지 확인
4. 관리자 동의가 부여되었는지 확인

### 파일을 찾을 수 없음

**증상**: "Excel 파일을 찾을 수 없습니다" 오류

**해결 방법**:
1. OneDrive에 파일이 존재하는지 확인
2. 파일 이름이 정확히 일치하는지 확인 (대소문자, 확장자 포함)
3. 파일이 삭제되지 않았는지 확인
4. 파일 접근 권한 확인

### 권한 오류

**증상**: "403 Forbidden" 또는 "401 Unauthorized" 오류

**해결 방법**:
1. Azure Portal에서 API 권한 확인
2. 관리자 동의가 부여되었는지 확인
3. 앱 등록의 인증 설정 확인 (리디렉션 URI 등)

### Excel 읽기/쓰기 오류

**증상**: "Excel 읽기 실패" 또는 "Excel 쓰기 실패" 오류

**해결 방법**:
1. 워크시트 이름이 정확히 일치하는지 확인
2. 범위(A1:M100 등)가 올바른지 확인
3. Excel 파일이 열려있지 않은지 확인 (다른 사용자가 편집 중이면 오류 발생 가능)

## 6단계: 보안 고려사항

### Client Secret 보안

- ✅ PropertiesService 사용 (권장)
- ❌ 코드에 직접 하드코딩 (비권장)
- ❌ 공개 저장소에 업로드 금지

### API 권한 최소화

필요한 최소한의 권한만 부여:
- `Files.ReadWrite.All` (OneDrive 파일 접근)
- 불필요한 권한은 제거

### 토큰 캐싱

시스템은 자동으로 토큰을 캐싱하여 성능을 최적화합니다. 토큰은 약 50분 후 자동으로 갱신됩니다.

## 추가 리소스

- [Microsoft Graph API 문서](https://docs.microsoft.com/graph/overview)
- [Excel REST API 문서](https://docs.microsoft.com/graph/api/resources/excel)
- [Azure AD 앱 등록 가이드](https://docs.microsoft.com/azure/active-directory/develop/quickstart-register-app)

