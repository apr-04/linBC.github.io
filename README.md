# 명함 신청 관리 시스템

Microsoft 365 OneDrive Excel과 Google Apps Script를 활용한 명함 신청 관리 시스템입니다.

## 주요 기능

1. **사용자 신청 페이지**: 명함 신청 폼 제출
2. **신청 내역 조회**: 이메일과 신청ID로 상태 조회
3. **초안 확인 페이지**: 초안 확인 및 제작 승인/수정 요청
4. **관리자 페이지**: 신청 목록 관리, 초안 파일 첨부, 상태 변경

## 시스템 구조

### OneDrive Excel 파일 구조

**파일 이름**: `명함 신청 관리 DB.xlsx`  
**워크시트 이름**: `신청목록`

| 열 | 이름 | 설명 |
|---|---|---|
| A | 신청ID | 고유 ID (예: LN-2025-1001) |
| B | Status | 상태 (신청, 초안전달, 수정요청, 제작, 지급완료, 삭제) |
| C | 신청자계정 | 신청자 이메일 |
| D | 등록자 | 신청자 이름 |
| E | 통 | 신청 개수 |
| F | 기존동일여부 | 예/아니오 |
| G | 변호사여부 | 변호사/Staff |
| H | 담당변호사 | Staff 선택 시 입력 |
| I | 비고 | 요청사항 |
| J | 등록일 | 신청 날짜 (자동입력) |
| K | 첨부파일 | 초안 파일 OneDrive 링크 |
| L | 처리자 | 관리자 이름 |
| M | 처리일자 | 상태 변경 날짜 |

## 설치 및 설정 가이드

⚠️ **중요**: 이 시스템은 Microsoft 365 OneDrive를 데이터베이스로 사용합니다.

### 사전 준비

1. **OneDrive 설정**: `ONEDRIVE_SETUP.md` 파일을 먼저 참고하세요.
2. **Azure AD 앱 등록**: Client ID와 Tenant ID가 이미 있으므로, Client Secret만 생성하면 됩니다.

### 1단계: OneDrive Excel 파일 생성

1. OneDrive에서 새 Excel 파일 생성
2. 파일 이름을 "명함 신청 관리 DB.xlsx"로 변경
3. 첫 번째 워크시트 이름을 "신청목록"으로 변경
4. 첫 번째 행에 헤더 입력:
   ```
   신청ID | Status | 신청자계정 | 등록자 | 통 | 기존동일여부 | 변호사여부 | 담당변호사 | 비고 | 등록일 | 첨부파일 | 처리자 | 처리일자
   ```

### 2단계: Google Apps Script 프로젝트 생성

1. Google Apps Script 편집기에서 새 프로젝트 생성
2. 다음 파일들을 생성:
   - `Code.gs` (메인 코드)
   - `MicrosoftGraphAPI.gs` (Microsoft Graph API 통신 모듈)
   - `UserForm.html` (명함 신청 폼)
   - `CheckStatus.html` (신청 내역 조회)
   - `ReviewDraft.html` (초안 확인)
   - `AdminPage.html` (관리자 페이지)
3. 각 파일에 제공된 내용을 복사하여 붙여넣기

### 3단계: Microsoft Graph API 설정

**중요**: `ONEDRIVE_SETUP.md` 파일의 3단계를 따라 Client Secret을 설정하세요.

1. Azure Portal에서 Client Secret 생성
2. Google Apps Script의 PropertiesService에 Client Secret 저장:
   ```javascript
   function setClientSecret() {
     var properties = PropertiesService.getScriptProperties();
     properties.setProperty('MS_CLIENT_SECRET', '여기에_Client_Secret_값_입력');
   }
   ```

### 4단계: 설정 변경

`Code.gs` 파일 상단의 `CONFIG` 객체를 수정:

```javascript
const CONFIG = {
  EXCEL_FILE_NAME: '명함 신청 관리 DB.xlsx', // OneDrive Excel 파일 이름
  WORKSHEET_NAME: '신청목록', // 워크시트 이름
  ADMIN_EMAIL: 'bysong@law-lin.com',  // 관리자 이메일
  ADMIN_PASSWORD: 'admin123',  // 관리자 비밀번호 (반드시 변경!)
  ONEDRIVE_FOLDER_PATH: '/명함 초안 파일' // OneDrive 폴더 경로
};
```

### 5단계: 웹 앱 배포

1. Apps Script 편집기에서 **배포 > 새 배포** 클릭
2. **유형 선택**에서 "웹 앱" 선택
3. 설정:
   - **설명**: "명함 신청 관리 시스템"
   - **실행 사용자**: "나"
   - **액세스 권한**: "모든 사용자" 또는 "내 조직의 사용자"
4. **배포** 클릭
5. 웹 앱 URL 복사 (예: `https://script.google.com/macros/s/...`)

### 6단계: 권한 설정

1. 첫 실행 시 권한 요청이 나타나면 **권한 검토** 클릭
2. Google 계정 선택
3. **고급** > **안전하지 않은 페이지로 이동** 클릭
4. **허용** 클릭

### 7단계: OneDrive 폴더 설정

1. OneDrive에서 "명함 초안 파일" 폴더 생성 (또는 CONFIG에서 지정한 이름)
2. 폴더는 Microsoft Graph API를 통해 자동으로 접근됩니다

## 사용 방법

### 사용자 페이지 접근

1. **명함 신청**: `[웹앱URL]?page=form`
2. **신청 내역 조회**: `[웹앱URL]?page=check`
3. **초안 확인**: `[웹앱URL]?page=review&id=[신청ID]` (이메일 링크로 접근)

### 관리자 페이지 접근

1. `[웹앱URL]?page=admin`
2. 비밀번호 입력 후 로그인
3. 신청 목록 조회 및 관리

## 워크플로우

1. **사용자 신청**
   - 사용자가 신청 폼 작성 및 제출
   - 자동으로 신청ID 생성 (LN-YYYY-NNNN)
   - 관리자와 신청자에게 이메일 알림 발송

2. **관리자 초안 준비**
   - 관리자 페이지에서 신청 내역 확인
   - 초안 파일 업로드 (PDF, 이미지 등)
   - "초안 확인 요청" 버튼 클릭
   - 상태가 '초안전달'로 변경되고 신청자에게 이메일 발송

3. **사용자 초안 확인**
   - 이메일 링크 클릭하여 초안 확인 페이지 접근
   - 초안 파일 확인
   - **제작 신청** 또는 **추가 수정 요청** 선택

4. **제작 승인 시**
   - 상태가 '제작'으로 변경
   - 관리자에게 제작 승인 알림 이메일 발송

5. **수정 요청 시**
   - 상태가 '수정요청'으로 변경
   - 관리자에게 수정 요청 알림 이메일 발송
   - 관리자는 2단계부터 다시 진행

6. **제작 완료**
   - 관리자가 상태를 '지급완료'로 변경

## 이메일 알림

시스템은 다음 상황에서 자동으로 이메일을 발송합니다:

- **신청 접수**: 관리자와 신청자에게 알림
- **초안 전달**: 신청자에게 초안 확인 요청 (초안 확인 페이지 링크 포함)
- **제작 승인**: 관리자에게 제작 승인 알림
- **수정 요청**: 관리자에게 수정 요청 알림

## 보안 고려사항

1. **관리자 비밀번호**: `Code.gs`의 `ADMIN_PASSWORD`를 반드시 강력한 비밀번호로 변경
2. **웹 앱 접근 권한**: 필요에 따라 "내 조직의 사용자"로 제한
3. **Google Drive 폴더**: 초안 파일은 링크로만 접근 가능하도록 설정됨

## 문제 해결

### 이메일이 발송되지 않는 경우
- Apps Script의 실행 시간 할당량 확인
- 관리자 이메일 주소 확인
- MailApp 서비스 권한 확인

### 파일 업로드가 안 되는 경우
- Microsoft Graph API 권한 확인 (`Files.ReadWrite.All`)
- 파일 크기 제한 확인 (OneDrive는 250GB 제한, 단일 파일은 250MB)
- OneDrive 폴더 경로가 올바른지 확인
- Client Secret이 올바르게 설정되었는지 확인

### Excel 파일을 찾을 수 없는 경우
- 파일 이름이 정확히 일치하는지 확인 (`CONFIG.EXCEL_FILE_NAME`)
- OneDrive에 파일이 존재하는지 확인
- Microsoft Graph API 권한이 올바르게 설정되었는지 확인

## 기술 스택

- **Google Apps Script**: 백엔드 로직 및 자동화
- **Microsoft 365 OneDrive**: 데이터베이스 (Excel 파일)
- **Microsoft Graph API**: OneDrive 접근
- **Azure AD**: 인증 및 권한 관리
- **HTML/CSS/JavaScript**: 프론트엔드 UI
- **HtmlService**: 웹 앱 서비스

## 라이선스

이 프로젝트는 내부 사용을 위한 것입니다.

## 지원

문제가 발생하거나 질문이 있으시면 관리자에게 문의하세요.

