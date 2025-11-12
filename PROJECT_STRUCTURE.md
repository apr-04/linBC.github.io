# 프로젝트 구조

## 파일 목록

```
google-apps-script/
├── Code.gs                 # 메인 Google Apps Script 코드
├── UserForm.html           # 사용자 신청 폼 페이지
├── CheckStatus.html        # 신청 내역 조회 페이지
├── ReviewDraft.html        # 초안 확인 페이지
├── AdminPage.html          # 관리자 페이지
├── README.md              # 프로젝트 개요 및 사용법
├── SETUP_GUIDE.md         # 상세 설정 가이드
└── PROJECT_STRUCTURE.md   # 이 파일
```

## 파일별 역할

### Code.gs
Google Apps Script의 메인 로직 파일입니다.

**주요 함수:**
- `doGet(e)`: 웹 앱 진입점, 페이지 라우팅
- `submitApplication(data)`: 신청 데이터 처리
- `checkApplicationStatus(email, applicationId)`: 신청 내역 조회
- `getApplicationForReview(applicationId)`: 초안 확인 페이지 데이터
- `approveProduction(applicationId)`: 제작 승인 처리
- `requestModification(applicationId, reason)`: 수정 요청 처리
- `adminLogin(password)`: 관리자 로그인 인증
- `getAdminApplicationList(statusFilter)`: 관리자 목록 조회
- `adminAttachDraft(...)`: 초안 파일 업로드
- `adminSendDraft(applicationId)`: 초안 전달 처리
- `adminChangeStatus(applicationId, newStatus)`: 상태 변경

**유틸리티 함수:**
- `getSheet()`: 시트 가져오기/생성
- `getNextSequence()`: 다음 신청ID 시퀀스 생성
- `getOrCreateDriveFolder()`: Drive 폴더 가져오기/생성
- `getWebAppUrl()`: 웹 앱 URL 가져오기

**이메일 함수:**
- `sendNewApplicationEmails(...)`: 신청 접수 알림
- `sendDraftReviewEmail(...)`: 초안 확인 요청
- `sendProductionApprovalEmail(...)`: 제작 승인 알림
- `sendModificationRequestEmail(...)`: 수정 요청 알림

### UserForm.html
명함 신청 폼 페이지입니다.

**기능:**
- 신청자 정보 입력 폼
- Staff 선택 시 담당 변호사 입력 필드 표시
- 폼 제출 및 결과 표시

**접근 URL:** `[웹앱URL]?page=form`

### CheckStatus.html
신청 내역 조회 페이지입니다.

**기능:**
- 이메일과 신청ID로 상태 조회
- 상태별 색상 표시
- 조회 결과 표시

**접근 URL:** `[웹앱URL]?page=check`

### ReviewDraft.html
명함 초안 확인 페이지입니다.

**기능:**
- 초안 파일 확인
- 제작 승인 버튼
- 수정 요청 버튼 (사유 입력 모달)

**접근 URL:** `[웹앱URL]?page=review&id=[신청ID]`

**보안:** 신청ID를 통한 접근 제어 (추가 인증 가능)

### AdminPage.html
관리자 페이지입니다.

**기능:**
- 비밀번호 로그인
- 신청 목록 조회 (상태별 필터링)
- 상세 정보 보기
- 초안 파일 업로드
- 초안 전달 처리
- 상태 변경

**접근 URL:** `[웹앱URL]?page=admin`

## 데이터 흐름

### 신청 프로세스
```
사용자 → UserForm.html → submitApplication() → Google Sheets
                                              → 이메일 알림
```

### 조회 프로세스
```
사용자 → CheckStatus.html → checkApplicationStatus() → Google Sheets → 결과 표시
```

### 초안 확인 프로세스
```
관리자 → AdminPage.html → adminAttachDraft() → Google Drive
                                              → Google Sheets (링크 저장)
                                              → adminSendDraft() → 이메일 발송
                                                                  → 상태 변경

사용자 → 이메일 링크 → ReviewDraft.html → approveProduction() 또는 requestModification()
                                                                    → 상태 변경
                                                                    → 이메일 알림
```

## 상태 전이도

```
신청 → 초안전달 → 제작 → 지급완료
              ↓
          수정요청 → (다시) 초안전달 → ...
```

## 보안 고려사항

1. **관리자 페이지**: 비밀번호 인증 (간단한 인증, 필요시 강화 가능)
2. **초안 확인 페이지**: 신청ID 기반 접근 (추가 이메일 인증 가능)
3. **Google Sheets**: 직접 수정 방지 (관리자 페이지 통해서만)
4. **파일 접근**: Google Drive 링크 공유 (링크로만 접근 가능)

## 확장 가능성

### 추가 기능 아이디어
- Excel 다운로드 기능
- 통계 대시보드
- 이메일 템플릿 커스터마이징
- 다국어 지원
- 모바일 앱 연동
- Slack/Teams 알림 연동

### 코드 개선 아이디어
- 에러 핸들링 강화
- 로깅 시스템 추가
- 캐싱 메커니즘 추가
- API 엔드포인트 추가

