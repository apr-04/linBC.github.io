# 법무법인 린 - 명함 관리 시스템

## 🚀 배포 가이드

### 1. Azure AD 앱 등록
1. [Azure Portal](https://portal.azure.com) 접속
2. Azure Active Directory > 앱 등록 > 새 등록
3. API 권한 추가 (Microsoft Graph):
   - Files.ReadWrite.All
   - Mail.Send
4. 인증서 및 암호에서 클라이언트 암호 생성

### 2. OneDrive 준비
1. OneDrive에 "명함 관리 DB.xlsx" 생성
2. "신청목록" 시트 생성
3. 헤더 행: