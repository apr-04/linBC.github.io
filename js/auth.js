// Azure AD 인증 모듈
// access_certi 파일의 인증 정보 사용

const AZURE_CONFIG = {
    clientId: '5b165094-6f98-4818-9667-f81e651ce7d0',
    tenantId: '20c53626-8c84-4cb2-b049-ae1f603445a6',
    authority: `https://login.microsoftonline.com/20c53626-8c84-4cb2-b049-ae1f603445a6`,
    redirectUri: window.location.origin,
    scopes: ['Files.ReadWrite.All', 'User.Read']
};

let accessToken = null;
let tokenExpiry = null;

/**
 * Azure AD로 로그인하여 액세스 토큰 획득
 * 리다이렉트 방식으로 인증
 */
async function authenticate() {
    try {
        // 리다이렉트 URI 저장 (인증 후 돌아올 위치)
        sessionStorage.setItem('authReturnUrl', window.location.href);
        
        // Azure AD 인증 URL 생성
        const authUrl = `${AZURE_CONFIG.authority}/oauth2/v2.0/authorize?` +
            `client_id=${AZURE_CONFIG.clientId}&` +
            `response_type=token&` +
            `redirect_uri=${encodeURIComponent(AZURE_CONFIG.redirectUri + window.location.pathname)}&` +
            `scope=${encodeURIComponent(AZURE_CONFIG.scopes.join(' '))}&` +
            `response_mode=fragment`;

        // 리다이렉트로 인증 페이지로 이동
        window.location.href = authUrl;
        
        // 이 함수는 리다이렉트되므로 여기까지 도달하지 않음
        return null;
    } catch (error) {
        console.error('Authentication error:', error);
        throw error;
    }
}

/**
 * 저장된 토큰 확인 및 갱신
 */
async function getAccessToken() {
    // 먼저 URL 해시에서 토큰 확인 (인증 리다이렉트 후)
    if (window.location.hash) {
        handleRedirect();
    }
    
    // 로컬 스토리지에서 토큰 확인
    const storedToken = localStorage.getItem('accessToken');
    const storedExpiry = localStorage.getItem('tokenExpiry');
    
    if (storedToken && storedExpiry && Date.now() < parseInt(storedExpiry)) {
        accessToken = storedToken;
        tokenExpiry = parseInt(storedExpiry);
        return accessToken;
    }
    
    // 토큰이 없거나 만료된 경우 재인증
    // 사용자에게 인증 필요 알림
    if (confirm('OneDrive 접근을 위해 Microsoft 계정으로 로그인이 필요합니다. 로그인하시겠습니까?')) {
        await authenticate();
    } else {
        throw new Error('Authentication required');
    }
    
    return null;
}

/**
 * 로그아웃
 */
function logout() {
    accessToken = null;
    tokenExpiry = null;
    localStorage.removeItem('accessToken');
    localStorage.removeItem('tokenExpiry');
}

/**
 * URL 해시에서 토큰 추출 (리다이렉트 방식)
 */
function handleRedirect() {
    const hash = window.location.hash.substring(1);
    if (hash) {
        const params = new URLSearchParams(hash);
        const token = params.get('access_token');
        if (token) {
            accessToken = token;
            const expiresIn = parseInt(params.get('expires_in')) * 1000;
            tokenExpiry = Date.now() + expiresIn;
            localStorage.setItem('accessToken', accessToken);
            localStorage.setItem('tokenExpiry', tokenExpiry.toString());
            
            // URL에서 해시 제거
            window.history.replaceState(null, null, window.location.pathname);
            return true;
        }
    }
    return false;
}

// 페이지 로드 시 리다이렉트 처리
if (window.location.hash) {
    handleRedirect();
}

