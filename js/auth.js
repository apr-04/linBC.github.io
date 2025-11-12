// Azure AD 인증 모듈 (MSAL.js 사용)
// access_certi 파일의 인증 정보 사용

// MSAL.js가 로드되었는지 확인
let msalInstance = null;

const AZURE_CONFIG = {
    clientId: '5b165094-6f98-4818-9667-f81e651ce7d0',
    tenantId: '20c53626-8c84-4cb2-b049-ae1f603445a6',
    authority: `https://login.microsoftonline.com/20c53626-8c84-4cb2-b049-ae1f603445a6`,
    redirectUri: window.location.origin + window.location.pathname,
    scopes: ['Files.ReadWrite.All', 'User.Read']
};

/**
 * MSAL 인스턴스 초기화
 */
function initializeMSAL() {
    if (typeof msal === 'undefined') {
        throw new Error('MSAL.js library is not loaded. Please include MSAL.js script.');
    }

    if (!msalInstance) {
        const msalConfig = {
            auth: {
                clientId: AZURE_CONFIG.clientId,
                authority: AZURE_CONFIG.authority,
                redirectUri: AZURE_CONFIG.redirectUri
            },
            cache: {
                cacheLocation: 'localStorage',
                storeAuthStateInCookie: false
            }
        };

        msalInstance = new msal.PublicClientApplication(msalConfig);
    }

    return msalInstance;
}

/**
 * Azure AD로 로그인하여 액세스 토큰 획득
 * Authorization Code Flow with PKCE 사용
 */
async function authenticate() {
    try {
        const msal = initializeMSAL();
        
        const loginRequest = {
            scopes: AZURE_CONFIG.scopes,
            prompt: 'select_account'
        };

        try {
            const response = await msal.loginPopup(loginRequest);
            return response.accessToken;
        } catch (error) {
            // 팝업이 차단된 경우 리다이렉트 방식 사용
            if (error.name === 'BrowserAuthError' || error.name === 'PopupWindowError') {
                await msal.loginRedirect(loginRequest);
                return null; // 리다이렉트되므로 여기까지 도달하지 않음
            }
            throw error;
        }
    } catch (error) {
        console.error('Authentication error:', error);
        throw error;
    }
}

/**
 * 저장된 토큰 확인 및 갱신
 */
async function getAccessToken() {
    try {
        const msal = initializeMSAL();
        
        // 계정 확인
        const accounts = msal.getAllAccounts();
        
        if (accounts.length === 0) {
            // 로그인 필요
            if (confirm('OneDrive 접근을 위해 Microsoft 계정으로 로그인이 필요합니다. 로그인하시겠습니까?')) {
                return await authenticate();
            } else {
                throw new Error('Authentication required');
            }
        }

        // 토큰 획득 시도
        const tokenRequest = {
            scopes: AZURE_CONFIG.scopes,
            account: accounts[0]
        };

        try {
            const response = await msal.acquireTokenSilent(tokenRequest);
            return response.accessToken;
        } catch (error) {
            // Silent 토큰 획득 실패 시 인터랙티브 로그인
            if (error.name === 'InteractionRequiredAuthError' || error.name === 'BrowserAuthError') {
                try {
                    const response = await msal.acquireTokenPopup(tokenRequest);
                    return response.accessToken;
                } catch (popupError) {
                    // 팝업 실패 시 리다이렉트
                    await msal.acquireTokenRedirect(tokenRequest);
                    return null;
                }
            }
            throw error;
        }
    } catch (error) {
        console.error('Error getting access token:', error);
        throw error;
    }
}

/**
 * 로그아웃
 */
async function logout() {
    try {
        const msal = initializeMSAL();
        await msal.logoutPopup();
    } catch (error) {
        // 팝업 실패 시 리다이렉트
        const msal = initializeMSAL();
        await msal.logoutRedirect();
    }
}

// 페이지 로드 시 MSAL 초기화 및 리다이렉트 처리
if (typeof msal !== 'undefined') {
    initializeMSAL().handleRedirectPromise().then((response) => {
        if (response) {
            console.log('Login successful');
        }
    }).catch((error) => {
        console.error('Error handling redirect:', error);
    });
}

=======
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