// OneDrive API 접근 모듈
// Microsoft Graph API 사용

const GRAPH_API_BASE = 'https://graph.microsoft.com/v1.0';

/**
 * OneDrive에 파일 업로드
 * @param {string} filePath - OneDrive 내 파일 경로 (예: /명함신청/명함DB.xlsx)
 * @param {File|Blob} file - 업로드할 파일
 * @returns {Promise<Object>} 업로드된 파일 정보
 */
async function uploadFileToOneDrive(filePath, file) {
    const token = await getAccessToken();
    
    // 파일을 ArrayBuffer로 변환
    const arrayBuffer = await file.arrayBuffer();
    
    // Microsoft Graph API를 사용하여 파일 업로드
    const response = await fetch(`${GRAPH_API_BASE}/me/drive/root:${filePath}:/content`, {
        method: 'PUT',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': file.type || 'application/octet-stream'
        },
        body: arrayBuffer
    });
    
    if (!response.ok) {
        const error = await response.text();
        throw new Error(`Failed to upload file: ${error}`);
    }
    
    return await response.json();
}

/**
 * OneDrive에서 파일 다운로드
 * @param {string} filePath - OneDrive 내 파일 경로
 * @returns {Promise<Blob>} 파일 데이터
 */
async function downloadFileFromOneDrive(filePath) {
    const token = await getAccessToken();
    
    const response = await fetch(`${GRAPH_API_BASE}/me/drive/root:${filePath}:/content`, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${token}`
        }
    });
    
    if (!response.ok) {
        const error = await response.text();
        throw new Error(`Failed to download file: ${error}`);
    }
    
    return await response.blob();
}

/**
 * OneDrive에서 파일 정보 조회
 * @param {string} filePath - OneDrive 내 파일 경로
 * @returns {Promise<Object>} 파일 메타데이터
 */
async function getFileInfo(filePath) {
    const token = await getAccessToken();
    
    const response = await fetch(`${GRAPH_API_BASE}/me/drive/root:${filePath}:`, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${token}`
        }
    });
    
    if (!response.ok) {
        const error = await response.text();
        throw new Error(`Failed to get file info: ${error}`);
    }
    
    return await response.json();
}

/**
 * OneDrive에 이미지 파일 업로드 및 공유 링크 생성
 * @param {string} folderPath - 폴더 경로 (예: /명함신청/초안)
 * @param {File} imageFile - 이미지 파일
 * @returns {Promise<Object>} 파일 정보 및 다운로드 URL
 */
async function uploadImageAndGetLink(folderPath, imageFile) {
    const token = await getAccessToken();
    
    // 파일명 생성
    const timestamp = Date.now();
    const fileName = `${timestamp}_${imageFile.name}`;
    const filePath = `${folderPath}/${fileName}`;
    
    // 파일 업로드
    const uploadResult = await uploadFileToOneDrive(filePath, imageFile);
    
    // 공유 링크 생성 (읽기 전용)
    const shareResponse = await fetch(`${GRAPH_API_BASE}/me/drive/items/${uploadResult.id}/createLink`, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            type: 'view',
            scope: 'anonymous'
        })
    });
    
    if (!shareResponse.ok) {
        throw new Error('Failed to create share link');
    }
    
    const shareData = await shareResponse.json();
    
    return {
        fileId: uploadResult.id,
        fileName: fileName,
        downloadUrl: shareData.link.webUrl,
        directDownloadUrl: `${GRAPH_API_BASE}/me/drive/items/${uploadResult.id}/content`
    };
}

/**
 * OneDrive 폴더 목록 조회
 * @param {string} folderPath - 폴더 경로
 * @returns {Promise<Array>} 파일 목록
 */
async function listFilesInFolder(folderPath) {
    const token = await getAccessToken();
    
    const response = await fetch(`${GRAPH_API_BASE}/me/drive/root:${folderPath}:/children`, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${token}`
        }
    });
    
    if (!response.ok) {
        const error = await response.text();
        throw new Error(`Failed to list files: ${error}`);
    }
    
    const data = await response.json();
    return data.value;
}


