// Excel 파일 읽기/쓰기 모듈
// Microsoft Graph API Excel 엔드포인트 사용

const EXCEL_FILE_PATH = '/명함신청/명함DB.xlsx';
const GRAPH_API_BASE = 'https://graph.microsoft.com/v1.0';

/**
 * Excel 파일에서 모든 행 읽기
 * @returns {Promise<Array>} Excel 데이터 배열
 */
async function readExcelData() {
    const token = await getAccessToken();
    
    try {
        // Microsoft Graph API Excel 엔드포인트 사용
        const fileInfo = await getFileInfo(EXCEL_FILE_PATH);
        
        // Excel 워크시트 읽기
        const response = await fetch(`${GRAPH_API_BASE}/me/drive/items/${fileInfo.id}/workbook/worksheets/Sheet1/usedRange`, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });
        
        if (!response.ok) {
            // Excel 파일이 없거나 형식이 다른 경우, 파일을 다운로드하여 파싱
            return await readExcelAsFile();
        }
        
        const data = await response.json();
        return parseExcelRange(data);
    } catch (error) {
        console.error('Error reading Excel via Graph API:', error);
        // 대체 방법: 파일 다운로드 후 SheetJS로 파싱
        return await readExcelAsFile();
    }
}

/**
 * Excel 파일을 다운로드하여 SheetJS로 파싱
 */
async function readExcelAsFile() {
    try {
        const blob = await downloadFileFromOneDrive(EXCEL_FILE_PATH);
        
        // SheetJS 라이브러리 사용 (CDN에서 로드 필요)
        if (typeof XLSX === 'undefined') {
            throw new Error('XLSX library not loaded. Please include SheetJS library.');
        }
        
        const arrayBuffer = await blob.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        
        return jsonData;
    } catch (error) {
        console.error('Error reading Excel file:', error);
        // 파일이 없는 경우 빈 배열 반환
        return [];
    }
}

/**
 * Graph API Excel 범위 데이터 파싱
 */
function parseExcelRange(rangeData) {
    if (!rangeData.values || rangeData.values.length === 0) {
        return [];
    }
    
    const headers = rangeData.values[0];
    const rows = [];
    
    for (let i = 1; i < rangeData.values.length; i++) {
        const row = {};
        rangeData.values[i].forEach((value, index) => {
            row[headers[index]] = value;
        });
        rows.push(row);
    }
    
    return rows;
}

/**
 * Excel 파일에 새 행 추가
 * @param {Object} rowData - 추가할 행 데이터
 */
async function addRowToExcel(rowData) {
    const token = await getAccessToken();
    
    try {
        const fileInfo = await getFileInfo(EXCEL_FILE_PATH);
        
        // Excel 워크시트에 행 추가
        const response = await fetch(`${GRAPH_API_BASE}/me/drive/items/${fileInfo.id}/workbook/worksheets/Sheet1/tables/Table1/rows/add`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                values: [[
                    rowData.신청ID || '',
                    rowData.신청자계정 || '',
                    rowData.통 || '',
                    rowData.첨부파일 || '',
                    rowData.비고 || '',
                    rowData.Status || '신청',
                    rowData.등록자 || '',
                    rowData.등록일 || new Date().toISOString().split('T')[0],
                    rowData.처리자 || '',
                    rowData.처리일자 || '',
                    rowData.기존동일여부 || ''
                ]]
            })
        });
        
        if (!response.ok) {
            throw new Error('Failed to add row via Graph API');
        }
        
        return await response.json();
    } catch (error) {
        console.error('Error adding row via Graph API:', error);
        // 대체 방법: 전체 파일 다운로드 → 수정 → 업로드
        return await addRowByFileUpdate(rowData);
    }
}

/**
 * 파일 다운로드 → 수정 → 업로드 방식으로 행 추가
 */
async function addRowByFileUpdate(rowData) {
    try {
        // 기존 데이터 읽기
        let existingData = await readExcelAsFile();
        
        // 헤더 확인
        if (existingData.length === 0) {
            // 빈 파일인 경우 헤더 추가
            existingData = [[
                '신청ID', '신청자계정', '통', '첨부파일', '비고', 'Status',
                '등록자', '등록일', '처리자', '처리일자', '기존동일여부'
            ]];
        }
        
        // 새 행 추가
        const newRow = [
            rowData.신청ID || '',
            rowData.신청자계정 || '',
            rowData.통 || '',
            rowData.첨부파일 || '',
            rowData.비고 || '',
            rowData.Status || '신청',
            rowData.등록자 || '',
            rowData.등록일 || new Date().toISOString().split('T')[0],
            rowData.처리자 || '',
            rowData.처리일자 || '',
            rowData.기존동일여부 || ''
        ];
        
        existingData.push(newRow);
        
        // SheetJS로 Excel 파일 생성
        if (typeof XLSX === 'undefined') {
            throw new Error('XLSX library not loaded');
        }
        
        const worksheet = XLSX.utils.aoa_to_sheet(existingData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
        
        const excelBuffer = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
        const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        
        // OneDrive에 업로드
        await uploadFileToOneDrive(EXCEL_FILE_PATH, blob);
        
        return { success: true };
    } catch (error) {
        console.error('Error adding row by file update:', error);
        throw error;
    }
}

/**
 * Excel 파일에서 특정 행 업데이트
 * @param {string} applicationId - 신청ID
 * @param {Object} updateData - 업데이트할 데이터
 */
async function updateRowInExcel(applicationId, updateData) {
    try {
        // 전체 데이터 읽기
        const allData = await readExcelAsFile();
        
        if (allData.length === 0) {
            throw new Error('Excel file is empty');
        }
        
        // 헤더 찾기
        const headers = allData[0];
        const 신청IDIndex = headers.indexOf('신청ID');
        
        if (신청IDIndex === -1) {
            throw new Error('신청ID column not found');
        }
        
        // 해당 신청ID 찾기
        let rowIndex = -1;
        for (let i = 1; i < allData.length; i++) {
            if (allData[i][신청IDIndex] === applicationId) {
                rowIndex = i;
                break;
            }
        }
        
        if (rowIndex === -1) {
            throw new Error(`Application ID ${applicationId} not found`);
        }
        
        // 데이터 업데이트
        Object.keys(updateData).forEach(key => {
            const colIndex = headers.indexOf(key);
            if (colIndex !== -1) {
                allData[rowIndex][colIndex] = updateData[key];
            }
        });
        
        // 파일 다시 저장
        if (typeof XLSX === 'undefined') {
            throw new Error('XLSX library not loaded');
        }
        
        const worksheet = XLSX.utils.aoa_to_sheet(allData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
        
        const excelBuffer = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
        const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        
        await uploadFileToOneDrive(EXCEL_FILE_PATH, blob);
        
        return { success: true };
    } catch (error) {
        console.error('Error updating row:', error);
        throw error;
    }
}

/**
 * Excel 파일에서 특정 신청ID로 행 찾기
 * @param {string} applicationId - 신청ID
 * @returns {Promise<Object|null>} 행 데이터
 */
async function findRowByApplicationId(applicationId) {
    const allData = await readExcelAsFile();
    
    if (allData.length === 0) {
        return null;
    }
    
    const headers = allData[0];
    const 신청IDIndex = headers.indexOf('신청ID');
    
    if (신청IDIndex === -1) {
        return null;
    }
    
    for (let i = 1; i < allData.length; i++) {
        if (allData[i][신청IDIndex] === applicationId) {
            const row = {};
            headers.forEach((header, index) => {
                row[header] = allData[i][index] || '';
            });
            return row;
        }
    }
    
    return null;
}
