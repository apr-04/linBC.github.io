const { getGraphClient } = require('./graph');

class ExcelService {
  constructor() {
    this.driveId = process.env.DRIVE_ID;
    this.workbookPath = process.env.WORKBOOK_PATH;
    this.worksheetName = process.env.WORKSHEET_NAME;
  }

  generateApplicationId() {
    const now = new Date();
    const year = now.getFullYear();
    const random = Math.floor(Math.random() * 9000) + 1000;
    return `LN-${year}-${random}`;
  }

  async addRow(data) {
    try {
      const client = await getGraphClient();
      const applicationId = this.generateApplicationId();
      const now = new Date().toISOString();

      const rowData = [
        applicationId,
        '신청',
        data.email,
        data.name,
        data.quantity,
        data.sameAsExisting,
        data.isLawyer,
        data.lawyerName || '',
        data.remarks || '',
        now,
        '',
        '',
        ''
      ];

      const endpoint = `/drives/${this.driveId}/items/root:${this.workbookPath}:/workbook/worksheets/${this.worksheetName}/tables/Table1/rows/add`;
      
      await client.api(endpoint).post({
        values: [rowData]
      });

      return { success: true, applicationId };
    } catch (error) {
      console.error('Excel 행 추가 오류:', error);
      throw error;
    }
  }

  async getAllRows() {
    try {
      const client = await getGraphClient();
      const endpoint = `/drives/${this.driveId}/items/root:${this.workbookPath}:/workbook/worksheets/${this.worksheetName}/usedRange`;
      
      const result = await client.api(endpoint).get();
      const headers = result.values[0];
      const rows = result.values.slice(1);
      
      return rows.map(row => {
        const obj = {};
        headers.forEach((header, index) => {
          obj[header] = row[index];
        });
        return obj;
      });
    } catch (error) {
      console.error('Excel 데이터 조회 오류:', error);
      throw error;
    }
  }

  async getApplicationByIdAndEmail(applicationId, email) {
    try {
      const allRows = await this.getAllRows();
      return allRows.find(row => 
        row['신청ID'] === applicationId && row['신청자계정'] === email
      );
    } catch (error) {
      console.error('신청 조회 오류:', error);
      throw error;
    }
  }

  async getApplicationById(applicationId) {
    try {
      const allRows = await this.getAllRows();
      return allRows.find(row => row['신청ID'] === applicationId);
    } catch (error) {
      console.error('신청 조회 오류:', error);
      throw error;
    }
  }

  async updateStatus(applicationId, status, adminName = '') {
    try {
      const client = await getGraphClient();
      const allRows = await this.getAllRows();
      const rowIndex = allRows.findIndex(row => row['신청ID'] === applicationId);
      
      if (rowIndex === -1) {
        throw new Error('신청을 찾을 수 없습니다.');
      }

      const actualRowIndex = rowIndex + 2;
      const now = new Date().toISOString();

      const endpoint = `/drives/${this.driveId}/items/root:${this.workbookPath}:/workbook/worksheets/${this.worksheetName}/range(address='B${actualRowIndex}')`;
      await client.api(endpoint).patch({ values: [[status]] });

      const endpoint2 = `/drives/${this.driveId}/items/root:${this.workbookPath}:/workbook/worksheets/${this.worksheetName}/range(address='L${actualRowIndex}:M${actualRowIndex}')`;
      await client.api(endpoint2).patch({ values: [[adminName, now]] });

      return { success: true };
    } catch (error) {
      console.error('상태 업데이트 오류:', error);
      throw error;
    }
  }

  async updateAttachment(applicationId, fileLink, adminName = '') {
    try {
      const client = await getGraphClient();
      const allRows = await this.getAllRows();
      const rowIndex = allRows.findIndex(row => row['신청ID'] === applicationId);
      
      if (rowIndex === -1) {
        throw new Error('신청을 찾을 수 없습니다.');
      }

      const actualRowIndex = rowIndex + 2;
      const now = new Date().toISOString();

      const endpoint = `/drives/${this.driveId}/items/root:${this.workbookPath}:/workbook/worksheets/${this.worksheetName}/range(address='K${actualRowIndex}:M${actualRowIndex}')`;
      
      await client.api(endpoint).patch({
        values: [[fileLink, adminName, now]]
      });

      return { success: true };
    } catch (error) {
      console.error('첨부파일 업데이트 오류:', error);
      throw error;
    }
  }

  async uploadFile(buffer, fileName, applicationId) {
    try {
      const client = await getGraphClient();
      const uploadFolder = '/명함초안';
      const fullFileName = `${applicationId}_${Date.now()}_${fileName}`;
      const filePath = `${uploadFolder}/${fullFileName}`;

      const uploadResult = await client.api(`/drives/${this.driveId}/root:${filePath}:/content`)
        .put(buffer);

      const shareLink = await client.api(`/drives/${this.driveId}/items/${uploadResult.id}/createLink`)
        .post({
          type: 'view',
          scope: 'organization'
        });

      return {
        fileId: uploadResult.id,
        fileName: fullFileName,
        webUrl: shareLink.link.webUrl
      };
    } catch (error) {
      console.error('파일 업로드 오류:', error);
      throw error;
    }
  }
}

module.exports = new ExcelService();