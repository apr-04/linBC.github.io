const excelService = require('../lib/excel');
const emailService = require('../lib/email');
const multiparty = require('multiparty');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, message: 'Method not allowed' });
  }

  const form = new multiparty.Form();

  form.parse(req, async (err, fields, files) => {
    if (err) {
      return res.status(400).json({
        success: false,
        message: '파일 업로드 오류'
      });
    }

    try {
      const applicationId = fields.applicationId[0];
      const file = files.file[0];

      if (!file) {
        return res.status(400).json({
          success: false,
          message: '파일을 업로드해주세요.'
        });
      }

      const application = await excelService.getApplicationById(applicationId);

      if (!application) {
        return res.status(404).json({
          success: false,
          message: '해당 신청을 찾을 수 없습니다.'
        });
      }

      const fs = require('fs');
      const fileBuffer = fs.readFileSync(file.path);

      const uploadResult = await excelService.uploadFile(
        fileBuffer,
        file.originalFilename,
        applicationId
      );

      await excelService.updateAttachment(applicationId, uploadResult.webUrl, '관리자');
      await excelService.updateStatus(applicationId, '초안전달', '관리자');

      emailService.sendDraftConfirmation(application['신청자계정'], applicationId)
        .catch(err => console.error('이메일 발송 오류:', err));

      res.json({
        success: true,
        message: '초안이 업로드되고 신청자에게 전달되었습니다.'
      });
    } catch (error) {
      console.error('초안 업로드 오류:', error);
      res.status(500).json({
        success: false,
        message: '업로드 중 오류가 발생했습니다.'
      });
    }
  });
};