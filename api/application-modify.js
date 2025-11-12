const excelService = require('../lib/excel');
const emailService = require('../lib/email');

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

  try {
    const { applicationId, email, reason } = req.body;

    if (!reason) {
      return res.status(400).json({
        success: false,
        message: '수정 요청 사유를 입력해주세요.'
      });
    }

    const application = await excelService.getApplicationByIdAndEmail(applicationId, email);

    if (!application) {
      return res.status(404).json({
        success: false,
        message: '해당 신청 내역을 찾을 수 없습니다.'
      });
    }

    await excelService.updateStatus(applicationId, '수정요청');

    emailService.sendModificationRequest(applicationId, application['등록자'], reason)
      .catch(err => console.error('이메일 발송 오류:', err));

    res.json({
      success: true,
      message: '수정 요청이 전달되었습니다.'
    });
  } catch (error) {
    console.error('수정 요청 오류:', error);
    res.status(500).json({
      success: false,
      message: '처리 중 오류가 발생했습니다.'
    });
  }
};