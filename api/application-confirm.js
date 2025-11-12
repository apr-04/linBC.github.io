const excelService = require('../lib/excel');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'GET') {
    return res.status(405).json({ success: false, message: 'Method not allowed' });
  }

  try {
    const { applicationId, email } = req.query;

    if (!applicationId || !email) {
      return res.status(400).json({
        success: false,
        message: '잘못된 접근입니다.'
      });
    }

    const application = await excelService.getApplicationByIdAndEmail(applicationId, email);

    if (!application) {
      return res.status(404).json({
        success: false,
        message: '해당 신청 내역을 찾을 수 없습니다.'
      });
    }

    if (application['Status'] !== '초안전달') {
      return res.status(400).json({
        success: false,
        message: '초안 확인 단계가 아닙니다.'
      });
    }

    res.json({
      success: true,
      data: {
        applicationId: application['신청ID'],
        name: application['등록자'],
        attachmentUrl: application['첨부파일']
      }
    });
  } catch (error) {
    console.error('확인 데이터 조회 오류:', error);
    res.status(500).json({
      success: false,
      message: '데이터 조회 중 오류가 발생했습니다.'
    });
  }
};