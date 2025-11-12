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
        message: '신청ID와 이메일을 입력해주세요.'
      });
    }

    const application = await excelService.getApplicationByIdAndEmail(applicationId, email);

    if (!application) {
      return res.status(404).json({
        success: false,
        message: '해당 신청 내역을 찾을 수 없습니다.'
      });
    }

    res.json({
      success: true,
      data: {
        applicationId: application['신청ID'],
        status: application['Status'],
        name: application['등록자'],
        quantity: application['통'],
        registeredDate: application['등록일'],
        processedDate: application['처리일자']
      }
    });
  } catch (error) {
    console.error('상태 조회 오류:', error);
    res.status(500).json({
      success: false,
      message: '상태 조회 중 오류가 발생했습니다.'
    });
  }
};