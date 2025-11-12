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
    const { applicationId } = req.query;
    const application = await excelService.getApplicationById(applicationId);

    if (!application) {
      return res.status(404).json({
        success: false,
        message: '해당 신청을 찾을 수 없습니다.'
      });
    }

    res.json({
      success: true,
      data: application
    });
  } catch (error) {
    console.error('상세 조회 오류:', error);
    res.status(500).json({
      success: false,
      message: '조회 중 오류가 발생했습니다.'
    });
  }
};