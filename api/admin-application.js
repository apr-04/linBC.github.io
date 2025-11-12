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
    const { status } = req.query;
    let applications = await excelService.getAllRows();

    if (status && status !== 'all') {
      applications = applications.filter(app => app['Status'] === status);
    }

    res.json({
      success: true,
      data: applications
    });
  } catch (error) {
    console.error('목록 조회 오류:', error);
    res.status(500).json({
      success: false,
      message: '목록 조회 중 오류가 발생했습니다.'
    });
  }
};