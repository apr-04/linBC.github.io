const excelService = require('../lib/excel');

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
    const { applicationId, status } = req.body;

    const validStatuses = ['신청', '초안전달', '수정요청', '제작', '지급완료', '삭제'];
    if (!validStatuses.includes(status)) {
      return res.status(400).json({
        success: false,
        message: '올바르지 않은 상태입니다.'
      });
    }

    await excelService.updateStatus(applicationId, status, '관리자');

    res.json({
      success: true,
      message: '상태가 변경되었습니다.'
    });
  } catch (error) {
    console.error('상태 변경 오류:', error);
    res.status(500).json({
      success: false,
      message: '상태 변경 중 오류가 발생했습니다.'
    });
  }
};