const nodemailer = require('nodemailer');

class EmailService {
  constructor() {
    this.transporter = nodemailer.createTransport({
      host: process.env.EMAIL_HOST,
      port: parseInt(process.env.EMAIL_PORT),
      secure: false,
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASSWORD
      }
    });
  }

  async sendApplicationConfirmation(toEmail, applicationId) {
    const baseUrl = process.env.BASE_URL || process.env.VERCEL_URL || 'http://localhost:3000';
    
    const mailOptions = {
      from: process.env.EMAIL_USER,
      to: toEmail,
      subject: '[법무법인 린] 명함 신청이 접수되었습니다',
      html: `
        <h2>명함 신청 접수 확인</h2>
        <p>안녕하세요,</p>
        <p>귀하의 명함 신청이 정상적으로 접수되었습니다.</p>
        <p><strong>신청ID:</strong> ${applicationId}</p>
        <p>처리 상태는 <a href="${baseUrl}/status.html">여기</a>에서 확인하실 수 있습니다.</p>
        <br>
        <p>감사합니다.</p>
        <p>법무법인 린</p>
      `
    };

    return await this.transporter.sendMail(mailOptions);
  }

  async sendNewApplicationAlert(applicationData) {
    const baseUrl = process.env.BASE_URL || process.env.VERCEL_URL || 'http://localhost:3000';
    
    const mailOptions = {
      from: process.env.EMAIL_USER,
      to: process.env.ADMIN_EMAIL,
      subject: '[명함관리] 새로운 명함 신청',
      html: `
        <h2>새로운 명함 신청</h2>
        <ul>
          <li><strong>신청ID:</strong> ${applicationData.applicationId}</li>
          <li><strong>신청자:</strong> ${applicationData.name}</li>
          <li><strong>이메일:</strong> ${applicationData.email}</li>
          <li><strong>통 수:</strong> ${applicationData.quantity}</li>
          <li><strong>변호사 여부:</strong> ${applicationData.isLawyer}</li>
          ${applicationData.lawyerName ? `<li><strong>담당변호사:</strong> ${applicationData.lawyerName}</li>` : ''}
          ${applicationData.remarks ? `<li><strong>비고:</strong> ${applicationData.remarks}</li>` : ''}
        </ul>
        <p><a href="${baseUrl}/admin.html">관리자 페이지로 이동</a></p>
      `
    };

    return await this.transporter.sendMail(mailOptions);
  }

  async sendDraftConfirmation(toEmail, applicationId) {
    const baseUrl = process.env.BASE_URL || process.env.VERCEL_URL || 'http://localhost:3000';
    const confirmLink = `${baseUrl}/confirm.html?id=${applicationId}&email=${encodeURIComponent(toEmail)}`;
    
    const mailOptions = {
      from: process.env.EMAIL_USER,
      to: toEmail,
      subject: '[법무법인 린] 명함 초안 확인 요망',
      html: `
        <h2>명함 초안 확인 요청</h2>
        <p>안녕하세요,</p>
        <p>명함 초안이 준비되었습니다. 아래 링크를 통해 확인해주세요.</p>
        <p><strong>신청ID:</strong> ${applicationId}</p>
        <p><a href="${confirmLink}" style="display:inline-block;padding:10px 20px;background-color:#0066cc;color:white;text-decoration:none;border-radius:5px;">초안 확인하기</a></p>
        <p>확인 후 제작 신청 또는 수정 요청을 선택해주세요.</p>
        <br>
        <p>감사합니다.</p>
        <p>법무법인 린</p>
      `
    };

    return await this.transporter.sendMail(mailOptions);
  }

  async sendProductionApproval(applicationId, applicantName) {
    const baseUrl = process.env.BASE_URL || process.env.VERCEL_URL || 'http://localhost:3000';
    
    const mailOptions = {
      from: process.env.EMAIL_USER,
      to: process.env.ADMIN_EMAIL,
      subject: '[명함관리] 제작 승인',
      html: `
        <h2>명함 제작 승인</h2>
        <p><strong>신청ID:</strong> ${applicationId}</p>
        <p><strong>신청자:</strong> ${applicantName}</p>
        <p>신청자가 초안을 승인하고 제작을 요청했습니다.</p>
        <p><a href="${baseUrl}/admin.html">관리자 페이지로 이동</a></p>
      `
    };

    return await this.transporter.sendMail(mailOptions);
  }

  async sendModificationRequest(applicationId, applicantName, modificationReason) {
    const baseUrl = process.env.BASE_URL || process.env.VERCEL_URL || 'http://localhost:3000';
    
    const mailOptions = {
      from: process.env.EMAIL_USER,
      to: process.env.ADMIN_EMAIL,
      subject: '[명함관리] 수정 요청',
      html: `
        <h2>명함 수정 요청</h2>
        <p><strong>신청ID:</strong> ${applicationId}</p>
        <p><strong>신청자:</strong> ${applicantName}</p>
        <p><strong>수정 요청 사유:</strong></p>
        <p>${modificationReason}</p>
        <p><a href="${baseUrl}/admin.html">관리자 페이지로 이동</a></p>
      `
    };

    return await this.transporter.sendMail(mailOptions);
  }
}

module.exports = new EmailService();