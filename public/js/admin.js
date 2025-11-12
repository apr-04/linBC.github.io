const API_BASE_URL = window.location.origin;
let currentApplicationId = '';

window.addEventListener('DOMContentLoaded', () => {
  loadApplications();
});

async function loadApplications() {
  const statusFilter = document.getElementById('statusFilter').value;
  const tbody = document.getElementById('applicationsTableBody');
  
  tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px;"><span class="loading"></span> 데이터를 불러오는 중...</td></tr>';

  try {
    const response = await fetch(`${API_BASE_URL}/api/admin-applications?status=${statusFilter}`);
    const result = await response.json();

    if (result.success) {
      displayApplications(result.data);
    } else {
      showAdminAlert('error', result.message || '데이터 로드 실패');
      tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px;">데이터를 불러올 수 없습니다.</td></tr>';
    }
  } catch (error) {
    console.error('Error:', error);
    showAdminAlert('error', '서버 연결 오류');
    tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px;">서버 연결 오류</td></tr>';
  }
}

function displayApplications(applications) {
  const tbody = document.getElementById('applicationsTableBody');
  
  if (applications.length === 0) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px;">신청 내역이 없습니다.</td></tr>';
    return;
  }

  tbody.innerHTML = applications.map(app => `
    <tr onclick="viewDetail('${app['신청ID']}')">
      <td>${app['신청ID']}</td>
      <td><span class="status-badge status-${app['Status']}">${app['Status']}</span></td>
      <td>${app['등록자']}</td>
      <td>${app['신청자계정']}</td>
      <td>${app['통']}통</td>
      <td>${app['변호사여부']}</td>
      <td>${formatDate(app['등록일'])}</td>
      <td><button onclick="event.stopPropagation(); viewDetail('${app['신청ID']}')">상세보기</button></td>
    </tr>
  `).join('');
}

async function viewDetail(applicationId) {
  currentApplicationId = applicationId;

  try {
    const response = await fetch(`${API_BASE_URL}/api/admin-detail?applicationId=${applicationId}`);
    const result = await response.json();

    if (result.success) {
      displayDetail(result.data);
      document.getElementById('detailModal').classList.remove('hidden');
    } else {
      showAdminAlert('error', '상세 정보를 불러올 수 없습니다.');
    }
  } catch (error) {
    console.error('Error:', error);
    showAdminAlert('error', '서버 연결 오류');
  }
}

function displayDetail(data) {
  const content = document.getElementById('detailContent');
  
  content.innerHTML = `
    <div style="background-color: #f8f9fa; padding: 20px; border-radius: 5px;">
      <table style="width: 100%;">
        <tr><td style="padding: 8px; font-weight: 600;">신청ID:</td><td style="padding: 8px;">${data['신청ID']}</td></tr>
        <tr><td style="padding: 8px; font-weight: 600;">상태:</td><td style="padding: 8px;"><span class="status-badge status-${data['Status']}">${data['Status']}</span></td></tr>
        <tr><td style="padding: 8px; font-weight: 600;">신청자:</td><td style="padding: 8px;">${data['등록자']}</td></tr>
        <tr><td style="padding: 8px; font-weight: 600;">이메일:</td><td style="padding: 8px;">${data['신청자계정']}</td></tr>
        <tr><td style="padding: 8px; font-weight: 600;">통 수:</td><td style="padding: 8px;">${data['통']}통</td></tr>
        <tr><td style="padding: 8px; font-weight: 600;">기존동일:</td><td style="padding: 8px;">${data['기존동일여부']}</td></tr>
        <tr><td style="padding: 8px; font-weight: 600;">변호사여부:</td><td style="padding: 8px;">${data['변호사여부']}</td></tr>
        ${data['담당변호사'] ? `<tr><td style="padding: 8px; font-weight: 600;">담당변호사:</td><td style="padding: 8px;">${data['담당변호사']}</td></tr>` : ''}
        ${data['비고'] ? `<tr><td style="padding: 8px; font-weight: 600;">비고:</td><td style="padding: 8px;">${data['비고']}</td></tr>` : ''}
        <tr><td style="padding: 8px; font-weight: 600;">신청일:</td><td style="padding: 8px;">${formatDate(data['등록일'])}</td></tr>
        ${data['첨부파일'] ? `<tr><td style="padding: 8px; font-weight: 600;">첨부파일:</td><td style="padding: 8px;"><a href="${data['첨부파일']}" target="_blank">파일 보기</a></td></tr>` : ''}
        ${data['처리자'] ? `<tr><td style="padding: 8px; font-weight: 600;">처리자:</td><td style="padding: 8px;">${data['처리자']}</td></tr>` : ''}
        ${data['처리일자'] ? `<tr><td style="padding: 8px; font-weight: 600;">처리일:</td><td style="padding: 8px;">${formatDate(data['처리일자'])}</td></tr>` : ''}
      </table>
    </div>
  `;

  document.getElementById('statusChange').value = data['Status'];
}

function closeModal() {
  document.getElementById('detailModal').classList.add('hidden');
  currentApplicationId = '';
}

async function uploadDraft() {
  const fileInput = document.getElementById('draftFile');
  const file = fileInput.files[0];

  if (!file) {
    alert('파일을 선택해주세요.');
    return;
  }

  if (!confirm('초안을 업로드하고 신청자에게 전달하시겠습니까?')) {
    return;
  }

  const formData = new FormData();
  formData.append('file', file);
  formData.append('applicationId', currentApplicationId);

  try {
    const response = await fetch(`${API_BASE_URL}/api/admin-upload`, {
      method: 'POST',
      body: formData
    });

    const result = await response.json();

    if (result.success) {
      showAdminAlert('success', '초안이 업로드되고 신청자에게 전달되었습니다.');
      fileInput.value = '';
      closeModal();
      loadApplications();
    } else {
      showAdminAlert('error', result.message || '업로드 실패');
    }
  } catch (error) {
    console.error('Error:', error);
    showAdminAlert('error', '서버 연결 오류');
  }
}

async function updateStatus() {
  const newStatus = document.getElementById('statusChange').value;

  if (!newStatus) {
    alert('상태를 선택해주세요.');
    return;
  }

  if (!confirm(`상태를 "${newStatus}"로 변경하시겠습니까?`)) {
    return;
  }

  try {
    const response = await fetch(`${API_BASE_URL}/api/admin-status`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        applicationId: currentApplicationId,
        status: newStatus
      })
    });

    const result = await response.json();

    if (result.success) {
      showAdminAlert('success', '상태가 변경되었습니다.');
      closeModal();
      loadApplications();
    } else {
      showAdminAlert('error', result.message || '상태 변경 실패');
    }
  } catch (error) {
    console.error('Error:', error);
    showAdminAlert('error', '서버 연결 오류');
  }
}

function formatDate(dateString) {
  if (!dateString) return '-';
  const date = new Date(dateString);
  return date.toLocaleString('ko-KR');
}

function showAdminAlert(type, message) {
  const alertBox = document.getElementById('adminAlert');
  alertBox.className = `alert alert-${type === 'success' ? 'success' : 'error'}`;
  alertBox.innerHTML = message;
  alertBox.classList.remove('hidden');

  setTimeout(() => {
    alertBox.classList.add('hidden');
  }, 5000);
}