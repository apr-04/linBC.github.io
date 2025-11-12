const API_BASE_URL = window.location.origin;

let currentApplicationId = '';
let currentEmail = '';

window.addEventListener('DOMContentLoaded', async () => {
  const urlParams = new URLSearchParams(window.location.search);
  currentApplicationId = urlParams.get('id');
  currentEmail = urlParams.get('email');

  if (!currentApplicationId || !currentEmail) {
    showAlert('error', '잘못된 접근입니다.');
    document.getElementById('loadingBox').classList.add('hidden');
    return;
  }

  await loadConfirmationData();
});

async function loadConfirmationData() {
  try {
    const response = await fetch(
      `${API_BASE_URL}/api/application-confirm?applicationId=${encodeURIComponent(currentApplicationId)}&email=${encodeURIComponent(currentEmail)}`
    );

    const result = await response.json();

    if (result.success) {
      displayConfirmationData(result.data);
      document.getElementById('loadingBox').classList.add('hidden');
      document.getElementById('contentBox').classList.remove('hidden');
    } else {
      showAlert('error', result.message || '데이터를 불러올 수 없습니다.');
      document.getElementById('loadingBox').classList.add('hidden');
    }
  } catch (error) {
    console.error('Error:', error);
    showAlert('error', '서버 연결 오류가 발생했습니다.');
    document.getElementById('loadingBox').classList.add('hidden');
  }
}

function displayConfirmationData(data) {
  document.getElementById('applicationId').textContent = data.applicationId;
  document.getElementById('applicantName').textContent = data.name;

  const fileContent = document.getElementById('fileContent');
  
  if (data.attachmentUrl) {
    const fileExtension = data.attachmentUrl.split('.').pop().toLowerCase();
    
    if (['jpg', 'jpeg', 'png', 'gif'].includes(fileExtension)) {
      fileContent.innerHTML = `<img src="${data.attachmentUrl}" alt="명함 초안">`;
    } else if (fileExtension === 'pdf') {
      fileContent.innerHTML = `
        <iframe src="${data.attachmentUrl}" style="width: 100%; height: 600px; border: none;"></iframe>
      `;
    } else {
      fileContent.innerHTML = `<p>파일을 다운로드하여 확인해주세요.</p>`;
    }
    
    fileContent.innerHTML += `
      <a href="${data.attachmentUrl}" target="_blank">📥 파일 다운로드 / 새 창에서 보기</a>
    `;
  } else {
    fileContent.innerHTML = '<p>첨부된 파일이 없습니다.</p>';
  }
}

function showModifyForm() {
  document.getElementById('modifyForm').classList.remove('hidden');
  document.getElementById('modifyReason').focus();
}

function hideModifyForm() {
  document.getElementById('modifyForm').classList.add('hidden');
  document.getElementById('modifyReason').value = '';
}

async function approveProduction() {
  if (!confirm('제작을 승인하시겠습니까?')) {
    return;
  }

  const approveBtn = document.getElementById('approveBtn');
  approveBtn.disabled = true;
  approveBtn.innerHTML = '<span class="loading"></span> 처리 중...';

  try {
    const response = await fetch(`${API_BASE_URL}/api/application-approve`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        applicationId: currentApplicationId,
        email: currentEmail
      })
    });

    const result = await response.json();

    if (result.success) {
      showAlert('success', '제작이 승인되었습니다. 완료 후 연락드리겠습니다.');
      document.getElementById('actionBox').classList.add('hidden');
    } else {
      showAlert('error', result.message || '처리 중 오류가 발생했습니다.');
      approveBtn.disabled = false;
      approveBtn.innerHTML = '✅ 제작 신청';
    }
  } catch (error) {
    console.error('Error:', error);
    showAlert('error', '서버 연결 오류가 발생했습니다.');
    approveBtn.disabled = false;
    approveBtn.innerHTML = '✅ 제작 신청';
  }
}

async function submitModification() {
  const reason = document.getElementById('modifyReason').value.trim();

  if (!reason) {
    alert('수정 요청 사유를 입력해주세요.');
    return;
  }

  if (!confirm('수정 요청을 제출하시겠습니까?')) {
    return;
  }

  try {
    const response = await fetch(`${API_BASE_URL}/api/application-modify`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        applicationId: currentApplicationId,
        email: currentEmail,
        reason: reason
      })
    });

    const result = await response.json();

    if (result.success) {
      showAlert('success', '수정 요청이 전달되었습니다. 수정된 초안은 이메일로 다시 발송됩니다.');
      document.getElementById('actionBox').classList.add('hidden');
    } else {
      showAlert('error', result.message || '처리 중 오류가 발생했습니다.');
    }
  } catch (error) {
    console.error('Error:', error);
    showAlert('error', '서버 연결 오류가 발생했습니다.');
  }
}

function showAlert(type, message) {
  const alertBox = document.getElementById('alertBox');
  alertBox.className = `alert alert-${type === 'success' ? 'success' : 'error'}`;
  alertBox.innerHTML = message;
  alertBox.classList.remove('hidden');
  alertBox.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}