const API_BASE_URL = window.location.origin;

window.addEventListener('DOMContentLoaded', () => {
  const urlParams = new URLSearchParams(window.location.search);
  const id = urlParams.get('id');
  const email = urlParams.get('email');

  if (id && email) {
    document.getElementById('applicationId').value = id;
    document.getElementById('email').value = email;
    document.getElementById('statusForm').dispatchEvent(new Event('submit'));
  }
});

document.getElementById('statusForm').addEventListener('submit', async (e) => {
  e.preventDefault();

  const searchBtn = document.getElementById('searchBtn');
  searchBtn.disabled = true;
  searchBtn.innerHTML = '<span class="loading"></span> 조회 중...';

  const applicationId = document.getElementById('applicationId').value.trim();
  const email = document.getElementById('email').value.trim();

  try {
    const response = await fetch(
      `${API_BASE_URL}/api/application-status?applicationId=${encodeURIComponent(applicationId)}&email=${encodeURIComponent(email)}`
    );

    const result = await response.json();

    if (result.success) {
      displayResult(result.data);
      document.getElementById('alertBox').classList.add('hidden');
    } else {
      showAlert('error', result.message || '조회 중 오류가 발생했습니다.');
      document.getElementById('statusResult').classList.add('hidden');
    }
  } catch (error) {
    console.error('Error:', error);
    showAlert('error', '서버 연결 오류가 발생했습니다.');
    document.getElementById('statusResult').classList.add('hidden');
  } finally {
    searchBtn.disabled = false;
    searchBtn.innerHTML = '조회하기';
  }
});

function displayResult(data) {
  document.getElementById('resultId').textContent = data.applicationId;
  document.getElementById('resultName').textContent = data.name;
  document.getElementById('resultQuantity').textContent = data.quantity + '통';
  document.getElementById('resultDate').textContent = formatDate(data.registeredDate);
  
  const statusBadge = document.getElementById('resultStatus');
  statusBadge.textContent = data.status;
  statusBadge.className = `status-badge status-${data.status}`;

  if (data.processedDate) {
    document.getElementById('processedDateItem').classList.remove('hidden');
    document.getElementById('resultProcessedDate').textContent = formatDate(data.processedDate);
  } else {
    document.getElementById('processedDateItem').classList.add('hidden');
  }

  document.getElementById('statusResult').classList.remove('hidden');
  document.getElementById('statusResult').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function formatDate(dateString) {
  if (!dateString) return '-';
  const date = new Date(dateString);
  return date.toLocaleString('ko-KR');
}

function showAlert(type, message) {
  const alertBox = document.getElementById('alertBox');
  alertBox.className = `alert alert-${type === 'success' ? 'success' : 'error'}`;
  alertBox.innerHTML = message;
  alertBox.classList.remove('hidden');
  alertBox.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}