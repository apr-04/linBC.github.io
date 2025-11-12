const API_BASE_URL = window.location.origin;

document.querySelectorAll('input[name="isLawyer"]').forEach(radio => {
  radio.addEventListener('change', (e) => {
    const lawyerNameGroup = document.getElementById('lawyerNameGroup');
    if (e.target.value === 'Staff') {
      lawyerNameGroup.classList.remove('hidden');
      document.getElementById('lawyerName').required = true;
    } else {
      lawyerNameGroup.classList.add('hidden');
      document.getElementById('lawyerName').required = false;
      document.getElementById('lawyerName').value = '';
    }
  });
});

document.getElementById('applicationForm').addEventListener('submit', async (e) => {
  e.preventDefault();

  const submitBtn = document.getElementById('submitBtn');
  submitBtn.disabled = true;
  submitBtn.innerHTML = '<span class="loading"></span> 처리 중...';

  const formData = {
    name: document.getElementById('name').value.trim(),
    email: document.getElementById('email').value.trim(),
    quantity: parseInt(document.getElementById('quantity').value),
    sameAsExisting: document.querySelector('input[name="sameAsExisting"]:checked').value,
    isLawyer: document.querySelector('input[name="isLawyer"]:checked').value,
    lawyerName: document.getElementById('lawyerName').value.trim(),
    remarks: document.getElementById('remarks').value.trim()
  };

  try {
    const response = await fetch(`${API_BASE_URL}/api/application-submit`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(formData)
    });

    const result = await response.json();

    if (result.success) {
      showAlert('success', `
        신청이 완료되었습니다!<br>
        <strong>신청ID: ${result.applicationId}</strong><br>
        확인 메일이 발송되었습니다.
      `);
      
      document.getElementById('applicationForm').reset();
      document.getElementById('lawyerNameGroup').classList.add('hidden');

      setTimeout(() => {
        location.href = `status.html?id=${result.applicationId}&email=${formData.email}`;
      }, 3000);
    } else {
      showAlert('error', result.message || '신청 처리 중 오류가 발생했습니다.');
      submitBtn.disabled = false;
      submitBtn.innerHTML = '신청하기';
    }
  } catch (error) {
    console.error('Error:', error);
    showAlert('error', '서버 연결 오류가 발생했습니다.');
    submitBtn.disabled = false;
    submitBtn.innerHTML = '신청하기';
  }
});

function showAlert(type, message) {
  const alertBox = document.getElementById('alertBox');
  alertBox.className = `alert alert-${type === 'success' ? 'success' : 'error'}`;
  alertBox.innerHTML = message;
  alertBox.classList.remove('hidden');
  alertBox.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}