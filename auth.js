/* ==========================================
   Simple Password Gate
   ========================================== */

(function() {
  const CORRECT_PASSWORD = 'Client216';
  const AUTH_KEY = 'portal_authenticated';

  // Check if already authenticated
  if (sessionStorage.getItem(AUTH_KEY) === 'true') return;

  // Create overlay
  const overlay = document.createElement('div');
  overlay.id = 'authGate';
  overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:#f5f5f5;z-index:99999;display:flex;align-items:center;justify-content:center;';

  overlay.innerHTML = `
    <div style="background:#fff;border-radius:12px;box-shadow:0 4px 24px rgba(0,0,0,0.12);padding:40px;max-width:380px;width:90%;text-align:center;">
      <div style="font-size:28px;margin-bottom:8px;">&#9730;</div>
      <div style="font-size:20px;font-weight:700;color:#333;margin-bottom:4px;">Client Portal</div>
      <div style="font-size:13px;color:#999;margin-bottom:24px;">powered by <span style="color:#e67e22;font-weight:700;">Arches</span></div>
      <p style="font-size:13px;color:#666;margin-bottom:20px;">This site is for authorized users only.<br>Please enter the access password.</p>
      <input type="password" id="authPassword" placeholder="Password"
        style="width:100%;padding:10px 14px;border:1px solid #ddd;border-radius:6px;font-size:14px;margin-bottom:12px;box-sizing:border-box;outline:none;">
      <div id="authError" style="color:#c62828;font-size:12px;margin-bottom:12px;display:none;">Incorrect password. Please try again.</div>
      <button id="authSubmit"
        style="width:100%;padding:10px;background:#e67e22;color:#fff;border:none;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;">
        Enter
      </button>
    </div>
  `;

  // Hide page content until authenticated
  document.documentElement.style.visibility = 'hidden';

  document.addEventListener('DOMContentLoaded', function() {
    document.body.appendChild(overlay);
    document.documentElement.style.visibility = 'visible';

    // Only show the gate, hide everything else
    const app = document.querySelector('.app');
    if (app) app.style.display = 'none';

    const pwInput = document.getElementById('authPassword');
    const errMsg = document.getElementById('authError');
    const submitBtn = document.getElementById('authSubmit');

    function tryLogin() {
      if (pwInput.value === CORRECT_PASSWORD) {
        sessionStorage.setItem(AUTH_KEY, 'true');
        overlay.remove();
        if (app) app.style.display = '';
      } else {
        errMsg.style.display = 'block';
        pwInput.value = '';
        pwInput.focus();
      }
    }

    submitBtn.addEventListener('click', tryLogin);
    pwInput.addEventListener('keydown', function(e) {
      if (e.key === 'Enter') tryLogin();
    });

    pwInput.focus();
  });
})();
