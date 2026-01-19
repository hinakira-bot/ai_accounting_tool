document.addEventListener('DOMContentLoaded', () => {
    // --- Auth & Config ---
    const CLIENT_ID = '353694435064-r6mlbk3mm2mflhl2mot2n94dpuactscc.apps.googleusercontent.com';
    const SCOPES = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file';
    let tokenClient;
    let accessToken = sessionStorage.getItem('access_token');
    let tokenExpiration = sessionStorage.getItem('token_expiration');

    // DOM Elements
    const authBtn = document.getElementById('auth-btn');
    const settingsBtn = document.getElementById('settings-btn');
    const settingsModal = document.getElementById('settings-modal');
    const saveSettingsBtn = document.getElementById('save-settings-btn');
    const closeModalSpan = document.querySelector('.close-modal');
    const loginOverlay = document.getElementById('login-overlay');
    const overlayLoginBtn = document.getElementById('overlay-login-btn');

    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const statusArea = document.getElementById('status-area');
    const resultsArea = document.getElementById('results-area');
    const resultsTable = document.querySelector('#results-table tbody');
    const sendBtn = document.getElementById('send-btn');
    const resetBtn = document.getElementById('reset-btn');
    const addMoreBtn = document.getElementById('add-more-btn');
    const addMoreInput = document.getElementById('add-more-input');

    let extractedData = [];

    // --- Google Identity Services Initialization ---
    window.onload = function () {
        if (typeof google === 'undefined') {
            console.error("Google Identity Services not loaded");
            return;
        }
        tokenClient = google.accounts.oauth2.initTokenClient({
            client_id: CLIENT_ID,
            scope: SCOPES,
            callback: (tokenResponse) => {
                if (tokenResponse.access_token) {
                    accessToken = tokenResponse.access_token;
                    // Token is valid for ~1 hour (3599 seconds)
                    const expiresIn = tokenResponse.expires_in;
                    const expirationTime = new Date().getTime() + (expiresIn * 1000);

                    sessionStorage.setItem('access_token', accessToken);
                    sessionStorage.setItem('token_expiration', expirationTime);

                    onLoginSuccess();
                }
            },
        });

        // Check if we have a valid token in session
        if (accessToken && tokenExpiration && new Date().getTime() < parseInt(tokenExpiration)) {
            onLoginSuccess();
        } else {
            loginOverlay.classList.remove('hidden'); // Show login overlay if not logged in
        }
    }

    function handleLogin() {
        if (tokenClient) {
            tokenClient.requestAccessToken();
        }
    }

    function onLoginSuccess() {
        loginOverlay.classList.add('hidden');
        authBtn.textContent = 'ログアウト';
        authBtn.onclick = handleLogout;
        settingsBtn.style.display = 'block';

        // Check local storage for settings
        const apiKey = localStorage.getItem('gemini_api_key');
        const sheetId = localStorage.getItem('spreadsheet_id');

        if (!apiKey || !sheetId) {
            openSettings(); // Prompt user to setup if missing
        }
    }

    function handleLogout() {
        const token = sessionStorage.getItem('access_token');
        if (token) {
            google.accounts.oauth2.revoke(token, () => { console.log('Revoked'); });
        }
        sessionStorage.removeItem('access_token');
        sessionStorage.removeItem('token_expiration');
        location.reload();
    }

    // --- Settings Logic ---
    function openSettings() {
        document.getElementById('api-key-input').value = localStorage.getItem('gemini_api_key') || '';
        document.getElementById('spreadsheet-id-input').value = localStorage.getItem('spreadsheet_id') || '';
        settingsModal.classList.remove('hidden');
    }

    function saveSettings() {
        const key = document.getElementById('api-key-input').value.trim();
        const sheet = document.getElementById('spreadsheet-id-input').value.trim();

        if (key && sheet) {
            localStorage.setItem('gemini_api_key', key);
            localStorage.setItem('spreadsheet_id', sheet);
            settingsModal.classList.add('hidden');
            alert('設定を保存しました');
        } else {
            alert('両方の項目を入力してください');
        }
    }

    // UI Event Listeners
    authBtn.onclick = handleLogin;
    overlayLoginBtn.onclick = handleLogin;
    settingsBtn.onclick = openSettings;
    saveSettingsBtn.onclick = saveSettings;
    closeModalSpan.onclick = () => settingsModal.classList.add('hidden');
    window.onclick = (event) => {
        if (event.target == settingsModal) settingsModal.classList.add('hidden');
    }

    // --- File Handling (Modified for Auth) ---
    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('active'); });
    dropZone.addEventListener('dragleave', () => { dropZone.classList.remove('active'); });
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('active');
        handleFiles(e.dataTransfer.files);
    });
    dropZone.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', () => handleFiles(fileInput.files));
    addMoreBtn.addEventListener('click', () => addMoreInput.click());
    addMoreInput.addEventListener('change', () => handleFiles(addMoreInput.files, true));

    function handleFiles(files, append = false) {
        if (files.length === 0) return;

        const apiKey = localStorage.getItem('gemini_api_key');
        const sheetId = localStorage.getItem('spreadsheet_id');

        if (!apiKey || !sheetId) {
            alert('設定画面でAPIキーとスプレッドシートIDを設定してください。');
            openSettings();
            return;
        }

        if (!append) dropZone.classList.add('hidden');
        statusArea.classList.remove('hidden');
        document.getElementById('status-text').textContent = `${files.length}枚のファイルを解析中...`;

        const formData = new FormData();
        for (let i = 0; i < files.length; i++) {
            formData.append('files', files[i]);
        }

        // Append Config
        formData.append('gemini_api_key', apiKey);
        formData.append('spreadsheet_id', sheetId);
        formData.append('access_token', accessToken); // Add OAuth token

        fetch('/api/analyze', {
            method: 'POST',
            body: formData
        })
            .then(response => response.json())
            .then(data => {
                if (data.error) throw new Error(data.error);
                extractedData = append ? [...extractedData, ...data] : data;
                renderResults(extractedData);
                statusArea.classList.add('hidden');
                resultsArea.classList.remove('hidden');
            })
            .catch(err => {
                console.error(err);
                alert(`エラー: ${err.message}`);
                statusArea.classList.add('hidden');
                if (!append && extractedData.length === 0) dropZone.classList.remove('hidden');
            });
    }

    function renderResults(data) {
        resultsTable.innerHTML = '';
        let hasDuplicate = false;
        data.forEach((item, index) => {
            if (item.is_duplicate) hasDuplicate = true;
            const row = document.createElement('tr');
            if (item.is_duplicate) row.classList.add('duplicate');
            row.innerHTML = `
                <td><input type="text" value="${item.date || ''}" data-index="${index}" data-key="date"></td>
                <td><input type="text" value="${item.debit_account || ''}" data-index="${index}" data-key="debit_account"></td>
                <td><input type="text" value="${item.credit_account || ''}" data-index="${index}" data-key="credit_account"></td>
                <td><input type="number" value="${item.amount || 0}" data-index="${index}" data-key="amount"></td>
                <td><input type="text" value="${item.counterparty || ''}" data-index="${index}" data-key="counterparty"></td>
                <td><input type="text" value="${item.memo || ''}" data-index="${index}" data-key="memo"></td>
            `;
            resultsTable.appendChild(row);
        });

        const duplicateAlert = document.getElementById('duplicate-alert');
        duplicateAlert.style.display = hasDuplicate ? 'flex' : 'none';

        resultsTable.querySelectorAll('input').forEach(input => {
            input.addEventListener('change', (e) => {
                extractedData[e.target.dataset.index][e.target.dataset.key] = e.target.value;
            });
        });
    }

    sendBtn.addEventListener('click', () => {
        const apiKey = localStorage.getItem('gemini_api_key');
        const sheetId = localStorage.getItem('spreadsheet_id');

        sendBtn.disabled = true;
        sendBtn.textContent = '送信中...';

        fetch('/api/save', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                data: extractedData,
                gemini_api_key: apiKey,
                spreadsheet_id: sheetId,
                access_token: accessToken
            })
        })
            .then(response => response.json())
            .then(data => {
                if (data.error) throw new Error(data.error);
                alert('スプレッドシートに保存しました！');
                resetUI();
            })
            .catch(err => {
                console.error(err);
                alert(`保存エラー: ${err.message}`);
            })
            .finally(() => {
                sendBtn.disabled = false;
                sendBtn.textContent = 'スプレッドシートに書き込む';
            });
    });

    resetBtn.addEventListener('click', resetUI);
    function resetUI() {
        resultsArea.classList.add('hidden');
        dropZone.classList.remove('hidden');
        fileInput.value = '';
        extractedData = [];
    }
});
