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

    const addManualBtn = document.getElementById('add-manual-btn');
    let accountOptions = [];

    // --- Fetch Accounts for Autocomplete ---
    function fetchAccounts() {
        if (accountOptions.length > 0) return;
        fetch('/api/accounts')
            .then(res => res.json())
            .then(data => {
                if (data.accounts) accountOptions = data.accounts;
                setupDatalist();
            })
            .catch(err => console.error("Account fetch failed", err));
    }

    function setupDatalist() {
        let dl = document.getElementById('account-datalist');
        if (!dl) {
            dl = document.createElement('datalist');
            dl.id = 'account-datalist';
            document.body.appendChild(dl);
        }
        dl.innerHTML = accountOptions.map(acc => `<option value="${acc}">`).join('');
    }

    addManualBtn.addEventListener('click', () => {
        // Fetch accounts if not yet done (lazy load)
        fetchAccounts();

        // Check Payment Type Selector
        // (Removed in favor of Split Button Auto-Detect)
        let creditAcc = '';


        // Add row with optional pre-filled credit account
        const newItem = {
            date: new Date().toISOString().split('T')[0], // Today
            debit_account: '',
            credit_account: creditAcc,
            amount: 0,
            counterparty: '',
            memo: '',
            source: 'manual' // Flag for backend routing
        };

        extractedData.push(newItem);
        renderResults(extractedData);

        // Focus on the first input of the new row
        setTimeout(() => {
            const inputs = resultsTable.querySelectorAll('input');
            if (inputs.length > 0) inputs[inputs.length - 6].focus(); // approximate
        }, 100);
    });

    // --- New: Start Manual Entry from Landing ---
    const startManualBtn = document.getElementById('start-manual-btn');
    if (startManualBtn) {
        startManualBtn.addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent dropZone file picker
            e.preventDefault();
            fetchAccounts(); // Ensure accounts are loaded

            // Hide drop zone, show results
            dropZone.classList.add('hidden');
            resultsArea.classList.remove('hidden');

            // Add initial empty row
            const newItem = {
                date: new Date().toISOString().split('T')[0],
                debit_account: '',
                credit_account: '',
                amount: 0,
                counterparty: '',
                memo: '',
                source: 'manual'
            };
            extractedData = [newItem];
            renderResults(extractedData);

            // Focus
            setTimeout(() => {
                const inputs = resultsTable.querySelectorAll('input');
                if (inputs.length > 0) inputs[0].focus();
            }, 100);
        });
    }

    function renderResults(data) {
        resultsTable.innerHTML = '';
        let hasDuplicate = false;

        // Ensure datalist exists whenever rendering
        if (accountOptions.length > 0) setupDatalist();
        else fetchAccounts();

        data.forEach((item, index) => {
            if (item.is_duplicate) hasDuplicate = true;
            const row = document.createElement('tr');
            if (item.is_duplicate) row.classList.add('duplicate');
            if (item.source === 'manual') row.style.backgroundColor = '#f0f8ff20'; // Slight highlight for manual

            row.innerHTML = `
                <td><input type="date" value="${item.date || ''}" data-index="${index}" data-key="date" style="width:110px;"></td>
                <td><input type="text" list="account-datalist" value="${item.debit_account || ''}" data-index="${index}" data-key="debit_account" placeholder="借方科目"></td>
                <td><input type="text" list="account-datalist" value="${item.credit_account || ''}" data-index="${index}" data-key="credit_account" placeholder="貸方科目"></td>
                <td><input type="number" value="${item.amount || 0}" data-index="${index}" data-key="amount"></td>
                <td><input type="text" value="${item.counterparty || ''}" data-index="${index}" data-key="counterparty" placeholder="取引先"></td>
                <td><input type="text" value="${item.memo || ''}" data-index="${index}" data-key="memo" placeholder="摘要"></td>
                <td>
                    <button class="delete-row-btn" data-index="${index}" style="background:none; border:none; color:#f66; cursor:pointer;">×</button>
                </td>
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

        resultsTable.querySelectorAll('.delete-row-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const idx = parseInt(e.target.dataset.index);
                extractedData.splice(idx, 1);
                renderResults(extractedData);
            });
        });
    }

    // --- Dropdown Menu Logic ---
    const menuBtn = document.getElementById('auto-predict-menu-btn');
    const dropdown = document.getElementById('auto-predict-dropdown');

    if (menuBtn && dropdown) {
        menuBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            dropdown.style.display = dropdown.style.display === 'block' ? 'none' : 'block';
        });

        // Close dropdown when clicking outside
        document.addEventListener('click', () => {
            dropdown.style.display = 'none';
        });

        // Handle Menu Item Clicks
        dropdown.querySelectorAll('a').forEach(link => {
            link.addEventListener('click', (e) => {
                e.preventDefault();
                const mode = e.target.dataset.mode; // unpaid, cash, bank, auto
                runAutoPredict(mode);
            });
        });
    }

    function runAutoPredict(mode) {
        // FORCE SYNC (Keep this robust logic)
        const inputs = resultsTable.querySelectorAll('input');
        inputs.forEach(input => {
            const idx = parseInt(input.dataset.index);
            const key = input.dataset.key;
            if (!isNaN(idx) && key && extractedData[idx]) extractedData[idx][key] = input.value;
        });

        // 1. Pre-fill Credit Account based on Mode (Local Override)
        let forcedCredit = null;
        if (mode === 'unpaid') forcedCredit = '未払金';
        if (mode === 'cash') forcedCredit = '現金';
        if (mode === 'bank') forcedCredit = '普通預金';

        // 2. Find Targets
        // If mode is specific (e.g. Cash), we target rows where Credit is Empty OR matches target (to refine) and Debit is Empty.
        // Actually, simpler: Target rows where Debit is Empty. And set Credit to forced value if empty or different?
        // Let's stick to safe logic: Target rows with missing accounts.

        const targets = extractedData.map((item, idx) => ({ ...item, index: idx }))
            .filter(item => (!item.debit_account || !item.credit_account) && (item.counterparty || item.memo));

        if (targets.length === 0) {
            alert("自動判定できる行がありません。");
            return;
        }

        // Apply Forced Credit locally BEFORE sending
        if (forcedCredit) {
            targets.forEach(t => {
                // If credit is empty, fill it. 
                // Determine if we should overwrite existing? User said "Predict as Cash".
                // Let's safe-overwrite: If currently empty, fill it.
                if (!extractedData[t.index].credit_account) {
                    extractedData[t.index].credit_account = forcedCredit;
                    t.credit_account = forcedCredit; // Update target object for API
                }
            });
            // Re-render to show "Cash" immediately? Maybe better to wait for full result.
        }

        const apiKey = localStorage.getItem('gemini_api_key');
        const sheetId = localStorage.getItem('spreadsheet_id');
        if (!apiKey || !sheetId) { alert("設定不足です"); return; }

        menuBtn.textContent = '判定中...';
        menuBtn.disabled = true;

        fetch('/api/predict', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                data: targets.map(t => ({
                    index: t.index,
                    counterparty: t.counterparty,
                    memo: t.memo,
                    debit: t.debit_account,
                    credit: t.credit_account
                })),
                gemini_api_key: apiKey,
                spreadsheet_id: sheetId,
                access_token: accessToken
            })
        })
            .then(res => res.json())
            .then(predictions => {
                if (predictions.error) throw new Error(predictions.error);
                let count = 0;
                predictions.forEach(p => {
                    const targetRow = extractedData[p.index];
                    if (targetRow) {
                        if (p.debit && !targetRow.debit_account) targetRow.debit_account = p.debit;
                        if (p.credit && !targetRow.credit_account) targetRow.credit_account = p.credit;
                        count++;
                    }
                });
                renderResults(extractedData);
                alert(`${count}件の科目を判定しました！`);
            })
            .catch(err => {
                console.error(err);
                alert(`エラー: ${err.message}`);
            })
            .finally(() => {
                menuBtn.disabled = false;
                menuBtn.textContent = '✨ 科目を自動判定 ▼';
            });
    }



    sendBtn.addEventListener('click', () => {
        const apiKey = localStorage.getItem('gemini_api_key');
        const sheetId = localStorage.getItem('spreadsheet_id');

        sendBtn.disabled = true;
        sendBtn.textContent = '送信中...';

        // Filter out rows with Amount 0 or empty
        const validData = extractedData.filter(item => {
            const val = parseInt(item.amount);
            return !isNaN(val) && val !== 0;
        });

        if (validData.length === 0) {
            alert("金額が0円以外のデータがありません。保存を中止します。");
            sendBtn.disabled = false;
            sendBtn.textContent = 'スプレッドシートに書き込む';
            return;
        }

        fetch('/api/save', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                data: validData,
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
