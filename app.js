// ==========================================
// 1. CONFIGURATION & STATE
// ==========================================
const VERSION = "v0";
const MSAL_CLIENT_ID = "11fc1bf4-b0e6-46ef-a5c0-5b372b960d0f";
const ONEDRIVE_BASE_PATH = `:/Office-HK/PrimeDrive/001_TodoBoard/database/todolist:`; // 待辦資料來源在 OneDrive 內的相對路徑 (對應專案 database/todolist)
const GRAPH_SCOPES = ["User.Read", "Files.ReadWrite"];

// Global State
let msalInstance;
let accountId = null;
let graphAccessToken = null;
let currentCardFileId = null;
let boardSettings = { Labels: {}, Assignees: [], CardOrder: [] };

// 在記憶體中保存當前各個 Card 的原始解析物件 (為了維持 Frontmatter / Description 的原始資訊)
// Key: fileId (OneDrive 圖形 API 的檔案 ID), Value: ParsedCardObject
const cardsStateMap = new Map();

// UI Elements
const ui = {
    loginBtn: document.getElementById('loginBtn'),
    logoutBtn: document.getElementById('logoutBtn'),
    refreshBtn: document.getElementById('refreshBtn'),
    userInfo: document.getElementById('userInfo'),
    userName: document.getElementById('userName'),
    welcomeState: document.getElementById('welcomeState'),
    loadingState: document.getElementById('loadingState'),
    loadingText: document.getElementById('loadingText'),
    boardContainer: document.getElementById('boardContainer'),
    modalCloseBtn: document.getElementById('modalCloseBtn'),
    cardDetailModal: document.getElementById('cardDetailModal')
};


// ==========================================
// 2. INITIALIZATION (MSAL AUTH)
// ==========================================
async function initMsal() {
    const msalConfig = {
        auth: {
            clientId: MSAL_CLIENT_ID,
            authority: "https://login.microsoftonline.com/common",
            redirectUri: window.location.origin + window.location.pathname // 自動判斷當前的 localhost 或 Github Pages 網址
        },
        cache: {
            cacheLocation: "sessionStorage",
            storeAuthStateInCookie: false,
        }
    };

    msalInstance = new msal.PublicClientApplication(msalConfig);
    await msalInstance.initialize();

    // 處理重新導向的登入結果
    const response = await msalInstance.handleRedirectPromise();
    if (response) {
        handleAuthResponse(response);
    } else {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            msalInstance.setActiveAccount(accounts[0]);
            accountId = accounts[0].homeAccountId;
            ui.userName.innerText = accounts[0].name || accounts[0].username;
            showAuthenticatedUI();
            loadBoardData();
        }
    }

    // 綁定按鈕
    ui.loginBtn.addEventListener('click', () => {
        msalInstance.loginPopup({ scopes: GRAPH_SCOPES }).then(handleAuthResponse).catch(err => {
            console.error(err);
            alert("登入失敗: " + err.message);
        });
    });

    ui.logoutBtn.addEventListener('click', () => {
        msalInstance.logoutPopup().then(() => {
            window.location.reload();
        });
    });

    ui.refreshBtn.addEventListener('click', loadBoardData);

    if (ui.modalCloseBtn) {
        ui.modalCloseBtn.addEventListener('click', () => {
            if (ui.cardDetailModal) ui.cardDetailModal.classList.add('hidden');
        });
    }
}

function handleAuthResponse(response) {
    if (response !== null) {
        msalInstance.setActiveAccount(response.account);
        accountId = response.account.homeAccountId;
        ui.userName.innerText = response.account.name || response.account.username;
        showAuthenticatedUI();
        loadBoardData();
    }
}

async function getGraphToken() {
    const request = {
        scopes: GRAPH_SCOPES,
        account: msalInstance.getActiveAccount()
    };
    try {
        const response = await msalInstance.acquireTokenSilent(request);
        return response.accessToken;
    } catch (e) {
        if (e instanceof msal.InteractionRequiredAuthError) {
            const response = await msalInstance.acquireTokenPopup(request);
            return response.accessToken;
        }
        throw e;
    }
}

function showAuthenticatedUI() {
    ui.loginBtn.classList.add('hidden');
    ui.userInfo.classList.remove('hidden');
    ui.welcomeState.classList.add('hidden');
    ui.boardContainer.classList.remove('hidden');
}


// ==========================================
// 3. GRAPH API (ONEDRIVE ACCESS)
// ==========================================
async function fetchGraph(endpoint, options = {}) {
    const token = await getGraphToken();
    const headers = new Headers(options.headers || {});
    headers.append("Authorization", `Bearer ${token}`);

    // 如果 endpoint 已經是 /items/ 開頭，就直接接在 drive 後面
    // 如果是路徑 (如 /root:/...)，則也是直接接
    const url = `https://graph.microsoft.com/v1.0/me/drive${endpoint}`;

    const fetchParams = {
        ...options,
        headers: headers
    };

    const res = await fetch(url, fetchParams);

    if (!res.ok) {
        let errMessage;
        try {
            const errDecoded = await res.json();
            errMessage = errDecoded.error.message;
        } catch (e) { errMessage = await res.text(); }
        console.error("Graph API Error:", res.status, errMessage);
        throw new Error(`Graph Error: ${res.status} - ${errMessage}`);
    }

    // Some endpoints return 204 No Content
    if (res.status === 204) return null;

    const contentType = res.headers.get("content-type") || "";
    if (contentType.includes("application/json")) {
        return await res.json();
    } else {
        // Fallback to text for markdown files, text files, or octet-streams
        return await res.text();
    }
}

// ==========================================
// 4. DATA LOADING & PARSING
// ==========================================

async function loadBoardSettings() {
    try {
        const textParams = { headers: { 'Accept': 'text/plain, text/markdown, */*' } };
        // 修正讀取路徑以配合 ONEDRIVE_BASE_PATH
        const settingsPath = ONEDRIVE_BASE_PATH.replace(/:$/, '') + '/Settings.md:';
        const rawContent = await fetchGraph(`/root${settingsPath}/content`, textParams);
        if (rawContent) {
            const parsed = parseMarkdown(rawContent, 'Settings.md');
            if (parsed.meta && parsed.meta.Labels) {
                boardSettings.Labels = parsed.meta.Labels;
            }
            if (parsed.meta && parsed.meta.Assignees) {
                let a = parsed.meta.Assignees;
                if (typeof a === 'string') {
                    try { a = JSON.parse(a); } catch (e) { a = [a]; }
                }
                boardSettings.Assignees = Array.isArray(a) ? a : [a];
            }
            if (parsed.meta && parsed.meta.CardOrder) {
                let c = parsed.meta.CardOrder;
                if (typeof c === 'string') {
                    try { c = JSON.parse(c); } catch (e) { c = []; }
                }
                boardSettings.CardOrder = Array.isArray(c) ? c : [];
            }
        }
    } catch (err) {
        console.warn('無法讀取設定檔，將使用預設設定。', err);
    }
}

async function loadBoardData() {
    ui.loadingState.classList.remove('hidden');
    ui.boardContainer.innerHTML = '';
    cardsStateMap.clear();

    try {
        ui.loadingText.innerText = "讀取全域設定...";
        await loadBoardSettings();

        ui.loadingText.innerText = "正在列出資料夾...";
        // 1. Get Lists (Folders inside TodoList)
        // 取得清單(資料夾)。先發送抓 TodoList 資料夾目錄下的 items 的指令
        const listsData = await fetchGraph(`/root${ONEDRIVE_BASE_PATH}/children`);

        let listsHtml = '';
        // 以 _ 開頭的資料夾為系統用途 (如 _Deleted)，不視為清單
        let folders = (listsData.value || []).filter(item => item.folder && !item.name.startsWith('_'));

        // 依 CardOrder 排序清單欄位
        if (boardSettings.CardOrder.length > 0) {
            const order = boardSettings.CardOrder;
            folders.sort((a, b) => {
                const idxA = order.indexOf(a.name);
                const idxB = order.indexOf(b.name);
                // 未定義順序的排到最後
                const posA = idxA === -1 ? 9999 : idxA;
                const posB = idxB === -1 ? 9999 : idxB;
                return posA - posB;
            });
        }

        // Render Lists UI Columns
        folders.forEach(folder => {
            listsHtml += `
                <div class="list-column bg-slate-800/80 backdrop-blur border border-white/5 rounded-xl w-[350px] shrink-0 flex flex-col max-h-full snap-center shadow-lg" data-list-id="${folder.id}" data-list-name="${escapeHtml(folder.name)}">
                    <div class="px-4 py-3 border-b border-white/5 flex items-center justify-between group list-drag-handle cursor-grab">
                        <h2 class="font-bold tracking-wide text-white flex items-center gap-2">
                            <i class="fa-solid fa-folder-closed text-blue-500/80 text-sm"></i> ${folder.name}
                        </h2>
                        <i class="fa-solid fa-grip-vertical text-slate-600 opacity-0 group-hover:opacity-100 shrink-0"></i>
                    </div>
                    <!-- Cards will go here -->
                    <div class="cards-container flex-1 p-3 overflow-y-auto overflow-x-hidden custom-scrollbar space-y-3 min-h-[50px]">
                    </div>
                    <!-- Add Card Interface -->
                    <div class="px-3 pb-3 pt-1 shrink-0">
                        <div id="addCardBtn_${folder.id}" class="w-full py-1.5 px-3 rounded-lg flex items-center gap-2 text-slate-400 hover:text-white hover:bg-white/10 transition-colors text-sm font-medium cursor-pointer select-none" onclick="toggleAddCardInput('${folder.id}', true)">
                            <i class="fa-solid fa-plus"></i> 新增卡片
                        </div>
                        <div id="addCardForm_${folder.id}" class="hidden flex-col gap-2 bg-slate-800/80 p-2 rounded-lg border border-slate-700">
                            <input type="text" id="addCardInput_${folder.id}" class="bg-slate-900 border border-slate-700 focus:border-blue-500 rounded px-2.5 py-1.5 text-sm text-white outline-none w-full placeholder:text-slate-500" placeholder="為這張卡片輸入標題...">
                            <div class="flex items-center gap-2 justify-end">
                                <button class="text-slate-400 hover:text-white text-xs px-3 py-1.5 rounded transition-colors" onclick="toggleAddCardInput('${folder.id}', false)">取消</button>
                                <button class="bg-blue-600 hover:bg-blue-500 text-white text-xs px-4 py-1.5 rounded font-medium transition-colors shadow-sm" onclick="submitNewCard('${folder.id}')">新增</button>
                            </div>
                        </div>
                    </div>
                </div>
            `;
        });

        // 最右側的「新增清單」欄位
        listsHtml += `
            <div class="add-list-column w-[350px] shrink-0 snap-center">
                <div id="addListBtn" class="w-full py-3 px-4 rounded-xl flex items-center gap-2 bg-white/5 hover:bg-white/10 border border-white/5 text-slate-300 hover:text-white transition-colors text-sm font-medium cursor-pointer select-none" onclick="toggleAddListInput(true)">
                    <i class="fa-solid fa-plus"></i> 新增清單
                </div>
                <div id="addListForm" class="hidden flex-col gap-2 bg-slate-800/80 backdrop-blur p-3 rounded-xl border border-slate-700">
                    <input type="text" id="addListInput" class="bg-slate-900 border border-slate-700 focus:border-blue-500 rounded px-2.5 py-1.5 text-sm text-white outline-none w-full placeholder:text-slate-500" placeholder="輸入清單名稱...">
                    <div class="flex items-center gap-2 justify-end">
                        <button class="text-slate-400 hover:text-white text-xs px-3 py-1.5 rounded transition-colors" onclick="toggleAddListInput(false)">取消</button>
                        <button id="addListSubmitBtn" class="bg-blue-600 hover:bg-blue-500 text-white text-xs px-4 py-1.5 rounded font-medium transition-colors shadow-sm" onclick="submitNewList()">新增</button>
                    </div>
                </div>
            </div>
        `;
        ui.boardContainer.innerHTML = listsHtml;

        // 2. 對於每個 Folder, 非同步獲取其底下的 .md 檔案 (Cards)
        for (const folder of folders) {
            ui.loadingText.innerText = `讀取 [${folder.name}] 中的卡片...`;
            const cardsData = await fetchGraph(`/items/${folder.id}/children`);
            const mdFiles = (cardsData.value || []).filter(item => item.file && item.name.endsWith('.md'));

            const listContainer = document.querySelector(`.list-column[data-list-id="${folder.id}"] .cards-container`);

            for (const file of mdFiles) {
                try {
                    // Fetch text content
                    // Often OneDrive markdown returns application/octet-stream or similar, fetchGraph is updated to handle it.
                    const textParams = { headers: { 'Accept': 'text/plain, text/markdown, */*' } };
                    const rawContent = await fetchGraph(`/items/${file.id}/content`, textParams);

                    // Parse the markdown string
                    const cardObj = parseMarkdown(rawContent, file.name);
                    cardsStateMap.set(file.id, cardObj);

                    // Build Card HTML
                    listContainer.insertAdjacentHTML('beforeend', buildCardHtml(file.id, file.name, cardObj));
                } catch (cardErr) {
                    console.error(`讀取或解析卡片失敗: ${file.name}`, cardErr);
                    // Add an error card to the UI so the user knows it exists but failed to load
                    listContainer.insertAdjacentHTML('beforeend', `
                        <div class="card bg-red-900/40 border border-red-600/50 p-3 rounded-xl shadow cursor-default">
                            <h3 class="font-bold text-red-200 text-[15px] truncate">⚠️ 讀取失敗: ${file.name}</h3>
                            <p class="text-xs text-red-300 mt-1">請查看主控台了解詳情</p>
                        </div>
                    `);
                }
            }
        }

        // 3. 全面初始化 SortableJS (套用所有 Drag Drop 規則)
        initSortable();

    } catch (err) {
        console.error("裝載失敗", err);
        // 若發生 404 表達查無資料夾
        if (err.message.includes('404')) {
            alert("找不到 TodoList 資料夾！請確認您的 OneDrive 中是否存在：\n" + ONEDRIVE_BASE_PATH.replace(/:/g, ''));
        } else {
            alert("讀取目錄失敗，請查看主控台日誌。");
        }
    } finally {
        ui.loadingState.classList.add('hidden');
    }
}

// 解析 Frontmatter (YAML metadata)
function parseFrontmatter(fmString) {
    const meta = {};
    if (!fmString) return meta;
    const lines = fmString.split(/\r?\n/);
    let currentObjKey = null;

    lines.forEach(line => {
        // match basic key-value
        const match = line.match(/^([A-Za-z0-9_]+):\s*(.*)$/);
        if (match) {
            let key = match[1].trim();
            let val = match[2].trim();
            currentObjKey = key;

            if (val) {
                if ((val.startsWith('[') && val.endsWith(']')) || (val.startsWith('{') && val.endsWith('}'))) {
                    try { val = JSON.parse(val); } catch (e) { }
                } else if (val.startsWith('"') && val.endsWith('"')) {
                    val = val.substring(1, val.length - 1);
                }
            }
            meta[key] = val;
        }
        // handle indented children
        else if (currentObjKey && line.match(/^\s+([^:]+):\s*(.*)$/)) {
            if (typeof meta[currentObjKey] !== 'object' || meta[currentObjKey] === null) {
                meta[currentObjKey] = {};
            }

            const childMatch = line.match(/^\s+([^:]+):\s*(.*)$/);
            let childKey = childMatch[1].trim();
            let childVal = childMatch[2].trim();

            if (childVal.startsWith('"') && childVal.endsWith('"')) {
                childVal = childVal.substring(1, childVal.length - 1);
            }
            if (childKey.startsWith('"') && childKey.endsWith('"')) {
                childKey = childKey.substring(1, childKey.length - 1);
            }

            meta[currentObjKey][childKey] = childVal;
        }
    });
    return meta;
}

// 將 Markdown 字串解析為 JSON 物件
function parseMarkdown(rawContent, filename) {
    // Frontmatter Extraction
    const fmMatch = rawContent.match(/^(-{3,}\n[\s\S]*?\n-{3,}\n)/);
    let frontmatter = "";
    let content = rawContent;
    if (fmMatch) {
        frontmatter = fmMatch[0];
        content = rawContent.substring(fmMatch[0].length);
    }

    // 將內文分出一行一行
    const lines = content.split(/\r?\n/);
    let descriptionLines = [];
    let groups = [];
    let currentGroup = null;
    let inGroup = false;

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (line.startsWith('## ')) {
            inGroup = true;
            const rawTitle = line.substring(3).trim();
            const meta = extractMetadata(rawTitle);
            currentGroup = {
                title: meta.cleanText,
                originalTitle: rawTitle,
                startDate: meta.startDate,
                dueDate: meta.dueDate,
                assignees: meta.assignees,
                items: []
            };
            groups.push(currentGroup);
        } else if (inGroup) {
            // 解析是否為 `- [ ] foo` 或是 `- [x] foo`
            const checkboxMatch = line.match(/^\s*-\s*\[(x|X| )\]\s+(.*)/);
            if (checkboxMatch) {
                const rawText = checkboxMatch[2];
                const meta = extractMetadata(rawText);
                currentGroup.items.push({
                    checked: checkboxMatch[1].toLowerCase() === 'x',
                    text: meta.cleanText,
                    originalText: rawText,
                    startDate: meta.startDate,
                    dueDate: meta.dueDate,
                    assignees: meta.assignees
                });
            }
        } else {
            descriptionLines.push(line);
        }
    }

    // 簡單清理頂部尾部的空行
    while (descriptionLines.length > 0 && descriptionLines[0].trim() === '') descriptionLines.shift();
    if (descriptionLines.length > 0 && descriptionLines[0].trim().toLowerCase() === '# description') {
        descriptionLines.shift();
        while (descriptionLines.length > 0 && descriptionLines[0].trim() === '') descriptionLines.shift();
    }
    while (descriptionLines.length > 0 && descriptionLines[descriptionLines.length - 1].trim() === '') descriptionLines.pop();

    return {
        // 若有抓到 Title 可以從 Frontmatter 抽，不然預設用檔名
        title: filename.replace('.md', ''),
        frontmatter,
        meta: parseFrontmatter(frontmatter),
        description: descriptionLines.join('\n'),
        groups
    };
}

// 輔助函式：從文字中萃取日期與負責人
function extractMetadata(text) {
    let cleanText = text;
    let startDate = null;
    let dueDate = null;
    let assignees = [];

    // 1. 匹配 [@負責人]
    const assigneeRegex = /@([\w\u4e00-\u9fa5]+)/g;
    let match;
    while ((match = assigneeRegex.exec(cleanText)) !== null) {
        assignees.push(match[1]);
    }
    cleanText = cleanText.replace(assigneeRegex, '');

    // 2. 匹配 [YYYY-MM-DD ~ YYYY-MM-DD]
    // 支援 [~ 2024-05-10], [2024-05-01 ~], [2024-05-01~2024-05-10] 等
    const rangeRegex = /\[\s*(\d{4}[-/]\d{1,2}[-/]\d{1,2})?\s*[~-]\s*(\d{4}[-/]\d{1,2}[-/]\d{1,2})?\s*\]/g;
    const rangeMatch = Array.from(cleanText.matchAll(rangeRegex));
    if (rangeMatch && rangeMatch.length > 0) {
        // 取最後一個匹配到的日期區間
        const lastMatch = rangeMatch[rangeMatch.length - 1];
        startDate = lastMatch[1] || null;
        dueDate = lastMatch[2] || null;
        cleanText = cleanText.replace(lastMatch[0], '');
    } else {
        // 3. 備用匹配： Start: YYYY-MM-DD, Due: YYYY-MM-DD
        const startMatch = cleanText.match(/(?:Start|開始|S):\s*(\d{4}[-/]\d{1,2}[-/]\d{1,2})/i);
        if (startMatch) {
            startDate = startMatch[1];
            cleanText = cleanText.replace(startMatch[0], '');
        }
        const dueMatch = cleanText.match(/(?:Due|到期|D):\s*(\d{4}[-/]\d{1,2}[-/]\d{1,2})/i);
        if (dueMatch) {
            dueDate = dueMatch[1];
            cleanText = cleanText.replace(dueMatch[0], '');
        }
    }

    return {
        cleanText: cleanText.trim(),
        startDate,
        dueDate,
        assignees
    };
}

// 日期顯示美化
function formatDateDisplay(dateStr) {
    if (!dateStr) return '';
    try {
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return dateStr;
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        const currentYear = new Date().getFullYear();
        if (year === currentYear) {
            return `${month}/${day}`;
        } else {
            return `${year}/${month}/${day}`;
        }
    } catch (e) { return dateStr; }
}

// 建立負責人徽章 UI
function buildAssigneeBadges(assignees) {
    if (!assignees || assignees.length === 0) return '';
    let html = '<div class="flex items-center gap-1 ml-2">';
    assignees.forEach(name => {
        // 抓名字第一個字當作頭像
        const initial = name.substring(0, 1).toUpperCase();
        html += `
            <div class="h-5 flex items-center bg-blue-500/20 border border-blue-500/30 rounded-full px-1.5 text-[10px] text-blue-300 whitespace-nowrap" title="${escapeHtml(name)}">
                <div class="w-3.5 h-3.5 bg-blue-500 rounded-full flex items-center justify-center text-white mr-1 opacity-80 font-bold">${escapeHtml(initial)}</div>
                ${escapeHtml(name)}
            </div>
        `;
    });
    html += '</div>';
    return html;
}

// 建立日期徽章 UI
function buildDateBadges(startDate, dueDate) {
    if (!startDate && !dueDate) return '';
    let html = '<div class="flex items-center gap-1 text-[11px] font-medium ml-2 shrink-0">';

    if (startDate && dueDate) {
        html += `<span class="bg-slate-700/80 text-slate-300 px-1.5 py-0.5 rounded border border-slate-600/50" title="開始與到期日"><i class="fa-regular fa-calendar shrink-0 mr-1 opacity-70"></i>${formatDateDisplay(startDate)} - ${formatDateDisplay(dueDate)}</span>`;
    } else if (startDate) {
        html += `<span class="bg-slate-700/80 text-emerald-400 px-1.5 py-0.5 rounded border border-emerald-500/30" title="開始日"><i class="fa-solid fa-play shrink-0 mr-1 opacity-70"></i>${formatDateDisplay(startDate)}</span>`;
    } else if (dueDate) {
        html += `<span class="bg-slate-700/80 text-orange-400 px-1.5 py-0.5 rounded border border-orange-500/30" title="到期日"><i class="fa-solid fa-flag-checkered shrink-0 mr-1 opacity-70"></i>${formatDateDisplay(dueDate)}</span>`;
    }

    html += '</div>';
    return html;
}


// 將資料模型畫為 HTML 元素
function buildCardHtml(fileId, filename, cardObj) {
    let cardLabelsHtml = '';
    const labels = cardObj.meta.Labels;
    if (Array.isArray(labels) && labels.length > 0) {
        cardLabelsHtml = '<div class="flex flex-wrap gap-1 mb-2">';
        labels.forEach(l => {
            const hexColor = boardSettings.Labels[l] || '#6366f1';
            const style = `background-color: ${hexColor}33; color: ${hexColor}; border-color: ${hexColor}4D;`;
            cardLabelsHtml += `<span class="px-1.5 py-0.5 rounded text-[10px] font-bold border" style="${style}">${escapeHtml(l)}</span>`;
        });
        cardLabelsHtml += '</div>';
    }

    let totalTodos = 0;
    let completedTodos = 0;
    cardObj.groups.forEach(group => {
        totalTodos += group.items.length;
        group.items.forEach(item => {
            if (item.checked) completedTodos++;
        });
    });

    const cardDueDate = cardObj.meta.DueDate || null;
    let extraInfoHtml = '';
    if (cardDueDate || totalTodos > 0) {
        extraInfoHtml = '<div class="flex items-center gap-3 mt-3 pt-2 border-t border-slate-600/50 text-xs text-slate-400">';
        
        if (cardDueDate) {
            const formattedDate = formatDateDisplay(cardDueDate);
            // 判斷是否過期
            const isOverdue = new Date(cardDueDate) < new Date() && new Date(cardDueDate).toDateString() !== new Date().toDateString();
            const dateColorClass = isOverdue ? 'text-red-400 font-bold' : 'text-slate-400';
            const dateIconColorClass = isOverdue ? 'text-red-500' : 'text-slate-500';
            
            extraInfoHtml += `
                <div class="flex items-center gap-1.5 ${dateColorClass}" title="結束日">
                    <i class="fa-regular fa-clock ${dateIconColorClass}"></i>
                    <span>${formattedDate}</span>
                </div>
            `;
        }

        if (totalTodos > 0) {
            const progressColor = completedTodos === totalTodos ? 'text-emerald-400 font-bold' : 'text-slate-400';
            const progressIconColor = completedTodos === totalTodos ? 'text-emerald-500' : 'text-slate-500';
            extraInfoHtml += `
                <div class="flex items-center gap-1.5 ${progressColor}" title="待辦事項進度">
                    <i class="fa-solid fa-check-square ${progressIconColor}"></i>
                    <span>${completedTodos}/${totalTodos}</span>
                </div>
            `;
        }
        extraInfoHtml += '</div>';
    }

    return `
        <div class="card bg-slate-700/40 border border-slate-600/50 p-3 rounded-xl shadow cursor-default flex flex-col" data-file-id="${fileId}" data-filename="${filename}">
            <div class="flex items-center justify-between mb-2">
                <h3 class="font-bold text-white text-[15px] truncate cursor-grab card-drag-handle flex-1" title="${filename}">
                    ${escapeHtml(cardObj.meta.Title || cardObj.title)}
                </h3>
                <button class="text-slate-400 hover:text-white transition-colors bg-white/5 hover:bg-white/20 w-7 h-7 rounded flex items-center justify-center shrink-0 ml-2" onclick="openCardDetail('${fileId}')" title="詳細檢視">
                    <i class="fa-solid fa-up-right-and-down-left-from-center text-xs"></i>
                </button>
            </div>
            ${cardLabelsHtml}
            <div class="flex-1"></div>
            ${extraInfoHtml}
        </div>
    `;
}

function escapeHtml(str) {
    if (!str) return '';
    return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

function renderMarkdownBasic(text) {
    if (!text || text.trim() === '') return '<span class="italic text-slate-500">沒有提供詳細說明...</span>';
    let html = escapeHtml(text);
    // Render links
    html = html.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2" target="_blank" class="text-blue-400 hover:text-blue-300 hover:underline transition-colors"><i class="fa-solid fa-link mr-1 text-[11px] opacity-70"></i>$1</a>');
    // Render Headers
    html = html.replace(/^#\s+(.*)$/gm, '<h1 class="text-lg font-bold text-white mb-3 mt-5 border-b border-slate-700/50 pb-2"><i class="fa-solid fa-caret-right text-blue-500 mr-2 text-sm"></i>$1</h1>');
    html = html.replace(/^##\s+(.*)$/gm, '<h2 class="text-base font-bold text-slate-200 mb-2 mt-4"><i class="fa-solid fa-caret-right text-slate-500 mr-2 text-sm"></i>$1</h2>');
    // Render basic bullet points
    html = html.replace(/^\s*-\s+(.*)$/gm, '<div class="flex items-start gap-2 mb-1.5"><div class="w-1.5 h-1.5 rounded-full bg-slate-500 mt-2 shrink-0"></div><span class="flex-1 min-w-0">$1</span></div>');
    return html;
}

// 展開詳細檢視 Modal
function openCardDetail(fileId) {
    currentCardFileId = fileId;
    const cardObj = cardsStateMap.get(fileId);
    if (!cardObj) return;

    document.getElementById('modalTitle').innerText = cardObj.meta.Title || cardObj.title;

    const descElem = document.getElementById('modalDescription');
    const descEditor = document.getElementById('modalDescriptionEditor');
    const descTextarea = document.getElementById('modalDescriptionTextarea');
    const editBtn = document.getElementById('modalEditDescriptionBtn');
    
    descElem.innerHTML = renderMarkdownBasic(cardObj.description);
    descTextarea.value = cardObj.description || '';
    
    // Reset edit state
    descElem.classList.remove('hidden');
    editBtn.classList.remove('hidden');
    descEditor.classList.add('hidden');
    descEditor.classList.remove('flex');

    const labelsElem = document.getElementById('modalLabels');
    labelsElem.innerHTML = '';
    const labels = cardObj.meta.Labels || [];
    if (Array.isArray(labels) && labels.length > 0) {
        labels.forEach((l, idx) => {
            const hexColor = boardSettings.Labels[l] || '#6366f1';
            const style = `background-color: ${hexColor}33; color: ${hexColor}; border-color: ${hexColor}4D;`;
            labelsElem.insertAdjacentHTML('beforeend', `<span class="px-2.5 py-1 rounded text-xs font-bold tracking-wide shadow-sm flex items-center gap-1 border" style="${style}">${escapeHtml(l)} <i class="fa-solid fa-xmark cursor-pointer hover:opacity-70 ml-1" onclick="removeLabel(${idx})"></i></span>`);
        });
    } else {
        labelsElem.innerHTML = '<span class="text-slate-500 text-xs italic" id="emptyLabelText">無標籤</span>';
    }

    const assigneesElem = document.getElementById('modalAssignees');
    assigneesElem.innerHTML = '';
    const assignees = cardObj.meta.Assignees || [];
    if (Array.isArray(assignees) && assignees.length > 0) {
        assignees.forEach((name, idx) => {
            const initial = name.substring(0, 1).toUpperCase();
            assigneesElem.insertAdjacentHTML('beforeend', `
                <div class="flex items-center gap-2 bg-slate-700/30 p-1.5 rounded-lg border border-slate-700/50">
                    <div class="w-6 h-6 bg-blue-500 shadow-sm shadow-blue-500/20 rounded-full flex items-center justify-center text-white font-bold text-xs shrink-0">${escapeHtml(initial)}</div>
                    <span class="text-sm text-slate-300 truncate font-medium flex-1">${escapeHtml(name)}</span>
                    <i class="fa-solid fa-xmark text-slate-500 hover:text-red-400 cursor-pointer transition-colors text-xs shrink-0" onclick="removeAssignee(${idx})" title="移除負責人"></i>
                </div>
            `);
        });
    } else {
        assigneesElem.innerHTML = '<span class="text-slate-500 text-xs italic">未指派</span>';
    }

    const startDateInput = document.getElementById('modalStartDateInput');
    const dueDateInput = document.getElementById('modalDueDateInput');

    if (startDateInput) {
        if (startDateInput._flatpickr) {
            startDateInput._flatpickr.setDate(cardObj.meta.StartDate || '', false);
            if (dueDateInput && dueDateInput._flatpickr) {
                dueDateInput._flatpickr.set('minDate', cardObj.meta.StartDate || null);
            }
        } else {
            startDateInput.value = cardObj.meta.StartDate || '';
        }
    }

    if (dueDateInput) {
        if (dueDateInput._flatpickr) {
            dueDateInput._flatpickr.setDate(cardObj.meta.DueDate || '', false);
            if (startDateInput && startDateInput._flatpickr) {
                startDateInput._flatpickr.set('maxDate', cardObj.meta.DueDate || null);
            }
        } else {
            dueDateInput.value = cardObj.meta.DueDate || '';
        }
    }

    // Hide popovers if open
    const popover = document.getElementById('labelSelectorPopover');
    if (popover) {
        popover.classList.add('hidden');
        renderLabelPopoverList();
    }
    const assigneePopover = document.getElementById('assigneeSelectorPopover');
    if (assigneePopover) {
        assigneePopover.classList.add('hidden');
        renderAssigneePopoverList();
    }

    renderCardAttachments(fileId);
    renderModalChecklists(cardObj);

    if (ui.cardDetailModal) ui.cardDetailModal.classList.remove('hidden');
}

// ==========================================
// MODAL CHECKLISTS LOGIC
// ==========================================

// --- Utility: Serialize from in-memory state and save to OneDrive ---
async function commitCardChangesFromMemory(fileId) {
    const cardObj = cardsStateMap.get(fileId);
    if (!cardObj) return;

    let mdString = "";
    if (cardObj.frontmatter) mdString += cardObj.frontmatter;
    if (cardObj.description) mdString += cardObj.description + "\n\n";

    for (const group of cardObj.groups) {
        mdString += `## ${group.originalTitle}\n`;
        for (const item of group.items) {
            mdString += `- [${item.checked ? 'x' : ' '}] ${item.originalText}\n`;
        }
        mdString += "\n";
    }

    try {
        const fetchOptions = {
            method: 'PUT',
            headers: { 'Content-Type': 'text/plain' },
            body: mdString.trim() + "\n"
        };
        await fetchGraph(`/items/${fileId}/content`, fetchOptions);
    } catch (err) {
        console.error("儲存卡片內容失敗 (from memory): ", err);
        alert("存檔失敗，請確認網路與連線狀態！");
    }
}

// --- Utility: Rebuild a single board card's DOM from in-memory state ---
function rebuildBoardCard(fileId) {
    const cardObj = cardsStateMap.get(fileId);
    if (!cardObj) return;
    const cardElem = document.querySelector(`.card[data-file-id="${fileId}"]`);
    if (!cardElem) return;
    const filename = cardElem.getAttribute('data-filename');
    const newHtml = buildCardHtml(fileId, filename, cardObj);
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = newHtml;
    const newCardElem = tempDiv.firstElementChild;
    cardElem.replaceWith(newCardElem);
}

// --- Utility: Update/insert/remove date range string in originalText or originalTitle ---
function updateDateRangeInText(text, newStart, newEnd) {
    // Remove existing date range [... ~ ...]
    let cleaned = text.replace(/\s*\[\s*(\d{4}[-/]\d{1,2}[-/]\d{1,2})?\s*[~-]\s*(\d{4}[-/]\d{1,2}[-/]\d{1,2})?\s*\]/g, '').trim();
    if (newStart || newEnd) {
        cleaned += ` [${newStart || ''} ~ ${newEnd || ''}]`;
    }
    return cleaned;
}

// --- Utility: Update/insert/remove @assignee markers in originalText or originalTitle ---
function updateAssigneesInText(text, newAssignees) {
    // Remove all existing @name markers
    let cleaned = text.replace(/\s*@[\w\u4e00-\u9fa5]+/g, '').trim();
    // Append new assignees
    if (newAssignees && newAssignees.length > 0) {
        cleaned += ' ' + newAssignees.map(a => '@' + a).join(' ');
    }
    return cleaned;
}

function renderModalChecklists(cardObj) {
    const container = document.getElementById('modalChecklists');
    if (!container) return;

    if (!cardObj.groups || cardObj.groups.length === 0) {
        container.innerHTML = `
            <div class="text-center py-6 border border-dashed border-slate-700/50 rounded-lg">
                <p class="text-sm font-semibold text-slate-500 mb-2">此卡片尚無待辦清單</p>
                <div class="flex items-center justify-center gap-2 max-w-xs mx-auto mt-3 border-t border-slate-700/50 pt-4">
                    <input type="text" id="newGroupInputFirst" class="bg-slate-900 border border-slate-700 focus:border-blue-500 rounded px-2.5 py-1.5 text-xs text-white flex-1 outline-none transition-colors placeholder:text-slate-500" placeholder="輸入新清單名稱..." onkeyup="if(event.key==='Enter') window.addChecklistGroupFromModal('${currentCardFileId}', this.value)">
                    <button class="text-white hover:text-blue-200 bg-blue-600 hover:bg-blue-500 transition-colors text-xs font-medium px-4 py-1.5 rounded shadow-sm shrink-0" onclick="window.addChecklistGroupFromModal('${currentCardFileId}', document.getElementById('newGroupInputFirst').value)">
                        <i class="fa-solid fa-plus mr-1"></i>建立
                    </button>
                </div>
            </div>`;
        return;
    }

    let html = '';
    cardObj.groups.forEach((g, groupIndex) => {
        const total = g.items.length;
        const checkedCount = g.items.filter(i => i.checked).length;
        const progressPct = total > 0 ? Math.round((checkedCount / total) * 100) : 0;

        // Group date display
        let groupDateBadge = '';
        if (g.startDate || g.dueDate) {
            const ds = g.startDate ? formatDateDisplay(g.startDate) : '';
            const de = g.dueDate ? formatDateDisplay(g.dueDate) : '';
            groupDateBadge = `<span class="ml-2 text-[10px] px-1.5 py-0.5 rounded bg-slate-700 text-slate-400 border border-slate-600 shrink-0"><i class="fa-regular fa-calendar mr-1"></i>${ds}${ds && de ? ' - ' : ''}${de}</span>`;
        }

        let itemsHtml = '';
        g.items.forEach((item, itemIndex) => {
            const isChecked = item.checked ? 'checked' : '';
            const textClass = item.checked ? 'line-through text-slate-500' : 'text-slate-300';

            let badges = '';
            if (item.assignees && item.assignees.length > 0) {
                badges += buildAssigneeBadges(item.assignees);
            }

            // Item date badge
            let dateBadge = '';
            if (item.startDate || item.dueDate) {
                const ds = item.startDate ? formatDateDisplay(item.startDate) : '';
                const de = item.dueDate ? formatDateDisplay(item.dueDate) : '';
                dateBadge = `<span class="ml-1 text-[10px] px-1.5 py-0.5 rounded bg-slate-700 text-slate-400 border border-slate-600 shrink-0"><i class="fa-regular fa-clock mr-1"></i>${ds}${ds && de ? ' - ' : ''}${de}</span>`;
            }

            itemsHtml += `
                <div class="modal-todo-item flex items-start gap-3 mb-2.5 group/item" data-group-index="${groupIndex}" data-item-index="${itemIndex}">
                    <i class="fa-solid fa-grip-vertical text-slate-600 opacity-0 group-hover/item:opacity-100 cursor-grab mt-1 shrink-0 modal-item-drag-handle"></i>
                    <input type="checkbox" class="mt-1 w-4 h-4 rounded border-slate-600 bg-slate-800 text-blue-500 focus:ring-blue-500 focus:ring-offset-slate-800 cursor-pointer transition-colors shrink-0" ${isChecked} onchange="toggleTodoFromModal(this, '${currentCardFileId}', ${groupIndex}, ${itemIndex})">
                    <span class="flex-1 text-[14px] ${textClass} transition-all select-none">${escapeHtml(item.text)}</span>
                    ${badges}
                    ${dateBadge}
                    <button class="text-slate-600 hover:text-blue-400 transition-colors opacity-0 group-hover/item:opacity-100 shrink-0" onclick="editItemAssignees('${currentCardFileId}', ${groupIndex}, ${itemIndex})" title="編輯負責人"><i class="fa-solid fa-user-pen text-xs"></i></button>
                    <button class="text-slate-600 hover:text-blue-400 transition-colors opacity-0 group-hover/item:opacity-100 shrink-0" onclick="editItemDates('${currentCardFileId}', ${groupIndex}, ${itemIndex})" title="編輯日期"><i class="fa-regular fa-calendar-plus text-xs"></i></button>
                    <button class="text-slate-600 hover:text-red-400 transition-colors opacity-0 group-hover/item:opacity-100 shrink-0" onclick="deleteTodoItemFromModal('${currentCardFileId}', ${groupIndex}, ${itemIndex})" title="刪除"><i class="fa-solid fa-xmark text-xs"></i></button>
                </div>
            `;
        });

        const allChecked = total > 0 && checkedCount === total;
        const groupCheckIconClass = allChecked ? 'fa-solid text-blue-500' : 'fa-regular text-slate-500';

        html += `
            <div class="modal-checklist-group bg-slate-800/50 rounded-lg border border-slate-700/50 shadow-sm relative overflow-hidden" data-group-index="${groupIndex}">
                <div class="flex items-center justify-between px-4 py-3 bg-slate-900/40 border-b border-slate-700/50 modal-group-drag-handle cursor-grab group/gheader">
                    <div class="flex items-center min-w-0 flex-1">
                        <i class="fa-solid fa-grip-vertical text-slate-600 opacity-0 group-hover/gheader:opacity-100 mr-2 shrink-0"></i>
                        <div class="mr-2 flex items-center hover:bg-white/10 p-0.5 rounded cursor-pointer" onclick="event.stopPropagation(); toggleGroupFromModal('${currentCardFileId}', ${groupIndex})" title="全選 / 取消全選">
                            <i class="${groupCheckIconClass} fa-square-check hover:text-blue-400 transition-colors"></i>
                        </div>
                        <h4 class="font-bold text-slate-200 text-[15px] truncate">${escapeHtml(g.title)}</h4>
                        ${groupDateBadge}
                    </div>
                    <div class="flex items-center gap-2 ml-2 shrink-0">
                        <button class="text-slate-500 hover:text-blue-400 transition-colors" onclick="editGroupAssignees('${currentCardFileId}', ${groupIndex})" title="編輯群組負責人"><i class="fa-solid fa-user-pen text-xs"></i></button>
                        <button class="text-slate-500 hover:text-blue-400 transition-colors" onclick="editGroupDates('${currentCardFileId}', ${groupIndex})" title="編輯群組日期"><i class="fa-regular fa-calendar-plus text-xs"></i></button>
                        <span class="text-xs font-semibold text-slate-400">${progressPct}%</span>
                    </div>
                </div>
                <!-- Progress Bar -->
                <div class="w-full bg-slate-700 h-1 overflow-hidden">
                    <div class="bg-blue-500 h-1 transition-all duration-300" style="width: ${progressPct}%"></div>
                </div>
                
                <!-- Items list -->
                <div class="modal-items-container p-3 flex flex-col min-h-[30px]">
                    ${itemsHtml}
                </div>
                
                <!-- Add new item -->
                <div class="px-3 pb-3 flex gap-2">
                    <input type="text" id="newTodoInput_${groupIndex}" class="bg-slate-900 border border-slate-700 focus:border-blue-500 rounded px-2.5 py-1.5 text-xs text-white flex-1 outline-none transition-colors placeholder:text-slate-500" placeholder="新增待辦項目..." onkeyup="if(event.key==='Enter') addTodoItemFromModal('${currentCardFileId}', ${groupIndex}, this.value)">
                    <button class="bg-slate-700 hover:bg-blue-500/30 text-white hover:text-blue-300 border border-slate-600 hover:border-blue-500/30 px-3 py-1.5 rounded text-xs font-medium transition-colors shadow-sm" onclick="addTodoItemFromModal('${currentCardFileId}', ${groupIndex}, document.getElementById('newTodoInput_${groupIndex}').value)">新增</button>
                </div>
            </div>
        `;
    });

    html += `
        <div class="mt-4 flex items-center justify-center gap-2 max-w-xs mx-auto">
             <input type="text" id="newGroupInputBottom" class="bg-slate-900 border border-slate-700 focus:border-blue-500 rounded px-2.5 py-1.5 text-xs text-white flex-1 outline-none transition-colors placeholder:text-slate-500" placeholder="輸入新清單名稱..." onkeyup="if(event.key==='Enter') window.addChecklistGroupFromModal('${currentCardFileId}', this.value)">
             <button class="text-slate-400 hover:text-white text-xs font-medium px-4 py-1.5 bg-slate-800 hover:bg-slate-700/80 rounded border border-slate-700 transition-colors shrink-0" onclick="window.addChecklistGroupFromModal('${currentCardFileId}', document.getElementById('newGroupInputBottom').value)">
                  <i class="fa-solid fa-plus mr-1"></i>新增
             </button>
        </div>
    `;

    container.innerHTML = html;
    initModalSortable();
}

// --- Initialize SortableJS for groups and items inside the modal ---
function initModalSortable() {
    const container = document.getElementById('modalChecklists');
    if (!container) return;

    // Group-level drag (drag entire checklist groups to reorder)
    if (container._sortableGroups) container._sortableGroups.destroy();
    container._sortableGroups = new Sortable(container, {
        animation: 150,
        handle: '.modal-group-drag-handle',
        draggable: '.modal-checklist-group',
        ghostClass: 'ghost-group',
        onEnd: function () {
            reorderGroupsFromModal();
        }
    });

    // Item-level drag within each group (and across groups)
    container.querySelectorAll('.modal-items-container').forEach(itemsContainer => {
        if (itemsContainer._sortableItems) itemsContainer._sortableItems.destroy();
        itemsContainer._sortableItems = new Sortable(itemsContainer, {
            group: 'modal-items',
            animation: 150,
            handle: '.modal-item-drag-handle',
            draggable: '.modal-todo-item',
            ghostClass: 'ghost-item',
            onEnd: function () {
                reorderItemsFromModal();
            }
        });
    });
}

// --- After group drag: rebuild in-memory groups array from DOM order ---
function reorderGroupsFromModal() {
    const fileId = currentCardFileId;
    if (!fileId) return;
    const cardObj = cardsStateMap.get(fileId);
    if (!cardObj) return;

    const container = document.getElementById('modalChecklists');
    const groupElems = container.querySelectorAll('.modal-checklist-group');
    const newGroups = [];
    groupElems.forEach(gElem => {
        const oldIndex = parseInt(gElem.getAttribute('data-group-index'));
        if (!isNaN(oldIndex) && cardObj.groups[oldIndex]) {
            newGroups.push(cardObj.groups[oldIndex]);
        }
    });

    cardObj.groups = newGroups;
    commitCardChangesFromMemory(fileId);
    rebuildBoardCard(fileId);
    renderModalChecklists(cardObj);
}

// --- After item drag: rebuild items arrays from DOM order ---
function reorderItemsFromModal() {
    const fileId = currentCardFileId;
    if (!fileId) return;
    const cardObj = cardsStateMap.get(fileId);
    if (!cardObj) return;

    const container = document.getElementById('modalChecklists');
    const groupElems = container.querySelectorAll('.modal-checklist-group');

    groupElems.forEach((gElem, newGroupIdx) => {
        const origGroupIdx = parseInt(gElem.getAttribute('data-group-index'));
        const itemElems = gElem.querySelectorAll('.modal-todo-item');
        const newItems = [];
        itemElems.forEach(iElem => {
            const origGI = parseInt(iElem.getAttribute('data-group-index'));
            const origII = parseInt(iElem.getAttribute('data-item-index'));
            if (!isNaN(origGI) && !isNaN(origII) && cardObj.groups[origGI] && cardObj.groups[origGI].items[origII]) {
                newItems.push(cardObj.groups[origGI].items[origII]);
            }
        });
        if (!isNaN(origGroupIdx) && cardObj.groups[origGroupIdx]) {
            cardObj.groups[origGroupIdx].items = newItems;
        }
    });

    commitCardChangesFromMemory(fileId);
    rebuildBoardCard(fileId);
    renderModalChecklists(cardObj);
}

// --- Toggle checkbox from modal (uses index-based lookup) ---
window.toggleTodoFromModal = async function (checkboxElem, fileId, groupIndex, itemIndex) {
    const isChecked = checkboxElem.checked;
    const cardObj = cardsStateMap.get(fileId);
    if (!cardObj) return;

    const group = cardObj.groups[groupIndex];
    if (group && group.items[itemIndex]) {
        group.items[itemIndex].checked = isChecked;
    }

    await commitCardChangesFromMemory(fileId);
    rebuildBoardCard(fileId);
    renderModalChecklists(cardObj);
}

// --- Toggle all items in a group from modal (select all / deselect all) ---
window.toggleGroupFromModal = async function (fileId, groupIndex) {
    const cardObj = cardsStateMap.get(fileId);
    if (!cardObj) return;
    const group = cardObj.groups[groupIndex];
    if (!group || group.items.length === 0) return;

    if (!window.confirm('確定要將群組內全數待辦事項變更完成狀態嗎？')) return;

    // 決定目標狀態：若有任何未完成項則全部打勾，否則全部取消
    const hasUnchecked = group.items.some(item => !item.checked);
    const targetChecked = hasUnchecked;

    group.items.forEach(item => {
        item.checked = targetChecked;
    });

    await commitCardChangesFromMemory(fileId);
    rebuildBoardCard(fileId);
    renderModalChecklists(cardObj);
};

// --- Add new todo item (fixed: uses in-memory groups) ---
window.addTodoItemFromModal = async function (fileId, groupIndex, newItemText) {
    if (!newItemText || !newItemText.trim() || !fileId) return;

    const inputElem = document.getElementById(`newTodoInput_${groupIndex}`);
    if (inputElem) inputElem.disabled = true;

    const cardObj = cardsStateMap.get(fileId);
    if (!cardObj) return;

    const trimmed = newItemText.trim();
    const group = cardObj.groups[groupIndex];
    if (!group) return;

    group.items.push({
        checked: false,
        text: trimmed,
        originalText: trimmed,
        startDate: null,
        dueDate: null,
        assignees: []
    });

    try {
        await commitCardChangesFromMemory(fileId);
        rebuildBoardCard(fileId);
        renderModalChecklists(cardObj);
    } catch (err) {
        console.error("Failed to add todo", err);
        alert("新增失敗");
    } finally {
        if (inputElem) {
            inputElem.disabled = false;
            inputElem.value = '';
            inputElem.focus();
        }
    }
};

// --- Add new checklist group (fixed: uses inline input) ---
window.addChecklistGroupFromModal = async function (fileId, groupTitle) {
    if (!groupTitle || !groupTitle.trim() || !fileId) return;

    const cardObj = cardsStateMap.get(fileId);
    if (!cardObj) return;

    const trimmedTitle = groupTitle.trim();
    if (!cardObj.groups) cardObj.groups = [];

    cardObj.groups.push({
        title: trimmedTitle,
        originalTitle: trimmedTitle,
        startDate: null,
        dueDate: null,
        assignees: [],
        items: []
    });

    try {
        await commitCardChangesFromMemory(fileId);
        rebuildBoardCard(fileId);
        renderModalChecklists(cardObj);
    } catch (err) {
        console.error("Failed to add group", err);
        alert("新增清單失敗");
    }
};

// --- Delete a todo item with confirmation ---
window.deleteTodoItemFromModal = async function (fileId, groupIndex, itemIndex) {
    const cardObj = cardsStateMap.get(fileId);
    if (!cardObj) return;
    const group = cardObj.groups[groupIndex];
    if (!group || !group.items[itemIndex]) return;

    const itemText = group.items[itemIndex].text;
    if (!confirm(`確定要刪除待辦事項「${itemText}」？`)) return;

    group.items.splice(itemIndex, 1);

    try {
        await commitCardChangesFromMemory(fileId);
        rebuildBoardCard(fileId);
        renderModalChecklists(cardObj);
    } catch (err) {
        console.error("Failed to delete todo", err);
        alert("刪除失敗");
    }
};

// --- Edit dates for a checklist group ---
window.editGroupDates = function (fileId, groupIndex) {
    const cardObj = cardsStateMap.get(fileId);
    if (!cardObj) return;
    const group = cardObj.groups[groupIndex];
    if (!group) return;

    showDateEditorPopup(group.startDate || '', group.dueDate || '', async (newStart, newEnd) => {
        group.startDate = newStart || null;
        group.dueDate = newEnd || null;
        group.originalTitle = updateDateRangeInText(group.originalTitle, newStart, newEnd);

        await commitCardChangesFromMemory(fileId);
        rebuildBoardCard(fileId);
        renderModalChecklists(cardObj);
    });
};

// --- Edit dates for a single todo item ---
window.editItemDates = function (fileId, groupIndex, itemIndex) {
    const cardObj = cardsStateMap.get(fileId);
    if (!cardObj) return;
    const group = cardObj.groups[groupIndex];
    if (!group || !group.items[itemIndex]) return;
    const item = group.items[itemIndex];

    showDateEditorPopup(item.startDate || '', item.dueDate || '', async (newStart, newEnd) => {
        item.startDate = newStart || null;
        item.dueDate = newEnd || null;
        item.originalText = updateDateRangeInText(item.originalText, newStart, newEnd);

        await commitCardChangesFromMemory(fileId);
        rebuildBoardCard(fileId);
        renderModalChecklists(cardObj);
    });
};

// --- Edit assignees for a checklist group ---
window.editGroupAssignees = function (fileId, groupIndex) {
    const cardObj = cardsStateMap.get(fileId);
    if (!cardObj) return;
    const group = cardObj.groups[groupIndex];
    if (!group) return;

    showAssigneeEditorPopup(group.assignees || [], async (newAssignees) => {
        group.assignees = newAssignees;
        group.originalTitle = updateAssigneesInText(group.originalTitle, newAssignees);
        group.title = extractMetadata(group.originalTitle).cleanText;

        await commitCardChangesFromMemory(fileId);
        rebuildBoardCard(fileId);
        renderModalChecklists(cardObj);
    });
};

// --- Edit assignees for a single todo item ---
window.editItemAssignees = function (fileId, groupIndex, itemIndex) {
    const cardObj = cardsStateMap.get(fileId);
    if (!cardObj) return;
    const group = cardObj.groups[groupIndex];
    if (!group || !group.items[itemIndex]) return;
    const item = group.items[itemIndex];

    showAssigneeEditorPopup(item.assignees || [], async (newAssignees) => {
        item.assignees = newAssignees;
        item.originalText = updateAssigneesInText(item.originalText, newAssignees);
        item.text = extractMetadata(item.originalText).cleanText;

        await commitCardChangesFromMemory(fileId);
        rebuildBoardCard(fileId);
        renderModalChecklists(cardObj);
    });
};

// --- Assignee editor popup ---
function showAssigneeEditorPopup(currentAssignees, onSave) {
    const existing = document.getElementById('assigneeEditorPopup');
    if (existing) existing.remove();

    let selected = [...(currentAssignees || [])];

    const popup = document.createElement('div');
    popup.id = 'assigneeEditorPopup';
    popup.className = 'fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-[100]';
    popup.innerHTML = `
        <div class="bg-slate-800 border border-slate-600 rounded-xl p-5 shadow-2xl w-80" onclick="event.stopPropagation()">
            <h4 class="text-sm font-bold text-white mb-3 flex items-center gap-2"><i class="fa-solid fa-user-pen text-blue-400"></i> 編輯負責人</h4>
            <input type="text" id="assigneeEditorFilter" class="bg-slate-900 border border-slate-700 focus:border-blue-500 rounded px-2.5 py-1.5 text-sm text-white w-full outline-none transition-colors placeholder:text-slate-500 mb-3" placeholder="搜尋或輸入新人員後按 Enter...">
            <div id="assigneeEditorList" class="max-h-48 overflow-y-auto custom-scrollbar space-y-1 mb-4"></div>
            <div id="assigneeEditorSelected" class="flex flex-wrap gap-1.5 mb-4 min-h-[28px]"></div>
            <div class="flex gap-2 justify-end">
                <button id="assigneeEditorCancelBtn" class="text-slate-400 hover:text-white text-xs px-3 py-1.5 rounded border border-slate-700 hover:bg-slate-700 transition-colors">取消</button>
                <button id="assigneeEditorSaveBtn" class="bg-blue-600 hover:bg-blue-500 text-white text-xs px-4 py-1.5 rounded font-medium transition-colors shadow-sm">儲存</button>
            </div>
        </div>
    `;
    document.body.appendChild(popup);

    const filterInput = document.getElementById('assigneeEditorFilter');
    const listElem = document.getElementById('assigneeEditorList');
    const selectedElem = document.getElementById('assigneeEditorSelected');

    function renderList() {
        listElem.innerHTML = '';
        const filterText = filterInput.value.trim().toLowerCase();
        let allNames = [...(boardSettings.Assignees || [])];
        // Include selected names that might not be in settings
        selected.forEach(a => { if (!allNames.includes(a)) allNames.push(a); });
        const filtered = filterText ? allNames.filter(a => a.toLowerCase().includes(filterText)) : allNames;

        filtered.forEach(name => {
            const isApplied = selected.includes(name);
            const initial = name.substring(0, 1).toUpperCase();
            const rowClass = isApplied ? 'bg-slate-700/80' : 'hover:bg-slate-700/50';
            const iconHtml = isApplied ? '<i class="fa-solid fa-check text-blue-400 text-xs mr-2"></i>' : '<i class="fa-solid fa-check text-transparent text-xs mr-2"></i>';

            const row = document.createElement('div');
            row.className = `flex items-center px-2 py-1.5 rounded transition-colors cursor-pointer ${rowClass}`;
            row.innerHTML = `${iconHtml}<div class="w-5 h-5 bg-blue-500 rounded-full flex items-center justify-center text-white font-bold text-[10px] shrink-0 mr-2">${escapeHtml(initial)}</div><span class="text-sm text-slate-300 truncate flex-1">${escapeHtml(name)}</span>`;
            row.addEventListener('click', () => {
                if (selected.includes(name)) {
                    selected = selected.filter(a => a !== name);
                } else {
                    selected.push(name);
                }
                renderList();
                renderSelected();
            });
            listElem.appendChild(row);
        });
    }

    function renderSelected() {
        selectedElem.innerHTML = '';
        if (selected.length === 0) {
            selectedElem.innerHTML = '<span class="text-slate-500 text-xs italic">未指派</span>';
            return;
        }
        selected.forEach(name => {
            const initial = name.substring(0, 1).toUpperCase();
            selectedElem.insertAdjacentHTML('beforeend', `
                <div class="flex items-center gap-1 bg-blue-500/15 border border-blue-500/30 rounded-full px-2 py-0.5 text-[11px] text-blue-300">
                    <div class="w-4 h-4 bg-blue-500 rounded-full flex items-center justify-center text-white font-bold text-[9px]">${escapeHtml(initial)}</div>
                    ${escapeHtml(name)}
                </div>
            `);
        });
    }

    renderList();
    renderSelected();
    filterInput.focus();

    filterInput.addEventListener('input', renderList);
    filterInput.addEventListener('keyup', (e) => {
        if (e.key === 'Enter') {
            const val = filterInput.value.trim();
            if (val && !selected.includes(val)) {
                selected.push(val);
                // Auto-add to global settings if new
                if (!boardSettings.Assignees.includes(val)) {
                    boardSettings.Assignees.push(val);
                    saveAssigneesToSettings();
                }
                filterInput.value = '';
                renderList();
                renderSelected();
            }
        }
    });

    popup.addEventListener('click', () => popup.remove());
    document.getElementById('assigneeEditorCancelBtn').addEventListener('click', () => popup.remove());
    document.getElementById('assigneeEditorSaveBtn').addEventListener('click', () => {
        popup.remove();
        onSave(selected);
    });
}

// --- Date editor popup (uses Flatpickr) ---
function showDateEditorPopup(currentStart, currentEnd, onSave) {
    // Remove any existing popup
    const existing = document.getElementById('dateEditorPopup');
    if (existing) existing.remove();

    const popup = document.createElement('div');
    popup.id = 'dateEditorPopup';
    popup.className = 'fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-[100]';
    popup.innerHTML = `
        <div class="bg-slate-800 border border-slate-600 rounded-xl p-5 shadow-2xl w-80" onclick="event.stopPropagation()">
            <h4 class="text-sm font-bold text-white mb-4 flex items-center gap-2"><i class="fa-regular fa-calendar text-blue-400"></i> 編輯日期區間</h4>
            <div class="flex flex-col gap-3 mb-4">
                <div class="flex items-center gap-2">
                    <span class="text-slate-400 text-xs w-10 text-right shrink-0">起始</span>
                    <input type="text" id="dateEditorStart" class="bg-slate-900 border border-slate-700 focus:border-blue-500 rounded px-2.5 py-1.5 text-sm text-white flex-1 outline-none transition-colors" placeholder="YYYY-MM-DD" value="${currentStart || ''}">
                </div>
                <div class="flex items-center gap-2">
                    <span class="text-slate-400 text-xs w-10 text-right shrink-0">到期</span>
                    <input type="text" id="dateEditorEnd" class="bg-slate-900 border border-slate-700 focus:border-blue-500 rounded px-2.5 py-1.5 text-sm text-white flex-1 outline-none transition-colors" placeholder="YYYY-MM-DD" value="${currentEnd || ''}">
                </div>
            </div>
            <div class="flex gap-2 justify-end">
                <button id="dateEditorClearBtn" class="text-slate-400 hover:text-white text-xs px-3 py-1.5 rounded border border-slate-700 hover:bg-slate-700 transition-colors">清除日期</button>
                <button id="dateEditorCancelBtn" class="text-slate-400 hover:text-white text-xs px-3 py-1.5 rounded border border-slate-700 hover:bg-slate-700 transition-colors">取消</button>
                <button id="dateEditorSaveBtn" class="bg-blue-600 hover:bg-blue-500 text-white text-xs px-4 py-1.5 rounded font-medium transition-colors shadow-sm">儲存</button>
            </div>
        </div>
    `;
    document.body.appendChild(popup);

    // Init Flatpickr on the date inputs (含交叉日期驗證)
    const startInput = document.getElementById('dateEditorStart');
    const endInput = document.getElementById('dateEditorEnd');

    const fpStart = flatpickr(startInput, {
        dateFormat: "Y-m-d", theme: "dark",
        defaultDate: currentStart || null,
        maxDate: currentEnd || null,
        onChange: function (selectedDates, dateStr) {
            if (fpEnd) fpEnd.set('minDate', dateStr || null);
        }
    });
    const fpEnd = flatpickr(endInput, {
        dateFormat: "Y-m-d", theme: "dark",
        defaultDate: currentEnd || null,
        minDate: currentStart || null,
        onChange: function (selectedDates, dateStr) {
            if (fpStart) fpStart.set('maxDate', dateStr || null);
        }
    });

    // Close on backdrop click
    popup.addEventListener('click', () => popup.remove());

    document.getElementById('dateEditorCancelBtn').addEventListener('click', () => popup.remove());

    document.getElementById('dateEditorClearBtn').addEventListener('click', () => {
        popup.remove();
        onSave('', '');
    });

    document.getElementById('dateEditorSaveBtn').addEventListener('click', () => {
        const newStart = startInput.value.trim();
        const newEnd = endInput.value.trim();
        popup.remove();
        onSave(newStart, newEnd);
    });
}

// Checkbox 直接點擊更新與群組功能已被移除


// ==========================================
// 5. DRAG AND DROP CAPABILITIES (SortableJS)
// ==========================================
function initSortable() {
    // 0. 清單欄位拖曳排序
    new Sortable(ui.boardContainer, {
        animation: 200,
        handle: '.list-drag-handle',
        draggable: '.list-column',
        ghostClass: 'ghost-card',
        dragClass: 'drag-card',
        direction: 'horizontal',
        onEnd: function () {
            saveListOrder();
        }
    });

    // A. 建立卡片在跨清單之間的拖拉
    document.querySelectorAll('.cards-container').forEach(container => {
        new Sortable(container, {
            group: 'board-cards',
            animation: 150,
            handle: '.card-drag-handle',
            ghostClass: 'ghost-card',
            dragClass: 'drag-card',
            onEnd: function (evt) {
                // 如果卡片移動到了不同的清單
                if (evt.to !== evt.from) {
                    const cardElem = evt.item;
                    const fileId = cardElem.getAttribute('data-file-id');
                    const targetListElem = evt.to.closest('.list-column');
                    const targetFolderId = targetListElem.getAttribute('data-list-id');

                    moveCardInOneDrive(fileId, targetFolderId);
                }
            }
        });
    });

    // B. Group & C. Items 拖曳已移除此看板層級
}

// 儲存清單欄位排列順序到 Settings.md (未指定 order 時依畫面上的欄位順序)
async function saveListOrder(order) {
    if (!order) {
        order = [];
        ui.boardContainer.querySelectorAll('.list-column').forEach(el => {
            const name = el.getAttribute('data-list-name');
            if (name) order.push(name);
        });
    }
    boardSettings.CardOrder = order;

    try {
        const textParams = { headers: { 'Accept': 'text/plain, text/markdown, */*' } };
        const settingsPath = ONEDRIVE_BASE_PATH.replace(/:$/, '') + '/Settings.md:';
        let rawContent = await fetchGraph(`/root${settingsPath}/content`, textParams).catch(() => null);

        const orderJson = JSON.stringify(order);
        if (!rawContent) {
            rawContent = `---\nCardOrder: ${orderJson}\n---\n# 系統設定\n`;
        } else {
            rawContent = updateFrontmatterString(rawContent, 'CardOrder', orderJson);
        }

        await fetchGraph(`/root${settingsPath}/content`, {
            method: 'PUT',
            headers: { 'Content-Type': 'text/plain' },
            body: rawContent.trim() + "\n"
        });
    } catch (err) {
        console.error('儲存清單排序失敗', err);
    }
}

// reinitItemsSortable 已移除


// ==========================================
// 6. SYNCHRONIZE CHANGES TO ONEDRIVE
// ==========================================

// 存檔邏輯已重構為僅透過 Modal 操作

// 拖曳「整張卡片」到另個「清單 (資料夾)」時
async function moveCardInOneDrive(fileId, targetFolderId) {
    try {
        const cardElem = document.querySelector(`.card[data-file-id="${fileId}"]`);
        if (cardElem) cardElem.classList.add('opacity-50');

        const fetchOptions = {
            method: 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                parentReference: {
                    id: targetFolderId
                }
            })
        };

        await fetchGraph(`/items/${fileId}`, fetchOptions);

        if (cardElem) cardElem.classList.remove('opacity-50');
    } catch (err) {
        console.error("移動檔案失敗: ", err);
        alert("無法將檔案搬移到新清單！可能權限不足。");
        // 如果失敗我們重整還原，以防資料不一致
        loadBoardData();
    }
}

// 取得 (不存在則建立) parentId 底下名為 name 的資料夾，回傳資料夾 id
async function ensureFolder(parentId, name) {
    try {
        const existing = await fetchGraph(`/items/${parentId}:/${encodeURIComponent(name)}`);
        if (existing && existing.id) return existing.id;
    } catch (err) {
        if (!err.message.includes('404')) throw err;
    }
    const created = await fetchGraph(`/items/${parentId}/children`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" })
    });
    return created.id;
}

// 刪除卡片：不直接刪除 .md，而是移至 TodoList/_Deleted/{原清單名稱}/ (同名時自動改名)
window.deleteCurrentCard = async function () {
    if (!currentCardFileId) return;
    const fileId = currentCardFileId;
    const cardObj = cardsStateMap.get(fileId);
    const cardElem = document.querySelector(`.card[data-file-id="${fileId}"]`);
    const listElem = cardElem ? cardElem.closest('.list-column') : null;
    const listName = listElem ? listElem.getAttribute('data-list-name') : '未分類';
    const title = cardObj ? cardObj.title : '';

    if (!confirm(`確定要刪除卡片「${title}」嗎？\n（檔案會移至 _Deleted/${listName} 資料夾保存，不會真正刪除）`)) return;

    try {
        const root = await fetchGraph(`/root${ONEDRIVE_BASE_PATH}`);
        const deletedRootId = await ensureFolder(root.id, '_Deleted');
        const targetFolderId = await ensureFolder(deletedRootId, listName);

        await fetchGraph(`/items/${fileId}?@microsoft.graph.conflictBehavior=rename`, {
            method: 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ parentReference: { id: targetFolderId } })
        });

        cardsStateMap.delete(fileId);
        if (cardElem) cardElem.remove();
        if (ui.modalCloseBtn) ui.modalCloseBtn.click();
        currentCardFileId = null;
    } catch (err) {
        console.error("刪除卡片失敗", err);
        alert("刪除卡片失敗，請確認網路連線與 OneDrive 權限！");
    }
};


// ==========================================
// 7. MODAL EDITING CAPABILITIES (Dates & Labels)
// ==========================================

function renderLabelPopoverList() {
    const listElem = document.getElementById('popoverLabelList');
    if (!listElem) return;
    listElem.innerHTML = '';

    if (!currentCardFileId) return;
    const cardObj = cardsStateMap.get(currentCardFileId);
    if (!cardObj) return;

    const activeLabels = cardObj.meta.Labels || [];
    let allLabels = Object.keys(boardSettings.Labels);

    // Include active labels that might not be in settings
    activeLabels.forEach(l => {
        if (!allLabels.includes(l)) allLabels.push(l);
    });

    allLabels.forEach(lName => {
        const isApplied = activeLabels.includes(lName);
        const hexColor = boardSettings.Labels[lName] || '#6366f1';

        const rowClass = isApplied ? 'bg-slate-700/80' : 'hover:bg-slate-700/50';
        const iconClass = isApplied ? `<i class="fa-solid fa-check text-white text-xs mr-2"></i>` : `<i class="fa-solid fa-check text-transparent text-xs mr-2"></i>`;
        const style = `background-color: ${hexColor}33; color: ${hexColor}; border-color: ${hexColor}4D;`;

        function ensureHex7(color) {
            if (!color || !color.startsWith('#')) return '#6366f1';
            let c = color.trim();
            if (c.length === 4) return '#' + c[1] + c[1] + c[2] + c[2] + c[3] + c[3];
            return c.substring(0, 7);
        }
        const pickerValue = ensureHex7(hexColor);

        listElem.insertAdjacentHTML('beforeend', `
            <div class="flex items-center px-2 py-1.5 rounded transition-colors ${rowClass}">
                <div class="flex-1 flex items-center cursor-pointer min-w-0" onclick="toggleLabelFromSettings('${escapeHtml(lName)}')">
                    ${iconClass}
                    <span class="px-2 py-0.5 rounded text-[11px] font-bold border block flex-1 truncate text-left" style="${style}">${escapeHtml(lName)}</span>
                </div>
                <div class="ml-2 relative flex items-center shrink-0">
                    <input type="color" value="${pickerValue}" 
                           class="w-6 h-6 p-0 border border-slate-600 rounded cursor-pointer shrink-0 bg-transparent flex items-center justify-center overflow-hidden"
                           onchange="updateLabelColorFromPopover('${escapeHtml(lName)}', this.value)" 
                           title="更換標籤顏色">
                </div>
            </div>
        `);
    });
}

window.toggleLabelFromSettings = async function (labelName) {
    if (!currentCardFileId) return;
    const cardObj = cardsStateMap.get(currentCardFileId);
    if (!cardObj) return;

    let labels = cardObj.meta.Labels || [];
    if (!Array.isArray(labels)) labels = [labels];

    if (labels.includes(labelName)) {
        labels = labels.filter(l => l !== labelName);
    } else {
        labels.push(labelName);
    }

    cardObj.meta.Labels = labels;
    await saveCardMetaAndRefresh(currentCardFileId, 'Labels', JSON.stringify(labels));
    renderLabelPopoverList();
};

window.updateLabelColorFromPopover = async function (labelName, newHexColor) {
    if (!labelName || !newHexColor) return;
    boardSettings.Labels[labelName] = newHexColor;

    // Soft UI Update (all label spans in board and modal)
    const hexColor = newHexColor;
    const styleText = `background-color: ${hexColor}33; color: ${hexColor}; border-color: ${hexColor}4D;`;
    document.querySelectorAll('span.border').forEach(span => {
        if (span.innerText.trim() === labelName) {
            span.style.cssText = styleText;
        }
    });

    // Hard Save to Settings.md
    try {
        const textParams = { headers: { 'Accept': 'text/plain, text/markdown, */*' } };
        const settingsPath = ONEDRIVE_BASE_PATH.replace(/:$/, '') + '/Settings.md:';
        let rawContent = await fetchGraph(`/root${settingsPath}/content`, textParams).catch(() => null);

        if (!rawContent) {
            rawContent = `---\nLabels:\n  ${labelName}: "${newHexColor}"\n---\n# 系統設定\n`;
        } else {
            rawContent = updateYamlChildString(rawContent, 'Labels', labelName, `"${newHexColor}"`);
        }

        await fetchGraph(`/root${settingsPath}/content`, {
            method: 'PUT',
            headers: { 'Content-Type': 'text/plain' },
            body: rawContent.trim() + "\n"
        });
    } catch (err) {
        console.error('儲存標籤色彩失敗', err);
    }
};

function updateYamlChildString(fmString, parentKey, childKey, childValue) {
    let lines = fmString.split(/\r?\n/);
    let inParent = false;
    let foundChild = false;
    let parentIndent = "  ";

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (line.match(new RegExp(`^${parentKey}:\\s*$`, 'i'))) {
            inParent = true;
            continue;
        }

        if (inParent) {
            if (line.match(/^[A-Za-z0-9_]+:/) || line === '---' || line === '...') {
                lines.splice(i, 0, `${parentIndent}${childKey}: ${childValue}`);
                foundChild = true;
                break;
            }

            const childMatch = line.match(/^(\s+)([^:]+):\s*(.*)$/);
            if (childMatch) {
                parentIndent = childMatch[1];
                let currentChildKey = childMatch[2].trim();
                if (currentChildKey.startsWith('"') && currentChildKey.endsWith('"')) {
                    currentChildKey = currentChildKey.substring(1, currentChildKey.length - 1);
                }
                if (currentChildKey === childKey) {
                    lines[i] = `${parentIndent}${childKey}: ${childValue}`;
                    foundChild = true;
                    break;
                }
            }
        }
    }

    if (!foundChild && !inParent) {
        for (let i = lines.length - 1; i >= 0; i--) {
            if (lines[i] === '---') {
                lines.splice(i, 0, `${parentKey}:\n  ${childKey}: ${childValue}`);
                break;
            }
        }
    } else if (!foundChild && inParent) {
        let lastIdx = lines.length - 1;
        while (lastIdx >= 0 && (lines[lastIdx].trim() === '' || lines[lastIdx] === '---')) lastIdx--;
        lines.splice(lastIdx + 1, 0, `${parentIndent}${childKey}: ${childValue}`);
    }

    return lines.join('\n');
}

// Helper: update frontmatter string directly
function updateFrontmatterString(fmString, key, value) {
    if (!fmString) fmString = "---\n---\n";
    let lines = fmString.split(/\r?\n/);
    let found = false;
    for (let i = 0; i < lines.length; i++) {
        // match key ignoring case to be safe, but preserve original case of others
        if (lines[i].match(new RegExp(`^${key}:`, 'i'))) {
            if (value === '') {
                lines.splice(i, 1);
            } else {
                lines[i] = `${key}: ${value}`;
            }
            found = true;
            break;
        }
    }
    if (!found && value !== '') {
        for (let i = lines.length - 1; i >= 0; i--) {
            if (lines[i] === '---') { // insert before the closing ---
                lines.splice(i, 0, `${key}: ${value}`);
                found = true;
                break;
            }
        }
    }
    if (!found && value !== '' && lines.length <= 2) { // just in case block isn't closed properly
        return `---\n${key}: ${value}\n---\n`;
    }
    return lines.join('\n');
}

// 共通存檔與重整函數
async function saveCardMetaAndRefresh(fileId, key, value) {
    const cardObj = cardsStateMap.get(fileId);
    if (!cardObj) return;

    // 更新記憶體物件
    cardObj.frontmatter = updateFrontmatterString(cardObj.frontmatter, key, value);

    // 呼叫存檔，並刷新看板與 Modal UI
    await commitCardChangesFromMemory(fileId);
    rebuildBoardCard(fileId);
    openCardDetail(fileId); // refresh modal UI
}

let isSavingCard = false;
async function updateCardMetaSingle(fileId, key, value) {
    if (isSavingCard) {
        // wait for the current save to finish then proceed -> serial queue
        await new Promise(r => setTimeout(r, 300));
        return updateCardMetaSingle(fileId, key, value);
    }

    isSavingCard = true;
    try {
        const cardObj = cardsStateMap.get(fileId);
        if (!cardObj) return;
        cardObj.meta[key] = value;
        if (!value) {
            await saveCardMetaAndRefresh(fileId, key, '');
        } else {
            await saveCardMetaAndRefresh(fileId, key, value);
        }
    } finally {
        isSavingCard = false;
    }
}

async function addLabelToCard(fileId, newLabel) {
    const cardObj = cardsStateMap.get(fileId);
    if (!cardObj) return;
    if (!cardObj.meta.Labels) cardObj.meta.Labels = [];
    if (!Array.isArray(cardObj.meta.Labels)) cardObj.meta.Labels = [cardObj.meta.Labels];

    if (!cardObj.meta.Labels.includes(newLabel)) {
        cardObj.meta.Labels.push(newLabel);
        await saveCardMetaAndRefresh(fileId, 'Labels', JSON.stringify(cardObj.meta.Labels));
    }
}

window.removeLabel = async function (idx) {
    if (!currentCardFileId) return;
    const cardObj = cardsStateMap.get(currentCardFileId);
    if (!cardObj || !Array.isArray(cardObj.meta.Labels)) return;

    cardObj.meta.Labels.splice(idx, 1);
    await saveCardMetaAndRefresh(currentCardFileId, 'Labels', JSON.stringify(cardObj.meta.Labels));
};

// ==========================================
// 8. MODAL EDITING CAPABILITIES (Assignees)
// ==========================================

function renderAssigneePopoverList() {
    const listElem = document.getElementById('popoverAssigneeList');
    if (!listElem) return;
    listElem.innerHTML = '';

    if (!currentCardFileId) return;
    const cardObj = cardsStateMap.get(currentCardFileId);
    if (!cardObj) return;

    const activeAssignees = cardObj.meta.Assignees || [];
    let allAssignees = [...(boardSettings.Assignees || [])];

    // Include active assignees that might not be in settings
    activeAssignees.forEach(a => {
        if (!allAssignees.includes(a)) allAssignees.push(a);
    });

    // Filter by search input
    const filterInput = document.getElementById('popoverAssigneeInput');
    const filterText = filterInput ? filterInput.value.trim().toLowerCase() : '';
    const filtered = filterText ? allAssignees.filter(a => a.toLowerCase().includes(filterText)) : allAssignees;

    filtered.forEach(name => {
        const isApplied = activeAssignees.includes(name);
        const initial = name.substring(0, 1).toUpperCase();
        const rowClass = isApplied ? 'bg-slate-700/80' : 'hover:bg-slate-700/50';
        const iconClass = isApplied ? `<i class="fa-solid fa-check text-blue-400 text-xs mr-2"></i>` : `<i class="fa-solid fa-check text-transparent text-xs mr-2"></i>`;

        listElem.insertAdjacentHTML('beforeend', `
            <div class="flex items-center px-2 py-1.5 rounded transition-colors cursor-pointer ${rowClass}" onclick="toggleAssigneeFromSettings('${escapeHtml(name)}')">
                ${iconClass}
                <div class="w-5 h-5 bg-blue-500 rounded-full flex items-center justify-center text-white font-bold text-[10px] shrink-0 mr-2">${escapeHtml(initial)}</div>
                <span class="text-sm text-slate-300 truncate flex-1">${escapeHtml(name)}</span>
            </div>
        `);
    });
}

window.toggleAssigneeFromSettings = async function (name) {
    if (!currentCardFileId) return;
    const cardObj = cardsStateMap.get(currentCardFileId);
    if (!cardObj) return;

    let assignees = cardObj.meta.Assignees || [];
    if (!Array.isArray(assignees)) assignees = [assignees];

    if (assignees.includes(name)) {
        assignees = assignees.filter(a => a !== name);
    } else {
        assignees.push(name);
        // Auto-add to global settings if not already present
        if (!boardSettings.Assignees.includes(name)) {
            boardSettings.Assignees.push(name);
            saveAssigneesToSettings();
        }
    }

    cardObj.meta.Assignees = assignees;
    await saveCardMetaAndRefresh(currentCardFileId, 'Assignees', JSON.stringify(assignees));
    renderAssigneePopoverList();
};

window.removeAssignee = async function (idx) {
    if (!currentCardFileId) return;
    const cardObj = cardsStateMap.get(currentCardFileId);
    if (!cardObj || !Array.isArray(cardObj.meta.Assignees)) return;

    cardObj.meta.Assignees.splice(idx, 1);
    await saveCardMetaAndRefresh(currentCardFileId, 'Assignees', JSON.stringify(cardObj.meta.Assignees));
};

async function saveAssigneesToSettings() {
    try {
        const textParams = { headers: { 'Accept': 'text/plain, text/markdown, */*' } };
        const settingsPath = ONEDRIVE_BASE_PATH.replace(/:$/, '') + '/Settings.md:';
        let rawContent = await fetchGraph(`/root${settingsPath}/content`, textParams).catch(() => null);

        const assigneesJson = JSON.stringify(boardSettings.Assignees);
        if (!rawContent) {
            rawContent = `---\nAssignees: ${assigneesJson}\n---\n# 系統設定\n`;
        } else {
            rawContent = updateFrontmatterString(rawContent, 'Assignees', assigneesJson);
        }

        await fetchGraph(`/root${settingsPath}/content`, {
            method: 'PUT',
            headers: { 'Content-Type': 'text/plain' },
            body: rawContent.trim() + "\n"
        });
    } catch (err) {
        console.error('儲存負責人清單失敗', err);
    }
}

// Bind Event Listeners on DOMContentLoaded
document.addEventListener('DOMContentLoaded', () => {
    // Dates (Flatpickr)
    const startDateInput = document.getElementById('modalStartDateInput');
    const dueDateInput = document.getElementById('modalDueDateInput');

    let startPicker, duePicker;
    if (startDateInput) {
        startPicker = flatpickr(startDateInput, {
            dateFormat: "Y-m-d",
            theme: "dark",
            onChange: async function (selectedDates, dateStr) {
                if (duePicker) duePicker.set('minDate', dateStr || null);
                if (!currentCardFileId) return;
                await updateCardMetaSingle(currentCardFileId, 'StartDate', dateStr);
            }
        });
    }
    if (dueDateInput) {
        duePicker = flatpickr(dueDateInput, {
            dateFormat: "Y-m-d",
            theme: "dark",
            onChange: async function (selectedDates, dateStr) {
                if (startPicker) startPicker.set('maxDate', dateStr || null);
                if (!currentCardFileId) return;
                await updateCardMetaSingle(currentCardFileId, 'DueDate', dateStr);
            }
        });
    }

    // Labels Popover
    const addLabelBtn = document.getElementById('modalAddLabelBtn');
    const popover = document.getElementById('labelSelectorPopover');
    const popoverInput = document.getElementById('popoverLabelInput');

    if (addLabelBtn && popover) {
        addLabelBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            popover.classList.toggle('hidden');
            if (!popover.classList.contains('hidden')) {
                if (popoverInput) popoverInput.focus();
                renderLabelPopoverList();
            }
        });

        document.addEventListener('click', (e) => {
            if (!popover.contains(e.target) && !addLabelBtn.contains(e.target)) {
                popover.classList.add('hidden');
            }
        });
    }

    if (popoverInput) {
        popoverInput.addEventListener('keyup', async (e) => {
            if (e.key === 'Enter') {
                const val = popoverInput.value.trim();
                if (val && currentCardFileId) {
                    await window.toggleLabelFromSettings(val);
                    popoverInput.value = '';
                }
            }
        });
    }

    // Assignees Popover
    const addAssigneeBtn = document.getElementById('modalAddAssigneeBtn');
    const assigneePopover = document.getElementById('assigneeSelectorPopover');
    const assigneePopoverInput = document.getElementById('popoverAssigneeInput');

    if (addAssigneeBtn && assigneePopover) {
        addAssigneeBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            assigneePopover.classList.toggle('hidden');
            // Close label popover if open
            if (popover) popover.classList.add('hidden');
            if (!assigneePopover.classList.contains('hidden')) {
                if (assigneePopoverInput) assigneePopoverInput.focus();
                renderAssigneePopoverList();
            }
        });

        document.addEventListener('click', (e) => {
            if (!assigneePopover.contains(e.target) && !addAssigneeBtn.contains(e.target)) {
                assigneePopover.classList.add('hidden');
            }
        });
    }

    if (assigneePopoverInput) {
        assigneePopoverInput.addEventListener('input', () => {
            renderAssigneePopoverList();
        });
        assigneePopoverInput.addEventListener('keyup', async (e) => {
            if (e.key === 'Enter') {
                const val = assigneePopoverInput.value.trim();
                if (val && currentCardFileId) {
                    await window.toggleAssigneeFromSettings(val);
                    assigneePopoverInput.value = '';
                    renderAssigneePopoverList();
                }
            }
        });
    }

    // Attachments Upload handler
    const uploadInput = document.getElementById('modalAttachmentInput');
    if (uploadInput) {
        uploadInput.addEventListener('change', async function(e) {
            const files = e.target.files;
            if (!files || files.length === 0 || !currentCardFileId) return;

            const cardObj = cardsStateMap.get(currentCardFileId);
            const cardElem = document.querySelector(`.card[data-file-id="${currentCardFileId}"]`);
            if (!cardObj || !cardElem) return;

            const listElem = cardElem.closest('.list-column');
            if (!listElem) return;
            const listId = listElem.getAttribute('data-list-id');
            if (!listId) {
                alert("無法取得該卡片所屬的清單目錄 ID！");
                return;
            }

            const uploadBtn = uploadInput.previousElementSibling;
            const originalHtml = uploadBtn.innerHTML;
            uploadBtn.innerHTML = '<i class="fa-solid fa-circle-notch fa-spin"></i> 上傳中...';
            uploadBtn.disabled = true;

            try {
                if (!cardObj.meta.Attachments) cardObj.meta.Attachments = [];
                if (!Array.isArray(cardObj.meta.Attachments)) cardObj.meta.Attachments = [cardObj.meta.Attachments];

                for (let i = 0; i < files.length; i++) {
                    const file = files[i];
                    const fetchOptions = {
                        method: 'PUT',
                        headers: {
                            'Content-Type': file.type || 'application/octet-stream'
                        },
                        body: file
                    };
                    
                    // Create explicitly under '_Upload' folder via Graph API path addressing relative to List root
                    const res = await fetchGraph(`/items/${listId}:/_Upload/${encodeURIComponent(file.name)}:/content`, fetchOptions);
                    
                    if (res && res.id) {
                        cardObj.meta.Attachments.push({
                            name: file.name,
                            url: res.webUrl || '#',
                            id: res.id
                        });
                    }
                }

                await saveCardMetaAndRefresh(currentCardFileId, 'Attachments', JSON.stringify(cardObj.meta.Attachments));
                renderCardAttachments(currentCardFileId);

            } catch (err) {
                console.error("附件上傳失敗", err);
                alert("附件上傳失敗，請確認檔案大小或連線狀態！");
            } finally {
                uploadBtn.innerHTML = originalHtml;
                uploadBtn.disabled = false;
                uploadInput.value = '';
            }
        });
    }
});


window.clearDate = async function (key) {
    if (!currentCardFileId) return;

    const inputId = key === 'StartDate' ? 'modalStartDateInput' : 'modalDueDateInput';
    const inputElem = document.getElementById(inputId);

    // Clear flatpickr visually; this will automatically trigger the onChange event which saves the data.
    if (inputElem && inputElem._flatpickr) {
        inputElem._flatpickr.clear();
    } else if (inputElem) {
        inputElem.value = '';
        await updateCardMetaSingle(currentCardFileId, key, '');
    }
};

// ==========================================
// 8. ADD NEW CARD CAPABILITIES
// ==========================================
window.toggleAddCardInput = function(folderId, show) {
    const btn = document.getElementById(`addCardBtn_${folderId}`);
    const form = document.getElementById(`addCardForm_${folderId}`);
    const input = document.getElementById(`addCardInput_${folderId}`);
    
    if (show) {
        btn.classList.add('hidden');
        form.classList.remove('hidden');
        form.classList.add('flex');
        input.focus();
        
        // Setup keyboard handlers if not already bound
        if (!input.dataset.hasKeyHandler) {
            input.addEventListener('keydown', function(e) {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    submitNewCard(folderId);
                } else if (e.key === 'Escape') {
                    toggleAddCardInput(folderId, false);
                }
            });
            input.dataset.hasKeyHandler = 'true';
        }
    } else {
        btn.classList.remove('hidden');
        form.classList.add('hidden');
        form.classList.remove('flex');
        input.value = '';
    }
};

window.submitNewCard = async function(folderId) {
    const input = document.getElementById(`addCardInput_${folderId}`);
    const title = input.value.trim();
    if (!title) return;
    
    const form = document.getElementById(`addCardForm_${folderId}`);
    const submitBtn = form.querySelector('button.bg-blue-600');
    
    // Disable input during request
    input.disabled = true;
    submitBtn.disabled = true;
    submitBtn.innerHTML = '<i class="fa-solid fa-circle-notch fa-spin"></i>';
    
    try {
        const filename = `${title}.md`;
        // Create an empty markdown file with frontmatter title
        const content = `---\nTitle: "${title.replace(/"/g, '\\"')}"\n---\n`;
        
        const fetchOptions = {
            method: 'PUT',
            headers: { 'Content-Type': 'text/plain' },
            body: content
        };
        
        // Create the file in the designated folder via Microsoft Graph
        const res = await fetchGraph(`/items/${folderId}:/${encodeURIComponent(filename)}:/content`, fetchOptions);
        
        if (res && res.id) {
            // Success: Parse the new card to match state structure and insert into memory state
            const cardObj = parseMarkdown(content, filename);
            cardsStateMap.set(res.id, cardObj);
            
            // Build the card HTML and inject it securely
            const listContainer = document.querySelector(`.list-column[data-list-id="${folderId}"] .cards-container`);
            if (listContainer) {
                listContainer.insertAdjacentHTML('beforeend', buildCardHtml(res.id, filename, cardObj));
            }
            
            toggleAddCardInput(folderId, false);
        } else {
            throw new Error("No File ID returned from OneDrive");
        }
    } catch (err) {
        console.error("Failed to create card", err);
        alert("新增卡片時發生錯誤，請確認網路連線與 OneDrive 權限！");
    } finally {
        input.disabled = false;
        submitBtn.disabled = false;
        submitBtn.innerText = '新增';
    }
};

// ==========================================
// 9. ADD NEW LIST CAPABILITIES
// ==========================================
window.toggleAddListInput = function (show) {
    const btn = document.getElementById('addListBtn');
    const form = document.getElementById('addListForm');
    const input = document.getElementById('addListInput');

    if (show) {
        btn.classList.add('hidden');
        form.classList.remove('hidden');
        form.classList.add('flex');
        input.focus();

        if (!input.dataset.hasKeyHandler) {
            input.addEventListener('keydown', function (e) {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    submitNewList();
                } else if (e.key === 'Escape') {
                    toggleAddListInput(false);
                }
            });
            input.dataset.hasKeyHandler = 'true';
        }
    } else {
        btn.classList.remove('hidden');
        form.classList.add('hidden');
        form.classList.remove('flex');
        input.value = '';
    }
};

// 新增清單：在 TodoList 根目錄建立資料夾，並加到 CardOrder 最後
window.submitNewList = async function () {
    const input = document.getElementById('addListInput');
    const submitBtn = document.getElementById('addListSubmitBtn');
    const name = input.value.trim();
    if (!name) return;

    if (/[\\/:*?"<>|#%]/.test(name) || name.startsWith('_') || name.startsWith('.')) {
        alert('清單名稱不可包含 \\ / : * ? " < > | # % 等字元，也不可以 _ 或 . 開頭。');
        return;
    }

    const currentOrder = [];
    ui.boardContainer.querySelectorAll('.list-column').forEach(el => {
        const n = el.getAttribute('data-list-name');
        if (n) currentOrder.push(n);
    });
    if (currentOrder.includes(name)) {
        alert(`清單「${name}」已存在！`);
        return;
    }

    input.disabled = true;
    submitBtn.disabled = true;
    submitBtn.innerHTML = '<i class="fa-solid fa-circle-notch fa-spin"></i>';

    try {
        await fetchGraph(`/root${ONEDRIVE_BASE_PATH}/children`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" })
        });
        await saveListOrder([...currentOrder, name]);
        await loadBoardData();
    } catch (err) {
        console.error("Failed to create list", err);
        if (err.message.includes('409')) {
            alert(`OneDrive 中已有名為「${name}」的資料夾！`);
        } else {
            alert("新增清單時發生錯誤，請確認網路連線與 OneDrive 權限！");
        }
        input.disabled = false;
        submitBtn.disabled = false;
        submitBtn.innerText = '新增';
    }
};

function renderCardAttachments(fileId) {
    const cardObj = cardsStateMap.get(fileId);
    if (!cardObj) return;

    const container = document.getElementById('modalAttachmentsContainer');
    const list = document.getElementById('modalAttachmentsList');
    if (!container || !list) return;
    
    if (!cardObj.meta.Attachments || cardObj.meta.Attachments.length === 0) {
        container.classList.add('hidden');
        list.innerHTML = '';
        return;
    }

    container.classList.remove('hidden');
    let html = '';
    cardObj.meta.Attachments.forEach((att, idx) => {
        const name = typeof att === 'string' ? att : att.name;
        const url = typeof att === 'string' ? '#' : (att.url || '#');
        
        const isImage = name.match(/\.(jpg|jpeg|png|gif|webp|svg)$/i);
        const iconClass = isImage ? 'fa-image text-purple-400' : 'fa-file-lines text-slate-400';
        
        let linkAttr = url === '#' ? `onclick="alert('無法取得檔案連結')"` : `href="${url}" target="_blank"`;

        html += `
            <div class="flex items-center justify-between bg-slate-800/80 border border-slate-700/50 p-2.5 rounded-lg group transition-colors hover:bg-slate-800">
                <a ${linkAttr} class="flex items-center gap-3 min-w-0 flex-1 hover:opacity-80 transition-opacity" title="點擊檢視或下載">
                    <div class="w-8 h-8 rounded bg-slate-900 flex items-center justify-center shrink-0 border border-slate-700 shadow-sm">
                        <i class="fa-solid ${iconClass}"></i>
                    </div>
                    <div class="truncate text-sm font-medium text-slate-300">
                        ${escapeHtml(name)}
                    </div>
                </a>
                <button class="text-slate-600 hover:text-red-400 opacity-0 group-hover:opacity-100 transition-all p-2 shrink-0" onclick="removeAttachment(${idx})" title="解除附件關聯">
                    <i class="fa-solid fa-trash-can"></i>
                </button>
            </div>
        `;
    });
    list.innerHTML = html;
}

window.removeAttachment = async function(idx) {
    if (!currentCardFileId) return;
    const cardObj = cardsStateMap.get(currentCardFileId);
    if (!cardObj || !cardObj.meta.Attachments) return;

    if (!confirm("確定要解除此附件的關聯嗎？\\n（實際檔案仍會保留在 OneDrive 的 _Upload 資料夾中）")) return;

    cardObj.meta.Attachments.splice(idx, 1);
    await saveCardMetaAndRefresh(currentCardFileId, 'Attachments', JSON.stringify(cardObj.meta.Attachments));
    renderCardAttachments(currentCardFileId);
};

// Start application logic
initMsal();

window.toggleDescriptionEdit = function(show) {
    const descElem = document.getElementById('modalDescription');
    const descEditor = document.getElementById('modalDescriptionEditor');
    const editBtn = document.getElementById('modalEditDescriptionBtn');
    const descTextarea = document.getElementById('modalDescriptionTextarea');

    if (show) {
        descElem.classList.add('hidden');
        editBtn.classList.add('hidden');
        descEditor.classList.remove('hidden');
        descEditor.classList.add('flex');
        descTextarea.focus();
    } else {
        descElem.classList.remove('hidden');
        editBtn.classList.remove('hidden');
        descEditor.classList.add('hidden');
        descEditor.classList.remove('flex');
        
        // Reset textarea to current saved text if cancelled
        if (currentCardFileId) {
            const cardObj = cardsStateMap.get(currentCardFileId);
            if (cardObj) {
                descTextarea.value = cardObj.description || '';
            }
        }
    }
};

window.saveDescriptionEdit = async function() {
    if (!currentCardFileId) return;
    
    const descTextarea = document.getElementById('modalDescriptionTextarea');
    const newDesc = descTextarea.value;
    
    const cardObj = cardsStateMap.get(currentCardFileId);
    if (!cardObj) return;
    
    // Update internal state
    cardObj.description = newDesc;
    
    const saveBtn = document.querySelector('#modalDescriptionEditor button.bg-blue-600');
    let originalText = '儲存';
    if (saveBtn) {
        originalText = saveBtn.innerText;
        saveBtn.disabled = true;
        saveBtn.innerHTML = '<i class="fa-solid fa-circle-notch fa-spin"></i>';
    }

    try {
        // Save to OneDrive via commitCardChangesFromMemory
        await commitCardChangesFromMemory(currentCardFileId);
        
        // Re-render UI
        const descElem = document.getElementById('modalDescription');
        descElem.innerHTML = renderMarkdownBasic(cardObj.description);
        
        // Rebuild board card to sync memory state visually everywhere safely
        rebuildBoardCard(currentCardFileId);
        
        // Close editor
        toggleDescriptionEdit(false);
    } catch (err) {
        console.error("Failed to save description", err);
        alert("儲存說明失敗，請檢查網路連線！");
    } finally {
        if (saveBtn) {
            saveBtn.disabled = false;
            saveBtn.innerText = originalText;
        }
    }
};
