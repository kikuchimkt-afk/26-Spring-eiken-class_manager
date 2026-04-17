// 日程情報（動的に追加可能）
let sessionsInfo = [
    { id: 'day1', date: '4月18日(土)', title: 'Day 1: 基礎力チェックと語彙' },
    { id: 'day2', date: '4月25日(土)', title: 'Day 2: リーディング演習' },
    { id: 'day3', date: '5月2日(土)', title: 'Day 3: リスニング演習' },
    { id: 'day4', date: '5月9日(土)', title: 'Day 4: 過去問演習と総括' },
    { id: 'day5', date: '5月16日(土)', title: 'Day 5: 予想問題演習' },
    { id: 'day6', date: '5月23日(土)', title: 'Day 6: 最終チェック' },
];

// モックデータ: 参加者リスト（初期値）
let participantsList = [
    { id: 'p1', name: '山田 太郎', grade: '2級', hasTablet: true, schoolYear: '高2' },
    { id: 'p2', name: '佐藤 花子', grade: '準2級', hasTablet: false, schoolYear: '高1' },
    { id: 'p3', name: '鈴木 健吉', grade: '3級', hasTablet: true, schoolYear: '中3' },
];

// 過去問の選択肢を生成する（古い順）
function getPastPaperOptions(grade) {
    let options = ['<option value="">未選択/その他</option>'];
    
    for (let year = 2018; year <= 2025; year++) {
        // 準2級プラスは2025年新設のためスキップ
        if (grade === '準2級プラス' && year < 2025) continue;
        
        for (let num = 1; num <= 3; num++) {
            // 旧仕様通り、3級の2024年度以降のみ「新形式」表記を付与
            let isNewFormat = (year >= 2024 && grade === '3級');
            
            let label = `第${num}回${isNewFormat ? '（新形式）' : ''}`;
            options.push(`<option value="${year}年度 ${label}">${year}年度 ${label}</option>`);
            
            let satLabel = `第${num}回（準会場${isNewFormat ? '/新形式' : ''}）`;
            options.push(`<option value="${year}年度 ${satLabel}">${year}年度 ${satLabel}</option>`);
        }
    }
    
    return options.join('\n');
}

// ====== GAS Web API URL ======
const GAS_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbz2lwmv77SBY8kDSwJK5wlawVUg0CAOKmQxTpX-GNLi5DjnXYRn2jbYqI-U2I1__ofC/exec";

// アプリの状態
let currentSessionId = null;
let appData = {};
let newParticipantIds = new Set(); // 未確認の新規参加者を追跡（localStorageに永続化）

// 初期化
function init() {
    // 保存済みの日程を読み込む
    const savedSessions = localStorage.getItem('eikenClassManagerSessions');
    if (savedSessions) {
        sessionsInfo = JSON.parse(savedSessions);
    }
    // 未確認のNEWバッジを復元
    const savedNewIds = localStorage.getItem('eikenNewParticipantIds');
    if (savedNewIds) {
        newParticipantIds = new Set(JSON.parse(savedNewIds));
    }
    loadData();
    renderSidebar();
}

// アプリ初期化・データ同期
async function loadData() {
    // 1. オフラインのバックアップをとりあえず読み込む
    const savedP = localStorage.getItem('eikenClassManagerParticipants');
    if (savedP) {
        participantsList = JSON.parse(savedP);
    }

    const saved = localStorage.getItem('eikenClassManagerData');
    if (saved) {
        appData = JSON.parse(saved);
        sessionsInfo.forEach(session => {
            if (!appData[session.id]) {
                appData[session.id] = { generalHomework: '', generalNotes: '', participants: {} };
            }
        });
    } else {
        sessionsInfo.forEach(session => {
            appData[session.id] = { generalHomework: '', generalNotes: '', participants: {} };
        });
    }
    
    // 2. 起動時に自動でクラウド同期（フォーム回答＆成績）
    await syncWithCloud();
}

// ====== クラウド同期機能 ======
async function syncWithCloud() {
    showLoading("クラウドから最新のデータを同期中...");
    try {
        const response = await fetch(GAS_WEB_APP_URL);
        const data = await response.json();
        
        // ★ 1. まずクラウドの成績データを先にマージ（空でないもののみ）
        if (data.appData && Object.keys(data.appData).length > 0) {
            Object.keys(data.appData).forEach(sessionId => {
                if (appData[sessionId]) {
                    const cloudSession = data.appData[sessionId];
                    // 参加者データがある場合のみ上書き
                    if (cloudSession.participants && Object.keys(cloudSession.participants).length > 0) {
                        appData[sessionId] = cloudSession;
                    } else {
                        // 参加者以外のフィールド（宿題・メモ）だけ反映
                        appData[sessionId].generalHomework = cloudSession.generalHomework || appData[sessionId].generalHomework;
                        appData[sessionId].generalNotes = cloudSession.generalNotes || appData[sessionId].generalNotes;
                    }
                }
            });
        }
        
        // ★ 2. その後にフォーム回答を処理（新規参加者を追加）
        newParticipantIds = new Set();
        if (data.formResponses && data.formResponses.length > 0) {
            processFormResponses(data.formResponses);
        }
        
        // 新規参加者がいれば通知を表示し、localStorageに保存
        if (newParticipantIds.size > 0) {
            showNotification(`🆕 新しい申し込みが ${newParticipantIds.size} 件あります！`);
            localStorage.setItem('eikenNewParticipantIds', JSON.stringify([...newParticipantIds]));
        }
        
        localStorage.setItem('eikenClassManagerData', JSON.stringify(appData));
        localStorage.setItem('eikenClassManagerParticipants', JSON.stringify(participantsList));
        
        if (!currentSessionId && sessionsInfo.length > 0) {
            currentSessionId = sessionsInfo[0].id;
        }
        renderSidebar();
        if (currentSessionId) {
            renderMainContent();
            updateHeader();
        }
    } catch(e) {
        console.error("クラウド同期エラー:", e);
        // エラーでも起動は続行（オフラインモード）
    } finally {
        hideLoading();
    }
}

function updateHeader() {
    if (!currentSessionId) return;
    const sessionInfo = sessionsInfo.find(s => s.id === currentSessionId);
    if (sessionInfo) {
        document.getElementById('currentDateTitle').textContent = sessionInfo.date;
        document.getElementById('currentSessionInfo').textContent = '';
        document.getElementById('emptyState').style.display = 'none';
        document.getElementById('mainScrollable').style.display = 'block';
    }
}

function showLoading(msg) {
    const overlay = document.getElementById('loadingOverlay');
    const text = document.getElementById('loadingText');
    if (overlay && text) {
        text.textContent = msg || "通信中...";
        overlay.classList.add('active');
    }
}

function hideLoading() {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.classList.remove('active');
    }
}

function processFormResponses(formResponses) {
    formResponses.forEach((row, i) => {
        const keys = Object.keys(row);
        let nameKey = keys.find(k => k.includes('氏名'));
        let yearKey = keys.find(k => k.includes('学年'));
        let gradeKey = keys.find(k => k.includes('級'));
        let tabletKey = keys.find(k => k.includes('タブレット'));
        let attendKey = keys.find(k => k.includes('参加したい日') || k.includes('参加'));

        if (!nameKey) return;
        
        const rawName = String(row[nameKey] || '');
        if (!rawName.trim()) return;
        const name = rawName.replace(/[\s　]+/g, ' ');
        
        const schoolYear = yearKey ? String(row[yearKey] || '') : '';
        const gradeStr = gradeKey ? String(row[gradeKey]) : '';
        let grade = '未選択';
        if (gradeStr.includes('5級')) grade = '5級';
        else if (gradeStr.includes('4級')) grade = '4級';
        else if (gradeStr.includes('3級')) grade = '3級';
        else if (gradeStr.includes('準2級プラス') || gradeStr.includes('準2級+')) grade = '準2級プラス';
        else if (gradeStr.includes('準2級')) grade = '準2級';
        else if (gradeStr.includes('2級')) grade = '2級';
        else if (gradeStr.includes('準1級')) grade = '準1級';
        else if (gradeStr.includes('1級')) grade = '1級';
        
        // 「希望しない」= 自分のタブレットを持参する → hasTablet: true
        // 「希望する」 = 貸し出し希望（持参しない）  → hasTablet: false
        const hasTablet = tabletKey ? String(row[tabletKey]).includes('希望しない') : false;
        const rawAttend = attendKey ? String(row[attendKey] || '') : '';

        // 新規か更新かを判定
        let existing = participantsList.find(p => p.name === name);
        if (!existing) {
            existing = {
                id: `p_ext_${Date.now()}_${i}`,
                name: name,
                grade: grade,
                hasTablet: hasTablet,
                schoolYear: schoolYear
            };
            participantsList.push(existing);
            newParticipantIds.add(existing.id); // 新規としてマーク
        } else {
            existing.grade = grade;
            existing.hasTablet = hasTablet;
            existing.schoolYear = schoolYear;
        }

        // 参加希望日に応じて日誌に割り当てる
        Object.keys(appData).forEach(sessionId => {
            const sessionData = sessionsInfo.find(s => s.id === sessionId);
            if (!sessionData) return;
            
            if (rawAttend) {
                // 参加希望日の指定がある場合：マッチする日程だけに追加
                const dateStr = sessionData.date.substring(0, sessionData.date.indexOf('(') > -1 ? sessionData.date.indexOf('(') : sessionData.date.length);
                const wantsToAttend = rawAttend.includes(dateStr);
                
                if (wantsToAttend) {
                    if (!appData[sessionId].participants[existing.id]) {
                        appData[sessionId].participants[existing.id] = {
                            attended: true, rpContent: '', rpScore: '', apContent: '', apScore: '', remarks: ''
                        };
                    }
                } else if (!wantsToAttend && appData[sessionId].participants[existing.id]) {
                    const pData = appData[sessionId].participants[existing.id];
                    if (!pData.rpScore && !pData.apScore && !pData.remarks) {
                        delete appData[sessionId].participants[existing.id];
                    }
                }
            } else {
                // 参加希望日の指定がない場合：全日程に追加
                if (!appData[sessionId].participants[existing.id]) {
                    appData[sessionId].participants[existing.id] = {
                        attended: true, rpContent: '', rpScore: '', apContent: '', apScore: '', remarks: ''
                    };
                }
            }
        });
    });
}

// データの保存（クラウド＆ローカル保存）
// ★ フォーム+iframe送信方式でCORSを完全回避
async function saveData() {
    if (!currentSessionId) return;

    appData[currentSessionId].generalHomework = document.getElementById('generalHomework').value;
    appData[currentSessionId].generalNotes = document.getElementById('generalNotes').value;

    if (appData[currentSessionId] && appData[currentSessionId].participants) {
        Object.keys(appData[currentSessionId].participants).forEach(id => {
            const attendElem = document.getElementById(`attend_${id}`);
            if (attendElem) {
                const attended = attendElem.checked;
                const rpContent = document.getElementById(`rpContent_${id}`).value;
                const rpScore = document.getElementById(`rpScore_${id}`).value;
                const apContent = document.getElementById(`apContent_${id}`).value;
                const apScore = document.getElementById(`apScore_${id}`).value;
                const remarks = document.getElementById(`remarks_${id}`).value;

                appData[currentSessionId].participants[id] = {
                    attended, rpContent, rpScore, apContent, apScore, remarks
                };
            }
        });
    }

    // オフライン動作用の一時保存
    localStorage.setItem('eikenClassManagerData', JSON.stringify(appData));
    localStorage.setItem('eikenClassManagerParticipants', JSON.stringify(participantsList));
    
    // 保存メッセージの表示
    const status = document.getElementById('saveStatus');
    status.textContent = 'クラウドへ保存中...';
    status.classList.add('show');
    
    try {
        await saveToCloud(appData);
        status.textContent = 'クラウドへ保存完了 ✓';
        setTimeout(() => { status.classList.remove('show'); }, 3000);
    } catch (e) {
        console.error("クラウド保存エラー:", e);
        status.textContent = '※オフライン保存のみ完了';
        setTimeout(() => { status.classList.remove('show'); }, 3000);
    }
}

// フォーム＋隠しiframeでGASにPOST送信（CORS完全回避）
function saveToCloud(data) {
    return new Promise((resolve) => {
        // 既存のiframe/formがあれば削除
        const oldIframe = document.getElementById('gas_save_frame');
        if (oldIframe) oldIframe.remove();
        const oldForm = document.getElementById('gas_save_form');
        if (oldForm) oldForm.remove();
        
        // 隠しiframeを作成
        const iframe = document.createElement('iframe');
        iframe.id = 'gas_save_frame';
        iframe.name = 'gas_save_frame';
        iframe.style.display = 'none';
        document.body.appendChild(iframe);
        
        // フォームを作成
        const form = document.createElement('form');
        form.id = 'gas_save_form';
        form.method = 'POST';
        form.action = GAS_WEB_APP_URL;
        form.target = 'gas_save_frame'; // iframeに送信
        
        // データをhiddenフィールドとして追加
        const input = document.createElement('input');
        input.type = 'hidden';
        input.name = 'data';
        input.value = JSON.stringify(data);
        form.appendChild(input);
        
        document.body.appendChild(form);
        form.submit();
        
        // 送信完了後にクリーンアップ
        setTimeout(() => {
            iframe.remove();
            form.remove();
            resolve();
        }, 3000);
    });
}

// サイドバーの描画
function renderSidebar() {
    const list = document.getElementById('dateList');
    list.innerHTML = '';

    sessionsInfo.forEach(session => {
        const li = document.createElement('li');
        li.onclick = () => selectSession(session.id);
        
        // アクティブ状態の保持
        if (currentSessionId === session.id) {
            li.className = 'active';
        }

        li.innerHTML = `
            <div>
                <span class="date-title">${session.date}</span>
            </div>
        `;
        list.appendChild(li);
    });
}

// セッション（日程）の選択
async function selectSession(sessionId) {
    await saveData(); // 今の画面を保存してから遷移
    currentSessionId = sessionId;
    
    const sessionInfo = sessionsInfo.find(s => s.id === sessionId);
    document.getElementById('currentDateTitle').textContent = sessionInfo.date;
    document.getElementById('currentSessionInfo').textContent = '';

    document.getElementById('emptyState').style.display = 'none';
    document.getElementById('mainScrollable').style.display = 'block';

    renderSidebar(); // アクティブハイライトの更新
    renderMainContent();
}

// サイドバーの開閉トグル
function toggleSidebar() {
    const container = document.querySelector('.app-container');
    container.classList.toggle('sidebar-hidden');
}

// 参加者情報の直接編集
function updateParticipantInfo(id, field, value) {
    const p = participantsList.find(x => x.id === id);
    if (p) {
        if (field === 'hasTablet') {
            p[field] = value === 'true';
        } else {
            p[field] = value;
        }
        localStorage.setItem('eikenClassManagerParticipants', JSON.stringify(participantsList));
        
        // 受験級が変更された場合はプルダウンを再描画する
        if (field === 'grade') {
            saveData();
            renderMainContent();
        }
    }
}

// 当日飛び入り参加者の手動追加
function addParticipant() {
    if (!currentSessionId) {
        alert('参加者を追加したい日程をサイドバーから選択してください。');
        return;
    }
    
    saveData(); // 現在の内容を一時保存
    
    const newId = 'p_manual_' + Date.now();
    participantsList.push({
        id: newId,
        name: '新規参加',
        grade: '3級',
        hasTablet: true,
        schoolYear: ''
    });
    
    // 現在選択中の日程だけに参加者を追加
    if (!appData[currentSessionId]) {
        appData[currentSessionId] = { generalHomework: '', generalNotes: '', participants: {} };
    }
    appData[currentSessionId].participants[newId] = {
        attended: true,
        rpContent: '',
        rpScore: '',
        apContent: '',
        apScore: '',
        remarks: ''
    };
    
    localStorage.setItem('eikenClassManagerParticipants', JSON.stringify(participantsList));
    localStorage.setItem('eikenClassManagerData', JSON.stringify(appData));
    
    renderMainContent();
    
    // 追加した参加者の名前入力欄にフォーカスを当てる（少し遅延させる）
    setTimeout(() => {
        const trs = document.getElementById('participantTableBody').querySelectorAll('tr');
        if (trs.length > 0) {
            const lastTr = trs[trs.length - 1];
            const nameInput = lastTr.querySelector('input[type="text"]');
            if (nameInput) nameInput.select();
        }
    }, 100);
}

// モーダル管理用
let participantToDelete = null;

// 参加者の削除（モーダル表示：この日 or 全日程の選択）
function deleteParticipant(id) {
    const p = participantsList.find(x => x.id === id);
    if (!p) return;
    
    participantToDelete = id;
    document.getElementById('deleteModalMessage').innerHTML = `<strong>${p.name}</strong> さんをどの範囲で削除しますか？`;
    document.getElementById('deleteModal').style.display = 'flex';
}

// モーダル閉じる
function closeDeleteModal() {
    document.getElementById('deleteModal').style.display = 'none';
    participantToDelete = null;
}

// この日のみ削除
function deleteFromCurrentSession() {
    if (!participantToDelete || !currentSessionId) return;
    const id = participantToDelete;

    // 現在の日程からのみ削除
    if (appData[currentSessionId] && appData[currentSessionId].participants[id]) {
        delete appData[currentSessionId].participants[id];
    }
    
    localStorage.setItem('eikenClassManagerData', JSON.stringify(appData));
    closeDeleteModal();
    renderMainContent();
}

// 全日程から削除
function deleteFromAllSessions() {
    if (!participantToDelete) return;
    const id = participantToDelete;

    // 参加者リストから完全削除
    participantsList = participantsList.filter(x => x.id !== id);
    
    // 全日程のデータから削除
    Object.keys(appData).forEach(sessionId => {
        if (appData[sessionId].participants[id]) {
            delete appData[sessionId].participants[id];
        }
    });
    
    localStorage.setItem('eikenClassManagerParticipants', JSON.stringify(participantsList));
    localStorage.setItem('eikenClassManagerData', JSON.stringify(appData));
    closeDeleteModal();
    renderMainContent();
}

// メインコンテンツ（参加者リストやフォーム）の描画
function renderMainContent() {
    const data = appData[currentSessionId];
    
    // 全体記録の復元
    document.getElementById('generalHomework').value = data.generalHomework || '';
    document.getElementById('generalNotes').value = data.generalNotes || '';

    // 参加者テーブルの描画
    const tbody = document.getElementById('participantTableBody');
    tbody.innerHTML = '';
    
    let renderedCount = 0;

    participantsList.forEach(p => {
        const pData = data.participants[p.id];
        if (!pData) return; // その日に割り当てられていない生徒は表示しない
        
        renderedCount++;
        const tr = document.createElement('tr');
        
        tr.innerHTML = `
            <td>
                <label class="toggle-switch">
                    <input type="checkbox" id="attend_${p.id}" ${pData.attended ? 'checked' : ''} onchange="toggleAttendance('${p.id}')">
                    <span class="slider"></span>
                </label>
            </td>
            <td>
                <input type="text" value="${p.name}" onchange="updateParticipantInfo('${p.id}', 'name', this.value)" style="width: 90px; padding: 4px; border: 1px solid var(--border-color); border-radius: 4px; font-weight: bold;">
                <button type="button" onclick="showSchedulePopup('${p.id}')" style="background:none;border:none;cursor:pointer;font-size:14px;padding:2px;vertical-align:middle;" title="参加日程を確認">📅</button>
                ${newParticipantIds.has(p.id) ? '<span class="badge-new">NEW</span>' : ''}
            </td>
            <td>
                <input type="text" value="${p.schoolYear || ''}" onchange="updateParticipantInfo('${p.id}', 'schoolYear', this.value)" style="width: 55px; text-align: center; border: 1px solid var(--border-color); border-radius: 4px; padding: 4px; font-size: 0.9em;">
            </td>
            <td>
                <select onchange="updateParticipantInfo('${p.id}', 'grade', this.value)" style="width: 75px; padding: 4px; border: 1px solid var(--border-color); border-radius: 4px;">
                    ${['1級', '準1級', '2級', '準2級プラス', '準2級', '3級', '4級', '5級'].map(g => `<option value="${g}" ${p.grade === g ? 'selected' : ''}>${g}</option>`).join('')}
                </select>
            </td>
            <td>
                <select onchange="updateParticipantInfo('${p.id}', 'hasTablet', this.value)" style="width: 60px; padding: 4px; border: 1px solid var(--border-color); border-radius: 4px; color: ${p.hasTablet ? 'var(--primary-color)' : 'var(--text-color)'}; font-weight: ${p.hasTablet ? 'bold' : 'normal'};">
                    <option value="true" ${p.hasTablet ? 'selected' : ''}>持参</option>
                    <option value="false" ${!p.hasTablet ? 'selected' : ''}>なし</option>
                </select>
            </td>
            <td>
                <select id="rpContent_${p.id}" ${!pData.attended ? 'disabled' : ''} style="width: 210px; font-size: 13px; padding: 4px; border: 1px solid var(--border-color); border-radius: 4px;">
                    ${getPastPaperOptions(p.grade).replace(`value="${pData.rpContent || ''}"`, `value="${pData.rpContent || ''}" selected`)}
                </select>
            </td>
            <td>
                <select id="rpScore_${p.id}" ${!pData.attended ? 'disabled' : ''} style="width: 60px; padding: 4px; border: 1px solid var(--border-color); border-radius: 4px; text-align: center;">
                    <option value="" ${!pData.rpScore ? 'selected' : ''}>-</option>
                    <option value="5" ${pData.rpScore === '5' ? 'selected' : ''}>5</option>
                    <option value="4" ${pData.rpScore === '4' ? 'selected' : ''}>4</option>
                    <option value="3" ${pData.rpScore === '3' ? 'selected' : ''}>3</option>
                    <option value="2" ${pData.rpScore === '2' ? 'selected' : ''}>2</option>
                    <option value="1" ${pData.rpScore === '1' ? 'selected' : ''}>1</option>
                </select>
            </td>
            <td>
                <select id="apContent_${p.id}" ${!pData.attended ? 'disabled' : ''} style="width: 210px; font-size: 13px; padding: 4px; border: 1px solid var(--border-color); border-radius: 4px;">
                    ${getPastPaperOptions(p.grade).replace(`value="${pData.apContent || ''}"`, `value="${pData.apContent || ''}" selected`)}
                </select>
            </td>
            <td>
                <select id="apScore_${p.id}" ${!pData.attended ? 'disabled' : ''} style="width: 60px; padding: 4px; border: 1px solid var(--border-color); border-radius: 4px; text-align: center;">
                    <option value="" ${!pData.apScore ? 'selected' : ''}>-</option>
                    <option value="5" ${pData.apScore === '5' ? 'selected' : ''}>5</option>
                    <option value="4" ${pData.apScore === '4' ? 'selected' : ''}>4</option>
                    <option value="3" ${pData.apScore === '3' ? 'selected' : ''}>3</option>
                    <option value="2" ${pData.apScore === '2' ? 'selected' : ''}>2</option>
                    <option value="1" ${pData.apScore === '1' ? 'selected' : ''}>1</option>
                </select>
            </td>
            <td>
                <input type="text" id="remarks_${p.id}" value="${pData.remarks || ''}" placeholder="特記事項・個別宿題" ${!pData.attended ? 'disabled' : ''}>
            </td>
            <td>
                <button type="button" onclick="deleteParticipant('${p.id}')" style="background: none; border: none; font-size: 20px; cursor: pointer; opacity: 0.7; padding: 8px;" title="この参加者を削除">🗑️</button>
            </td>
        `;
        
        if (!pData.attended) {
            tr.style.opacity = '0.5';
        }
        
        tbody.appendChild(tr);
    });
    
    document.getElementById('participantCount').textContent = `${renderedCount}名`;
}

// 欠席時の入力制限
function toggleAttendance(id) {
    const isChecked = document.getElementById(`attend_${id}`).checked;
    
    // UIの切り替え
    const tr = document.getElementById(`attend_${id}`).closest('tr');
    if (isChecked) {
        tr.style.opacity = '1';
    } else {
        tr.style.opacity = '0.5';
    }
    
    document.getElementById(`rpContent_${id}`).disabled = !isChecked;
    document.getElementById(`rpScore_${id}`).disabled = !isChecked;
    document.getElementById(`apContent_${id}`).disabled = !isChecked;
    document.getElementById(`apScore_${id}`).disabled = !isChecked;
    document.getElementById(`remarks_${id}`).disabled = !isChecked;
}
// ====== 通知バナー ======
function showNotification(message) {
    const banner = document.getElementById('notificationBanner');
    const msgEl = document.getElementById('notifMessage');
    if (banner && msgEl) {
        msgEl.textContent = message;
        banner.classList.add('show');
    }
}

function dismissNotification() {
    const banner = document.getElementById('notificationBanner');
    if (banner) {
        banner.classList.remove('show');
    }
    // Newバッジもクリア（永続化も解除）
    newParticipantIds.clear();
    localStorage.removeItem('eikenNewParticipantIds');
    if (currentSessionId) {
        renderMainContent();
    }
}

// ====== 参加日程ポップアップ ======
function showSchedulePopup(participantId) {
    const p = participantsList.find(x => x.id === participantId);
    if (!p) return;
    
    document.getElementById('schedulePopupName').textContent = p.name;
    
    const listEl = document.getElementById('schedulePopupList');
    listEl.innerHTML = '';
    
    // 今日の日付を取得（月日で比較用）
    const today = new Date();
    const todayStr = `${today.getMonth() + 1}月${today.getDate()}日`;
    
    sessionsInfo.forEach(session => {
        const isRegistered = appData[session.id] && appData[session.id].participants[participantId];
        
        // 日付文字列から月日を抽出して過去・未来を判定
        const dateMatch = session.date.match(/(\d+)月(\d+)日/);
        let isPast = false;
        if (dateMatch) {
            const sessionDate = new Date(today.getFullYear(), parseInt(dateMatch[1]) - 1, parseInt(dateMatch[2]));
            const todayMidnight = new Date(today.getFullYear(), today.getMonth(), today.getDate());
            isPast = sessionDate < todayMidnight;
        }
        
        const li = document.createElement('div');
        li.className = 'schedule-item' + (isRegistered ? ' registered' : ' not-registered') + (isPast ? ' past' : '');
        li.innerHTML = `
            <span class="schedule-status">${isRegistered ? '✅' : '—'}</span>
            <span class="schedule-date">${session.date}</span>
            ${isPast ? '<span class="schedule-past-label">済</span>' : ''}
        `;
        listEl.appendChild(li);
    });
    
    document.getElementById('schedulePopup').style.display = 'flex';
}

function closeSchedulePopup() {
    document.getElementById('schedulePopup').style.display = 'none';
}

// ====== 日程追加 ======
function showAddSessionModal() {
    document.getElementById('newSessionDate').value = '';
    document.getElementById('addSessionModal').style.display = 'flex';
}

function closeAddSessionModal() {
    document.getElementById('addSessionModal').style.display = 'none';
}

function addSession() {
    const dateValue = document.getElementById('newSessionDate').value;
    
    if (!dateValue) {
        alert('日付を選択してください。');
        return;
    }
    
    // yyyy-mm-dd → "X月Y日(曜日)" に変換
    const d = new Date(dateValue + 'T00:00:00');
    const dayNames = ['日', '月', '火', '水', '木', '金', '土'];
    const dateStr = `${d.getMonth() + 1}月${d.getDate()}日(${dayNames[d.getDay()]})`;
    
    const newId = 'day_' + Date.now();
    const newSession = {
        id: newId,
        date: dateStr,
        title: ''
    };
    
    sessionsInfo.push(newSession);
    appData[newId] = { generalHomework: '', generalNotes: '', participants: {} };
    
    // 保存
    localStorage.setItem('eikenClassManagerSessions', JSON.stringify(sessionsInfo));
    localStorage.setItem('eikenClassManagerData', JSON.stringify(appData));
    
    closeAddSessionModal();
    renderSidebar();
}

// 起動
init();
