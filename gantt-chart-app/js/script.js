document.addEventListener('DOMContentLoaded', () => {

    // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
    // ★ ステップ2でコピーした、あなたのウェブアプリのURLに書き換えてください ★
    const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwL8qS8XkIBjeKLKU5pqBTia8RTI-Ap5Ofu_yyN37GwDtY-SgxjhqKdkL-A8U3sJSuDYQ/exec";
    // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

    // --- DOM要素の取得 ---
    const loadingOverlay = document.getElementById('loading-overlay');
    const taskListContainer = document.getElementById('task-list');
    const chartHeaderContainer = document.getElementById('chart-header-container');
    const chartGridContainer = document.getElementById('chart-grid');
    const startYearInput = document.getElementById('start-year');
    const startMonthInput = document.getElementById('start-month');
    const durationMonthsInput = document.getElementById('duration-months');
    const updateTimelineBtn = document.getElementById('update-timeline-btn');
    
    // モーダル関連
    const addTaskBtn = document.getElementById('addTaskBtn');
    const addTaskModal = document.getElementById('add-task-modal');
    const addTaskForm = document.getElementById('add-task-form');
    const addCancelBtn = document.getElementById('add-cancel-btn');

    const editTaskModal = document.getElementById('edit-task-modal');
    const editTaskForm = document.getElementById('edit-task-form');
    const editTaskId = document.getElementById('edit-task-id');
    const deleteTaskBtn = document.getElementById('delete-task-btn');
    const editCancelBtn = document.getElementById('edit-cancel-btn');

    const settingsBtn = document.getElementById('settings-btn');
    const settingsModal = document.getElementById('settings-modal');
    const settingsForm = document.getElementById('settings-form');
    const teamsWebhookUrlInput = document.getElementById('teams-webhook-url');
    const settingsCancelBtn = document.getElementById('settings-cancel-btn');

    // --- アプリケーションの状態 ---
    let timelineMonths = [];
    let tasks = []; // 初期データは空にして、APIから取得する

    // --- ローディング表示の制御 ---
    const showLoading = () => loadingOverlay.classList.remove('hidden');
    const hideLoading = () => loadingOverlay.classList.add('hidden');

    /** タイムラインのヘッダーとグリッドの列を生成 */
    const generateTimeline = () => {
        // ... (この関数の中身は変更なし)
        const startYear = parseInt(startYearInput.value);
        const startMonth = parseInt(startMonthInput.value);
        const duration = parseInt(durationMonthsInput.value);
        timelineMonths = [];
        chartHeaderContainer.innerHTML = '';
        for (let i = 0; i < duration; i++) {
            const currentDate = new Date(startYear, startMonth - 1 + i);
            const year = currentDate.getFullYear();
            const month = currentDate.getMonth() + 1;
            const monthString = `${year}-${String(month).padStart(2, '0')}`;
            timelineMonths.push(monthString);
            const headerCell = document.createElement('div');
            headerCell.textContent = `${year}/${month}`;
            chartHeaderContainer.appendChild(headerCell);
        }
        const gridColumnCount = timelineMonths.length;
        const columnStyle = `repeat(${gridColumnCount}, 100px)`;
        const backgroundSizeStyle = '100px 100%';
        const backgroundStyle = `linear-gradient(to right, #1f4068 1px, transparent 1px)`;
        chartHeaderContainer.style.gridTemplateColumns = columnStyle;
        chartHeaderContainer.style.backgroundImage = backgroundStyle;
        chartHeaderContainer.style.backgroundSize = backgroundSizeStyle;
        chartGridContainer.style.gridTemplateColumns = columnStyle;
        chartGridContainer.style.backgroundImage = backgroundStyle;
        chartGridContainer.style.backgroundSize = backgroundSizeStyle;
    };
    
    /** ガントチャートのタスク部分を描画 */
    const renderTasks = () => {
        // ... (この関数の中身は変更なし)
        taskListContainer.innerHTML = '';
        chartGridContainer.innerHTML = '';
        if (timelineMonths.length === 0) return;
        const timelineStart = new Date(timelineMonths[0] + '-01');
        const timelineEnd = new Date(timelineMonths[timelineMonths.length - 1] + '-01');
        tasks.forEach((task, index) => {
            const taskElement = document.createElement('div');
            taskElement.className = 'task-row';
            taskElement.textContent = task.name;
            taskListContainer.appendChild(taskElement);
            const taskStart = new Date(task.start + '-01');
            const taskEnd = new Date(task.end + '-01');
            if (taskEnd < timelineStart || taskStart > timelineEnd) return;
            const visibleStart = new Date(Math.max(taskStart, timelineStart));
            const visibleEnd = new Date(Math.min(taskEnd, timelineEnd));
            const visibleStartMonthString = `${visibleStart.getFullYear()}-${String(visibleStart.getMonth() + 1).padStart(2, '0')}`;
            const visibleEndMonthString = `${visibleEnd.getFullYear()}-${String(visibleEnd.getMonth() + 1).padStart(2, '0')}`;
            const startCol = timelineMonths.indexOf(visibleStartMonthString) + 1;
            const endCol = timelineMonths.indexOf(visibleEndMonthString) + 1;
            if (startCol > 0 && endCol > 0 && endCol >= startCol) {
                const taskBarElement = document.createElement('div');
                taskBarElement.className = 'task-bar';
                taskBarElement.textContent = task.name;
                taskBarElement.addEventListener('click', () => openEditModal(task.id));
                taskBarElement.style.setProperty('--start', startCol);
                taskBarElement.style.setProperty('--end', endCol + 1);
                taskBarElement.style.setProperty('--row', index + 1);
                const progressElement = document.createElement('div');
                progressElement.className = 'progress';
                progressElement.style.setProperty('--progress', `${task.progress}%`);
                taskBarElement.appendChild(progressElement);
                chartGridContainer.appendChild(taskBarElement);
            }
        });
    };

    /** 編集モーダルを開く関数 */
    const openEditModal = (taskId) => {
        const task = tasks.find(t => t.id == taskId); // IDの型が異なる可能性があるので `==` を使用
        if (!task) return;
        editTaskId.value = task.id;
        document.getElementById('edit-task-name').value = task.name;
        const [startY, startM] = task.start.split('-');
        const [endY, endM] = task.end.split('-');
        document.getElementById('edit-task-start-year').value = startY;
        document.getElementById('edit-task-start-month').value = startM;
        document.getElementById('edit-task-end-year').value = endY;
        document.getElementById('edit-task-end-month').value = endM;
        document.getElementById('edit-task-progress').value = task.progress;
        editTaskModal.classList.remove('hidden');
    };

    /** Teamsへ通知を送信する関数 */
    const sendTeamsNotification = async (task) => {
        // ... (この関数の中身は変更なし)
        const webhookUrl = localStorage.getItem('teamsWebhookUrl');
        if (!webhookUrl) {
            console.log("Teams Webhook URLが設定されていません。");
            return;
        }
        const card = {
            "@type": "MessageCard", "@context": "http://schema.org/extensions", "themeColor": "0076D7",
            "summary": `新しいタスクが追加されました: ${task.name}`,
            "sections": [{"activityTitle": "新しいタスクが追加されました！", "activitySubtitle": "ガントチャートアプリより",
                "facts": [{"name": "タスク名", "value": task.name}, {"name": "期間", "value": `${task.start} 〜 ${task.end}`}, {"name": "進捗", "value": `${task.progress}%`}],
                "markdown": true
            }]
        };
        try {
            await fetch(webhookUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(card), mode: 'no-cors' });
            console.log("Teamsへの通知を試みました。");
        } catch (error) {
            console.error('Teamsへの通知送信中にエラーが発生しました:', error);
            alert('Teamsへの通知送信中にエラーが発生しました。');
        }
    };

    /** アプリ全体を初期化・再描画する */
    const initializeApp = async () => {
        showLoading();
        generateTimeline();
        try {
            const response = await fetch(SCRIPT_URL);
            tasks = await response.json();
            renderTasks();
        } catch (error) {
            console.error("データの取得に失敗しました:", error);
            alert("データの取得に失敗しました。");
        } finally {
            hideLoading();
        }
    };

    // --- イベントリスナー ---
    updateTimelineBtn.addEventListener('click', initializeApp);

    // モーダル開閉
    addTaskBtn.addEventListener('click', () => addTaskModal.classList.remove('hidden'));
    addCancelBtn.addEventListener('click', () => addTaskModal.classList.add('hidden'));
    addTaskModal.addEventListener('click', (e) => { if (e.target === addTaskModal) addTaskModal.classList.add('hidden'); });
    editCancelBtn.addEventListener('click', () => editTaskModal.classList.add('hidden'));
    editTaskModal.addEventListener('click', (e) => { if (e.target === editTaskModal) editTaskModal.classList.add('hidden'); });
    settingsBtn.addEventListener('click', () => {
        teamsWebhookUrlInput.value = localStorage.getItem('teamsWebhookUrl') || '';
        settingsModal.classList.remove('hidden');
    });
    settingsCancelBtn.addEventListener('click', () => settingsModal.classList.add('hidden'));
    settingsModal.addEventListener('click', (e) => { if (e.target === settingsModal) settingsModal.classList.add('hidden'); });

    // フォーム送信
    addTaskForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        showLoading();
        const newTaskData = {
            name: document.getElementById('task-name').value,
            start: `${document.getElementById('task-start-year').value}-${String(document.getElementById('task-start-month').value).padStart(2, '0')}`,
            end: `${document.getElementById('task-end-year').value}-${String(document.getElementById('task-end-month').value).padStart(2, '0')}`,
            progress: parseInt(document.getElementById('task-progress').value)
        };
        try {
            const response = await fetch(SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'add', task: newTaskData }) });
            const result = await response.json();
            if (result.success) {
                await sendTeamsNotification(result.task);
                await initializeApp();
            } else {
                throw new Error(result.message);
            }
        } catch (error) {
            console.error("タスクの追加に失敗しました:", error);
            alert("タスクの追加に失敗しました。");
        } finally {
            addTaskModal.classList.add('hidden');
            addTaskForm.reset();
            hideLoading();
        }
    });

    editTaskForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        showLoading();
        const updatedTaskData = {
            id: document.getElementById('edit-task-id').value,
            name: document.getElementById('edit-task-name').value,
            start: `${document.getElementById('edit-task-start-year').value}-${String(document.getElementById('edit-task-start-month').value).padStart(2, '0')}`,
            end: `${document.getElementById('edit-task-end-year').value}-${String(document.getElementById('edit-task-end-month').value).padStart(2, '0')}`,
            progress: parseInt(document.getElementById('edit-task-progress').value)
        };
        try {
            const response = await fetch(SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'update', task: updatedTaskData }) });
            const result = await response.json();
            if (result.success) {
                await initializeApp();
            } else {
                throw new Error(result.message);
            }
        } catch (error) {
            console.error("タスクの更新に失敗しました:", error);
            alert("タスクの更新に失敗しました。");
        } finally {
            editTaskModal.classList.add('hidden');
            hideLoading();
        }
    });
    
    deleteTaskBtn.addEventListener('click', async () => {
        const taskId = document.getElementById('edit-task-id').value;
        if (!confirm(`このタスクを本当に削除しますか？`)) return;
        showLoading();
        try {
            const response = await fetch(SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'delete', id: taskId }) });
            const result = await response.json();
            if (result.success) {
                await initializeApp();
            } else {
                throw new Error(result.message);
            }
        } catch (error) {
            console.error("タスクの削除に失敗しました:", error);
            alert("タスクの削除に失敗しました。");
        } finally {
            editTaskModal.classList.add('hidden');
            hideLoading();
        }
    });

    settingsForm.addEventListener('submit', (e) => {
        e.preventDefault();
        localStorage.setItem('teamsWebhookUrl', teamsWebhookUrlInput.value);
        settingsModal.classList.add('hidden');
        alert('設定を保存しました。');
    });

    // --- 初期化実行 ---
    initializeApp();

    // --- ウィンドウリサイズ時の処理 ---
    let resizeTimer;
    window.addEventListener('resize', () => {
        clearTimeout(resizeTimer);
        resizeTimer = setTimeout(initializeApp, 250);
    });
});
