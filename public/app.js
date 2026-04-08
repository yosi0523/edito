// ===== Firebase 설정 =====
const firebaseConfig = {
    apiKey: "AIzaSyDLry0f9aE24wWfa8_CBJuJAxEwYzNaWyU",
    authDomain: "edito-f8256.firebaseapp.com",
    projectId: "edito-f8256",
    storageBucket: "edito-f8256.firebasestorage.app",
    messagingSenderId: "704063878825",
    appId: "1:704063878825:web:8976ea2f8eb6661aeb5bb8"
};

firebase.initializeApp(firebaseConfig);
const auth = firebase.auth();
const db = firebase.firestore();
const storage = firebase.storage();

// ===== 상태 관리 =====
let currentUser = null;
let currentYear = new Date().getFullYear();
let currentMonth = new Date().getMonth();
let events = [];
let memos = [];
let checklists = [];
let ddays = [];
let selectedDate = null;
let selectedColor = '#4A90D9';
let selectedDdayColor = '#E74C3C';
let pendingPhotos = [];
let existingPhotos = [];
let currentTab = 'calendar';
let currentView = 'month';
let isDragging = false;
let memoPinned = false;

// ===== DOM 요소 =====
const $ = (id) => document.getElementById(id);
const $$ = (sel) => document.querySelectorAll(sel);

// ===== 인증 상태 감시 =====
auth.onAuthStateChanged((user) => {
    if ($('loading-screen')) $('loading-screen').classList.add('hidden');
    if (user) {
        currentUser = user;
        if ($('login-screen')) $('login-screen').classList.add('hidden');
        if ($('main-screen')) $('main-screen').classList.remove('hidden');
        if ($('user-avatar')) $('user-avatar').src = user.photoURL || '';
        $$('.u-avatar').forEach(img => img.src = user.photoURL || '');
        if ($('user-name')) $('user-name').textContent = user.displayName || user.email;
        loadEvents();
        loadMemos();
        loadChecklists();
        loadDdays();
        renderCalendar();
    } else {
        currentUser = null;
        if ($('login-screen')) $('login-screen').classList.remove('hidden');
        if ($('main-screen')) $('main-screen').classList.add('hidden');
    }
});

// ===== 로그인 =====
$('google-login-btn').addEventListener('click', () => {
    const provider = new firebase.auth.GoogleAuthProvider();
    auth.signInWithPopup(provider).catch((error) => {
        alert('로그인 실패: ' + error.message);
    });
});

// ===== 로그아웃 =====
$('logout-btn').addEventListener('click', () => {
    auth.signOut();
});

// ===== 유저 메뉴 =====
$('user-menu-btn').addEventListener('click', (e) => {
    e.stopPropagation();
    $('user-menu').classList.toggle('hidden');
});

$$('.umenu-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
        e.stopPropagation();
        $('user-menu').classList.toggle('hidden');
    });
});

document.addEventListener('click', () => {
    $('user-menu').classList.add('hidden');
});

// ===== Firestore 데이터 로드 =====
function loadEvents() {
    if (!currentUser) return;
    db.collection('users').doc(currentUser.uid).collection('events')
        .onSnapshot((snapshot) => {
            events = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
            renderCalendar();
            if (selectedDate) renderDayDetail(selectedDate);
            renderTodayWidget();
        });
}

function loadMemos() {
    if (!currentUser) return;
    db.collection('users').doc(currentUser.uid).collection('memos')
        .orderBy('updatedAt', 'desc')
        .onSnapshot((snapshot) => {
            memos = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
            renderMemoList();
        });
}

function loadChecklists() {
    if (!currentUser) return;
    db.collection('users').doc(currentUser.uid).collection('checklists')
        .orderBy('updatedAt', 'desc')
        .onSnapshot((snapshot) => {
            checklists = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
            renderChecklistList();
        });
}

function loadDdays() {
    if (!currentUser) return;
    db.collection('users').doc(currentUser.uid).collection('ddays')
        .orderBy('date', 'asc')
        .onSnapshot((snapshot) => {
            ddays = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
            renderDdayList();
            renderTodayWidget();
        });
}

// ===== 월별 헤더 컬러 =====
const monthColors = [
    '#E74C3C', // 1월 - 빨강
    '#E91E63', // 2월 - 핑크
    '#9B59B6', // 3월 - 보라
    '#4A90D9', // 4월 - 파랑
    '#1ABC9C', // 5월 - 민트
    '#2ECC71', // 6월 - 초록
    '#27AE60', // 7월 - 진초록
    '#F39C12', // 8월 - 주황
    '#E67E22', // 9월 - 오렌지
    '#D35400', // 10월 - 갈색
    '#8E44AD', // 11월 - 자주
    '#2C3E50', // 12월 - 남색
];

function applyMonthColor() {
    const color = monthColors[currentMonth];
    document.documentElement.style.setProperty('--header-color', color);
}

// ===== 캘린더 렌더링 =====
function renderCalendar() {
    const grid = $('calendar-grid');
    if (!grid) return;
    grid.innerHTML = '';

    applyMonthColor();

    const monthNames = ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'];
    if ($('current-title')) $('current-title').textContent = `${currentYear}년 ${monthNames[currentMonth]}`;

    const firstDay = new Date(currentYear, currentMonth, 1).getDay();
    const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
    const daysInPrevMonth = new Date(currentYear, currentMonth, 0).getDate();

    const today = new Date();
    const todayStr = `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`;

    // 이전 달
    for (let i = firstDay - 1; i >= 0; i--) {
        const day = daysInPrevMonth - i;
        const date = new Date(currentYear, currentMonth - 1, day);
        const dateStr = formatDate(date);
        grid.appendChild(createCalendarCell(day, dateStr, true));
    }

    // 이번 달
    for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = `${currentYear}-${String(currentMonth+1).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
        const isToday = dateStr === todayStr;
        grid.appendChild(createCalendarCell(day, dateStr, false, isToday));
    }

    // 다음 달 (6줄 고정)
    const totalCells = grid.children.length;
    const targetCells = 42;
    for (let day = 1; day <= targetCells - totalCells; day++) {
        const date = new Date(currentYear, currentMonth + 1, day);
        const dateStr = formatDate(date);
        grid.appendChild(createCalendarCell(day, dateStr, true));
    }
}

function createCalendarCell(day, dateStr, isOtherMonth, isToday = false) {
    const cell = document.createElement('div');
    cell.className = 'calendar-cell';
    cell.dataset.date = dateStr;
    if (isOtherMonth) cell.classList.add('other-month');
    if (isToday) cell.classList.add('today');

    const dateNum = document.createElement('div');
    dateNum.className = 'date-num';
    dateNum.textContent = day;
    cell.appendChild(dateNum);

    // 해당 날짜의 일정 표시
    const dayEvents = events.filter(e => e.date === dateStr);
    if (dayEvents.length > 0) {
        const dotContainer = document.createElement('div');
        dotContainer.className = 'event-dot-container';

        const maxShow = 3;
        dayEvents.slice(0, maxShow).forEach(event => {
            const dot = document.createElement('div');
            dot.className = 'event-dot';
            dot.style.background = event.color || '#4A90D9';
            dot.textContent = event.title;

            // PC 드래그 앤 드롭
            dot.setAttribute('draggable', 'true');
            dot.addEventListener('mousedown', (e) => {
                e.stopPropagation();
            });
            dot.addEventListener('dragstart', (e) => {
                isDragging = true;
                e.dataTransfer.setData('text/plain', event.id);
                e.dataTransfer.effectAllowed = 'move';
                setTimeout(() => dot.classList.add('dragging'), 0);
            });
            dot.addEventListener('dragend', () => {
                setTimeout(() => { isDragging = false; }, 50);
                dot.classList.remove('dragging');
                $$('.calendar-cell.drag-over').forEach(c => c.classList.remove('drag-over'));
            });

            dotContainer.appendChild(dot);
        });

        if (dayEvents.length > maxShow) {
            const more = document.createElement('div');
            more.className = 'event-more';
            more.textContent = `+${dayEvents.length - maxShow}`;
            dotContainer.appendChild(more);
        }

        cell.appendChild(dotContainer);
    }

    // 드래그 앤 드롭 (드롭 대상)
    cell.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        cell.classList.add('drag-over');
    });
    cell.addEventListener('dragleave', (e) => {
        if (!cell.contains(e.relatedTarget)) {
            cell.classList.remove('drag-over');
        }
    });
    cell.addEventListener('drop', async (e) => {
        e.preventDefault();
        cell.classList.remove('drag-over');
        const eventId = e.dataTransfer.getData('text/plain');
        if (eventId && currentUser) {
            await db.collection('users').doc(currentUser.uid).collection('events')
                .doc(eventId).update({ date: dateStr });
        }
    });

    cell.addEventListener('click', () => {
        if (isDragging) return;
        selectedDate = dateStr;
        renderDayDetail(dateStr);
        $('day-detail').classList.remove('hidden');
    });

    return cell;
}

// ===== 오늘 위젯 =====
function renderTodayWidget() {
    const today = new Date();
    const todayStr = formatDate(today);
    const todayEvents = events.filter(e => e.date === todayStr);

    const widgetEvents = $('widget-events');
    if (widgetEvents) {
        if (todayEvents.length > 0) {
            widgetEvents.innerHTML = todayEvents.slice(0, 3).map(e =>
                `<div style="border-left:3px solid ${e.color || '#4A90D9'}; padding-left:8px; margin:4px 0; font-size:12px;">
                    ${escapeHtml(e.title)} ${e.startTime ? '<span style="color:#888">' + e.startTime + '</span>' : ''}
                </div>`
            ).join('');
        } else {
            widgetEvents.innerHTML = '<div style="font-size:12px;color:#888;padding:4px 0;">오늘 일정 없음</div>';
        }
    }

    const widgetDday = $('widget-dday');
    if (widgetDday && ddays.length > 0) {
        const closest = ddays.filter(d => d.date >= todayStr).sort((a,b) => a.date.localeCompare(b.date))[0];
        if (closest) {
            const diff = Math.ceil((new Date(closest.date) - today) / (1000*60*60*24));
            widgetDday.innerHTML = `<div style="font-size:12px;"><strong>${escapeHtml(closest.title)}</strong><br><span style="color:${closest.color || '#E74C3C'}; font-weight:700;">D-${diff === 0 ? 'Day' : diff}</span></div>`;
        }
    }
}

// ===== 일별 상세 =====
function renderDayDetail(dateStr) {
    const date = new Date(dateStr + 'T00:00:00');
    const dayNames = ['일', '월', '화', '수', '목', '금', '토'];
    $('day-detail-title').textContent = `${date.getMonth()+1}월 ${date.getDate()}일 (${dayNames[date.getDay()]})`;

    const list = $('day-events-list');
    const dayEvents = events.filter(e => e.date === dateStr)
        .sort((a, b) => (a.startTime || '').localeCompare(b.startTime || ''));

    if (dayEvents.length === 0) {
        list.innerHTML = '<div class="no-events">일정이 없습니다</div>';
    } else {
        list.innerHTML = dayEvents.map(event => {
            const timeStr = event.startTime ?
                `${event.startTime}${event.endTime ? ' ~ ' + event.endTime : ''}` : '종일';
            const photosHtml = event.photos && event.photos.length > 0 ?
                `<div class="day-event-photos">${event.photos.slice(0,3).map(p => `<img src="${p}" alt="">`).join('')}</div>` : '';
            return `
                <div class="day-event-item" onclick="openEditEvent('${event.id}')">
                    <div class="day-event-color" style="background:${event.color || '#4A90D9'}"></div>
                    <div class="day-event-info">
                        <div class="day-event-title">${escapeHtml(event.title)}</div>
                        <div class="day-event-time">${timeStr}</div>
                    </div>
                    ${photosHtml}
                </div>
            `;
        }).join('');
    }
}

// ===== 일정 모달 =====
function openNewEvent(dateStr) {
    $('event-modal-title').textContent = '일정 추가';
    $('event-form').reset();
    $('event-id').value = '';
    $('event-date').value = dateStr || formatDate(new Date());
    $('delete-event-btn').classList.add('hidden');
    $('event-photo-preview').innerHTML = '';
    if ($('repeat-end-group')) $('repeat-end-group').classList.add('hidden');
    pendingPhotos = [];
    existingPhotos = [];
    selectedColor = '#4A90D9';
    updateColorPicker();
    $('event-modal').classList.remove('hidden');
}

function openEditEvent(eventId) {
    const event = events.find(e => e.id === eventId);
    if (!event) return;

    $('event-modal-title').textContent = '일정 수정';
    $('event-id').value = event.id;
    $('event-title').value = event.title;
    $('event-date').value = event.date;
    $('event-start-time').value = event.startTime || '';
    $('event-end-time').value = event.endTime || '';
    if ($('event-memo')) $('event-memo').value = event.memo || '';
    if ($('event-repeat')) $('event-repeat').value = event.repeat || 'none';
    if ($('event-alarm')) $('event-alarm').value = event.alarm || 'none';
    if ($('event-repeat-end')) $('event-repeat-end').value = event.repeatEnd || '';
    if ($('repeat-end-group') && event.repeat && event.repeat !== 'none') {
        $('repeat-end-group').classList.remove('hidden');
    }
    selectedColor = event.color || '#4A90D9';
    updateColorPicker();
    $('delete-event-btn').classList.remove('hidden');

    pendingPhotos = [];
    existingPhotos = event.photos || [];
    renderPhotoPreview('event');

    $('event-modal').classList.remove('hidden');
}

// ===== 일정 저장 =====
$('event-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const eventId = $('event-id').value;
    const data = {
        title: $('event-title').value,
        date: $('event-date').value,
        startTime: $('event-start-time').value || null,
        endTime: $('event-end-time').value || null,
        color: selectedColor,
        memo: $('event-memo') ? $('event-memo').value : '',
        repeat: $('event-repeat') ? $('event-repeat').value : 'none',
        alarm: $('event-alarm') ? $('event-alarm').value : 'none',
        repeatEnd: $('event-repeat-end') ? $('event-repeat-end').value : null,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    };

    const photoUrls = [...existingPhotos];
    for (const file of pendingPhotos) {
        const url = await uploadPhoto(file);
        if (url) photoUrls.push(url);
    }
    data.photos = photoUrls;

    const ref = db.collection('users').doc(currentUser.uid).collection('events');
    if (eventId) {
        await ref.doc(eventId).update(data);
    } else {
        data.createdAt = firebase.firestore.FieldValue.serverTimestamp();
        await ref.add(data);
    }

    $('event-modal').classList.add('hidden');
});

// ===== 일정 삭제 =====
$('delete-event-btn').addEventListener('click', async () => {
    if (!confirm('이 일정을 삭제할까요?')) return;
    const eventId = $('event-id').value;
    await db.collection('users').doc(currentUser.uid).collection('events').doc(eventId).delete();
    $('event-modal').classList.add('hidden');
});

// 반복 설정 UI
if ($('event-repeat')) {
    $('event-repeat').addEventListener('change', (e) => {
        if ($('repeat-end-group')) {
            $('repeat-end-group').classList.toggle('hidden', e.target.value === 'none');
        }
    });
}

// ===== 메모 관련 =====
function renderMemoList() {
    const list = $('memo-list');
    if (!list) return;
    if (memos.length === 0) {
        list.innerHTML = '<div class="empty-state"><span>📝</span><span>메모가 없습니다</span></div>';
        return;
    }

    const filter = $('memo-filter') ? $('memo-filter').value : 'all';
    const filtered = filter === 'all' ? memos : memos.filter(m => m.category === filter);

    const sorted = [...filtered].sort((a, b) => {
        if (a.pinned && !b.pinned) return -1;
        if (!a.pinned && b.pinned) return 1;
        return 0;
    });

    list.innerHTML = sorted.map(memo => {
        const date = memo.updatedAt ? new Date(memo.updatedAt.seconds * 1000) : new Date();
        const dateStr = `${date.getFullYear()}.${date.getMonth()+1}.${date.getDate()}`;
        const photosHtml = memo.photos && memo.photos.length > 0 ?
            `<div class="memo-card-photos">${memo.photos.slice(0,4).map(p => `<img src="${p}" alt="">`).join('')}</div>` : '';
        const pinIcon = memo.pinned ? '★ ' : '';
        return `
            <div class="memo-card" onclick="openEditMemo('${memo.id}')">
                <div class="memo-card-title">${pinIcon}${escapeHtml(memo.title)}</div>
                <div class="memo-card-content">${escapeHtml(memo.content || '')}</div>
                ${photosHtml}
                <div class="memo-card-date">${dateStr}</div>
            </div>
        `;
    }).join('');
}

function openNewMemo() {
    $('memo-modal-title').textContent = '메모 추가';
    $('memo-form').reset();
    $('memo-id').value = '';
    if ($('memo-pinned')) $('memo-pinned').value = 'false';
    memoPinned = false;
    updatePinButton();
    $('delete-memo-btn').classList.add('hidden');
    $('memo-photo-preview').innerHTML = '';
    pendingPhotos = [];
    existingPhotos = [];
    $('memo-modal').classList.remove('hidden');
}

function openEditMemo(memoId) {
    const memo = memos.find(m => m.id === memoId);
    if (!memo) return;

    $('memo-modal-title').textContent = '메모 수정';
    $('memo-id').value = memo.id;
    if ($('memo-title')) $('memo-title').value = memo.title;
    if ($('memo-content')) $('memo-content').value = memo.content || '';
    if ($('memo-category')) $('memo-category').value = memo.category || 'default';
    memoPinned = memo.pinned || false;
    if ($('memo-pinned')) $('memo-pinned').value = memoPinned ? 'true' : 'false';
    updatePinButton();
    $('delete-memo-btn').classList.remove('hidden');

    pendingPhotos = [];
    existingPhotos = memo.photos || [];
    renderPhotoPreview('memo');

    $('memo-modal').classList.remove('hidden');
}

function updatePinButton() {
    if ($('pin-icon')) {
        $('pin-icon').textContent = memoPinned ? '★' : '☆';
    }
}

if ($('memo-pin-btn')) {
    $('memo-pin-btn').addEventListener('click', () => {
        memoPinned = !memoPinned;
        if ($('memo-pinned')) $('memo-pinned').value = memoPinned ? 'true' : 'false';
        updatePinButton();
    });
}

$('memo-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const memoId = $('memo-id').value;
    const data = {
        title: $('memo-title').value,
        content: $('memo-content') ? $('memo-content').value : '',
        category: $('memo-category') ? $('memo-category').value : 'default',
        pinned: memoPinned,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    };

    const photoUrls = [...existingPhotos];
    for (const file of pendingPhotos) {
        const url = await uploadPhoto(file);
        if (url) photoUrls.push(url);
    }
    data.photos = photoUrls;

    const ref = db.collection('users').doc(currentUser.uid).collection('memos');
    if (memoId) {
        await ref.doc(memoId).update(data);
    } else {
        data.createdAt = firebase.firestore.FieldValue.serverTimestamp();
        await ref.add(data);
    }

    $('memo-modal').classList.add('hidden');
});

$('delete-memo-btn').addEventListener('click', async () => {
    if (!confirm('이 메모를 삭제할까요?')) return;
    const memoId = $('memo-id').value;
    await db.collection('users').doc(currentUser.uid).collection('memos').doc(memoId).delete();
    $('memo-modal').classList.add('hidden');
});

if ($('memo-filter')) {
    $('memo-filter').addEventListener('change', () => renderMemoList());
}

// ===== 체크리스트 관련 =====
function renderChecklistList() {
    const list = $('checklist-list');
    if (!list) return;
    if (checklists.length === 0) {
        list.innerHTML = '<div class="empty-state"><span>✅</span><span>할일이 없습니다</span></div>';
        return;
    }
    list.innerHTML = checklists.map(cl => {
        const itemsHtml = (cl.items || []).map((item, idx) => `
            <label class="checklist-row" data-cl-id="${cl.id}" data-idx="${idx}">
                <input type="checkbox" ${item.checked ? 'checked' : ''} onchange="toggleChecklistItem('${cl.id}', ${idx}, this.checked)">
                <span class="${item.checked ? 'checked-text' : ''}">${escapeHtml(item.text)}</span>
            </label>
        `).join('');
        return `
            <div class="checklist-card">
                <div class="checklist-card-header" onclick="openEditChecklist('${cl.id}')">
                    <span class="checklist-card-title">${escapeHtml(cl.title)}</span>
                    <span class="checklist-card-edit">수정</span>
                </div>
                <div class="checklist-card-items">${itemsHtml}</div>
            </div>
        `;
    }).join('');
}

async function toggleChecklistItem(clId, idx, checked) {
    const cl = checklists.find(c => c.id === clId);
    if (!cl || !cl.items) return;
    cl.items[idx].checked = checked;
    await db.collection('users').doc(currentUser.uid).collection('checklists')
        .doc(clId).update({ items: cl.items });
}

function openNewChecklist() {
    $('checklist-modal-title').textContent = '할일 추가';
    $('checklist-form').reset();
    $('checklist-id').value = '';
    $('checklist-items').innerHTML = '';
    $('delete-checklist-btn').classList.add('hidden');
    addChecklistItemRow();
    $('checklist-modal').classList.remove('hidden');
}

function openEditChecklist(clId) {
    const cl = checklists.find(c => c.id === clId);
    if (!cl) return;
    $('checklist-modal-title').textContent = '할일 수정';
    $('checklist-id').value = cl.id;
    $('checklist-title').value = cl.title;
    $('checklist-items').innerHTML = '';
    (cl.items || []).forEach(item => addChecklistItemRow(item.text, item.checked));
    $('delete-checklist-btn').classList.remove('hidden');
    $('checklist-modal').classList.remove('hidden');
}

function addChecklistItemRow(text = '', checked = false) {
    const container = $('checklist-items');
    if (!container) return;
    const row = document.createElement('div');
    row.className = 'checklist-item-row';
    row.style.cssText = 'display:flex;gap:8px;margin-bottom:6px;align-items:center;';
    row.innerHTML = `
        <input type="checkbox" ${checked ? 'checked' : ''} style="width:18px;height:18px;">
        <input type="text" value="${escapeHtml(text)}" placeholder="항목 입력" style="flex:1;padding:8px;border:1px solid #E8E8E8;border-radius:6px;font-size:14px;">
        <button type="button" onclick="this.parentElement.remove()" style="border:none;background:none;color:#E74C3C;font-size:18px;cursor:pointer;">&times;</button>
    `;
    container.appendChild(row);
}

if ($('add-checklist-item')) {
    $('add-checklist-item').addEventListener('click', () => addChecklistItemRow());
}

if ($('checklist-form')) {
    $('checklist-form').addEventListener('submit', async (e) => {
        e.preventDefault();
        const clId = $('checklist-id').value;
        const rows = $$('.checklist-item-row');
        const items = Array.from(rows).map(row => ({
            text: row.querySelector('input[type="text"]').value,
            checked: row.querySelector('input[type="checkbox"]').checked
        })).filter(i => i.text.trim());

        const data = {
            title: $('checklist-title').value,
            items: items,
            updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        };

        const ref = db.collection('users').doc(currentUser.uid).collection('checklists');
        if (clId) {
            await ref.doc(clId).update(data);
        } else {
            data.createdAt = firebase.firestore.FieldValue.serverTimestamp();
            await ref.add(data);
        }
        $('checklist-modal').classList.add('hidden');
    });
}

if ($('delete-checklist-btn')) {
    $('delete-checklist-btn').addEventListener('click', async () => {
        if (!confirm('이 할일을 삭제할까요?')) return;
        const clId = $('checklist-id').value;
        await db.collection('users').doc(currentUser.uid).collection('checklists').doc(clId).delete();
        $('checklist-modal').classList.add('hidden');
    });
}

// ===== D-day 관련 =====
function renderDdayList() {
    const list = $('dday-list');
    if (!list) return;
    if (ddays.length === 0) {
        list.innerHTML = '<div class="empty-state"><span>🎯</span><span>D-day가 없습니다</span></div>';
        return;
    }
    const today = new Date();
    today.setHours(0,0,0,0);
    list.innerHTML = ddays.map(dd => {
        const target = new Date(dd.date + 'T00:00:00');
        const diff = Math.ceil((target - today) / (1000*60*60*24));
        const label = diff === 0 ? 'D-Day' : diff > 0 ? `D-${diff}` : `D+${Math.abs(diff)}`;
        return `
            <div class="memo-card" onclick="openEditDday('${dd.id}')">
                <div class="memo-card-title" style="display:flex;justify-content:space-between;">
                    <span>${escapeHtml(dd.title)}</span>
                    <span style="color:${dd.color || '#E74C3C'};font-weight:700;">${label}</span>
                </div>
                <div class="memo-card-date">${dd.date}</div>
            </div>
        `;
    }).join('');
}

function openNewDday() {
    $('dday-modal-title').textContent = 'D-day 추가';
    $('dday-form').reset();
    $('dday-id').value = '';
    $('delete-dday-btn').classList.add('hidden');
    selectedDdayColor = '#E74C3C';
    updateDdayColorPicker();
    $('dday-modal').classList.remove('hidden');
}

function openEditDday(ddId) {
    const dd = ddays.find(d => d.id === ddId);
    if (!dd) return;
    $('dday-modal-title').textContent = 'D-day 수정';
    $('dday-id').value = dd.id;
    $('dday-title').value = dd.title;
    $('dday-date').value = dd.date;
    selectedDdayColor = dd.color || '#E74C3C';
    updateDdayColorPicker();
    $('delete-dday-btn').classList.remove('hidden');
    $('dday-modal').classList.remove('hidden');
}

function updateDdayColorPicker() {
    $$('#dday-color-picker .color-option').forEach(el => {
        el.classList.toggle('selected', el.dataset.color === selectedDdayColor);
    });
}

$$('#dday-color-picker .color-option').forEach(el => {
    el.addEventListener('click', () => {
        selectedDdayColor = el.dataset.color;
        updateDdayColorPicker();
    });
});

if ($('dday-form')) {
    $('dday-form').addEventListener('submit', async (e) => {
        e.preventDefault();
        const ddId = $('dday-id').value;
        const data = {
            title: $('dday-title').value,
            date: $('dday-date').value,
            color: selectedDdayColor,
            updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        };
        const ref = db.collection('users').doc(currentUser.uid).collection('ddays');
        if (ddId) {
            await ref.doc(ddId).update(data);
        } else {
            data.createdAt = firebase.firestore.FieldValue.serverTimestamp();
            await ref.add(data);
        }
        $('dday-modal').classList.add('hidden');
    });
}

if ($('delete-dday-btn')) {
    $('delete-dday-btn').addEventListener('click', async () => {
        if (!confirm('이 D-day를 삭제할까요?')) return;
        const ddId = $('dday-id').value;
        await db.collection('users').doc(currentUser.uid).collection('ddays').doc(ddId).delete();
        $('dday-modal').classList.add('hidden');
    });
}

// ===== 사진 업로드 =====
async function uploadPhoto(file) {
    try {
        const path = `users/${currentUser.uid}/photos/${Date.now()}_${file.name}`;
        const ref = storage.ref(path);
        await ref.put(file);
        return await ref.getDownloadURL();
    } catch (error) {
        console.error('사진 업로드 실패:', error);
        return null;
    }
}

function renderPhotoPreview(type) {
    const preview = $(`${type}-photo-preview`);
    if (!preview) return;
    preview.innerHTML = '';

    existingPhotos.forEach((url, index) => {
        const item = document.createElement('div');
        item.className = 'photo-preview-item';
        item.innerHTML = `
            <img src="${url}" alt="">
            <button type="button" class="remove-photo" onclick="removeExistingPhoto(${index}, '${type}')">&times;</button>
        `;
        preview.appendChild(item);
    });

    pendingPhotos.forEach((file, index) => {
        const item = document.createElement('div');
        item.className = 'photo-preview-item';
        const img = document.createElement('img');
        img.alt = '';
        const reader = new FileReader();
        reader.onload = (e) => { img.src = e.target.result; };
        reader.readAsDataURL(file);
        item.appendChild(img);

        const removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'remove-photo';
        removeBtn.textContent = '×';
        removeBtn.onclick = () => { pendingPhotos.splice(index, 1); renderPhotoPreview(type); };
        item.appendChild(removeBtn);

        preview.appendChild(item);
    });
}

function removeExistingPhoto(index, type) {
    existingPhotos.splice(index, 1);
    renderPhotoPreview(type);
}

if ($('event-photo')) {
    $('event-photo').addEventListener('change', (e) => {
        pendingPhotos.push(...Array.from(e.target.files));
        renderPhotoPreview('event');
        e.target.value = '';
    });
}

if ($('memo-photo')) {
    $('memo-photo').addEventListener('change', (e) => {
        pendingPhotos.push(...Array.from(e.target.files));
        renderPhotoPreview('memo');
        e.target.value = '';
    });
}

// ===== 컬러 피커 =====
$$('#color-picker .color-option').forEach(el => {
    el.addEventListener('click', () => {
        selectedColor = el.dataset.color;
        updateColorPicker();
    });
});

function updateColorPicker() {
    $$('#color-picker .color-option').forEach(el => {
        el.classList.toggle('selected', el.dataset.color === selectedColor);
    });
}

// ===== 네비게이션 =====
if ($('prev-btn')) {
    $('prev-btn').addEventListener('click', () => {
        currentMonth--;
        if (currentMonth < 0) { currentMonth = 11; currentYear--; }
        renderCalendar();
    });
}

if ($('next-btn')) {
    $('next-btn').addEventListener('click', () => {
        currentMonth++;
        if (currentMonth > 11) { currentMonth = 0; currentYear++; }
        renderCalendar();
    });
}

if ($('today-btn')) {
    $('today-btn').addEventListener('click', () => {
        const today = new Date();
        currentYear = today.getFullYear();
        currentMonth = today.getMonth();
        renderCalendar();
    });
}

// ===== 뷰 모드 전환 =====
$$('.view-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        currentView = btn.dataset.view;
        $$('.view-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');

        if ($('month-view')) $('month-view').classList.toggle('hidden', currentView !== 'month');
        if ($('week-view')) $('week-view').classList.toggle('hidden', currentView !== 'week');
        if ($('day-view')) $('day-view').classList.toggle('hidden', currentView !== 'day');
    });
});

// ===== 탭 전환 =====
$$('.nav-tab').forEach(tab => {
    tab.addEventListener('click', () => {
        const target = tab.dataset.tab;
        currentTab = target;

        $$('.nav-tab').forEach(t => t.classList.remove('active'));
        $$(`.nav-tab[data-tab="${target}"]`).forEach(t => t.classList.add('active'));

        // 헤더 & 콘텐츠 전환
        ['calendar', 'checklist', 'memo', 'dday'].forEach(section => {
            const header = $(`header-${section}`);
            const content = $(`content-${section}`);
            if (header) header.classList.toggle('hidden', section !== target);
            if (content) content.classList.toggle('hidden', section !== target);
        });
    });
});

// ===== FAB 버튼 =====
if ($('fab-btn')) {
    $('fab-btn').addEventListener('click', () => {
        switch (currentTab) {
            case 'calendar': openNewEvent(); break;
            case 'memo': openNewMemo(); break;
            case 'checklist': openNewChecklist(); break;
            case 'dday': openNewDday(); break;
        }
    });
}

if ($('add-event-day-btn')) {
    $('add-event-day-btn').addEventListener('click', () => {
        $('day-detail').classList.add('hidden');
        openNewEvent(selectedDate);
    });
}

// ===== 모달 닫기 =====
if ($('close-day-detail')) $('close-day-detail').addEventListener('click', () => $('day-detail').classList.add('hidden'));
if ($('close-event-modal')) $('close-event-modal').addEventListener('click', () => $('event-modal').classList.add('hidden'));
if ($('close-memo-modal')) $('close-memo-modal').addEventListener('click', () => $('memo-modal').classList.add('hidden'));
if ($('close-checklist-modal')) $('close-checklist-modal').addEventListener('click', () => $('checklist-modal').classList.add('hidden'));
if ($('close-dday-modal')) $('close-dday-modal').addEventListener('click', () => $('dday-modal').classList.add('hidden'));

['day-detail', 'event-modal', 'memo-modal', 'checklist-modal', 'dday-modal'].forEach(id => {
    const el = $(id);
    if (el) {
        el.addEventListener('click', (e) => {
            if (e.target === el) el.classList.add('hidden');
        });
    }
});

// ===== 유틸 =====
function formatDate(date) {
    return `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`;
}

function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

// PWA 서비스워커 등록
if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js').catch(() => {});
}
