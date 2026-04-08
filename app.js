// ===== Firebase 설정 (여기에 본인 Firebase 설정값을 넣으세요) =====
const firebaseConfig = {
    apiKey: "AIzaSyDLry0f9aE24wWfa8_CBJuJAxEwYzNaWyU",
    authDomain: "edito-f8256.firebaseapp.com",
    projectId: "edito-f8256",
    storageBucket: "edito-f8256.firebasestorage.app",
    messagingSenderId: "704063878825",
    appId: "1:704063878825:web:8976ea2f8eb6661aeb5bb8"
};

// Firebase 초기화
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
let selectedDate = null;
let selectedColor = '#4A90D9';
let pendingPhotos = []; // 새로 추가할 사진 파일들
let existingPhotos = []; // 기존 사진 URL들
let currentTab = 'calendar';
let isDragging = false;

// ===== DOM 요소 =====
const $ = (id) => document.getElementById(id);
const $$ = (sel) => document.querySelectorAll(sel);

// ===== 인증 상태 감시 =====
auth.onAuthStateChanged((user) => {
    $('loading-screen').classList.add('hidden');
    if (user) {
        currentUser = user;
        $('login-screen').classList.add('hidden');
        $('main-screen').classList.remove('hidden');
        $('user-avatar').src = user.photoURL || '';
        $('user-avatar2').src = user.photoURL || '';
        $('user-name').textContent = user.displayName || user.email;
        loadEvents();
        loadMemos();
        renderCalendar();
    } else {
        currentUser = null;
        $('login-screen').classList.remove('hidden');
        $('main-screen').classList.add('hidden');
        $('memo-screen').classList.add('hidden');
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

$('user-menu-btn2')?.addEventListener('click', (e) => {
    e.stopPropagation();
    $('user-menu').classList.toggle('hidden');
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

// ===== 캘린더 렌더링 =====
function renderCalendar() {
    const grid = $('calendar-grid');
    grid.innerHTML = '';

    const monthNames = ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'];
    $('current-month').textContent = `${currentYear}년 ${monthNames[currentMonth]}`;

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
        const cell = createCalendarCell(day, dateStr, true);
        grid.appendChild(cell);
    }

    // 이번 달
    for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = `${currentYear}-${String(currentMonth+1).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
        const isToday = dateStr === todayStr;
        const cell = createCalendarCell(day, dateStr, false, isToday);
        grid.appendChild(cell);
    }

    // 다음 달
    const totalCells = grid.children.length;
    const remaining = Math.ceil(totalCells / 7) * 7 - totalCells;
    for (let day = 1; day <= remaining; day++) {
        const date = new Date(currentYear, currentMonth + 1, day);
        const dateStr = formatDate(date);
        const cell = createCalendarCell(day, dateStr, true);
        grid.appendChild(cell);
    }
}

function createCalendarCell(day, dateStr, isOtherMonth, isToday = false) {
    const cell = document.createElement('div');
    cell.className = 'calendar-cell';
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
            dot.addEventListener('dragstart', (e) => {
                isDragging = true;
                e.dataTransfer.setData('text/plain', event.id);
                e.dataTransfer.effectAllowed = 'move';
                setTimeout(() => dot.classList.add('dragging'), 0);
            });
            dot.addEventListener('dragend', () => {
                isDragging = false;
                dot.classList.remove('dragging');
                document.querySelectorAll('.calendar-cell.drag-over').forEach(c => c.classList.remove('drag-over'));
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
        // 자식 요소로 이동할 때 오작동 방지
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
        // 드래그 중에는 클릭 이벤트 무시
        if (isDragging) return;
        selectedDate = dateStr;
        renderDayDetail(dateStr);
        $('day-detail').classList.remove('hidden');
    });

    return cell;
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
    $('event-memo').value = event.memo || '';
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
        memo: $('event-memo').value,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    };

    // 사진 업로드
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

// ===== 메모 관련 =====
function renderMemoList() {
    const list = $('memo-list');
    if (memos.length === 0) {
        list.innerHTML = '<div class="empty-state"><span>📝</span><span>메모가 없습니다</span></div>';
        return;
    }

    list.innerHTML = memos.map(memo => {
        const date = memo.updatedAt ? new Date(memo.updatedAt.seconds * 1000) : new Date();
        const dateStr = `${date.getFullYear()}.${date.getMonth()+1}.${date.getDate()}`;
        const photosHtml = memo.photos && memo.photos.length > 0 ?
            `<div class="memo-card-photos">${memo.photos.slice(0,4).map(p => `<img src="${p}" alt="">`).join('')}</div>` : '';
        return `
            <div class="memo-card" onclick="openEditMemo('${memo.id}')">
                <div class="memo-card-title">${escapeHtml(memo.title)}</div>
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
    $('memo-title').value = memo.title;
    $('memo-content').value = memo.content || '';
    $('delete-memo-btn').classList.remove('hidden');

    pendingPhotos = [];
    existingPhotos = memo.photos || [];
    renderPhotoPreview('memo');

    $('memo-modal').classList.remove('hidden');
}

$('memo-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const memoId = $('memo-id').value;
    const data = {
        title: $('memo-title').value,
        content: $('memo-content').value,
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

// 사진 미리보기 (일정/메모 공통)
function renderPhotoPreview(type) {
    const preview = $(`${type}-photo-preview`);
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

$('event-photo').addEventListener('change', (e) => {
    pendingPhotos.push(...Array.from(e.target.files));
    renderPhotoPreview('event');
    e.target.value = '';
});

$('memo-photo').addEventListener('change', (e) => {
    pendingPhotos.push(...Array.from(e.target.files));
    renderPhotoPreview('memo');
    e.target.value = '';
});

// ===== 컬러 피커 =====
$$('.color-option').forEach(el => {
    el.addEventListener('click', () => {
        selectedColor = el.dataset.color;
        updateColorPicker();
    });
});

function updateColorPicker() {
    $$('.color-option').forEach(el => {
        el.classList.toggle('selected', el.dataset.color === selectedColor);
    });
}

// ===== 네비게이션 =====
$('prev-month').addEventListener('click', () => {
    currentMonth--;
    if (currentMonth < 0) { currentMonth = 11; currentYear--; }
    renderCalendar();
});

$('next-month').addEventListener('click', () => {
    currentMonth++;
    if (currentMonth > 11) { currentMonth = 0; currentYear++; }
    renderCalendar();
});

$('today-btn').addEventListener('click', () => {
    const today = new Date();
    currentYear = today.getFullYear();
    currentMonth = today.getMonth();
    renderCalendar();
});

// 탭 전환
$$('.nav-tab').forEach(tab => {
    tab.addEventListener('click', () => {
        const target = tab.dataset.tab;
        currentTab = target;

        $$('.nav-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');

        if (target === 'calendar') {
            $('main-screen').classList.remove('hidden');
            $('memo-screen').classList.add('hidden');
        } else {
            $('main-screen').classList.add('hidden');
            $('memo-screen').classList.remove('hidden');
        }

        // 다른 탭의 active도 동기화
        $$(`.nav-tab[data-tab="${target}"]`).forEach(t => t.classList.add('active'));
    });
});

// FAB 버튼
$('add-event-btn').addEventListener('click', () => openNewEvent());
$('add-memo-btn').addEventListener('click', () => openNewMemo());
$('add-event-day-btn').addEventListener('click', () => {
    $('day-detail').classList.add('hidden');
    openNewEvent(selectedDate);
});

// 모달 닫기
$('close-day-detail').addEventListener('click', () => $('day-detail').classList.add('hidden'));
$('close-event-modal').addEventListener('click', () => $('event-modal').classList.add('hidden'));
$('close-memo-modal').addEventListener('click', () => $('memo-modal').classList.add('hidden'));

// 모달 외부 클릭 닫기
['day-detail', 'event-modal', 'memo-modal'].forEach(id => {
    $(id).addEventListener('click', (e) => {
        if (e.target === $(id)) $(id).classList.add('hidden');
    });
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
