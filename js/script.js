(() => {
    const recordsKey = 'attendance-records';
    const metaKey = 'attendance-meta';
    const apiKey = 'attendance-api-url';
    const guardKey = 'guard-notes';

    const defaultRecords = [
        {
            id: 1,
            className: '#102',
            groupName: '12-6',
            expected: 33,
            present: 30,
            absentMorning: 1,
            absentAfternoon: 1,
            absentOther: 1,
            absentNames: '',
            leaveNames: '',
            notes: ''
        },
        {
            id: 2,
            className: '#205',
            groupName: '11-б',
            expected: 33,
            present: 33,
            absentMorning: 0,
            absentAfternoon: 0,
            absentOther: 0,
            absentNames: '',
            leaveNames: '',
            notes: ''
        },
        {
            id: 3,
            className: '#220',
            groupName: '12-б',
            expected: 34,
            present: 32,
            absentMorning: 1,
            absentAfternoon: 1,
            absentOther: 0,
            absentNames: '',
            leaveNames: 'Энхжин',
            notes: 'Чөлөөт'
        }
    ];

    let records = [];
    let editingId = null;
    let guardRecords = [];

    const form = document.getElementById('attendance-form');
    const attendanceBody = document.getElementById('attendance-body');
    const summaryExpected = document.getElementById('summary-expected');
    const summaryPresent = document.getElementById('summary-present');
    const summaryAbsent = document.getElementById('summary-absent');
    const totalExpected = document.getElementById('total-expected');
    const totalPresent = document.getElementById('total-present');
    const totalMorning = document.getElementById('total-morning');
    const totalAfternoon = document.getElementById('total-afternoon');
    const totalOther = document.getElementById('total-other');
    const absenceList = document.getElementById('absence-list');
    const printBtn = document.getElementById('print-btn');
    const exportBtn = document.getElementById('export-btn');
    const syncBtn = document.getElementById('sync-btn');
    const syncStatus = document.getElementById('sync-status');
    const apiUrlInput = document.getElementById('api-url');
    const saveApiBtn = document.getElementById('save-api');
    const clearBtn = document.getElementById('clear-form');
    const sessionDate = document.getElementById('session-date');
    const shiftSelect = document.getElementById('shift-select');
    const classroomLead = document.getElementById('classroom-lead');
    const guardForm = document.getElementById('guard-form');
    const guardTeacher = document.getElementById('guard-teacher');
    const guardShift = document.getElementById('guard-shift');
    const guardNote = document.getElementById('guard-note');
    const guardList = document.getElementById('guard-list');
    const guardClear = document.getElementById('guard-clear');

    const storage = (() => {
        try {
            return window.localStorage;
        } catch {
            return null;
        }
    })();

    const readStorage = (key) => {
        if (!storage) return null;
        const raw = storage.getItem(key);
        try {
            return raw ? JSON.parse(raw) : null;
        } catch {
            return null;
        }
    };

    const writeStorage = (key, value) => {
        if (!storage) return;
        storage.setItem(key, JSON.stringify(value));
    };

    const persistMeta = () => writeStorage(metaKey, {
        date: sessionDate.value,
        shift: shiftSelect.value,
        lead: classroomLead.value.trim()
    });

    const getMeta = () => ({
        date: sessionDate.value,
        shift: shiftSelect.value,
        lead: classroomLead.value.trim()
    });

    const loadMeta = () => {
        const saved = readStorage(metaKey);
        if (!saved) {
            return;
        }
        if (saved.date) {
            sessionDate.value = saved.date;
        }
        if (saved.shift) {
            shiftSelect.value = saved.shift;
        }
        if (saved.lead) {
            classroomLead.value = saved.lead;
        }
    };

    const loadRecords = () => {
        const saved = readStorage(recordsKey);
        records = saved && saved.length ? saved : defaultRecords;
    };

    const persistRecords = () => writeStorage(recordsKey, records);

    const loadGuardRecords = () => {
        const saved = readStorage(guardKey);
        guardRecords = Array.isArray(saved) ? saved : [];
    };

    const persistGuardRecords = () => writeStorage(guardKey, guardRecords);

    const normalizeNumber = (value) => {
        const number = parseInt(value, 10);
        return Number.isFinite(number) ? number : 0;
    };

    const renderRecords = () => {
        attendanceBody.innerHTML = '';
        records.forEach((record, index) => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${index + 1}</td>
                <td class="class-name">${record.className}</td>
                <td>${record.groupName}</td>
                <td class="count">${record.expected}</td>
                <td class="count">${record.present}</td>
                <td class="count">${record.absentMorning}</td>
                <td class="count">${record.absentAfternoon}</td>
                <td class="count">${record.absentOther}</td>
                <td>${record.absentNames || '—'}</td>
                <td>${record.leaveNames || '—'}</td>
                <td class="actions">
                    <button type="button" class="edit" data-id="${record.id}">Засах</button>
                    <button type="button" class="delete" data-id="${record.id}">Устгах</button>
                </td>
            `;
            attendanceBody.appendChild(row);
        });
        updateSummary();
        renderAbsenceList();
    };

    const updateSummary = () => {
        const totalExp = records.reduce((sum, record) => sum + record.expected, 0);
        const totalPres = records.reduce((sum, record) => sum + record.present, 0);
        const totalMorningAbs = records.reduce((sum, record) => sum + record.absentMorning, 0);
        const totalAfternoonAbs = records.reduce((sum, record) => sum + record.absentAfternoon, 0);
        const totalOtherAbs = records.reduce((sum, record) => sum + record.absentOther, 0);
        const totalAbs = Math.max(totalExp - totalPres, totalMorningAbs + totalAfternoonAbs + totalOtherAbs);

        summaryExpected.textContent = totalExp;
        summaryPresent.textContent = totalPres;
        summaryAbsent.textContent = totalAbs;
        totalExpected.textContent = totalExp;
        totalPresent.textContent = totalPres;
        totalMorning.textContent = totalMorningAbs;
        totalAfternoon.textContent = totalAfternoonAbs;
        totalOther.textContent = totalOtherAbs;
    };

    const renderAbsenceList = () => {
        absenceList.innerHTML = '';
        records.forEach((record) => {
            const entries = [];
            if (record.absentNames) {
                entries.push(`Тасалсан: ${record.absentNames}`);
            }
            if (record.leaveNames) {
                entries.push(`Чөлөөт: ${record.leaveNames}`);
            }
            if (!entries.length) {
                return;
            }
            const item = document.createElement('li');
            item.textContent = `${record.className} ${record.groupName} — ${entries.join(' | ')}`;
            absenceList.appendChild(item);
        });
    };

    const clearFormState = () => {
        editingId = null;
        form.reset();
    };

    const populateForm = (record) => {
        form.elements['className'].value = record.className;
        form.elements['groupName'].value = record.groupName;
        form.elements['expected'].value = record.expected;
        form.elements['present'].value = record.present;
        form.elements['absentMorning'].value = record.absentMorning;
        form.elements['absentAfternoon'].value = record.absentAfternoon;
        form.elements['absentOther'].value = record.absentOther;
        form.elements['absentNames'].value = record.absentNames;
        form.elements['leaveNames'].value = record.leaveNames;
        form.elements['notes'].value = record.notes;
        editingId = record.id;
    };

    const handleSubmit = (event) => {
        event.preventDefault();
        const data = {
            className: form.elements['className'].value.trim(),
            groupName: form.elements['groupName'].value.trim(),
            expected: normalizeNumber(form.elements['expected'].value),
            present: normalizeNumber(form.elements['present'].value),
            absentMorning: normalizeNumber(form.elements['absentMorning'].value),
            absentAfternoon: normalizeNumber(form.elements['absentAfternoon'].value),
            absentOther: normalizeNumber(form.elements['absentOther'].value),
            absentNames: form.elements['absentNames'].value.trim(),
            leaveNames: form.elements['leaveNames'].value.trim(),
            notes: form.elements['notes'].value.trim()
        };

        if (!data.className || !data.groupName) {
            alert('Ангийн нэр болон бүлэг заавал оруулна уу.');
            return;
        }

        if (data.present > data.expected) {
            data.present = data.expected;
        }

        const record = {
            ...data,
            id: editingId ?? Date.now()
        };

        if (editingId) {
            records = records.map((item) => (item.id === editingId ? record : item));
        } else {
            records = [...records, record];
        }

        persistRecords();
        renderRecords();
        clearFormState();
    };

    const handleTableActions = (event) => {
        const button = event.target.closest('button');
        if (!button || !button.dataset.id) {
            return;
        }
        const id = Number(button.dataset.id);
        if (button.classList.contains('delete')) {
            const confirmed = confirm('Энэ мөрийг устгах уу?');
            if (!confirmed) {
                return;
            }
            records = records.filter((record) => record.id !== id);
            persistRecords();
            renderRecords();
            return;
        }
        if (button.classList.contains('edit')) {
            const record = records.find((item) => item.id === id);
            if (!record) {
                return;
            }
            populateForm(record);
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }
    };

    const downloadJSON = () => {
        const payload = {
            meta: getMeta(),
            records
        };
        const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href = url;
        anchor.download = `attendance-${new Date().toISOString().slice(0, 10)}.json`;
        document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();
        URL.revokeObjectURL(url);
    };

    const showSyncMessage = (message, success = true) => {
        syncStatus.textContent = message;
        syncStatus.style.color = success ? 'var(--primary)' : 'var(--rose)';
    };

    const syncWithServer = async () => {
        const url = apiUrlInput.value.trim();
        if (!url) {
            showSyncMessage('API хаяг оруулна уу.', false);
            return;
        }
        try {
            showSyncMessage('Сервертэй холбож байна...');
            const response = await fetch(url, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ meta: getMeta(), records })
            });
            if (!response.ok) {
                throw new Error(`${response.status} ${response.statusText}`);
            }
            showSyncMessage('Өгөгдлийг амжилттай илгээлээ.');
        } catch (error) {
            showSyncMessage(`Алдаа: ${error.message}`, false);
        }
    };

    const loadApiUrl = () => {
        const savedUrl = storage ? storage.getItem(apiKey) : null;
        if (savedUrl) {
            apiUrlInput.value = savedUrl;
        } else {
            apiUrlInput.value = '/api/attendance';
        }
    };

    const saveApiUrl = () => {
        if (storage) {
            storage.setItem(apiKey, apiUrlInput.value.trim());
            showSyncMessage('API хаяг хадгалагдлаа.', true);
        }
    };

    const renderGuardRecords = () => {
        guardList.innerHTML = '';
        guardRecords.forEach((entry) => {
            const listItem = document.createElement('li');
            listItem.innerHTML = `
                <div>
                    <strong>${entry.teacher}</strong>
                    <span>${entry.shift} · ${entry.time}</span>
                    <p>${entry.note || '—'}</p>
                </div>
                <button type="button" data-id="${entry.id}">Устгах</button>
            `;
            guardList.appendChild(listItem);
        });
    };

    const handleGuardSubmit = (event) => {
        event.preventDefault();
        const teacherName = guardTeacher.value.trim();
        const shift = guardShift.value;
        const note = guardNote.value.trim();

        if (!teacherName && !note) {
            alert('Жижүүр багшийн нэр эсвэл тэмдэглэл шаардлагатай.');
            return;
        }

        const entry = {
            id: Date.now(),
            teacher: teacherName || 'Тодорхойгүй',
            shift,
            note,
            time: new Date().toLocaleTimeString('mn-MN', { hour: '2-digit', minute: '2-digit' })
        };

        guardRecords = [entry, ...guardRecords];
        persistGuardRecords();
        renderGuardRecords();
        guardForm.reset();
    };

    const handleGuardActions = (event) => {
        const button = event.target.closest('button');
        if (!button || !button.dataset.id) {
            return;
        }
        const id = Number(button.dataset.id);
        guardRecords = guardRecords.filter((entry) => entry.id !== id);
        persistGuardRecords();
        renderGuardRecords();
    };

    const init = () => {
        loadRecords();
        loadGuardRecords();
        loadMeta();
        loadApiUrl();
        bindEvents();
        renderRecords();
        renderGuardRecords();
    };

    const bindEvents = () => {
        form.addEventListener('submit', handleSubmit);
        attendanceBody.addEventListener('click', handleTableActions);
        guardForm.addEventListener('submit', handleGuardSubmit);
        guardList.addEventListener('click', handleGuardActions);
        clearBtn.addEventListener('click', () => {
            clearFormState();
            form.reset();
        });
        guardClear.addEventListener('click', () => guardForm.reset());
        printBtn.addEventListener('click', () => window.print());
        exportBtn.addEventListener('click', downloadJSON);
        syncBtn.addEventListener('click', syncWithServer);
        saveApiBtn.addEventListener('click', saveApiUrl);
        sessionDate.addEventListener('change', persistMeta);
        shiftSelect.addEventListener('change', persistMeta);
        classroomLead.addEventListener('input', persistMeta);
    };

    init();
})();
