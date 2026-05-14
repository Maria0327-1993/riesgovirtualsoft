// Risk Manager - ESTABLE v64
function removeAccents(str) {
    if (!str) return "";
    return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

const currentUserObj = localStorage.getItem('riskOps_currentUser');
if (!currentUserObj && !window.location.href.includes('login.html')) {
    window.location.href = 'login.html';
}
const currentUser = currentUserObj ? JSON.parse(currentUserObj) : null;

function updateClock() {
    const el = document.getElementById('liveClock');
    if (el) {
        const now = new Date();
        el.textContent = now.toLocaleTimeString('es-ES');
    }
}
setInterval(updateClock, 1000);

async function initApp() {
    if (currentUser) {
        document.querySelector('.user-name').textContent = currentUser.name;
        document.querySelector('.user-role').textContent = currentUser.role;
        const avatar = document.querySelector('.avatar');
        const clean = removeAccents(currentUser.name).trim();
        avatar.src = `assets/src/img/${clean}.png`;
        avatar.onerror = () => { avatar.src = `https://ui-avatars.com/api/?name=${encodeURIComponent(currentUser.name)}&background=0D8ABC&color=fff`; };
    }

    loadTasks();
    loadSchedule();
    loadTeletrabajo();

    // Nav Logic
    document.querySelectorAll('.nav-item').forEach(item => {
        item.onclick = (e) => {
            e.preventDefault();
            document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
            item.classList.add('active');
            document.querySelectorAll('.view-panel').forEach(v => v.style.display = 'none');
            
            if(item.id === 'navWorkspace') document.getElementById('view-workspace').style.display = 'block';
            if(item.id === 'navHorario') document.getElementById('view-horario').style.display = 'block';
            if(item.id === 'navTeletrabajo') document.getElementById('view-teletrabajo').style.display = 'block';
        };
    });
}

async function loadTasks() {
    try {
        const res = await fetch('Tareas Riesgo/Tareas de Riesgo.xlsx?v=' + Date.now());
        const data = await res.arrayBuffer();
        const wb = XLSX.read(data, {type: 'array'});
        const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        
        const grouped = {};
        json.forEach(t => {
            const set = t['Set '] || t['Set'] || 'Otros';
            if(!grouped[set]) grouped[set] = [];
            grouped[set].push(t);
        });

        const select = document.getElementById('activeSetSelect');
        select.innerHTML = '<option value="" disabled selected>Selecciona tu SET...</option>';
        Object.keys(grouped).sort().forEach(s => {
            select.innerHTML += `<option value="${s}">${s}</option>`;
        });
        select.onchange = (e) => renderTree(grouped[e.target.value]);
    } catch(e) {}
}

function renderTree(tasks) {
    const container = document.querySelector('.tree-container');
    container.innerHTML = tasks.map(t => `
        <div class="task-item" onclick="showDetail(${JSON.stringify(t).replace(/"/g, '&quot;')})">
            <i class='bx bx-file'></i> <span>${t.Tarea}</span>
        </div>
    `).join('');
}

function showDetail(t) {
    document.querySelector('.task-detail-content').innerHTML = `
        <div class="glass-panel" style="padding:20px;">
            <h3>${t.Tarea}</h3>
            <p style="margin-top:10px;">${t['Detalle de Tarea'] || 'Sin detalle'}</p>
            <p><strong>Horario:</strong> ${t.Horario || '-'}</p>
        </div>
    `;
}

async function loadSchedule() {
    try {
        const res = await fetch('Horario/Horario 2026.xlsx?v=' + Date.now());
        const data = await res.arrayBuffer();
        const wb = XLSX.read(data, {type: 'array'});
        const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header: 1});
        const weeks = json.map((r, i) => ({n: r[0], i})).filter(w => w.n && w.n.toString().includes('Semana'));
        const sel = document.getElementById('weekSelector');
        sel.innerHTML = weeks.map(w => `<option value="${w.i}">${w.n}</option>`).join('');
        sel.onchange = (e) => renderTable('scheduleTableBody', json, parseInt(e.target.value));
        renderTable('scheduleTableBody', json, weeks[0].i);
    } catch(e) {}
}

async function loadTeletrabajo() {
    try {
        const res = await fetch('Teletrabajo/Teletrabajo.xlsx?v=' + Date.now());
        const data = await res.arrayBuffer();
        const wb = XLSX.read(data, {type: 'array'});
        const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header: 1});
        const weeks = json.map((r, i) => ({n: r[0], i})).filter(w => w.n && w.n.toString().includes('Semana'));
        const sel = document.getElementById('teletrabajoWeekSelector');
        sel.innerHTML = weeks.map(w => `<option value="${w.i}">${w.n}</option>`).join('');
        sel.onchange = (e) => renderTable('teletrabajoTableBody', json, parseInt(e.target.value), true);
        renderTable('teletrabajoTableBody', json, weeks[0].i, true);
    } catch(e) {}
}

function renderTable(id, json, start, isTele) {
    const body = document.getElementById(id);
    body.innerHTML = '';
    for(let i = start + 1; i < json.length; i++) {
        if(!json[i][0] || json[i][0].toString().includes('Semana')) break;
        const tr = document.createElement('tr');
        tr.innerHTML = json[i].map((c, idx) => {
            if(isTele && idx > 0) {
                const txt = (c || '').toString().toUpperCase();
                if(txt.includes('T')) return `<td>🏠</td>`;
                if(txt.includes('O')) return `<td>🏢</td>`;
            }
            return `<td>${c || '-'}</td>`;
        }).join('');
        body.appendChild(tr);
    }
}

window.handleEndShift = () => {
    if(confirm("¿Finalizar turno?")) {
        localStorage.removeItem('riskOps_currentUser');
        window.location.href = 'login.html';
    }
};

window.onload = initApp;
