// RiskOps - Portal Logic v68
function removeAccents(str) {
    if (!str) return "";
    return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

const currentUser = JSON.parse(localStorage.getItem('riskOps_currentUser') || '{}');

function init() {
    if (currentUser.name) {
        document.querySelector('.user-name').textContent = currentUser.name;
        // Iniciales para el avatar círculo
        const initials = currentUser.name.split(' ').map(n => n[0]).join('').toUpperCase().substring(0,2);
        const avCircle = document.querySelector('.avatar-circle');
        if(avCircle) avCircle.textContent = initials;
    }

    loadTasks();
    loadTable('Horario/Horario 2026.xlsx', 'weekSelector', 'scheduleTableBody', false);

    // Navegación RiskOps
    const navMapping = {
        'navWorkspace': 'view-workspace',
        'navTurno': 'view-turno',
        'navDocs': 'view-docs',
        'navPermisos': 'view-permisos',
        'navSoporte': 'view-soporte'
    };

    document.querySelectorAll('.nav-item').forEach(item => {
        item.onclick = (e) => {
            e.preventDefault();
            const targetId = navMapping[item.id];
            if(!targetId) return;

            document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
            item.classList.add('active');

            document.querySelectorAll('.view-panel').forEach(v => v.classList.remove('active-view'));
            const targetView = document.getElementById(targetId);
            if(targetView) targetView.classList.add('active-view');
        };
    });

    setInterval(() => {
        const clock = document.getElementById('liveClock');
        if(clock) clock.textContent = new Date().toLocaleTimeString();
    }, 1000);
}

async function loadTasks() {
    try {
        const res = await fetch('Tareas Riesgo/Tareas de Riesgo.xlsx?v=' + Date.now());
        const wb = XLSX.read(await res.arrayBuffer(), {type: 'array'});
        const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        const grouped = {};
        json.forEach(t => {
            const s = t['Set '] || t['Set'] || 'Otros';
            if(!grouped[s]) grouped[s] = [];
            grouped[s].push(t);
        });
        const sel = document.getElementById('activeSetSelect');
        if(sel) {
            sel.innerHTML = '<option value="">Selecciona tu SET...</option>' + Object.keys(grouped).sort().map(s => `<option value="${s}">${s}</option>`).join('');
            sel.onchange = (e) => {
                const list = grouped[e.target.value];
                const tree = document.querySelector('.tree-container');
                if(tree) tree.innerHTML = list.map(t => `<div class="task-item" onclick="showTask(\'${t.Tarea}\',\'${(t['Detalle de Tarea']||'-').replace(/'/g, "\\'")}\')"><i class='bx bx-file'></i>${t.Tarea}</div>`).join('');
            };
        }
    } catch(e) {}
}

window.showTask = (n, d) => {
    const detail = document.getElementById('taskDetail');
    if(detail) detail.innerHTML = `<div class="glass-panel" style="padding:25px;"><h3>${n}</h3><p style="margin-top:15px; color:#9CA3AF;">${d}</p></div>`;
};

async function loadTable(path, selId, bodyId, isTele) {
    try {
        const res = await fetch(path + '?v=' + Date.now());
        const wb = XLSX.read(await res.arrayBuffer(), {type: 'array'});
        const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header: 1});
        const weeks = json.map((r, i) => ({n: r[0], i})).filter(w => w.n && w.n.toString().includes('Semana'));
        const sel = document.getElementById(selId);
        if(sel) {
            sel.innerHTML = weeks.map(w => `<option value="${w.i}">${w.n}</option>`).join('');
            sel.onchange = (e) => render(bodyId, json, parseInt(e.target.value), isTele);
            if(weeks.length > 0) render(bodyId, json, weeks[0].i, isTele);
        }
    } catch(e) {}
}

function render(id, json, start, isTele) {
    const b = document.getElementById(id);
    if(!b) return;
    b.innerHTML = '';
    for(let i = start + 1; i < json.length; i++) {
        if(!json[i] || !json[i][0] || json[i][0].toString().includes('Semana')) break;
        const tr = document.createElement('tr');
        tr.innerHTML = json[i].map(c => `<td>${c || '-'}</td>`).join('');
        b.appendChild(tr);
    }
}

window.handleEndShift = () => { if(confirm('¿Finalizar turno?')) { localStorage.removeItem('riskOps_currentUser'); window.location.href='login.html'; } };
init();
