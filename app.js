// RiskOps - Reversión Estable v69
function removeAccents(str) {
    if (!str) return "";
    return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

const currentUser = JSON.parse(localStorage.getItem('riskOps_currentUser') || '{}');

function init() {
    if (currentUser.name) {
        document.querySelector('.user-name').textContent = currentUser.name;
        const initials = currentUser.name.split(' ').map(n => n[0]).join('').toUpperCase().substring(0,2);
        const circle = document.querySelector('.avatar-circle');
        if(circle) circle.textContent = initials;
    }

    loadTasks();
    loadSchedule();
    
    // Navegación
    const navs = {
        'navWorkspace': 'view-workspace',
        'navTurno': 'view-turno',
        'navDocs': 'view-docs',
        'navPermisos': 'view-permisos',
        'navSoporte': 'view-soporte'
    };

    document.querySelectorAll('.nav-item').forEach(item => {
        item.onclick = (e) => {
            e.preventDefault();
            const target = navs[item.id];
            if(!target) return;
            document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
            item.classList.add('active');
            document.querySelectorAll('.view-panel').forEach(v => v.classList.remove('active-view'));
            document.getElementById(target).classList.add('active-view');
        };
    });

    // Theme Toggle
    const themeBtn = document.getElementById('themeToggle');
    if(themeBtn) {
        themeBtn.onclick = () => {
            const cur = document.documentElement.getAttribute('data-theme');
            const next = cur === 'dark' ? 'light' : 'dark';
            document.documentElement.setAttribute('data-theme', next);
            themeBtn.innerHTML = next === 'dark' ? '<i class="bx bx-moon"></i>' : '<i class="bx bx-sun"></i>';
        };
    }

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
                tree.innerHTML = list.map(t => `<div class="task-item" onclick="showTask(\'${t.Tarea}\',\'${(t['Detalle de Tarea']||'-').replace(/'/g, "\\'")}\')"><i class='bx bx-file'></i> ${t.Tarea}</div>`).join('');
            };
        }
    } catch(e) {}
}

window.showTask = (n, d) => {
    document.getElementById('taskDetail').innerHTML = `<div class="glass-panel" style="padding:25px;"><h3>${n}</h3><p style="margin-top:15px; color:var(--text-secondary); line-height:1.6;">${d}</p></div>`;
};

async function loadSchedule() {
    try {
        const res = await fetch('Horario/Horario 2026.xlsx?v=' + Date.now());
        const wb = XLSX.read(await res.arrayBuffer(), {type: 'array'});
        const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header: 1});
        const weeks = json.map((r, i) => ({n: r[0], i})).filter(w => w.n && w.n.toString().includes('Semana'));
        const sel = document.getElementById('weekSelector');
        if(sel) {
            sel.innerHTML = weeks.map(w => `<option value="${w.i}">${w.n}</option>`).join('');
            sel.onchange = (e) => render(json, parseInt(e.target.value));
            render(json, weeks[0].i);
        }
    } catch(e) {}
}

function render(json, start) {
    const b = document.getElementById('scheduleTableBody');
    if(!b) return;
    b.innerHTML = '';
    for(let i = start + 1; i < json.length; i++) {
        if(!json[i] || !json[i][0] || json[i][0].toString().includes('Semana')) break;
        const tr = document.createElement('tr');
        tr.innerHTML = json[i].map(c => `<td>${c || '-'}</td>`).join('');
        b.appendChild(tr);
    }
}

window.handleEndShift = () => { if(confirm('¿Salir de RiskOps?')) { localStorage.removeItem('riskOps_currentUser'); window.location.href='login.html'; } };
init();
