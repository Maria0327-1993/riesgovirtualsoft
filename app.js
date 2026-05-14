// Risk Manager - Restauración Total v66
function removeAccents(str) {
    if (!str) return "";
    return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

const currentUser = JSON.parse(localStorage.getItem('riskOps_currentUser') || '{}');

function init() {
    if (currentUser.name) {
        document.querySelector('.user-name').textContent = currentUser.name;
        document.querySelector('.user-role').textContent = currentUser.role;
        const av = document.querySelector('.avatar');
        const clean = removeAccents(currentUser.name).trim();
        av.src = `assets/src/img/${clean}.png`;
        av.onerror = () => av.src = `https://ui-avatars.com/api/?name=${encodeURIComponent(currentUser.name)}`;

        // Mostrar Aprobaciones si es Admin
        if (currentUser.role === 'Admin' || currentUser.role === 'Supervisor') {
            document.getElementById('navAprobaciones').style.display = 'flex';
        }
    }

    loadTasks();
    loadTable('Horario/Horario 2026.xlsx', 'weekSelector', 'scheduleTableBody', false);
    loadTable('Teletrabajo/Teletrabajo.xlsx', 'teleWeekSelector', 'teleTableBody', true);

    // Nav Switcher
    document.querySelectorAll('.nav-item').forEach(item => {
        item.onclick = (e) => {
            e.preventDefault();
            document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
            item.classList.add('active');
            document.querySelectorAll('.view-panel').forEach(v => v.classList.remove('active-view'));
            const viewId = item.id.replace('nav', 'view').toLowerCase();
            document.getElementById(viewId).classList.add('active-view');
        };
    });

    // Modo Claro / Oscuro
    const themeBtn = document.getElementById('themeToggle');
    themeBtn.onclick = () => {
        const current = document.documentElement.getAttribute('data-theme');
        const next = current === 'dark' ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', next);
        themeBtn.innerHTML = next === 'dark' ? '<i class="bx bx-sun"></i>' : '<i class="bx bx-moon"></i>';
    };

    setInterval(() => {
        document.getElementById('liveClock').textContent = new Date().toLocaleTimeString();
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
        sel.innerHTML = '<option value="">Selecciona tu SET...</option>' + Object.keys(grouped).sort().map(s => `<option value="${s}">${s}</option>`).join('');
        sel.onchange = (e) => {
            const list = grouped[e.target.value];
            document.querySelector('.tree-container').innerHTML = list.map(t => `<div class="task-item" onclick="showTask(\'${t.Tarea}\',\'${(t['Detalle de Tarea']||'-').replace(/'/g, "\\'")}\')"><i class='bx bx-file'></i>${t.Tarea}</div>`).join('');
        };
    } catch(e) {}
}

window.showTask = (n, d) => {
    document.getElementById('taskDetail').innerHTML = `<div class="glass-panel" style="padding:25px;"><h3>${n}</h3><p style="margin-top:15px; color:var(--text-secondary);">${d}</p></div>`;
};

async function loadTable(path, selId, bodyId, isTele) {
    try {
        const res = await fetch(path + '?v=' + Date.now());
        const wb = XLSX.read(await res.arrayBuffer(), {type: 'array'});
        const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header: 1});
        const weeks = json.map((r, i) => ({n: r[0], i})).filter(w => w.n && w.n.toString().includes('Semana'));
        const sel = document.getElementById(selId);
        sel.innerHTML = weeks.map(w => `<option value="${w.i}">${w.n}</option>`).join('');
        sel.onchange = (e) => render(bodyId, json, parseInt(e.target.value), isTele);
        render(bodyId, json, weeks[0].i, isTele);
    } catch(e) {}
}

function render(id, json, start, isTele) {
    const b = document.getElementById(id); b.innerHTML = '';
    for(let i = start + 1; i < json.length; i++) {
        if(!json[i][0] || json[i][0].toString().includes('Semana')) break;
        if(json[i][0] === 'GESTOR') continue;
        const tr = document.createElement('tr');
        tr.innerHTML = json[i].map((c, idx) => {
            if(isTele && idx > 0) {
                const txt = (c||'').toString().toUpperCase();
                if(txt.includes('T')) return '<td><div class="status-badge tele"><i class="bx bx-home"></i></div></td>';
                if(txt.includes('O')) return '<td><div class="status-badge office"><i class="bx bx-building"></i></div></td>';
            }
            return `<td>${c || '-'}</td>`;
        }).join('');
        b.appendChild(tr);
    }
}

window.handleEndShift = () => { if(confirm('¿Seguro que deseas salir?')) { localStorage.removeItem('riskOps_currentUser'); window.location.href='login.html'; } };
init();
