// Risk Manager - Rescate v65
function removeAccents(str) {
    if (!str) return "";
    return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

const user = JSON.parse(localStorage.getItem('riskOps_currentUser') || '{}');

function init() {
    if (user.name) {
        document.querySelector('.user-name').textContent = user.name;
        document.querySelector('.user-role').textContent = user.role;
        const av = document.querySelector('.avatar');
        const clean = removeAccents(user.name).trim();
        av.src = `assets/src/img/${clean}.png`;
        av.onerror = () => av.src = `https://ui-avatars.com/api/?name=${encodeURIComponent(user.name)}`;
    }

    loadTasks();
    loadTable('Horario/Horario 2026.xlsx', 'weekSelector', 'scheduleTableBody', false);
    loadTable('Teletrabajo/Teletrabajo.xlsx', 'teleWeekSelector', 'teleTableBody', true);

    document.querySelectorAll('.nav-item').forEach(item => {
        item.onclick = () => {
            document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
            item.classList.add('active');
            document.querySelectorAll('.view-panel').forEach(v => v.style.display = 'none');
            if(item.id === 'navWorkspace') document.getElementById('view-workspace').style.display = 'block';
            if(item.id === 'navHorario') document.getElementById('view-horario').style.display = 'block';
            if(item.id === 'navTeletrabajo') document.getElementById('view-teletrabajo').style.display = 'block';
        };
    });

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
            document.querySelector('.tree-container').innerHTML = list.map(t => `<div class="task-item" onclick="showTask(\'${t.Tarea}\',\'${t['Detalle de Tarea'] || '-'}\')"><i class='bx bx-file'></i>${t.Tarea}</div>`).join('');
        };
    } catch(e) {}
}

window.showTask = (n, d) => {
    document.getElementById('taskDetail').innerHTML = `<h3>${n}</h3><p>${d}</p>`;
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
        const tr = document.createElement('tr');
        tr.innerHTML = json[i].map((c, idx) => {
            if(isTele && idx > 0) {
                if((c||'').toString().toUpperCase().includes('T')) return '<td>🏠</td>';
                if((c||'').toString().toUpperCase().includes('O')) return '<td>🏢</td>';
            }
            return `<td>${c || '-'}</td>`;
        }).join('');
        b.appendChild(tr);
    }
}

window.handleEndShift = () => { if(confirm('¿Salir?')) { localStorage.removeItem('riskOps_currentUser'); window.location.href='login.html'; } };
init();
