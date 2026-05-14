// Risk Manager - App Logic v63 (ULTRA ROBUST)
function removeAccents(str) {
    if (!str) return "";
    return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

const currentUserObj = localStorage.getItem('riskOps_currentUser');
if (!currentUserObj && !window.location.href.includes('login.html')) {
    window.location.href = 'login.html';
}

let currentUser = null;
try {
    currentUser = currentUserObj ? JSON.parse(currentUserObj) : null;
} catch(e) {
    localStorage.removeItem('riskOps_currentUser');
    window.location.href = 'login.html';
}

function updateClock() {
    const el = document.getElementById('liveClock');
    if (!el) return;
    const now = new Date();
    el.textContent = now.getHours().toString().padStart(2, '0') + ':' + 
                     now.getMinutes().toString().padStart(2, '0') + ':' + 
                     now.getSeconds().toString().padStart(2, '0');
}
setInterval(updateClock, 1000);

async function initApp() {
    if (currentUser) {
        const nameEl = document.querySelector('.user-name');
        const roleEl = document.querySelector('.user-role');
        if (nameEl) nameEl.textContent = currentUser.name;
        if (roleEl) roleEl.textContent = currentUser.role;
        
        const avatarEl = document.querySelector('.avatar');
        if (avatarEl) {
            const clean = removeAccents(currentUser.name).trim();
            avatarEl.src = `assets/src/img/${clean}.png`;
            avatarEl.onerror = () => { avatarEl.src = `https://ui-avatars.com/api/?name=${encodeURIComponent(currentUser.name)}&background=0D8ABC&color=fff`; };
        }
    }

    loadExcelTasks();
    loadSchedule();
    loadTeletrabajo();
    renderDocs();
    
    document.querySelectorAll('.nav-item').forEach(item => {
        item.onclick = (e) => {
            e.preventDefault();
            document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
            item.classList.add('active');
            document.querySelectorAll('.view-panel').forEach(v => v.style.display = 'none');
            
            const txt = item.textContent;
            if(txt.includes('Inicio')) { document.getElementById('view-dashboard').style.display = 'block'; loadDashboardStats(); }
            else if(txt.includes('Tareas')) document.getElementById('view-workspace').style.display = 'block';
            else if(txt.includes('Horario')) document.getElementById('view-horario').style.display = 'block';
            else if(txt.includes('Teletrabajo')) document.getElementById('view-teletrabajo').style.display = 'block';
            else if(txt.includes('Documentación')) document.getElementById('view-docs').style.display = 'block';
            else if(txt.includes('Permisos')) document.getElementById('view-permisos').style.display = 'block';
            else if(txt.includes('Aprobaciones')) document.getElementById('view-aprobaciones').style.display = 'block';
        };
    });

    loadDashboardStats();
}

// LOGICA DE TAREAS (SETS) CON FILTRO
async function loadExcelTasks() {
    const container = document.querySelector('.tree-container');
    const select = document.getElementById('activeSetSelect');
    if(!container) return;
    try {
        const res = await fetch('Tareas Riesgo/Tareas de Riesgo.xlsx?v=' + Date.now());
        const data = await res.arrayBuffer();
        const wb = XLSX.read(data, {type: 'array'});
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet);
        
        const grouped = {};
        json.forEach(row => {
            const set = row['Set '] || row['Set'] || 'Otros';
            if(!grouped[set]) grouped[set] = [];
            grouped[set].push(row);
        });

        if(select) {
            select.innerHTML = '<option value="" disabled selected>Selecciona tu SET a trabajar...</option>';
            Object.keys(grouped).sort().forEach(s => {
                select.innerHTML += `<option value="${s}">${s}</option>`;
            });
            select.onchange = (e) => {
                const filtered = {};
                filtered[e.target.value] = grouped[e.target.value];
                renderTree(filtered);
            };
        }
    } catch(e) { console.error("Error cargando SETS:", e); }
}

function renderTree(grouped) {
    const container = document.querySelector('.tree-container');
    if(!container) return;
    container.innerHTML = '';
    Object.keys(grouped).forEach(set => {
        const div = document.createElement('div');
        div.className = 'tree-item open';
        div.innerHTML = `
            <div class="tree-header">
                <i class="bx bx-chevron-down"></i>
                <i class="bx bx-folder-open"></i>
                <span>${set}</span>
                <span class="badge pending">${grouped[set].length}</span>
            </div>
            <div class="tree-children show">
                ${grouped[set].map(t => `
                    <div class="task-item" onclick="selectTask(${JSON.stringify(t).replace(/"/g, '&quot;')})">
                        <i class="bx bx-file"></i>
                        <span>${t.Tarea || 'Sin nombre'}</span>
                    </div>
                `).join('')}
            </div>
        `;
        container.appendChild(div);
    });
}

function selectTask(task) {
    const detail = document.querySelector('.task-detail-content');
    if(detail) {
        detail.innerHTML = `
            <h3 style="color:var(--accent-primary); margin-bottom:15px;">${task.Tarea}</h3>
            <div class="glass-panel" style="padding:20px;">
                <p><strong>Detalle:</strong> ${task['Detalle de Tarea'] || '-'}</p>
                <p><strong>Horario:</strong> ${task.Horario || '-'}</p>
                <p><strong>Día:</strong> ${task['Día'] || '-'}</p>
            </div>
        `;
    }
}

// LOGICA DE TABLAS ROBUSTA
async function loadTableData(path, selectorId, bodyId, isTele) {
    try {
        const res = await fetch(path + '?v=' + Date.now());
        if(!res.ok) return;
        const data = await res.arrayBuffer();
        const wb = XLSX.read(data, {type: 'array'});
        const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header: 1});
        
        const weeks = [];
        json.forEach((row, i) => {
            if(row && row[0] && row[0].toString().toLowerCase().includes('semana')) {
                weeks.push({name: row[0], index: i});
            }
        });

        const selector = document.getElementById(selectorId);
        if(selector) {
            selector.innerHTML = weeks.map(w => `<option value="${w.index}">${w.name}</option>`).join('');
            selector.onchange = (e) => renderTable(bodyId, json, parseInt(e.target.value), isTele);
            if(weeks.length > 0) renderTable(bodyId, json, weeks[0].index, isTele);
        }
    } catch(e) { console.error("Error en tabla:", path, e); }
}

async function loadSchedule() { loadTableData('Horario/Horario 2026.xlsx', 'weekSelector', 'scheduleTableBody', false); }
async function loadTeletrabajo() { loadTableData('Teletrabajo/Teletrabajo.xlsx', 'teletrabajoWeekSelector', 'teletrabajoTableBody', true); }

function renderTable(bodyId, json, start, isTele) {
    const body = document.getElementById(bodyId);
    if(!body) return;
    body.innerHTML = '';
    for(let i = start + 1; i < json.length; i++) {
        const row = json[i];
        if(!row || !row[0] || row[0].toString().toLowerCase().includes('semana')) break;
        if(row[0] === 'GESTOR') continue;
        const tr = document.createElement('tr');
        tr.innerHTML = row.map((cell, idx) => {
            if(isTele && idx > 0) {
                const c = (cell || '').toString().toUpperCase();
                if(c.includes('T')) return `<td><div class="status-badge tele"><i class="bx bx-home"></i></div></td>`;
                if(c.includes('O')) return `<td><div class="status-badge office"><i class="bx bx-building"></i></div></td>`;
            }
            return `<td>${cell || '-'}</td>`;
        }).join('');
        body.appendChild(tr);
    }
}

function renderDocs() {
    const grid = document.querySelector('.docs-grid');
    const recent = document.getElementById('recent-docs-list');
    const files = ["Instructivo de validación de GGR Casino.pdf", "Política Procedimiento De Aprobación De Retiros.pdf", "VALIDACIÓN DE ABUSO DE BONOS EN CAMPAÑAS DE CRM.pdf"];
    const html = files.map(f => `
        <div class="glass-panel doc-card" onclick="window.open('Procesos/${f}')">
            <i class="bx bxs-file-pdf"></i>
            <span>${f.split('.')[0]}</span>
        </div>`).join('');
    if(grid) grid.innerHTML = html;
    if(recent) recent.innerHTML = html;
}

async function loadDashboardStats() {
    const t = document.getElementById('stat-teletrabajo');
    if(t) t.textContent = '12';
}

window.onload = initApp;
window.handleEndShift = function() {
    if(confirm("¿Seguro que deseas salir?")) {
        localStorage.removeItem("riskOps_currentUser");
        window.location.href = "login.html";
    }
};
