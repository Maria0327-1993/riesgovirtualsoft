// Risk Manager - App Logic v61
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
            else if(txt.includes('Horario')) { document.getElementById('view-horario').style.display = 'block'; loadSchedule(); }
            else if(txt.includes('Teletrabajo')) { document.getElementById('view-teletrabajo').style.display = 'block'; loadTeletrabajo(); }
            else if(txt.includes('Documentación')) document.getElementById('view-docs').style.display = 'block';
            else if(txt.includes('Permisos')) document.getElementById('view-permisos').style.display = 'block';
            else if(txt.includes('Aprobaciones')) document.getElementById('view-aprobaciones').style.display = 'block';
        };
    });

    loadDashboardStats();
}

async function loadSchedule() {
    try {
        const res = await fetch('Horario/Horario 2026.xlsx?v=' + Date.now());
        if(!res.ok) return;
        const data = await res.arrayBuffer();
        const wb = XLSX.read(data, {type: 'array'});
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, {header: 1});
        
        const selector = document.getElementById('weekSelector');
        if(selector) {
            const weeks = [];
            json.forEach((row, i) => {
                if(row[0] && row[0].toString().toLowerCase().includes('semana')) {
                    weeks.push({ name: row[0], index: i });
                }
            });

            selector.innerHTML = weeks.map(w => `<option value="${w.index}">${w.name}</option>`).join('');
            selector.onchange = (e) => renderGroupedTable('scheduleTableBody', json, parseInt(e.target.value), false);
            if(weeks.length > 0) renderGroupedTable('scheduleTableBody', json, weeks[0].index, false);
        }
    } catch(e) { console.error(e); }
}

async function loadTeletrabajo() {
    try {
        const res = await fetch('Teletrabajo/Teletrabajo.xlsx?v=' + Date.now());
        if(!res.ok) return;
        const data = await res.arrayBuffer();
        const wb = XLSX.read(data, {type: 'array'});
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, {header: 1});
        
        const selector = document.getElementById('teletrabajoWeekSelector');
        if(selector) {
            const weeks = [];
            json.forEach((row, i) => {
                if(row[0] && row[0].toString().toLowerCase().includes('semana')) {
                    weeks.push({ name: row[0], index: i });
                }
            });

            selector.innerHTML = weeks.map(w => `<option value="${w.index}">${w.name}</option>`).join('');
            selector.onchange = (e) => renderGroupedTable('teletrabajoTableBody', json, parseInt(e.target.value), true);
            if(weeks.length > 0) renderGroupedTable('teletrabajoTableBody', json, weeks[0].index, true);
        }
    } catch(e) { console.error(e); }
}

function renderGroupedTable(bodyId, json, startIndex, isTele) {
    const body = document.getElementById(bodyId);
    if(!body) return;
    body.innerHTML = '';
    
    for(let i = startIndex + 1; i < json.length; i++) {
        const row = json[i];
        if(!row[0] || row[0].toString().toLowerCase().includes('semana')) break;
        if(row[0] === 'GESTOR') continue;

        const tr = document.createElement('tr');
        tr.innerHTML = row.map((cell, idx) => {
            if(isTele && idx > 0) {
                // Lógica de íconos para Teletrabajo
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
    const archivos = ["Instructivo de validación de GGR Casino.pdf", "Política Procedimiento De Aprobación De Retiros.pdf", "VALIDACIÓN DE ABUSO DE BONOS EN CAMPAÑAS DE CRM.pdf"];
    const html = archivos.map(f => `
        <div class="glass-panel" style="padding:10px; display:flex; gap:10px; cursor:pointer; align-items:center;" onclick="window.open('Procesos/${f}')">
            <i class="bx bxs-file-pdf" style="color:#FF5A5A; font-size:20px;"></i>
            <span style="font-size:12px;">${f}</span>
        </div>`).join('');
    if(grid) grid.innerHTML = html;
    if(recent) recent.innerHTML = html;
}

async function loadDashboardStats() {
    const statT = document.getElementById('stat-teletrabajo');
    if(statT) statT.textContent = '8'; 
}

window.onload = initApp;
window.handleEndShift = function() {
    if(confirm("¿Seguro que deseas salir?")) {
        localStorage.removeItem("riskOps_currentUser");
        window.location.href = "login.html";
    }
};
