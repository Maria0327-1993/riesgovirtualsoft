// Helper para quitar acentos
function removeAccents(str) {
    if (!str) return "";
    return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

// Auth Check
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

let taskStateCache = {};
try {
    const cached = localStorage.getItem('riskOps_cache');
    if(cached) taskStateCache = JSON.parse(cached);
} catch(e) {}

let currentActiveTaskId = null;

// Live Clock Logic
function updateClock() {
    const clockElement = document.getElementById('liveClock');
    if (!clockElement) return;

    const now = new Date();
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');
    
    clockElement.textContent = `${hours}:${minutes}:${seconds}`;
}

setInterval(updateClock, 1000);
updateClock();

let allTasks = [];
let currentSelectedTask = null;

async function loadExcelTasks() {
    const container = document.querySelector('.tree-container');
    if(container) container.innerHTML = '<div style="padding: 20px; color: var(--text-secondary);"><i class="bx bx-loader-alt bx-spin"></i> Cargando Tareas...</div>';
    
    try {
        const url = encodeURI('Tareas Riesgo/Tareas de Riesgo.xlsx') + '?t=' + new Date().getTime();
        const response = await fetch(url);
        if(!response.ok) throw new Error("Error HTTP " + response.status);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, {type: 'array'});
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        
        const tasksBySet = {};
        allTasks = [];
        
        json.forEach((row, index) => {
            const set = row['Set '] || row['Set'] || 'Otros';
            const taskName = row['Tarea'];
            if (!tasksBySet[set]) tasksBySet[set] = [];
            const isDuplicate = tasksBySet[set].some(t => t.name === taskName);
            if (!isDuplicate) {
                tasksBySet[set].push({
                    id: index,
                    name: taskName,
                    detail: row['Detalle de Tarea'],
                    time: row['Horario'],
                    day: row['D�a']
                });
            }
            allTasks.push({ ...row, id: index });
        });

        const select = document.getElementById('activeSetSelect');
        if(select) {
            select.innerHTML = '<option value="" disabled selected>Selecciona tu SET a trabajar...</option><option value="Todos">Mostrar Todos</option>';
            const setsKeys = Object.keys(tasksBySet).sort();
            setsKeys.forEach(set => {
                select.innerHTML += \<option value="\\">\\</option>\;
            });
            
            select.addEventListener('change', (e) => {
                const val = e.target.value;
                if(val === 'Todos') renderTree(tasksBySet);
                else {
                    const filtered = {};
                    filtered[val] = tasksBySet[val];
                    renderTree(filtered);
                }
            });
        }
        renderTree(tasksBySet);
    } catch(e) { console.error(e); }
}

async function loadSchedule() {
    try {
        const response = await fetch('Horario/Horario 2026.xlsx?v=' + Date.now());
        if(!response.ok) return;
        const data = await response.arrayBuffer();
        const workbook = XLSX.read(data, {type: 'array'});
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, {header: 1});
        
        const selector = document.getElementById('weekSelector');
        if(selector) {
            selector.innerHTML = json.slice(1, 10).map((row, i) => \<option value="\\">\\</option>\).join('');
        }
    } catch(e) { console.error(e); }
}

async function loadTeletrabajo() {
    try {
        const response = await fetch('Teletrabajo/Teletrabajo.xlsx?v=' + Date.now());
        if(!response.ok) return;
        const data = await response.arrayBuffer();
        const workbook = XLSX.read(data, {type: 'array'});
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, {header: 1});
    } catch(e) { console.error(e); }
}

function renderTree(tasksBySet) {
    const container = document.querySelector('.tree-container');
    if(!container) return;
    container.innerHTML = '';
    const sets = Object.keys(tasksBySet).sort();
    sets.forEach(set => {
        const setDiv = document.createElement('div');
        setDiv.className = 'tree-item';
        const total = tasksBySet[set].length;
        setDiv.innerHTML = \<div class="tree-header" onclick="toggleTree(this)"><i class="bx bx-chevron-right"></i><span>\\</span><span class="badge pending">\\ Tareas</span></div><div class="tree-children">\\</div>\;
        container.appendChild(setDiv);
    });
    updateKPI();
}

function updateKPI() {
    const totalTasks = document.querySelectorAll('.task-item').length;
    const completed = document.querySelectorAll('.task-item .status-completed').length;
    let percentage = totalTasks > 0 ? Math.round((completed / totalTasks) * 100) : 0;
    const kpi = document.querySelector('.kpi-card');
    if(kpi) kpi.innerHTML = \<div class="kpi-circle">\\%</div><div class="kpi-stats"><p><strong>\\</strong> Asignadas</p></div>\;
}

function toggleTree(el) { el.classList.toggle('open'); const children = el.nextElementSibling; if(children) children.classList.toggle('show'); }


function initApp() {
    loadExcelTasks();
    loadSchedule();
    loadTeletrabajo();
    
    if (currentUser) {
        const userNameEl = document.querySelector('.user-name');
        const roleEl = document.querySelector('.user-role');
        if (userNameEl) userNameEl.textContent = currentUser.name;
        if (roleEl) roleEl.textContent = currentUser.role;
        
        const avatarEl = document.querySelector('.avatar');
        if (avatarEl && currentUser.name) {
            const cleanName = removeAccents(currentUser.name).trim();
            avatarEl.onerror = function() { this.src = 'https://ui-avatars.com/api/?name=' + encodeURIComponent(currentUser.name); };
            avatarEl.src = 'assets/src/img/' + cleanName + '.png';
        }

        if (currentUser.role === 'Admin' || currentUser.role === 'Supervisor') {
            document.querySelectorAll('.view-panel').forEach(v => v.style.display = 'none');
            const vDash = document.getElementById('view-dashboard');
            const nDash = document.getElementById('navDashboard');
            if(vDash) vDash.style.display = 'block';
            if(nDash) nDash.classList.add('active');
            loadDashboardStats();
        }
    }

    document.querySelectorAll('.nav-item').forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
            item.classList.add('active');
            document.querySelectorAll('.view-panel').forEach(v => v.style.display = 'none');
            if(item.textContent.includes('Inicio')) { document.getElementById('view-dashboard').style.display = 'block'; loadDashboardStats(); }
            else if(item.textContent.includes('Tareas')) document.getElementById('view-workspace').style.display = 'block';
            else if(item.textContent.includes('Horario')) document.getElementById('view-horario').style.display = 'block';
            else if(item.textContent.includes('Teletrabajo')) document.getElementById('view-teletrabajo').style.display = 'block';
        });
    });
}

window.onload = initApp;


window.handleEndShift = function() {
    if(confirm("�Seguro que deseas salir?")) {
        localStorage.removeItem("riskOps_currentUser");
        window.location.href = "login.html";
    }
};

async function loadDashboardStats() {
    try {
        const teleEl = document.getElementById('stat-teletrabajo');
        if(teleEl) teleEl.textContent = '...';
        const tasksEl = document.getElementById('stat-tareas');
        if(tasksEl) tasksEl.textContent = '0%';
        const supEl = document.getElementById('stat-supervisores');
        if(supEl) supEl.textContent = '1';
        const turnoEl = document.getElementById('stat-turno');
        if(turnoEl && currentUser) turnoEl.textContent = currentUser.shift;
    } catch(e) { console.error(e); }
}

window.toggleNotifications = function() {
    const drop = document.getElementById("notificationDropdown");
    if(drop) drop.style.display = drop.style.display === "none" ? "block" : "none";
};

window.openProfileModal = function() {
    const modal = document.getElementById("profileModal");
    if(modal) modal.classList.add("active");
};

window.closeModal = function(id) {
    const modal = document.getElementById(id);
    if(modal) modal.classList.remove("active");
};

window.markAllAsRead = function() {
    const count = document.getElementById("notificationCount");
    if(count) count.style.display = "none";
};

