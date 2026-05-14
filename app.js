// Risk Manager - App Logic v58
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
        // Populate Profile
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

        // Module Logic
        if (currentUser.role === 'Admin' || currentUser.role === 'Supervisor') {
            document.querySelectorAll('.view-panel').forEach(v => v.style.display = 'none');
            const vDash = document.getElementById('view-dashboard');
            if(vDash) vDash.style.display = 'block';
            document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
            const nDash = document.getElementById('navDashboard');
            if(nDash) nDash.classList.add('active');
            loadDashboardStats();
        }
    }

    // Nav Listeners
    document.querySelectorAll('.nav-item').forEach(item => {
        item.onclick = (e) => {
            e.preventDefault();
            document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
            item.classList.add('active');
            document.querySelectorAll('.view-panel').forEach(v => v.style.display = 'none');
            
            if(item.textContent.includes('Inicio')) { document.getElementById('view-dashboard').style.display = 'block'; loadDashboardStats(); }
            else if(item.textContent.includes('Tareas')) document.getElementById('view-workspace').style.display = 'block';
            else if(item.textContent.includes('Horario')) document.getElementById('view-horario').style.display = 'block';
            else if(item.textContent.includes('Teletrabajo')) document.getElementById('view-teletrabajo').style.display = 'block';
            else if(item.textContent.includes('Documentación')) document.getElementById('view-docs').style.display = 'block';
            else if(item.textContent.includes('Permisos')) document.getElementById('view-permisos').style.display = 'block';
            else if(item.textContent.includes('Aprobaciones')) document.getElementById('view-aprobaciones').style.display = 'block';
        };
    });
}

async function loadDashboardStats() {
    const teleEl = document.getElementById('stat-teletrabajo');
    if(teleEl) teleEl.textContent = '...';
    // Logic here for stats
}

window.onload = initApp;
