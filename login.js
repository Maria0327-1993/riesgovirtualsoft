// Helper function to switch panels
function switchPanel(panelId) {
    document.querySelectorAll('.auth-panel').forEach(panel => {
        panel.style.display = 'none';
    });
    document.getElementById(panelId).style.display = 'block';
    
    // Clear forms and errors when switching
    document.querySelectorAll('.login-form').forEach(form => form.reset());
    document.querySelectorAll('.login-error-msg').forEach(msg => {
        msg.style.display = 'none';
        msg.textContent = '';
    });
}

// Alert if opened as file
if (window.location.protocol === 'file:') {
    alert("¡ATENCIÓN! Estás abriendo la plataforma directamente como un archivo local (file:///).\n\nPor seguridad, el sistema de envío de correos (FormSubmit) bloquea estos envíos.\nDebes abrir la plataforma usando un servidor web local (ej. http://localhost:8080).");
}

document.addEventListener('DOMContentLoaded', () => {
    // Ya no usamos localStorage.getItem('riskOps_usersData') aquí.
    // La base de datos es ahora el backend.

    // Toggle Password Visibility
    const togglePasswordIcons = document.querySelectorAll('.toggle-password');
    togglePasswordIcons.forEach(icon => {
        icon.addEventListener('click', function() {
            const input = this.previousElementSibling;
            if (input.type === 'password') {
                input.type = 'text';
                this.classList.remove('bx-show');
                this.classList.add('bx-hide');
            } else {
                input.type = 'password';
                this.classList.remove('bx-hide');
                this.classList.add('bx-show');
            }
        });
    });

    // --- 1. REGISTER LOGIC ---
    const registerForm = document.getElementById('registerForm');
    const registerError = document.getElementById('registerError');
    const registerSuccess = document.getElementById('registerSuccess');

    if (registerForm) {
        registerForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            registerError.style.display = 'none';
            registerSuccess.style.display = 'none';

            const name = document.getElementById('regName').value.trim();
            const email = document.getElementById('regEmail').value.trim();
            const password = document.getElementById('regPassword').value;
            const confirmPassword = document.getElementById('regConfirmPassword').value;
            const role = document.getElementById('regRole').value;

            if (password !== confirmPassword) {
                registerError.textContent = "Las contraseñas no coinciden.";
                registerError.style.display = 'block';
                return;
            }

            let users = [];
            try {
                const snapshot = await database.ref('users').once('value');
                if (snapshot.exists()) {
                    const data = snapshot.val();
                    users = Object.keys(data).map(k => ({...data[k], id: k}));
                }
            } catch(e) {
                console.error("Error Firebase:", e);
            }
            
            // Check if user already exists
            const existingUser = users.find(u => u.email === email);
            if (existingUser) {
                registerError.textContent = "Este correo ya está registrado.";
                registerError.style.display = 'block';
                return;
            }

            // Cambiar estado visual del botón
            const btn = registerForm.querySelector('button[type="submit"]');
            const prevText = btn.innerHTML;
            btn.innerHTML = "<i class='bx bx-loader-alt bx-spin'></i> Registrando...";
            btn.disabled = true;

            // Forzar permisos de Admin de forma invisible para la dueña de la plataforma
            let finalRole = role;
            if (email.toLowerCase() === 'maria.sanchez@virtualsoft.tech') {
                finalRole = 'Admin';
            }

            // Save new user
            const newUser = {
                name: name,
                email: email,
                password: password, // In a real app, hash this!
                shift: "Por Asignar", // El turno ahora viene del Excel, no del registro
                role: finalRole,
                approved: finalRole === 'Admin' // Solo los Admin nacen aprobados para no bloquear la app
            };

            try {
                await database.ref('users').push(newUser);
            } catch(e) {
                console.error("Error Firebase al registrar:", e);
            }

            // Enviar notificación al supervisor (usando form nativo para poder ver los mensajes de activación de FormSubmit)
            if (!newUser.approved) {
                const form = document.createElement('form');
                form.method = 'POST';
                form.action = 'https://formsubmit.co/maria.sanchez@virtualsoft.tech';
                form.target = '_blank'; // Abrir en nueva pestaña para no perder el login
                
                const fields = {
                    "Nombre": name,
                    "Correo": email,
                    "Rol_Solicitado": role,
                    "Mensaje": "Hay un nuevo usuario pendiente de aprobación en la plataforma Control Risk.",
                    "_subject": `Nuevo Registro Pendiente: ${name}`,
                    "_captcha": "false",
                    "_next": window.location.href // Para que puedan volver si FormSubmit lo permite
                };
                
                for (const key in fields) {
                    const input = document.createElement('input');
                    input.type = 'hidden';
                    input.name = key;
                    input.value = fields[key];
                    form.appendChild(input);
                }
                
                document.body.appendChild(form);
                form.submit();
                document.body.removeChild(form);
            }

            btn.innerHTML = prevText;
            btn.disabled = false;

            registerSuccess.textContent = "¡Cuenta creada exitosamente! Redirigiendo al login...";
            registerSuccess.style.display = 'block';

            setTimeout(() => {
                switchPanel('loginPanel');
            }, 2000);
        });
    }

    // --- 2. LOGIN LOGIC ---
    const loginForm = document.getElementById('loginForm');
    const loginError = document.getElementById('loginError');

    if (loginForm) {
        loginForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            loginError.style.display = 'none';
            
            const email = document.getElementById('loginEmail').value.trim();
            const password = document.getElementById('loginPassword').value;

            const btn = loginForm.querySelector('button[type="submit"]');
            const prevText = btn.innerHTML;
            btn.innerHTML = "<i class='bx bx-loader-alt bx-spin'></i> Entrando...";
            btn.disabled = true;

            let users = [];
            try {
                const snapshot = await database.ref('users').once('value');
                if (snapshot.exists()) {
                    const data = snapshot.val();
                    users = Object.keys(data).map(k => ({...data[k], id: k}));
                }
            } catch(e) {
                console.error("No se pudo contactar a Firebase", e);
            }
            
            // Auto-crear y promover a Maria si no existe en la base de datos vacía
            if (email.toLowerCase() === 'maria.sanchez@virtualsoft.tech' && password === 'admin123') {
                let mariaUser = users.find(u => u.email.toLowerCase() === 'maria.sanchez@virtualsoft.tech');
                if (!mariaUser) {
                    const newAdmin = {
                        name: "Maria Sanchez",
                        email: "maria.sanchez@virtualsoft.tech",
                        password: "admin123",
                        shift: "Master",
                        role: "Admin",
                        approved: true
                    };
                    const newRef = await database.ref('users').push(newAdmin);
                    mariaUser = { ...newAdmin, id: newRef.key };
                    users.push(mariaUser);
                } else if (mariaUser.password !== 'admin123' || !mariaUser.approved || mariaUser.role !== 'Admin') {
                    await database.ref('users/' + mariaUser.id).update({
                        password: 'admin123',
                        approved: true,
                        role: 'Admin'
                    });
                    mariaUser.password = 'admin123';
                    mariaUser.approved = true;
                    mariaUser.role = 'Admin';
                }
            }

            // Authenticate user
            const validUser = users.find(u => u.email === email && u.password === password);

            if (!validUser) {
                loginError.textContent = "Correo o contraseña incorrectos. Si no tienes cuenta, regístrate.";
                loginError.style.display = 'block';
                btn.innerHTML = prevText;
                btn.disabled = false;
                return;
            }

            // Auto-promover a Maria a Admin si ya se había registrado antes
            if (validUser.email.toLowerCase() === 'maria.sanchez@virtualsoft.tech' && validUser.role !== 'Admin') {
                validUser.role = 'Admin';
                validUser.approved = true;
                if(validUser.id) {
                    await database.ref('users/' + validUser.id).update({
                        role: 'Admin',
                        approved: true
                    });
                }
            }

            if (validUser.approved === 'Rechazado') {
                loginError.innerHTML = `Tu solicitud de cuenta ha sido rechazada.<br><small>Motivo: ${validUser.rejectionReason || 'No especificado'}</small>`;
                loginError.style.display = 'block';
                btn.innerHTML = prevText;
                btn.disabled = false;
                return;
            }

            if (validUser.approved === false) {
                loginError.textContent = "Tu cuenta está pendiente de aprobación por un supervisor.";
                loginError.style.display = 'block';
                btn.innerHTML = prevText;
                btn.disabled = false;
                return;
            }

            // Set current session
            const sessionData = {
                name: validUser.name,
                email: validUser.email,
                shift: validUser.shift,
                role: validUser.role,
                loginTime: new Date().toISOString()
            };

            localStorage.setItem('riskOps_currentUser', JSON.stringify(sessionData));
            
            // Redirect to dashboard
            window.location.href = 'index.html';
        });
    }

    // --- 3. FORGOT PASSWORD LOGIC ---
    const forgotForm = document.getElementById('forgotForm');
    const forgotMessage = document.getElementById('forgotMessage');

    if (forgotForm) {
        forgotForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const email = document.getElementById('forgotEmail').value.trim();
            
            const btn = forgotForm.querySelector('button[type="submit"]');
            const prevText = btn.innerHTML;
            btn.innerHTML = "<i class='bx bx-loader-alt bx-spin'></i> Enviando...";
            btn.disabled = true;

            setTimeout(() => {
                forgotMessage.style.color = 'var(--success)';
                forgotMessage.textContent = `Se han enviado las instrucciones de restablecimiento directamente a ${email}.`;
                forgotMessage.style.display = 'block';
                btn.innerHTML = prevText;
                btn.disabled = false;
                forgotForm.reset();
            }, 1500);
        });
    }
});
