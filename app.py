from flask import Flask, render_template_string, request, redirect, url_for, flash
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
import os

app = Flask(__name__)
app.config['SECRET_KEY'] = 'clave-super-secreta-2024'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///tareas.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

# ========== MODELOS DE BASE DE DATOS ==========
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    tareas = db.relationship('Tarea', backref='owner', lazy=True)

class Tarea(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    descripcion = db.Column(db.String(200), nullable=False)
    completada = db.Column(db.Boolean, default=False)
    fecha_limite = db.Column(db.String(20))
    fecha_creacion = db.Column(db.String(50))
    urgente = db.Column(db.Boolean, default=False)
    vencida = db.Column(db.Boolean, default=False)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# ========== FUNCIONES AUXILIARES ==========
def verificar_urgencia(tareas):
    hoy = datetime.now().date()
    for tarea in tareas:
        if tarea.fecha_limite and not tarea.completada:
            try:
                fecha_lim = datetime.strptime(tarea.fecha_limite, '%Y-%m-%d').date()
                dias = (fecha_lim - hoy).days
                tarea.urgente = dias <= 1
                tarea.vencida = dias < 0
            except:
                tarea.urgente = False
                tarea.vencida = False
        else:
            tarea.urgente = False
            tarea.vencida = False
    return tareas

# ========== PLANTILLA HTML ==========
HTML = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>📋 Mis Tareas PRO</title>
    <style>
        :root {
            --bg-gradient-start: #667eea;
            --bg-gradient-end: #764ba2;
            --card-bg: #ffffff;
            --text-primary: #1a1a1a;
            --text-secondary: #666666;
            --input-bg: #ffffff;
            --input-border: #e0e0e0;
            --task-bg: #f8f9fa;
            --task-completed: #e9ecef;
            --urgent-color: #e74c3c;
            --success-color: #00b894;
            --edit-color: #3498db;
            --delete-color: #e74c3c;
        }
        
        [data-theme="dark"] {
            --bg-gradient-start: #1a1a2e;
            --bg-gradient-end: #16213e;
            --card-bg: #2d2d44;
            --text-primary: #ffffff;
            --text-secondary: #b0b0c0;
            --input-bg: #3a3a5c;
            --input-border: #4a4a6a;
            --task-bg: #3a3a5c;
            --task-completed: #2a2a4a;
        }
        
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            transition: background-color 0.3s, color 0.2s;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: linear-gradient(135deg, var(--bg-gradient-start) 0%, var(--bg-gradient-end) 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 16px;
        }
        
        .container {
            background: var(--card-bg);
            border-radius: 24px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.15);
            padding: 24px;
            width: 100%;
            max-width: 600px;
        }
        
        .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            flex-wrap: wrap;
            gap: 10px;
        }
        
        h1 {
            color: var(--text-primary);
            font-size: 1.8rem;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .user-info {
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .user-badge {
            color: var(--text-primary);
            font-weight: 500;
        }
        
        .btn-logout {
            background: var(--delete-color);
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 20px;
            cursor: pointer;
            text-decoration: none;
            font-size: 14px;
        }
        
        .theme-toggle {
            background: var(--task-bg);
            border: none;
            font-size: 24px;
            padding: 10px 16px;
            border-radius: 40px;
            cursor: pointer;
        }
        
        .stats {
            display: flex;
            justify-content: space-around;
            margin-bottom: 20px;
            padding: 12px;
            background: var(--task-bg);
            border-radius: 16px;
        }
        
        .stat-item {
            text-align: center;
        }
        
        .stat-number {
            font-size: 1.5rem;
            font-weight: bold;
            color: var(--text-primary);
        }
        
        .stat-label {
            font-size: 0.8rem;
            color: var(--text-secondary);
        }
        
        .form-container {
            background: var(--task-bg);
            border-radius: 16px;
            padding: 16px;
            margin-bottom: 20px;
        }
        
        .input-group {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-bottom: 12px;
        }
        
        .input-group input[type="text"] {
            flex: 2;
            min-width: 180px;
            padding: 12px 16px;
            border: 2px solid var(--input-border);
            border-radius: 14px;
            font-size: 16px;
            background: var(--input-bg);
            color: var(--text-primary);
        }
        
        .input-group input[type="date"] {
            flex: 1;
            min-width: 130px;
            padding: 12px 16px;
            border: 2px solid var(--input-border);
            border-radius: 14px;
            background: var(--input-bg);
            color: var(--text-primary);
        }
        
        .btn-primary {
            padding: 12px 24px;
            background: var(--success-color);
            color: white;
            border: none;
            border-radius: 14px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            width: 100%;
        }
        
        .task-list {
            display: flex;
            flex-direction: column;
            gap: 10px;
            max-height: 400px;
            overflow-y: auto;
        }
        
        .task-item {
            background: var(--task-bg);
            border-radius: 16px;
            padding: 16px;
            display: flex;
            align-items: flex-start;
            justify-content: space-between;
            gap: 12px;
            border-left: 5px solid transparent;
        }
        
        .task-item.urgente {
            border-left-color: var(--urgent-color);
            animation: pulse 2s infinite;
        }
        
        @keyframes pulse {
            0%, 100% { box-shadow: 0 0 0 0 rgba(231, 76, 60, 0.3); }
            50% { box-shadow: 0 0 0 6px rgba(231, 76, 60, 0); }
        }
        
        .task-item.completada {
            opacity: 0.6;
            background: var(--task-completed);
        }
        
        .task-content {
            flex: 1;
            cursor: pointer;
        }
        
        .task-description {
            font-size: 16px;
            font-weight: 500;
            color: var(--text-primary);
            margin-bottom: 6px;
        }
        
        .completada .task-description {
            text-decoration: line-through;
        }
        
        .task-meta {
            display: flex;
            flex-wrap: wrap;
            gap: 8px 16px;
            font-size: 12px;
            color: var(--text-secondary);
        }
        
        .urgent-badge {
            background: var(--urgent-color);
            color: white;
            padding: 3px 10px;
            border-radius: 20px;
            font-size: 11px;
            font-weight: bold;
        }
        
        .task-actions {
            display: flex;
            gap: 6px;
        }
        
        .btn-icon {
            width: 40px;
            height: 40px;
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            text-decoration: none;
            font-size: 18px;
            border: none;
            cursor: pointer;
            background: var(--input-bg);
            color: var(--text-primary);
        }
        
        .btn-complete {
            background: var(--success-color);
            color: white;
        }
        
        .btn-edit {
            background: var(--edit-color);
            color: white;
        }
        
        .btn-delete {
            background: var(--delete-color);
            color: white;
        }
        
        .empty-state {
            text-align: center;
            padding: 40px 20px;
            color: var(--text-secondary);
        }
        
        /* LOGIN STYLES */
        .login-container {
            max-width: 400px;
        }
        
        .login-form {
            display: flex;
            flex-direction: column;
            gap: 16px;
        }
        
        .login-form input {
            padding: 14px 16px;
            border: 2px solid var(--input-border);
            border-radius: 14px;
            font-size: 16px;
            background: var(--input-bg);
            color: var(--text-primary);
        }
        
        .login-link {
            text-align: center;
            margin-top: 16px;
            color: var(--text-secondary);
        }
        
        .login-link a {
            color: var(--success-color);
            text-decoration: none;
        }
        
        .flash-message {
            background: var(--urgent-color);
            color: white;
            padding: 12px;
            border-radius: 12px;
            margin-bottom: 16px;
        }
        
        @media (max-width: 480px) {
            .container { padding: 16px; }
            .task-item { flex-direction: column; }
            .task-actions { justify-content: flex-end; }
        }
    </style>
</head>
<body>
    <div class="container">
        {% if current_user.is_authenticated %}
        <!-- VISTA PRINCIPAL (USUARIO LOGUEADO) -->
        <div class="header">
            <h1>📋 Mis Tareas</h1>
            <div class="user-info">
                <span class="user-badge">👤 {{ current_user.username }}</span>
                <a href="/logout" class="btn-logout">Salir</a>
                <button class="theme-toggle" onclick="toggleTheme()" id="themeBtn">🌙</button>
            </div>
        </div>
        
        <div class="stats">
            <div class="stat-item">
                <div class="stat-number">{{ total }}</div>
                <div class="stat-label">Total</div>
            </div>
            <div class="stat-item">
                <div class="stat-number">{{ completadas }}</div>
                <div class="stat-label">Completadas</div>
            </div>
            <div class="stat-item">
                <div class="stat-number">{{ urgentes }}</div>
                <div class="stat-label">Urgentes</div>
            </div>
        </div>
        
        <div class="form-container">
            <form action="/añadir" method="POST">
                <div class="input-group">
                    <input type="text" name="descripcion" placeholder="¿Qué tienes que hacer?" required>
                    <input type="date" name="fecha_limite" value="{{ hoy }}">
                </div>
                <button type="submit" class="btn-primary">➕ Añadir Tarea</button>
            </form>
        </div>
        
        <div class="task-list">
            {% if tareas %}
                {% for tarea in tareas %}
                <div class="task-item {% if tarea.urgente %}urgente{% endif %} {% if tarea.completada %}completada{% endif %}" data-id="{{ tarea.id }}">
                    <div class="task-content" ondblclick="abrirModal({{ tarea.id }})">
                        <div class="task-description">{{ tarea.descripcion }}</div>
                        <div class="task-meta">
                            {% if tarea.fecha_limite %}
                            <span>📅 {{ tarea.fecha_limite }}</span>
                            {% if tarea.urgente and not tarea.completada %}
                            <span class="urgent-badge">⚡ ¡URGENTE!</span>
                            {% endif %}
                            {% endif %}
                            <span>🕐 {{ tarea.fecha_creacion[:10] }}</span>
                        </div>
                    </div>
                    <div class="task-actions">
                        <a href="/completar/{{ tarea.id }}" class="btn-icon btn-complete">
                            {% if tarea.completada %}↩️{% else %}✅{% endif %}
                        </a>
                        <button class="btn-icon btn-edit" onclick="abrirModal({{ tarea.id }})">✏️</button>
                        <a href="/eliminar/{{ tarea.id }}" class="btn-icon btn-delete" onclick="return confirm('¿Eliminar?')">🗑️</a>
                    </div>
                </div>
                {% endfor %}
            {% else %}
                <div class="empty-state">
                    <p>🎉 ¡No hay tareas!</p>
                    <p style="font-size: 14px; margin-top: 8px;">Añade una nueva tarea arriba</p>
                </div>
            {% endif %}
        </div>
        
        {% else %}
        <!-- VISTA DE LOGIN / REGISTRO -->
        <h1 style="color: var(--text-primary); margin-bottom: 24px;">🔐 {% if modo == 'login' %}Iniciar Sesión{% else %}Registrarse{% endif %}</h1>
        
        {% with messages = get_flashed_messages() %}
            {% if messages %}
                {% for message in messages %}
                    <div class="flash-message">{{ message }}</div>
                {% endfor %}
            {% endif %}
        {% endwith %}
        
        <form method="POST" class="login-form">
            <input type="text" name="username" placeholder="Usuario" required>
            <input type="password" name="password" placeholder="Contraseña" required>
            <button type="submit" class="btn-primary">{% if modo == 'login' %}Entrar{% else %}Registrarse{% endif %}</button>
        </form>
        
        <div class="login-link">
            {% if modo == 'login' %}
                ¿No tienes cuenta? <a href="/registro">Regístrate aquí</a>
            {% else %}
                ¿Ya tienes cuenta? <a href="/login">Inicia sesión</a>
            {% endif %}
        </div>
        {% endif %}
    </div>
    
    <script>
        function toggleTheme() {
            const body = document.body;
            const current = body.getAttribute('data-theme');
            const newTheme = current === 'dark' ? 'light' : 'dark';
            body.setAttribute('data-theme', newTheme);
            document.getElementById('themeBtn').textContent = newTheme === 'dark' ? '☀️' : '🌙';
            localStorage.setItem('theme', newTheme);
        }
        
        const saved = localStorage.getItem('theme') || 'light';
        document.body.setAttribute('data-theme', saved);
        if (document.getElementById('themeBtn')) {
            document.getElementById('themeBtn').textContent = saved === 'dark' ? '☀️' : '🌙';
        }
        
        const today = new Date().toISOString().split('T')[0];
        document.querySelectorAll('input[type="date"]').forEach(i => {
            if (!i.value) i.value = today;
            i.min = today;
        });
    </script>
</body>
</html>
"""

# ========== RUTAS DE AUTENTICACIÓN ==========
@app.route('/')
def index():
    if current_user.is_authenticated:
        return redirect(url_for('tareas'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('tareas'))
    
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        user = User.query.filter_by(username=username).first()
        
        if user and check_password_hash(user.password_hash, password):
            login_user(user)
            return redirect(url_for('tareas'))
        else:
            flash('Usuario o contraseña incorrectos')
    
    return render_template_string(HTML, modo='login', current_user=current_user)

@app.route('/registro', methods=['GET', 'POST'])
def registro():
    if current_user.is_authenticated:
        return redirect(url_for('tareas'))
    
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        if User.query.filter_by(username=username).first():
            flash('El usuario ya existe')
        else:
            user = User(
                username=username,
                password_hash=generate_password_hash(password)
            )
            db.session.add(user)
            db.session.commit()
            login_user(user)
            return redirect(url_for('tareas'))
    
    return render_template_string(HTML, modo='registro', current_user=current_user)

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

# ========== RUTAS DE TAREAS ==========
@app.route('/tareas')
@login_required
def tareas():
    tareas = Tarea.query.filter_by(user_id=current_user.id).all()
    tareas = verificar_urgencia(tareas)
    
    tareas.sort(key=lambda x: (x.completada, not x.urgente))
    
    total = len(tareas)
    completadas = sum(1 for t in tareas if t.completada)
    urgentes = sum(1 for t in tareas if t.urgente and not t.completada)
    hoy = datetime.now().strftime('%Y-%m-%d')
    
    return render_template_string(HTML, 
                                 tareas=tareas, 
                                 total=total, 
                                 completadas=completadas,
                                 urgentes=urgentes,
                                 hoy=hoy,
                                 current_user=current_user)

@app.route('/añadir', methods=['POST'])
@login_required
def añadir():
    descripcion = request.form.get('descripcion')
    fecha_limite = request.form.get('fecha_limite', '')
    
    if descripcion:
        tarea = Tarea(
            descripcion=descripcion,
            fecha_limite=fecha_limite if fecha_limite else None,
            fecha_creacion=datetime.now().strftime("%d/%m/%Y %H:%M"),
            user_id=current_user.id
        )
        db.session.add(tarea)
        db.session.commit()
    
    return redirect(url_for('tareas'))

@app.route('/completar/<int:id>')
@login_required
def completar(id):
    tarea = Tarea.query.get_or_404(id)
    if tarea.user_id == current_user.id:
        tarea.completada = not tarea.completada
        db.session.commit()
    return redirect(url_for('tareas'))

@app.route('/eliminar/<int:id>')
@login_required
def eliminar(id):
    tarea = Tarea.query.get_or_404(id)
    if tarea.user_id == current_user.id:
        db.session.delete(tarea)
        db.session.commit()
    return redirect(url_for('tareas'))

# ========== INICIALIZACIÓN ==========
with app.app_context():
    db.create_all()

if __name__ == '__main__':
    app.run(debug=True)