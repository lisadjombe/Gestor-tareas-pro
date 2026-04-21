from flask import Flask, render_template_string, request, redirect, url_for
import json
import os
from datetime import datetime

app = Flask(__name__)
ARCHIVO = "tareas.json"

def cargar_tareas():
    if os.path.exists(ARCHIVO):
        try:
            with open(ARCHIVO, 'r') as f:
                contenido = f.read().strip()
                if contenido:
                    return json.loads(contenido)
        except:
            pass
    return []

def guardar_tareas(tareas):
    with open(ARCHIVO, 'w') as f:
        json.dump(tareas, f, indent=4, ensure_ascii=False)

def verificar_urgencia(tareas):
    """Añade flag de urgente a tareas que vencen en 24h"""
    hoy = datetime.now().date()
    for tarea in tareas:
        if tarea.get('fecha_limite') and not tarea.get('completada', False):
            try:
                fecha_lim = datetime.strptime(tarea['fecha_limite'], '%Y-%m-%d').date()
                dias = (fecha_lim - hoy).days
                tarea['urgente'] = dias <= 1
                tarea['vencida'] = dias < 0
                tarea['dias_restantes'] = dias
            except:
                tarea['urgente'] = False
                tarea['vencida'] = False
        else:
            tarea['urgente'] = False
            tarea['vencida'] = False
    return tareas

HTML = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=yes">
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
            --shadow: 0 10px 30px rgba(0,0,0,0.15);
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
            --shadow: 0 10px 30px rgba(0,0,0,0.4);
        }
        
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            transition: background-color 0.3s, border-color 0.3s, color 0.2s;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, sans-serif;
            background: linear-gradient(135deg, var(--bg-gradient-start) 0%, var(--bg-gradient-end) 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 16px;
            margin: 0;
        }
        
        .container {
            background: var(--card-bg);
            border-radius: 24px;
            box-shadow: var(--shadow);
            padding: 24px;
            width: 100%;
            max-width: 600px;
            margin: 0 auto;
        }
        
        .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 24px;
            flex-wrap: wrap;
            gap: 12px;
        }
        
        h1 {
            color: var(--text-primary);
            font-size: clamp(1.5rem, 5vw, 2rem);
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .theme-toggle {
            background: var(--task-bg);
            border: none;
            font-size: 24px;
            padding: 10px 16px;
            border-radius: 40px;
            cursor: pointer;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border: 1px solid var(--input-border);
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
            font-size: clamp(1.2rem, 4vw, 1.8rem);
            font-weight: bold;
            color: var(--text-primary);
        }
        
        .stat-label {
            font-size: clamp(0.7rem, 2.5vw, 0.85rem);
            color: var(--text-secondary);
        }
        
        .form-container {
            background: var(--task-bg);
            border-radius: 16px;
            padding: 16px;
            margin-bottom: 20px;
        }
        
        .form-row {
            display: flex;
            flex-direction: column;
            gap: 12px;
        }
        
        .input-group {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
        }
        
        .input-group input[type="text"] {
            flex: 2;
            min-width: 180px;
            padding: 14px 16px;
            border: 2px solid var(--input-border);
            border-radius: 14px;
            font-size: 16px;
            background: var(--input-bg);
            color: var(--text-primary);
            outline: none;
        }
        
        .input-group input[type="date"] {
            flex: 1;
            min-width: 140px;
            padding: 14px 16px;
            border: 2px solid var(--input-border);
            border-radius: 14px;
            font-size: 14px;
            background: var(--input-bg);
            color: var(--text-primary);
            outline: none;
        }
        
        .btn-primary {
            padding: 14px 24px;
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
            gap: 12px;
            max-height: 400px;
            overflow-y: auto;
            padding-right: 4px;
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
            word-break: break-word;
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
            flex-shrink: 0;
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
        
        .modal {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.6);
            backdrop-filter: blur(4px);
            justify-content: center;
            align-items: center;
            z-index: 1000;
            padding: 16px;
        }
        
        .modal-content {
            background: var(--card-bg);
            padding: 24px;
            border-radius: 24px;
            width: 100%;
            max-width: 450px;
        }
        
        .modal-content h3 {
            color: var(--text-primary);
            margin-bottom: 20px;
        }
        
        .modal-input {
            width: 100%;
            padding: 14px 16px;
            border: 2px solid var(--input-border);
            border-radius: 14px;
            font-size: 16px;
            background: var(--input-bg);
            color: var(--text-primary);
            margin-bottom: 16px;
            outline: none;
        }
        
        .modal-actions {
            display: flex;
            gap: 12px;
        }
        
        .modal-actions button {
            flex: 1;
            padding: 14px;
            border: none;
            border-radius: 14px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
        }
        
        .btn-save {
            background: var(--success-color);
            color: white;
        }
        
        .btn-cancel {
            background: var(--task-bg);
            color: var(--text-primary);
        }
        
        .empty-state {
            text-align: center;
            padding: 40px 20px;
            color: var(--text-secondary);
        }
        
        .empty-state p:first-child {
            font-size: 48px;
            margin-bottom: 16px;
        }
        
        @media (max-width: 480px) {
            body {
                padding: 8px;
                align-items: flex-start;
            }
            
            .container {
                padding: 16px;
            }
            
            .input-group input[type="text"],
            .input-group input[type="date"] {
                width: 100%;
            }
            
            .task-item {
                flex-direction: column;
                align-items: stretch;
            }
            
            .task-actions {
                justify-content: flex-end;
                margin-top: 8px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📋 Mis Tareas</h1>
            <button class="theme-toggle" onclick="toggleTheme()" id="themeBtn">🌙</button>
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
            <form action="/añadir" method="POST" class="form-row">
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
                <div class="task-item {% if tarea.urgente %}urgente{% endif %} {% if tarea.completada %}completada{% endif %}" data-indice="{{ loop.index0 }}">
                    <div class="task-content" ondblclick="abrirModal({{ loop.index0 }})">
                        <div class="task-description">{{ tarea.descripcion }}</div>
                        <div class="task-meta">
                            {% if tarea.fecha_limite %}
                            <span>📅 {{ tarea.fecha_limite }}</span>
                            {% if tarea.urgente and not tarea.completada %}
                            <span class="urgent-badge">⚡ ¡URGENTE!</span>
                            {% endif %}
                            {% endif %}
                            <span>🕐 {{ tarea.fecha[:10] }}</span>
                        </div>
                    </div>
                    <div class="task-actions">
                        <a href="/completar/{{ loop.index0 }}" class="btn-icon btn-complete">
                            {% if tarea.completada %}↩️{% else %}✅{% endif %}
                        </a>
                        <button class="btn-icon btn-edit" onclick="abrirModal({{ loop.index0 }})">✏️</button>
                        <a href="/eliminar/{{ loop.index0 }}" class="btn-icon btn-delete" onclick="return confirm('¿Eliminar?')">🗑️</a>
                    </div>
                </div>
                {% endfor %}
            {% else %}
                <div class="empty-state">
                    <p>🎉</p>
                    <p>¡No hay tareas!</p>
                    <p style="font-size: 14px; margin-top: 8px;">Añade una nueva tarea arriba</p>
                </div>
            {% endif %}
        </div>
    </div>
    
    <div class="modal" id="editModal">
        <div class="modal-content">
            <h3>✏️ Editar Tarea</h3>
            <form id="editForm" action="/editar" method="POST">
                <input type="hidden" name="indice" id="editIndice">
                <input type="text" name="descripcion" id="editDescripcion" class="modal-input" placeholder="Descripción" required>
                <input type="date" name="fecha_limite" id="editFecha" class="modal-input">
                <div class="modal-actions">
                    <button type="submit" class="btn-save">💾 Guardar</button>
                    <button type="button" class="btn-cancel" onclick="cerrarModal()">❌ Cancelar</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
        function toggleTheme() {
            const body = document.body;
            const currentTheme = body.getAttribute('data-theme');
            const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
            body.setAttribute('data-theme', newTheme);
            document.getElementById('themeBtn').textContent = newTheme === 'dark' ? '☀️' : '🌙';
            localStorage.setItem('theme', newTheme);
        }
        
        const savedTheme = localStorage.getItem('theme') || 'light';
        document.body.setAttribute('data-theme', savedTheme);
        document.getElementById('themeBtn').textContent = savedTheme === 'dark' ? '☀️' : '🌙';
        
        const tareasData = {{ tareas | tojson }};
        
        function abrirModal(indice) {
            const tarea = tareasData[indice];
            document.getElementById('editIndice').value = indice;
            document.getElementById('editDescripcion').value = tarea.descripcion;
            document.getElementById('editFecha').value = tarea.fecha_limite || '';
            document.getElementById('editModal').style.display = 'flex';
        }
        
        function cerrarModal() {
            document.getElementById('editModal').style.display = 'none';
        }
        
        document.getElementById('editModal').addEventListener('click', function(e) {
            if (e.target === this) cerrarModal();
        });
        
        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape') cerrarModal();
        });
        
        const today = new Date().toISOString().split('T')[0];
        document.querySelectorAll('input[type="date"]').forEach(input => {
            if (!input.value) input.value = today;
            input.min = today;
        });
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    tareas = cargar_tareas()
    tareas = verificar_urgencia(tareas)
    
    # Ordenar: No completadas primero, luego urgentes
    tareas.sort(key=lambda x: (
        x.get('completada', False),
        not x.get('urgente', False)
    ))
    
    total = len(tareas)
    completadas = sum(1 for t in tareas if t.get('completada', False))
    urgentes = sum(1 for t in tareas if t.get('urgente', False) and not t.get('completada', False))
    
    hoy = datetime.now().strftime('%Y-%m-%d')
    
    return render_template_string(HTML, 
                                 tareas=tareas, 
                                 total=total, 
                                 completadas=completadas,
                                 urgentes=urgentes,
                                 hoy=hoy)

@app.route('/añadir', methods=['POST'])
def añadir():
    descripcion = request.form.get('descripcion')
    fecha_limite = request.form.get('fecha_limite', '')
    
    if descripcion:
        tareas = cargar_tareas()
        tareas.append({
            "descripcion": descripcion,
            "completada": False,
            "fecha": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "fecha_limite": fecha_limite if fecha_limite else None
        })
        guardar_tareas(tareas)
    
    return redirect(url_for('index'))

@app.route('/completar/<int:indice>')
def completar(indice):
    tareas = cargar_tareas()
    if 0 <= indice < len(tareas):
        tareas[indice]['completada'] = not tareas[indice]['completada']
        guardar_tareas(tareas)
    return redirect(url_for('index'))

@app.route('/eliminar/<int:indice>')
def eliminar(indice):
    tareas = cargar_tareas()
    if 0 <= indice < len(tareas):
        tareas.pop(indice)
        guardar_tareas(tareas)
    return redirect(url_for('index'))

@app.route('/editar', methods=['POST'])
def editar():
    indice = int(request.form.get('indice', -1))
    descripcion = request.form.get('descripcion')
    fecha_limite = request.form.get('fecha_limite', '')
    
    tareas = cargar_tareas()
    if 0 <= indice < len(tareas):
        if descripcion:
            tareas[indice]['descripcion'] = descripcion
        tareas[indice]['fecha_limite'] = fecha_limite if fecha_limite else None
        guardar_tareas(tareas)
    
    return redirect(url_for('index'))

if __name__ == '__main__':
    if not os.path.exists(ARCHIVO):
        guardar_tareas([])
    print("🚀 Servidor iniciado en http://127.0.0.1:5000")
    app.run(debug=True)