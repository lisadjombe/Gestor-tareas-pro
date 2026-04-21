def min(a, b):
    return a if a < b else b
from flask import Flask, render_template_string, request, redirect, url_for, session, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, timedelta
from functools import wraps
import secrets
import io
import json   # <--- AÑADE ESTO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

app = Flask(__name__)
app.config['SECRET_KEY'] = 'oficina-secreta-2024'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///oficina.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# ========== MODELOS ==========

class Usuario(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    nombre_completo = db.Column(db.String(100), nullable=False)
    departamento = db.Column(db.String(50), default='General')
    cargo = db.Column(db.String(50), default='Empleado')
    rol = db.Column(db.String(20), default='empleado')
    fecha_registro = db.Column(db.String(20), default=datetime.now().strftime("%d/%m/%Y"))
    tareas = db.relationship('Tarea', backref='asignado', lazy=True, foreign_keys='Tarea.usuario_id')
    reset_token = db.Column(db.String(100), nullable=True)
reset_token_expira = db.Column(db.String(20), nullable=True)

class Tarea(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    titulo = db.Column(db.String(200), nullable=False)
    descripcion = db.Column(db.Text)
    fecha_limite = db.Column(db.String(20))
    fecha_creacion = db.Column(db.String(20), default=datetime.now().strftime("%d/%m/%Y"))
    completada = db.Column(db.Boolean, default=False)
    prioridad = db.Column(db.String(20), default='media')
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    comentarios = db.relationship('Comentario', backref='tarea', lazy=True, cascade='all, delete-orphan')

class TareaPersonal(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    titulo = db.Column(db.String(200), nullable=False)
    descripcion = db.Column(db.Text)
    fecha_limite = db.Column(db.String(20))
    fecha_creacion = db.Column(db.String(20), default=datetime.now().strftime("%d/%m/%Y"))
    completada = db.Column(db.Boolean, default=False)
    prioridad = db.Column(db.String(20), default='media')
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    usuario = db.relationship('Usuario', backref='tareas_personales')

class NotaPersonal(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    titulo = db.Column(db.String(200), nullable=False)
    contenido = db.Column(db.Text)
    color = db.Column(db.String(20), default='#f39c12')
    fecha_creacion = db.Column(db.String(20), default=datetime.now().strftime("%d/%m/%Y %H:%M"))
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    usuario = db.relationship('Usuario', backref='notas_personales')

class Comentario(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    texto = db.Column(db.Text, nullable=False)
    fecha = db.Column(db.String(20), default=datetime.now().strftime("%d/%m/%Y %H:%M"))
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    tarea_id = db.Column(db.Integer, db.ForeignKey('tarea.id'), nullable=False)
    usuario = db.relationship('Usuario', backref='comentarios')

class Mensaje(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    contenido = db.Column(db.Text, nullable=False)
    fecha = db.Column(db.String(20), default=datetime.now().strftime("%d/%m/%Y %H:%M"))
    leido = db.Column(db.Boolean, default=False)
    emisor_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    receptor_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    emisor = db.relationship('Usuario', foreign_keys=[emisor_id], backref='mensajes_enviados')
    receptor = db.relationship('Usuario', foreign_keys=[receptor_id], backref='mensajes_recibidos')

class Cliente(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(100))
    telefono = db.Column(db.String(20))
    empresa = db.Column(db.String(100))
    direccion = db.Column(db.String(200))
    notas = db.Column(db.Text)
    fecha_creacion = db.Column(db.String(20), default=datetime.now().strftime("%d/%m/%Y"))
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    usuario = db.relationship('Usuario', backref='clientes')

class Interaccion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    tipo = db.Column(db.String(50), default='Nota')  # Llamada, Email, Reunión, Nota
    descripcion = db.Column(db.Text, nullable=False)
    fecha = db.Column(db.String(20), default=datetime.now().strftime("%d/%m/%Y %H:%M"))
    cliente_id = db.Column(db.Integer, db.ForeignKey('cliente.id'), nullable=False)
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    cliente = db.relationship('Cliente', backref='interacciones')
    usuario = db.relationship('Usuario', backref='interacciones')

class MensajeGrupal(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    contenido = db.Column(db.Text, nullable=False)
    fecha = db.Column(db.String(20), default=datetime.now().strftime("%d/%m/%Y %H:%M"))
    departamento = db.Column(db.String(50), nullable=False)
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    usuario = db.relationship('Usuario', backref='mensajes_grupales')

# ========== FUNCIONES AUXILIARES ==========

# ========== RECUPERAR CONTRASEÑA ==========
@app.route('/olvide-password', methods=['GET', 'POST'])
def olvide_password():
    if request.method == 'POST':
        username = request.form.get('username')
        user = Usuario.query.filter_by(username=username).first()
        
        if user:
            token = secrets.token_urlsafe(32)
            user.reset_token = token
            user.reset_token_expira = (datetime.now() + timedelta(hours=1)).strftime("%d/%m/%Y %H:%M")
            db.session.commit()
            
            # En producción enviarías email. Aquí mostramos el link
            reset_url = url_for('reset_password', token=token, _external=True)
            flash(f'🔗 Link de recuperación: {reset_url}')
        else:
            flash('❌ Usuario no encontrado')
    
    content = """
        <h2>🔐 Recuperar Contraseña</h2>
        <form method="POST" style="max-width: 400px;">
            <div class="form-group">
                <label>Usuario</label>
                <input type="text" name="username" class="form-control" required>
            </div>
            <button type="submit" class="btn btn-primary">Enviar Link</button>
        </form>
        <p><a href="/login">← Volver al login</a></p>
    """
    return base_html(content, "Recuperar Contraseña")

@app.route('/reset-password/<token>', methods=['GET', 'POST'])
def reset_password(token):
    user = Usuario.query.filter_by(reset_token=token).first()
    
    if not user:
        flash('❌ Link inválido o expirado')
        return redirect('/login')
    
    expira = datetime.strptime(user.reset_token_expira, "%d/%m/%Y %H:%M")
    if datetime.now() > expira:
        flash('❌ El link ha expirado')
        return redirect('/login')
    
    if request.method == 'POST':
        password = request.form.get('password')
        user.password_hash = generate_password_hash(password, method='pbkdf2:sha256')
        user.reset_token = None
        user.reset_token_expira = None
        db.session.commit()
        flash('✅ Contraseña actualizada. Ya puedes iniciar sesión.')
        return redirect('/login')
    
    content = """
        <h2>🔐 Nueva Contraseña</h2>
        <form method="POST" style="max-width: 400px;">
            <div class="form-group">
                <label>Nueva contraseña</label>
                <input type="password" name="password" class="form-control" required>
            </div>
            <button type="submit" class="btn btn-primary">Guardar</button>
        </form>
    """
    return base_html(content, "Nueva Contraseña")

# ========== EXPORTAR A EXCEL ==========
@app.route('/admin/exportar')
def exportar_excel():
    if 'user_id' not in session or session.get('rol') != 'admin':
        return redirect('/dashboard')
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Tareas"
    
    # Estilos
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1a73e8", end_color="1a73e8", fill_type="solid")
    
    headers = ['ID', 'Título', 'Descripción', 'Asignado a', 'Departamento', 'Prioridad', 'Fecha Límite', 'Estado', 'Fecha Creación']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
    
    tareas = Tarea.query.all()
    for row, t in enumerate(tareas, 2):
        ws.cell(row=row, column=1, value=t.id)
        ws.cell(row=row, column=2, value=t.titulo)
        ws.cell(row=row, column=3, value=t.descripcion or '')
        ws.cell(row=row, column=4, value=t.asignado.nombre_completo)
        ws.cell(row=row, column=5, value=t.asignado.departamento)
        ws.cell(row=row, column=6, value=t.prioridad)
        ws.cell(row=row, column=7, value=t.fecha_limite or '')
        ws.cell(row=row, column=8, value='Completada' if t.completada else 'Pendiente')
        ws.cell(row=row, column=9, value=t.fecha_creacion)
    
    # Ajustar ancho de columnas
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column].width = adjusted_width
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'tareas_{datetime.now().strftime("%Y%m%d")}.xlsx'
    )

# ========== FILTROS Y BÚSQUEDA ==========
@app.route('/admin/tareas-filtro')
def tareas_filtro():
    if 'user_id' not in session or session.get('rol') != 'admin':
        return redirect('/dashboard')
    
    estado = request.args.get('estado', 'todas')
    prioridad = request.args.get('prioridad', 'todas')
    departamento = request.args.get('departamento', 'todos')
    busqueda = request.args.get('busqueda', '')
    
    query = Tarea.query
    
    if estado == 'pendientes':
        query = query.filter_by(completada=False)
    elif estado == 'completadas':
        query = query.filter_by(completada=True)
    
    if prioridad != 'todas':
        query = query.filter_by(prioridad=prioridad)
    
    if departamento != 'todos':
        query = query.join(Usuario).filter(Usuario.departamento == departamento)
    
    if busqueda:
        query = query.filter(Tarea.titulo.contains(busqueda) | Tarea.descripcion.contains(busqueda))
    
    tareas = query.all()
    departamentos = db.session.query(Usuario.departamento).distinct().all()
    
    # Construir HTML con filtros y resultados
    filtros_html = f"""
    <div style="background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
        <form method="GET" style="display: flex; gap: 10px; flex-wrap: wrap;">
            <input type="text" name="busqueda" placeholder="🔍 Buscar..." value="{busqueda}" style="padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
            <select name="estado" style="padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                <option value="todas" {'selected' if estado == 'todas' else ''}>Todos los estados</option>
                <option value="pendientes" {'selected' if estado == 'pendientes' else ''}>Pendientes</option>
                <option value="completadas" {'selected' if estado == 'completadas' else ''}>Completadas</option>
            </select>
            <select name="prioridad" style="padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                <option value="todas" {'selected' if prioridad == 'todas' else ''}>Todas las prioridades</option>
                <option value="alta" {'selected' if prioridad == 'alta' else ''}>🔴 Alta</option>
                <option value="media" {'selected' if prioridad == 'media' else ''}>🟡 Media</option>
                <option value="baja" {'selected' if prioridad == 'baja' else ''}>🟢 Baja</option>
            </select>
            <select name="departamento" style="padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                <option value="todos">Todos los departamentos</option>
                {''.join([f'<option value="{d[0]}" {"selected" if departamento == d[0] else ""}>{d[0]}</option>' for d in departamentos])}
            </select>
            <button type="submit" class="btn btn-primary">Filtrar</button>
            <a href="/admin/tareas-filtro" class="btn" style="background: #95a5a6; color: white;">Limpiar</a>
        </form>
    </div>
    """
    
    tareas_html = ""
    for t in tareas:
        prioridad_badge = {'alta': '🔴', 'media': '🟡', 'baja': '🟢'}.get(t.prioridad, '')
        tareas_html += f"""
        <tr>
            <td>{t.id}</td>
            <td>{prioridad_badge} {t.titulo[:50]}...</td>
            <td>{t.asignado.nombre_completo}</td>
            <td>{t.asignado.departamento}</td>
            <td>{t.fecha_limite or '-'}</td>
            <td>{'✅' if t.completada else '⏳'}</td>
            <td>
                <a href="/tarea/{t.id}" class="btn btn-primary btn-sm">👁️ Ver</a>
            </td>
        </tr>
        """
    
    content = f"""
        <h2>🔍 Búsqueda y Filtros</h2>
        {filtros_html}
        
        <p>📊 {len(tareas)} resultados encontrados</p>
        
        <table>
            <thead>
                <tr>
                    <th>ID</th>
                    <th>Título</th>
                    <th>Asignado</th>
                    <th>Depto</th>
                    <th>Fecha Límite</th>
                    <th>Estado</th>
                    <th>Acción</th>
                </tr>
            </thead>
            <tbody>
                {tareas_html if tareas else '<tr><td colspan="7" style="text-align:center;">No hay resultados</td></tr>'}
            </tbody>
        </table>
        
        <p style="margin-top: 20px;">
            <a href="/admin/exportar" class="btn btn-success">📊 Exportar a Excel</a>
            <a href="/admin/calendario" class="btn btn-primary">📅 Ver Calendario</a>
        </p>
    """
    return base_html(content, "Filtros de Tareas")

# ========== VISTA DE TAREA CON COMENTARIOS ==========
@app.route('/tarea/<int:id>', methods=['GET', 'POST'])
def ver_tarea(id):
    if 'user_id' not in session:
        return redirect('/login')
    
    tarea = Tarea.query.get_or_404(id)
    
    # Verificar permisos
    if session.get('rol') != 'admin' and tarea.usuario_id != session['user_id']:
        flash('❌ No tienes permiso para ver esta tarea')
        return redirect('/dashboard')
    
    if request.method == 'POST':
        comentario = request.form.get('comentario')
        if comentario:
            c = Comentario(
                texto=comentario,
                usuario_id=session['user_id'],
                tarea_id=tarea.id
            )
            db.session.add(c)
            db.session.commit()
            flash('✅ Comentario añadido')
        return redirect(f'/tarea/{id}')
    
    comentarios_html = ""
    for c in tarea.comentarios:
        comentarios_html += f"""
        <div style="background: #f8f9fa; padding: 12px; border-radius: 8px; margin-bottom: 10px;">
            <strong>{c.usuario.nombre_completo}</strong> 
            <small style="color: #666;">{c.fecha}</small>
            <p style="margin-top: 5px;">{c.texto}</p>
        </div>
        """
    
    content = f"""
        <h2>📋 {tarea.titulo}</h2>
        
        <div style="background: #ecf0f1; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
            <p><strong>Asignada a:</strong> {tarea.asignado.nombre_completo} ({tarea.asignado.departamento})</p>
            <p><strong>Prioridad:</strong> {tarea.prioridad.upper()}</p>
            <p><strong>Fecha límite:</strong> {tarea.fecha_limite or 'Sin fecha'}</p>
            <p><strong>Estado:</strong> {'✅ Completada' if tarea.completada else '⏳ Pendiente'}</p>
            <p><strong>Descripción:</strong><br>{tarea.descripcion or 'Sin descripción'}</p>
            
            {f'<a href="/tarea/completar/{tarea.id}" class="btn btn-success">✅ Marcar como completada</a>' if not tarea.completada else ''}
        </div>
        
        <h3>💬 Comentarios ({len(tarea.comentarios)})</h3>
        
        <div style="margin-bottom: 20px;">
            {comentarios_html if tarea.comentarios else '<p>No hay comentarios aún</p>'}
        </div>
        
        <form method="POST">
            <div class="form-group">
                <label>Añadir comentario</label>
                <textarea name="comentario" class="form-control" rows="3" required></textarea>
            </div>
            <button type="submit" class="btn btn-primary">💬 Enviar comentario</button>
        </form>
        
        <p style="margin-top: 20px;">
            <a href="/dashboard">← Volver</a>
        </p>
    """
    return base_html(content, f"Tarea: {tarea.titulo[:30]}")

# ========== CALENDARIO SEMANAL ==========
@app.route('/admin/calendario')
def calendario():
    if 'user_id' not in session or session.get('rol') != 'admin':
        return redirect('/dashboard')
    
    # Obtener semana actual
    hoy = datetime.now().date()
    lunes = hoy - timedelta(days=hoy.weekday())
    semana = [lunes + timedelta(days=i) for i in range(7)]
    
    # Obtener tareas de la semana
    tareas_semana = []
    for tarea in Tarea.query.all():
        if tarea.fecha_limite:
            fecha_tarea = datetime.strptime(tarea.fecha_limite, '%Y-%m-%d').date()
            if lunes <= fecha_tarea <= lunes + timedelta(days=6):
                tareas_semana.append(tarea)
    
    # Construir tabla
    dias_html = ""
    for dia in semana:
        tareas_dia = [t for t in tareas_semana if datetime.strptime(t.fecha_limite, '%Y-%m-%d').date() == dia]
        es_hoy = dia == hoy
        
        tareas_html = ""
        for t in tareas_dia:
            prioridad_color = {'alta': '#e74c3c', 'media': '#f39c12', 'baja': '#27ae60'}.get(t.prioridad, '#3498db')
            tareas_html += f"""
            <div style="background: {prioridad_color}; color: white; padding: 5px; margin: 3px 0; border-radius: 4px; font-size: 12px;">
                <a href="/tarea/{t.id}" style="color: white; text-decoration: none;">
                    {t.titulo[:20]}... ({t.asignado.nombre_completo.split()[0]})
                </a>
            </div>
            """
        
        dias_html += f"""
        <td style="vertical-align: top; padding: 10px; border: 1px solid #ddd; {'background: #e8f4f8;' if es_hoy else ''}">
            <strong>{dia.strftime('%a')}<br>{dia.strftime('%d/%m')}</strong>
            {tareas_html if tareas_dia else '<p style="color: #999; font-size: 12px;">Sin tareas</p>'}
        </td>
        """
    
    content = f"""
        <h2>📅 Calendario Semanal</h2>
        <p>Semana del {lunes.strftime('%d/%m/%Y')} al {(lunes + timedelta(days=6)).strftime('%d/%m/%Y')}</p>
        
        <table style="width: 100%; border-collapse: collapse;">
            <tr>
                {dias_html}
            </tr>
        </table>
        
        <p style="margin-top: 20px;">
            <a href="/admin/tareas-filtro" class="btn btn-primary">← Volver a filtros</a>
            <a href="/admin/exportar" class="btn btn-success">📊 Exportar Excel</a>
        </p>
    """
    return base_html(content, "Calendario Semanal")

def base_html(content, titulo="Oficina"):
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>{titulo}</title>
        <style>
            * {{ margin: 0; padding: 0; box-sizing: border-box; }}

            body {{
            font-family: Arial, sans-serif;
            background: #f0f2f5;
            padding: 10px;
            margin: 0;
        }}

        .navbar {{
            background: #2c3e50;
            color: white;
            padding: 12px 15px;
            margin-bottom: 15px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 10px;
        }}

     .navbar a {{
            color: white;
            margin: 0 5px;
            text-decoration: none;
            font-size: 14px;
            white-space: nowrap;
        }}

        .navbar a:hover {{ text-decoration: underline; }}

        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            width: 100%;
            overflow-x: auto;
            word-wrap: break-word;
        }}

        .navbar div {{
            display: flex;
            flex-wrap: wrap;
            gap: 5px;
        }}

            table {{
            width: 100%;
            border-collapse: collapse;
            margin: 15px 0;
            display: block;
            overflow-x: auto;
            white-space: nowrap;
        }}
            th, td {{ padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }}
            th {{ background: #34495e; color: white; }}
            tr:hover {{ background: #f5f5f5; }}
            .btn {{ padding: 8px 15px; border: none; border-radius: 4px; cursor: pointer; text-decoration: none; display: inline-block; margin: 2px; font-size: 13px; }}
            .btn-primary {{ background: #3498db; color: white; }}
            .btn-success {{ background: #27ae60; color: white; }}
            .btn-warning {{ background: #f39c12; color: white; }}
            .btn-danger {{ background: #e74c3c; color: white; }}
            .btn-sm {{ padding: 4px 10px; font-size: 12px; }}
            .form-group {{ margin-bottom: 15px; }}
            .form-control {{ width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px; }}
            .badge {{ padding: 3px 8px; border-radius: 3px; font-size: 11px; }}
            .badge-admin {{ background: #e74c3c; color: white; }}
            .badge-empleado {{ background: #3498db; color: white; }}
            .badge-alta {{ background: #e74c3c; color: white; }}
            .badge-media {{ background: #f39c12; color: white; }}
            .badge-baja {{ background: #27ae60; color: white; }}
            .flash {{ padding: 12px; margin: 10px 0; border-radius: 4px; }}
            .flash-success {{ background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }}
            .stats {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 20px; }}
            .stat-card {{ background: #ecf0f1; padding: 15px; border-radius: 8px; text-align: center; }}
            .stat-number {{ font-size: 28px; font-weight: bold; color: #2c3e50; }}
            .stat-label {{ color: #7f8c8d; }}
            
            
            /* NUEVOS ESTILOS PARA MI ESPACIO */
            
                    /* CHAT */
            .chat-container, .chat-grupal-container {{
            display: flex;
            height: 500px;
            max-height: 70vh;
            border: 1px solid #ddd;
            border-radius: 8px;
            overflow: hidden;
        }}

         .chat-message-bubble {{
            max-width: 85%;
            padding: 10px 12px;
            border-radius: 18px;
            background: white;
            box-shadow: 0 1px 2px rgba(0,0,0,0.1);
            word-wrap: break-word;
        }}

        .chat-sidebar {{
            width: 250px;
            background: #f8f9fa;
            border-right: 1px solid #ddd;
            overflow-y: auto;
        }}
        .chat-sidebar-item {{
            padding: 15px;
            border-bottom: 1px solid #ddd;
            cursor: pointer;
            transition: background 0.2s;
        }}
        .chat-sidebar-item:hover {{ background: #e9ecef; }}
        .chat-sidebar-item.active {{ background: #1a73e8; color: white; }}
        .chat-sidebar-item .badge-new {{ background: #e74c3c; color: white; padding: 2px 8px; border-radius: 10px; font-size: 11px; margin-left: 5px; }}
        .chat-main {{
            flex: 1;
            display: flex;
            flex-direction: column;
        }}
        .chat-header {{
            padding: 15px;
            background: #f8f9fa;
            border-bottom: 1px solid #ddd;
            font-weight: bold;
        }}
        .chat-messages {{
            flex: 1;
            padding: 15px;
            overflow-y: auto;
            background: #fafafa;
        }}
        .chat-message {{
            margin-bottom: 15px;
            display: flex;
        }}
        .chat-message.sent {{
            justify-content: flex-end;
        }}
        .chat-message-bubble {{
            max-width: 70%;
            padding: 10px 15px;
            border-radius: 18px;
            background: white;
            box-shadow: 0 1px 2px rgba(0,0,0,0.1);
        }}
        .chat-message.sent .chat-message-bubble {{
            background: #1a73e8;
            color: white;
        }}
        .chat-message-time {{
            font-size: 10px;
            color: #999;
            margin-top: 5px;
        }}
        .chat-message.sent .chat-message-time {{
            color: #ddd;
        }}
        .chat-input {{
            padding: 15px;
            background: white;
            border-top: 1px solid #ddd;
            display: flex;
            gap: 10px;
        }}
        .chat-input input {{
            flex: 1;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 20px;
        }}
            .notas-grid {{ 
                display: grid; 
                grid-template-columns: repeat(auto-fill, minmax(250px, 1fr)); 
                gap: 15px; 
                margin: 20px 0; 
            }}
            .nota-card {{ 
                padding: 15px; 
                border-radius: 8px; 
                box-shadow: 0 2px 4px rgba(0,0,0,0.1); 
                min-height: 150px; 
                position: relative; 
            }}
            .nota-card h4 {{ 
                margin: 0 0 10px 0; 
            }}
            .nota-card p {{ 
                margin: 0; 
                color: #555; 
            }}
            .nota-actions {{ 
                position: absolute; 
                bottom: 10px; 
                right: 10px; 
                display: flex; 
                gap: 5px; 
            }}
            .color-picker {{ 
                display: flex; 
                gap: 5px; 
                margin: 10px 0; 
            }}
            .color-option {{ 
                width: 30px; 
                height: 30px; 
                border-radius: 50%; 
                cursor: pointer; 
                border: 2px solid transparent; 
            }}
            .color-option.selected {{ 
                border-color: #333; 
            }}
            .tabs {{ 
                display: flex; 
                gap: 10px; 
                margin-bottom: 20px; 
                border-bottom: 2px solid #ddd; 
            }}
            .tab {{ 
                padding: 10px 20px; 
                cursor: pointer; 
                border: none; 
                background: none; 
                font-size: 16px; 
            }}
            .tab.active {{ 
                border-bottom: 3px solid #1a73e8; 
                color: #1a73e8; 
                font-weight: bold; 
            }}
            .tab-content {{ 
                display: none; 
            }}
            .tab-content.active {{ 
                display: block; 
            }}
                    /* CLIENTES CRM */
        .clientes-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
            gap: 20px;
            margin: 20px 0;
        }}
        .cliente-card {{
            background: white;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border: 1px solid #e0e0e0;
        }}
        .cliente-header {{
            display: flex;
            align-items: center;
            gap: 10px;
            margin-bottom: 15px;
        }}
        .cliente-avatar {{
            width: 50px;
            height: 50px;
            background: #1a73e8;
            color: white;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 20px;
            font-weight: bold;
        }}
        .cliente-info h3 {{
            margin: 0;
            color: #333;
        }}
        .cliente-info p {{
            margin: 5px 0 0;
            color: #666;
            font-size: 14px;
        }}
        .cliente-body {{
            margin: 15px 0;
        }}
        .cliente-contacto {{
            display: flex;
            flex-direction: column;
            gap: 5px;
            margin-bottom: 10px;
        }}
        .cliente-contacto span {{
            display: flex;
            align-items: center;
            gap: 8px;
            color: #555;
            font-size: 14px;
        }}
        .cliente-footer {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-top: 15px;
            padding-top: 15px;
            border-top: 1px solid #eee;
        }}
        .interaccion-timeline {{
            margin: 20px 0;
        }}
        .interaccion-item {{
            display: flex;
            gap: 15px;
            padding: 15px;
            background: #f8f9fa;
            border-radius: 8px;
            margin-bottom: 10px;
        }}
        .interaccion-icon {{
            width: 40px;
            height: 40px;
            background: #1a73e8;
            color: white;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
        }}
        .interaccion-content {{
            flex: 1;
        }}
        .interaccion-header {{
            display: flex;
            justify-content: space-between;
            margin-bottom: 5px;
        }}
        .interaccion-tipo {{
            font-weight: bold;
            color: #1a73e8;
        }}
        .interaccion-fecha {{
            color: #999;
            font-size: 12px;
        }}

                /* CHAT GRUPAL */
        .chat-grupal-container {{
            display: flex;
            height: 500px;
            border: 1px solid #ddd;
            border-radius: 8px;
            overflow: hidden;
        }}
        .chat-grupal-sidebar {{
            width: 250px;
            background: #f8f9fa;
            border-right: 1px solid #ddd;
            overflow-y: auto;
        }}
        .chat-grupal-item {{
            padding: 15px;
            border-bottom: 1px solid #ddd;
            cursor: pointer;
        }}
        .chat-grupal-item:hover {{ background: #e9ecef; }}
        .chat-grupal-item.active {{ background: #1a73e8; color: white; }}
        .chat-grupal-main {{
            flex: 1;
            display: flex;
            flex-direction: column;
        }}
        .chat-grupal-header {{
            padding: 15px;
            background: #f8f9fa;
            border-bottom: 1px solid #ddd;
        }}
        .chat-grupal-messages {{
            flex: 1;
            padding: 15px;
            overflow-y: auto;
            background: #fafafa;
        }}
        .chat-grupal-input {{
            padding: 15px;
            background: white;
            border-top: 1px solid #ddd;
            display: flex;
            gap: 10px;
        }}
        
                /* RESPONSIVE */
        @media (max-width: 768px) {{
            .container {{
                padding: 15px;
            }}
            .navbar {{
                flex-direction: column;
                gap: 10px;
                text-align: center;
            }}
            .navbar div {{
                display: flex;
                flex-wrap: wrap;
                justify-content: center;
                gap: 5px;
            }}
            .stats {{
                grid-template-columns: repeat(2, 1fr);
            }}
            .chat-container, .chat-grupal-container {{
                flex-direction: column;
                height: auto;
            }}
            .chat-sidebar, .chat-grupal-sidebar {{
                width: 100%;
                max-height: 200px;
            }}
            table {{
                display: block;
                overflow-x: auto;
                white-space: nowrap;
            }}
            .clientes-grid {{
                grid-template-columns: 1fr;
            }}
            .form-row {{
                flex-direction: column;
            }}
            .btn {{
                padding: 10px 12px;
                font-size: 14px;
            }}
        }}
        
        @media (max-width: 480px) {{
            h1 {{
                font-size: 1.5rem;
            }}
            h2 {{
                font-size: 1.3rem;
            }}
            .stats {{
                grid-template-columns: 1fr;
            }}
            .stat-card {{
                padding: 10px;
            }}
            .stat-number {{
                font-size: 1.5rem;
            }}
            .navbar a {{
                padding: 5px 8px;
                font-size: 12px;
            }}
            .tabs {{
                flex-wrap: wrap;
            }}
            .tab {{
                padding: 8px 12px;
                font-size: 14px;
            }}
        }}    
        </style>
    </head>
    <body>
        {navbar_html() if session.get('user_id') else ''}
        <div class="container">
            {flash_html()}
            {content}
        </div>
    </body>
    </html>
    """
def navbar_html():
    user_id = session.get('user_id')
    no_leidos = 0
    if user_id:
        no_leidos = Mensaje.query.filter_by(receptor_id=user_id, leido=False).count()
    
    badge_chat = f' <span style="background:#e74c3c;color:white;padding:2px 8px;border-radius:10px;">{no_leidos}</span>' if no_leidos > 0 else ''
    
    return f"""
    <div class="navbar">
        <span><strong>🏢 OFICINA</strong> | 👤 {session.get('nombre')} ({session.get('rol')})</span>
        <div>
            <a href="/dashboard">📊 Dashboard</a>
            <a href="/mi-espacio">📌 Mi Espacio</a>
            <a href="/clientes">👥 Clientes</a>
            <a href="/chat">💬 Chat{badge_chat}</a>
            <a href="/chat-grupal">👥 Chat Grupal</a>   <!-- AQUÍ ESTÁ PARA TODOS -->
            {'''<a href="/admin/empleados">👥 Empleados</a>
            <a href="/admin/asignar">📋 Asignar Tarea</a>''' if session.get('rol') == 'admin' else ''}
            <a href="/logout" style="background: #e74c3c; color: white; padding: 6px 12px; border-radius: 20px; font-size: 12px; text-decoration: none;">🚪</a>
                    </div>
    </div>
    """


def flash_html():
    from flask import get_flashed_messages
    messages = get_flashed_messages()
    return ''.join([f'<div class="flash flash-success">{m}</div>' for m in messages]) if messages else ''

# ========== RUTAS ==========

# ========== ESPACIO PERSONAL (NOTAS Y TAREAS PROPIAS) ==========
@app.route('/mi-espacio')
def mi_espacio():
    if 'user_id' not in session:
        return redirect('/login')
    
    user = Usuario.query.get(session.get('user_id'))
    notas = NotaPersonal.query.filter_by(usuario_id=user.id).order_by(NotaPersonal.id.desc()).all()
    tareas_personales = TareaPersonal.query.filter_by(usuario_id=user.id).order_by(TareaPersonal.id.desc()).all()
    
    # Calendario personal
    hoy = datetime.now().date()
    lunes = hoy - timedelta(days=hoy.weekday())
    semana = [lunes + timedelta(days=i) for i in range(7)]
    
    tareas_semana = []
    for t in tareas_personales:
        if t.fecha_limite and not t.completada:
            try:
                fecha_t = datetime.strptime(t.fecha_limite, '%Y-%m-%d').date()
                if lunes <= fecha_t <= lunes + timedelta(days=6):
                    tareas_semana.append(t)
            except:
                pass
    
    # Construir calendario
    dias_html = ""
    for dia in semana:
        tareas_dia = [t for t in tareas_semana if datetime.strptime(t.fecha_limite, '%Y-%m-%d').date() == dia]
        es_hoy = dia == hoy
        
        tareas_dia_html = ""
        for t in tareas_dia:
            prioridad_color = {'alta': '#e74c3c', 'media': '#f39c12', 'baja': '#27ae60'}.get(t.prioridad, '#3498db')
            tareas_dia_html += f"""
            <div style="background: {prioridad_color}; color: white; padding: 3px 5px; margin: 2px 0; border-radius: 4px; font-size: 11px;">
                <a href="/tarea-personal/completar/{t.id}" style="color: white; text-decoration: none;" title="{t.titulo}">
                    {t.titulo[:15]}...
                </a>
            </div>
            """
        
        dias_html += f"""
        <td style="vertical-align: top; padding: 8px; border: 1px solid #ddd; {'background: #e8f4f8;' if es_hoy else ''}">
            <strong>{dia.strftime('%a')}<br>{dia.strftime('%d/%m')}</strong>
            {tareas_dia_html if tareas_dia else '<p style="color: #999; font-size: 10px;">-</p>'}
        </td>
        """
    
    # Notas HTML
    notas_html = ""
    colores = ['#f39c12', '#e74c3c', '#3498db', '#2ecc71', '#9b59b6', '#1abc9c']
    for nota in notas:
        notas_html += f"""
        <div class="nota-card" style="background: {nota.color}20; border-left: 5px solid {nota.color};">
            <h4>{nota.titulo}</h4>
            <p>{nota.contenido[:100]}{'...' if len(nota.contenido) > 100 else ''}</p>
            <small style="color: #999;">{nota.fecha_creacion}</small>
            <div class="nota-actions">
                <a href="/nota/editar/{nota.id}" class="btn btn-warning btn-sm">✏️</a>
                <a href="/nota/eliminar/{nota.id}" class="btn btn-danger btn-sm" onclick="return confirm('¿Eliminar nota?')">🗑️</a>
            </div>
        </div>
        """
    
    # Tareas personales HTML
    tareas_html = ""
    for t in tareas_personales:
        prioridad_badge = {'alta': '🔴', 'media': '🟡', 'baja': '🟢'}.get(t.prioridad, '')
        tareas_html += f"""
        <tr>
            <td>{prioridad_badge} {t.titulo[:40]}...</td>
            <td>{t.fecha_limite or '-'}</td>
            <td>{'✅' if t.completada else '⏳'}</td>
            <td>
                {f'<a href="/tarea-personal/completar/{t.id}" class="btn btn-success btn-sm">✅</a>' if not t.completada else ''}
                <a href="/tarea-personal/eliminar/{t.id}" class="btn btn-danger btn-sm" onclick="return confirm(\'¿Eliminar?\')">🗑️</a>
            </td>
        </tr>
        """
    
    content = f"""
        <h2>📌 Mi Espacio Personal - {user.nombre_completo}</h2>
        
        <div class="tabs">
            <button class="tab active" onclick="showTab('notas')">📓 Mis Notas</button>
            <button class="tab" onclick="showTab('tareas')">✅ Mis Tareas Personales</button>
            <button class="tab" onclick="showTab('calendario')">📅 Mi Calendario</button>
        </div>
        
        <!-- NOTAS -->
        <div id="tab-notas" class="tab-content active">
            <p><a href="/nota/nueva" class="btn btn-primary">➕ Nueva Nota</a></p>
            <div class="notas-grid">
                {notas_html if notas else '<p style="grid-column: 1/-1; text-align: center; color: #999;">No hay notas. ¡Crea una!</p>'}
            </div>
        </div>
        
        <!-- TAREAS PERSONALES -->
        <div id="tab-tareas" class="tab-content">
            <p><a href="/tarea-personal/nueva" class="btn btn-primary">➕ Nueva Tarea Personal</a></p>
            <table>
                <thead>
                    <tr>
                        <th>Título</th>
                        <th>Fecha Límite</th>
                        <th>Estado</th>
                        <th>Acciones</th>
                    </tr>
                </thead>
                <tbody>
                    {tareas_html if tareas_personales else '<tr><td colspan="4" style="text-align:center;">No hay tareas personales</td></tr>'}
                </tbody>
            </table>
        </div>
        
        <!-- CALENDARIO PERSONAL -->
        <div id="tab-calendario" class="tab-content">
            <p>Semana del {lunes.strftime('%d/%m/%Y')} al {(lunes + timedelta(days=6)).strftime('%d/%m/%Y')}</p>
            <table style="width: 100%; border-collapse: collapse;">
                <tr>{dias_html}</tr>
            </table>
            <p style="margin-top: 15px;"><a href="/tarea-personal/nueva" class="btn btn-primary">➕ Añadir tarea al calendario</a></p>
        </div>
        
        <script>
            function showTab(tabName) {{
                document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
                document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
                event.target.classList.add('active');
                document.getElementById('tab-' + tabName).classList.add('active');
            }}
        </script>
    """
    return base_html(content, "Mi Espacio Personal")

# ========== GESTIÓN DE NOTAS ==========
@app.route('/nota/nueva', methods=['GET', 'POST'])
def nueva_nota():
    if 'user_id' not in session:
        return redirect('/login')
    
    if request.method == 'POST':
        nota = NotaPersonal(
            titulo=request.form['titulo'],
            contenido=request.form['contenido'],
            color=request.form.get('color', '#f39c12'),
            usuario_id=session['user_id']
        )
        db.session.add(nota)
        db.session.commit()
        flash('✅ Nota creada')
        return redirect('/mi-espacio')
    
    colores = ['#f39c12', '#e74c3c', '#3498db', '#2ecc71', '#9b59b6', '#1abc9c']
    colores_html = "".join([f'<span class="color-option" style="background: {c};" onclick="selectColor(\'{c}\')"></span>' for c in colores])
    
    content = f"""
        <h2>📓 Nueva Nota</h2>
        <form method="POST" style="max-width: 500px;">
            <div class="form-group">
                <label>Título</label>
                <input type="text" name="titulo" class="form-control" required>
            </div>
            <div class="form-group">
                <label>Contenido</label>
                <textarea name="contenido" class="form-control" rows="6" required></textarea>
            </div>
            <div class="form-group">
                <label>Color</label>
                <div class="color-picker">
                    {colores_html}
                </div>
                <input type="hidden" name="color" id="colorSelected" value="#f39c12">
            </div>
            <button type="submit" class="btn btn-primary">Guardar Nota</button>
            <a href="/mi-espacio" class="btn" style="background: #95a5a6; color: white;">Cancelar</a>
        </form>
        <script>
            function selectColor(color) {{
                document.querySelectorAll('.color-option').forEach(c => c.classList.remove('selected'));
                event.target.classList.add('selected');
                document.getElementById('colorSelected').value = color;
            }}
        </script>
    """
    return base_html(content, "Nueva Nota")

@app.route('/nota/editar/<int:id>', methods=['GET', 'POST'])
def editar_nota(id):
    if 'user_id' not in session:
        return redirect('/login')
    
    nota = NotaPersonal.query.get_or_404(id)
    if nota.usuario_id != session['user_id']:
        flash('❌ No tienes permiso')
        return redirect('/mi-espacio')
    
    if request.method == 'POST':
        nota.titulo = request.form['titulo']
        nota.contenido = request.form['contenido']
        nota.color = request.form.get('color', nota.color)
        db.session.commit()
        flash('✅ Nota actualizada')
        return redirect('/mi-espacio')
    
    colores = ['#f39c12', '#e74c3c', '#3498db', '#2ecc71', '#9b59b6', '#1abc9c']
    colores_html = "".join([f'<span class="color-option {"selected" if c == nota.color else ""}" style="background: {c};" onclick="selectColor(\'{c}\')"></span>' for c in colores])
    
    content = f"""
        <h2>✏️ Editar Nota</h2>
        <form method="POST" style="max-width: 500px;">
            <div class="form-group">
                <label>Título</label>
                <input type="text" name="titulo" class="form-control" value="{nota.titulo}" required>
            </div>
            <div class="form-group">
                <label>Contenido</label>
                <textarea name="contenido" class="form-control" rows="6" required>{nota.contenido}</textarea>
            </div>
            <div class="form-group">
                <label>Color</label>
                <div class="color-picker">
                    {colores_html}
                </div>
                <input type="hidden" name="color" id="colorSelected" value="{nota.color}">
            </div>
            <button type="submit" class="btn btn-primary">Actualizar</button>
            <a href="/mi-espacio" class="btn" style="background: #95a5a6; color: white;">Cancelar</a>
        </form>
        <script>
            function selectColor(color) {{
                document.querySelectorAll('.color-option').forEach(c => c.classList.remove('selected'));
                event.target.classList.add('selected');
                document.getElementById('colorSelected').value = color;
            }}
        </script>
    """
    return base_html(content, "Editar Nota")

@app.route('/nota/eliminar/<int:id>')
def eliminar_nota(id):
    if 'user_id' not in session:
        return redirect('/login')
    
    nota = NotaPersonal.query.get_or_404(id)
    if nota.usuario_id == session['user_id']:
        db.session.delete(nota)
        db.session.commit()
        flash('✅ Nota eliminada')
    return redirect('/mi-espacio')

# ========== GESTIÓN DE TAREAS PERSONALES ==========
@app.route('/tarea-personal/nueva', methods=['GET', 'POST'])
def nueva_tarea_personal():
    if 'user_id' not in session:
        return redirect('/login')
    
    if request.method == 'POST':
        tarea = TareaPersonal(
            titulo=request.form['titulo'],
            descripcion=request.form.get('descripcion', ''),
            fecha_limite=request.form.get('fecha_limite', ''),
            prioridad=request.form.get('prioridad', 'media'),
            usuario_id=session['user_id']
        )
        db.session.add(tarea)
        db.session.commit()
        flash('✅ Tarea personal creada')
        return redirect('/mi-espacio')
    
    content = """
        <h2>✅ Nueva Tarea Personal</h2>
        <form method="POST" style="max-width: 500px;">
            <div class="form-group">
                <label>Título</label>
                <input type="text" name="titulo" class="form-control" required>
            </div>
            <div class="form-group">
                <label>Descripción</label>
                <textarea name="descripcion" class="form-control" rows="3"></textarea>
            </div>
            <div class="form-group">
                <label>Fecha límite</label>
                <input type="date" name="fecha_limite" class="form-control">
            </div>
            <div class="form-group">
                <label>Prioridad</label>
                <select name="prioridad" class="form-control">
                    <option value="alta">🔴 Alta</option>
                    <option value="media" selected>🟡 Media</option>
                    <option value="baja">🟢 Baja</option>
                </select>
            </div>
            <button type="submit" class="btn btn-primary">Crear Tarea</button>
            <a href="/mi-espacio" class="btn" style="background: #95a5a6; color: white;">Cancelar</a>
        </form>
    """
    return base_html(content, "Nueva Tarea Personal")

@app.route('/tarea-personal/completar/<int:id>')
def completar_tarea_personal(id):
    if 'user_id' not in session:
        return redirect('/login')
    
    tarea = TareaPersonal.query.get_or_404(id)
    if tarea.usuario_id == session['user_id']:
        tarea.completada = not tarea.completada
        db.session.commit()
    return redirect('/mi-espacio')

@app.route('/tarea-personal/eliminar/<int:id>')
def eliminar_tarea_personal(id):
    if 'user_id' not in session:
        return redirect('/login')
    
    tarea = TareaPersonal.query.get_or_404(id)
    if tarea.usuario_id == session['user_id']:
        db.session.delete(tarea)
        db.session.commit()
        flash('✅ Tarea eliminada')
    return redirect('/mi-espacio')

@app.route('/')
def index():
    if 'user_id' in session:
        return redirect('/dashboard')
    return redirect('/login')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        user = Usuario.query.filter_by(username=request.form['username']).first()
        if user and check_password_hash(user.password_hash, request.form['password']):
            session['user_id'] = user.id
            session['username'] = user.username
            session['nombre'] = user.nombre_completo
            session['rol'] = user.rol
            flash(f'✅ Bienvenido/a {user.nombre_completo}')
            return redirect('/dashboard')
        flash('❌ Usuario o contraseña incorrectos')
    
    content = """
        <h2>🔐 Acceso a la Oficina</h2>
        <form method="POST" style="max-width: 400px;">
            <div class="form-group">
                <label>Usuario</label>
                <input type="text" name="username" class="form-control" required>
            </div>
            <div class="form-group">
                <label>Contraseña</label>
                <input type="password" name="password" class="form-control" required>
            </div>
            <button type="submit" class="btn btn-primary">Entrar</button>
        </form>
        <p style="margin-top: 15px;"><a href="/registro">📝 Registrar nuevo empleado</a></p>
        <p></p>
        <p><a href="/olvide-password">¿Olvidaste tu contraseña?</a></p>
    """
    return base_html(content, "Login")

@app.route('/registro', methods=['GET', 'POST'])
def registro():
    if request.method == 'POST':
        if Usuario.query.filter_by(username=request.form['username']).first():
            flash('❌ El usuario ya existe')
        else:
            es_admin = Usuario.query.count() == 0
            user = Usuario(
                username=request.form['username'],
                password_hash=generate_password_hash(request.form['password'], method='pbkdf2:sha256'),
                nombre_completo=request.form['nombre_completo'],
                departamento=request.form.get('departamento', 'General'),
                cargo=request.form.get('cargo', 'Empleado'),
                rol='admin' if es_admin else 'empleado'
            )
            db.session.add(user)
            db.session.commit()
            
            if 'user_id' not in session:
                session['user_id'] = user.id
                session['username'] = user.username
                session['nombre'] = user.nombre_completo
                session['rol'] = user.rol
            
            flash('✅ Empleado registrado correctamente')
            return redirect('/dashboard')
    
    content = """
        <h2>📝 Registrar Nuevo Empleado</h2>
        <form method="POST" style="max-width: 500px;">
            <div class="form-group">
                <label>Usuario *</label>
                <input type="text" name="username" class="form-control" required>
            </div>
            <div class="form-group">
                <label>Contraseña *</label>
                <input type="password" name="password" class="form-control" required>
            </div>
            <div class="form-group">
                <label>Nombre Completo *</label>
                <input type="text" name="nombre_completo" class="form-control" required>
            </div>
            <div class="form-group">
                <label>Departamento</label>
                <select name="departamento" class="form-control">
                    <option value="Administración">Administración</option>
                    <option value="RRHH">Recursos Humanos</option>
                    <option value="Contabilidad">Contabilidad</option>
                    <option value="Ventas">Ventas</option>
                    <option value="IT">Tecnología</option>
                    <option value="General">General</option>
                </select>
            </div>
            <div class="form-group">
                <label>Cargo</label>
                <input type="text" name="cargo" class="form-control" value="Empleado">
            </div>
            <button type="submit" class="btn btn-success">Registrar Empleado</button>
            <a href="/dashboard" class="btn" style="background: #95a5a6; color: white;">Cancelar</a>
        </form>
    """
    return base_html(content, "Registro")

@app.route('/dashboard')
def dashboard():
    if 'user_id' not in session:
        return redirect('/login')
    
    user = Usuario.query.get(session.get('user_id'))
   
    if session.get('rol') == 'admin':
        # Dashboard de ADMIN - Versión PRO
        total_empleados = Usuario.query.count()
        total_tareas = Tarea.query.count()
        tareas_pendientes = Tarea.query.filter_by(completada=False).count()
        tareas_completadas = Tarea.query.filter_by(completada=True).count()
        
        # Clientes nuevos este mes
        hoy = datetime.now()
        inicio_mes = hoy.replace(day=1).strftime('%d/%m/%Y')
        clientes_mes = Cliente.query.filter(Cliente.fecha_creacion >= inicio_mes).count()
        
        # Tareas por departamento
        deptos = db.session.query(Usuario.departamento, db.func.count(Tarea.id)).join(Tarea).filter(Tarea.completada == False).group_by(Usuario.departamento).all()
        
        # Top 5 empleados
        empleados = Usuario.query.filter_by(rol='empleado').all()
        ranking = []
        for e in empleados:
            total = len(e.tareas)
            completadas = sum(1 for t in e.tareas if t.completada)
            prod = int((completadas / total * 100)) if total > 0 else 0
            ranking.append((e.nombre_completo, prod, completadas, total, e.departamento))
        ranking.sort(key=lambda x: x[1], reverse=True)
        top5 = ranking[:5]
        
        # Tareas urgentes
        hoy_str = datetime.now().strftime('%Y-%m-%d')
        manana = (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d')
        tareas_urgentes = Tarea.query.filter(
            Tarea.completada == False,
            Tarea.fecha_limite.in_([hoy_str, manana])
        ).limit(5).all()
        
        # Actividad reciente (últimos 5 clientes)
        clientes_recientes = Cliente.query.order_by(Cliente.id.desc()).limit(5).all()
        
        content = f"""
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 25px;">
                <h2 style="margin:0;">📊 Panel de Control</h2>
                <div>
                    <a href="/admin/exportar-todo" class="btn btn-success" style="margin-right:10px;">📊 Exportar Todo</a>
                    <span style="color:#666;">{hoy.strftime('%d de %B, %Y')}</span>
                </div>
            </div>
            
            <!-- KPIs -->
            <div style="display: grid; grid-template-columns: repeat(5, 1fr); gap: 15px; margin-bottom: 25px;">
                <div style="background: linear-gradient(135deg, #667eea, #764ba2); color: white; padding: 20px; border-radius: 12px;">
                    <div style="font-size: 14px; opacity: 0.9;">👥 Empleados</div>
                    <div style="font-size: 32px; font-weight: bold;">{total_empleados}</div>
                </div>
                <div style="background: linear-gradient(135deg, #f093fb, #f5576c); color: white; padding: 20px; border-radius: 12px;">
                    <div style="font-size: 14px; opacity: 0.9;">⏳ Pendientes</div>
                    <div style="font-size: 32px; font-weight: bold;">{tareas_pendientes}</div>
                </div>
                <div style="background: linear-gradient(135deg, #4facfe, #00f2fe); color: white; padding: 20px; border-radius: 12px;">
                    <div style="font-size: 14px; opacity: 0.9;">✅ Completadas</div>
                    <div style="font-size: 32px; font-weight: bold;">{tareas_completadas}</div>
                </div>
                <div style="background: linear-gradient(135deg, #43e97b, #38f9d7); color: white; padding: 20px; border-radius: 12px;">
                    <div style="font-size: 14px; opacity: 0.9;">🆕 Clientes (mes)</div>
                    <div style="font-size: 32px; font-weight: bold;">{clientes_mes}</div>
                </div>
                <div style="background: linear-gradient(135deg, #fa709a, #fee140); color: white; padding: 20px; border-radius: 12px;">
                    <div style="font-size: 14px; opacity: 0.9;">📈 Productividad</div>
                    <div style="font-size: 32px; font-weight: bold;">{int((tareas_completadas/total_tareas*100)) if total_tareas else 0}%</div>
                </div>
            </div>
            
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px;">
                <!-- Tareas por Departamento -->
                <div style="background: white; border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                    <h3 style="margin-bottom: 15px;">📋 Tareas Pendientes por Departamento</h3>
                    <div style="max-height: 250px; overflow-y: auto;">
                        <table style="width:100%; min-width:0;">
                            <thead><tr><th>Departamento</th><th>Pendientes</th><th>Barra</th></tr></thead>
                            <tbody>
                                {''.join([f'''
                                <tr>
                                    <td>{d[0]}</td>
                                    <td><strong>{d[1]}</strong></td>
                                    <td>
                                        <div style="background:#e0e0e0; height:8px; border-radius:4px; width:100px;">
                                            <div style="background:#1a73e8; height:8px; border-radius:4px; width:{min(d[1]*10, 100)}px;"></div>
                                        </div>
                                    </td>
                                </tr>
                                ''' for d in deptos]) if deptos else '<tr><td colspan="3">No hay datos</td></tr>'}
                            </tbody>
                        </table>
                    </div>
                </div>
                
                <!-- Top Empleados -->
                <div style="background: white; border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                    <h3 style="margin-bottom: 15px;">🏆 Top Productividad</h3>
                    <div style="max-height: 250px; overflow-y: auto;">
                        {''.join([f'''
                        <div style="display: flex; align-items: center; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid #eee;">
                            <div style="display: flex; align-items: center; gap: 10px;">
                                <div style="width: 40px; height: 40px; background: {"#ffd700" if i==0 else "#c0c0c0" if i==1 else "#cd7f32" if i==2 else "#1a73e8"}; color: white; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: bold;">
                                    {nombre[0]}
                                </div>
                                <div>
                                    <strong>{nombre}</strong>
                                    <div style="font-size: 12px; color: #666;">{depto}</div>
                                </div>
                            </div>
                            <div style="text-align: right;">
                                <div style="font-weight: bold; color: #1a73e8;">{prod}%</div>
                                <div style="font-size: 12px; color: #666;">{comp}/{total} tareas</div>
                            </div>
                        </div>
                        ''' for i, (nombre, prod, comp, total, depto) in enumerate(top5)]) if top5 else '<p>No hay empleados</p>'}
                    </div>
                </div>
            </div>
            
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
                <!-- Tareas Urgentes -->
                <div style="background: white; border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                    <h3 style="margin-bottom: 15px; color: #e74c3c;">⚡ Tareas Urgentes (hoy/mañana)</h3>
                    {''.join([f'''
                    <div style="padding: 12px; background: #fff5f5; border-radius: 8px; margin-bottom: 8px; border-left: 4px solid #e74c3c;">
                        <div style="display: flex; justify-content: space-between;">
                            <strong>{t.titulo[:40]}...</strong>
                            <span style="color:#e74c3c; font-size:12px;">{t.fecha_limite}</span>
                        </div>
                        <div style="font-size: 13px; color: #666;">👤 {t.asignado.nombre_completo}</div>
                    </div>
                    ''' for t in tareas_urgentes]) if tareas_urgentes else '<p style="color:#27ae60;">✅ No hay tareas urgentes</p>'}
                </div>
                
                <!-- Actividad Reciente -->
                <div style="background: white; border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                    <h3 style="margin-bottom: 15px;">🆕 Clientes Recientes</h3>
                    {''.join([f'''
                    <div style="padding: 12px; border-bottom: 1px solid #eee;">
                        <div style="display: flex; align-items: center; gap: 10px;">
                            <div style="width: 35px; height: 35px; background: #1a73e8; color: white; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 14px;">
                                {c.nombre[0].upper()}
                            </div>
                            <div style="flex:1;">
                                <strong>{c.nombre}</strong>
                                <div style="font-size: 12px; color: #666;">{c.empresa or 'Sin empresa'} • {c.fecha_creacion}</div>
                            </div>
                            <span style="font-size: 12px; color: #999;">👤 {c.usuario.nombre_completo.split()[0]}</span>
                        </div>
                    </div>
                    ''' for c in clientes_recientes]) if clientes_recientes else '<p>No hay clientes aún</p>'}
                    <a href="/clientes" style="display: block; text-align: center; margin-top: 15px; color: #1a73e8;">Ver todos los clientes →</a>
                </div>
            </div>
        """
    else:
        # Dashboard de EMPLEADO - Versión PRO
        tareas = Tarea.query.filter_by(usuario_id=user.id).all()
        pendientes = sum(1 for t in tareas if not t.completada)
        completadas = sum(1 for t in tareas if t.completada)
        productividad = int((completadas / len(tareas) * 100)) if tareas else 0
        
        # Tareas urgentes (hoy/mañana)
        hoy_str = datetime.now().strftime('%Y-%m-%d')
        manana = (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d')
        tareas_urgentes = [t for t in tareas if not t.completada and t.fecha_limite in [hoy_str, manana]]
        
        # Tareas recientes (últimas 5)
        tareas_recientes = sorted([t for t in tareas if not t.completada], key=lambda x: x.fecha_creacion, reverse=True)[:5]
        
        # Mis clientes
        mis_clientes = Cliente.query.filter_by(usuario_id=user.id).count()
        clientes_recientes = Cliente.query.filter_by(usuario_id=user.id).order_by(Cliente.id.desc()).limit(3).all()
        
        # Notas personales recientes
        mis_notas = NotaPersonal.query.filter_by(usuario_id=user.id).order_by(NotaPersonal.id.desc()).limit(3).all()
        
        # Mensajes no leídos
        mensajes_no_leidos = Mensaje.query.filter_by(receptor_id=user.id, leido=False).count()
        
        # Tareas por prioridad
        tareas_alta = sum(1 for t in tareas if not t.completada and t.prioridad == 'alta')
        tareas_media = sum(1 for t in tareas if not t.completada and t.prioridad == 'media')
        tareas_baja = sum(1 for t in tareas if not t.completada and t.prioridad == 'baja')
        
        # Frase motivacional según productividad
        if productividad >= 80:
            frase = "🌟 ¡Excelente trabajo! Eres muy productivo."
            emoji = "🏆"
        elif productividad >= 50:
            frase = "💪 ¡Vas por buen camino! Sigue así."
            emoji = "🚀"
        else:
            frase = "📋 ¡Ánimo! Cada tarea completada es un paso adelante."
            emoji = "🎯"
        
        content = f"""
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 25px;">
                <div>
                    <h2 style="margin:0;">👋 ¡Hola, {user.nombre_completo}!</h2>
                    <p style="color:#666; margin:5px 0 0 0;">{user.departamento} • {user.cargo}</p>
                </div>
                <div style="text-align: right;">
                    <div style="font-size: 32px;">{emoji}</div>
                    <p style="color:#666; font-style:italic; max-width:300px;">{frase}</p>
                </div>
            </div>
            
            <!-- KPIs Personales -->
            <div style="display: grid; grid-template-columns: repeat(5, 1fr); gap: 15px; margin-bottom: 25px;">
                <div style="background: linear-gradient(135deg, #667eea, #764ba2); color: white; padding: 20px; border-radius: 12px;">
                    <div style="font-size: 14px; opacity: 0.9;">📋 Mis Tareas</div>
                    <div style="font-size: 32px; font-weight: bold;">{len(tareas)}</div>
                    <div style="font-size: 12px; opacity:0.8;">{pendientes} pendientes</div>
                </div>
                <div style="background: linear-gradient(135deg, #f093fb, #f5576c); color: white; padding: 20px; border-radius: 12px;">
                    <div style="font-size: 14px; opacity: 0.9;">✅ Completadas</div>
                    <div style="font-size: 32px; font-weight: bold;">{completadas}</div>
                </div>
                <div style="background: linear-gradient(135deg, #4facfe, #00f2fe); color: white; padding: 20px; border-radius: 12px;">
                    <div style="font-size: 14px; opacity: 0.9;">📈 Productividad</div>
                    <div style="font-size: 32px; font-weight: bold;">{productividad}%</div>
                    <div style="margin-top:5px; background:rgba(255,255,255,0.3); height:6px; border-radius:3px;">
                        <div style="background:white; width:{productividad}%; height:6px; border-radius:3px;"></div>
                    </div>
                </div>
                <div style="background: linear-gradient(135deg, #43e97b, #38f9d7); color: white; padding: 20px; border-radius: 12px;">
                    <div style="font-size: 14px; opacity: 0.9;">👥 Mis Clientes</div>
                    <div style="font-size: 32px; font-weight: bold;">{mis_clientes}</div>
                </div>
                <div style="background: linear-gradient(135deg, #fa709a, #fee140); color: white; padding: 20px; border-radius: 12px;">
                    <div style="font-size: 14px; opacity: 0.9;">💬 Mensajes</div>
                    <div style="font-size: 32px; font-weight: bold;">{mensajes_no_leidos}</div>
                    <div style="font-size: 12px; opacity:0.8;">sin leer</div>
                </div>
            </div>
            
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px;">
                <!-- Tareas Urgentes -->
                <div style="background: white; border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                    <h3 style="margin-bottom: 15px; color: #e74c3c;">⚡ Tareas Urgentes</h3>
                    {''.join([f'''
                    <div style="padding: 12px; background: #fff5f5; border-radius: 8px; margin-bottom: 8px; border-left: 4px solid #e74c3c;">
                        <div style="display: flex; justify-content: space-between;">
                            <strong>{t.titulo[:35]}...</strong>
                            <span style="color:#e74c3c; font-size:12px;">{t.fecha_limite}</span>
                        </div>
                        <div style="margin-top:8px;">
                            <a href="/tarea/completar/{t.id}" class="btn btn-success btn-sm">✅ Completar</a>
                            <button class="btn btn-primary btn-sm" onclick="verTarea({t.id})">👁️ Ver</button>
                        </div>
                    </div>
                    ''' for t in tareas_urgentes]) if tareas_urgentes else '<p style="color:#27ae60; padding:12px; background:#f0fff0; border-radius:8px;">✅ No tienes tareas urgentes</p>'}
                </div>
                
                <!-- Resumen por Prioridad -->
                <div style="background: white; border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                    <h3 style="margin-bottom: 15px;">📊 Tareas Pendientes por Prioridad</h3>
                    <div style="margin-bottom: 20px;">
                        <div style="display: flex; justify-content: space-between; margin-bottom: 8px;">
                            <span>🔴 Alta</span>
                            <span><strong>{tareas_alta}</strong> tareas</span>
                        </div>
                        <div style="background:#e0e0e0; height:8px; border-radius:4px;">
                            <div style="background:#e74c3c; width:{min(tareas_alta*20, 100)}%; height:8px; border-radius:4px;"></div>
                        </div>
                    </div>
                    <div style="margin-bottom: 20px;">
                        <div style="display: flex; justify-content: space-between; margin-bottom: 8px;">
                            <span>🟡 Media</span>
                            <span><strong>{tareas_media}</strong> tareas</span>
                        </div>
                        <div style="background:#e0e0e0; height:8px; border-radius:4px;">
                            <div style="background:#f39c12; width:{min(tareas_media*20, 100)}%; height:8px; border-radius:4px;"></div>
                        </div>
                    </div>
                    <div>
                        <div style="display: flex; justify-content: space-between; margin-bottom: 8px;">
                            <span>🟢 Baja</span>
                            <span><strong>{tareas_baja}</strong> tareas</span>
                        </div>
                        <div style="background:#e0e0e0; height:8px; border-radius:4px;">
                            <div style="background:#27ae60; width:{min(tareas_baja*20, 100)}%; height:8px; border-radius:4px;"></div>
                        </div>
                    </div>
                </div>
            </div>
            
            <div style="display: grid; grid-template-columns: 2fr 1fr; gap: 20px; margin-bottom: 20px;">
                <!-- Próximas Tareas -->
                <div style="background: white; border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                    <h3 style="margin-bottom: 15px;">📌 Próximas Tareas</h3>
                    <div style="max-height: 250px; overflow-y: auto;">
                        {''.join([f'''
                        <div style="padding: 12px; border-bottom: 1px solid #eee; display: flex; align-items: center; gap: 10px;">
                            <div style="width: 8px; height: 8px; background: {"#e74c3c" if t.prioridad=="alta" else "#f39c12" if t.prioridad=="media" else "#27ae60"}; border-radius: 50%;"></div>
                            <div style="flex:1;">
                                <strong>{t.titulo[:40]}...</strong>
                                <div style="font-size: 12px; color: #666;">📅 {t.fecha_limite or 'Sin fecha'}</div>
                            </div>
                            <button class="btn btn-primary btn-sm" onclick="verTarea({t.id})">👁️</button>
                        </div>
                        ''' for t in tareas_recientes]) if tareas_recientes else '<p style="color:#999; text-align:center; padding:20px;">No hay tareas pendientes</p>'}
                    </div>
                    <a href="/mi-espacio" style="display: block; text-align: center; margin-top: 15px; color: #1a73e8;">Ver todas mis tareas →</a>
                </div>
                
                <!-- Accesos Rápidos -->
                <div style="background: white; border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                    <h3 style="margin-bottom: 15px;">⚡ Accesos Rápidos</h3>
                    <div style="display: flex; flex-direction: column; gap: 10px;">
                        <a href="/mi-espacio" style="padding: 15px; background: #f8f9fa; border-radius: 8px; text-decoration: none; color: #333; display: flex; align-items: center; gap: 10px;">
                            <span style="font-size: 24px;">📓</span>
                            <div>
                                <strong>Mi Espacio</strong>
                                <div style="font-size: 12px; color: #666;">Notas y tareas personales</div>
                            </div>
                        </a>
                        <a href="/clientes" style="padding: 15px; background: #f8f9fa; border-radius: 8px; text-decoration: none; color: #333; display: flex; align-items: center; gap: 10px;">
                            <span style="font-size: 24px;">👥</span>
                            <div>
                                <strong>Mis Clientes</strong>
                                <div style="font-size: 12px; color: #666;">{mis_clientes} clientes asignados</div>
                            </div>
                        </a>
                        <a href="/chat" style="padding: 15px; background: #f8f9fa; border-radius: 8px; text-decoration: none; color: #333; display: flex; align-items: center; gap: 10px;">
                            <span style="font-size: 24px;">💬</span>
                            <div>
                                <strong>Chat con Admin</strong>
                                {f'<span style="background:#e74c3c; color:white; padding:2px 8px; border-radius:10px; font-size:11px; margin-left:10px;">{mensajes_no_leidos} nuevo</span>' if mensajes_no_leidos > 0 else ''}
                            </div>
                        </a>
                        <a href="/chat-grupal" style="padding: 15px; background: #f8f9fa; border-radius: 8px; text-decoration: none; color: #333; display: flex; align-items: center; gap: 10px;">
                            <span style="font-size: 24px;">👥</span>
                            <div>
                                <strong>Chat {user.departamento}</strong>
                                <div style="font-size: 12px; color: #666;">Habla con tu equipo</div>
                            </div>
                        </a>
                    </div>
                </div>
            </div>
            
            <!-- Modal para ver tarea -->
            <div id="modalTarea" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:1000; justify-content:center; align-items:center;">
                <div style="background:white; padding:30px; border-radius:12px; max-width:500px; width:90%; max-height:80vh; overflow-y:auto;">
                    <h3 id="modalTitulo"></h3>
                    <p><strong>Fecha límite:</strong> <span id="modalFecha"></span></p>
                    <p><strong>Prioridad:</strong> <span id="modalPrioridad"></span></p>
                    <p><strong>Descripción completa:</strong></p>
                    <div id="modalDescripcion" style="background:#f5f5f5; padding:15px; border-radius:8px; margin:10px 0;"></div>
                    <div style="display: flex; gap: 10px;">
                        <button class="btn btn-primary" onclick="document.getElementById('modalTarea').style.display='none'">Cerrar</button>
                        <a href="#" id="modalCompletar" class="btn btn-success">✅ Completar</a>
                    </div>
                </div>
            </div>
            
            <script>
            const tareasData = """ + json.dumps([{
                'id': t.id,
                'titulo': t.titulo,
                'descripcion': t.descripcion or 'Sin descripción',
                'fecha_limite': t.fecha_limite or 'Sin fecha',
                'prioridad': t.prioridad,
                'completada': t.completada
            } for t in tareas]) + """;
            
            function verTarea(id) {{
                const tarea = tareasData.find(t => t.id === id);
                if (tarea) {{
                    document.getElementById('modalTitulo').textContent = tarea.titulo;
                    document.getElementById('modalFecha').textContent = tarea.fecha_limite;
                    document.getElementById('modalPrioridad').textContent = tarea.prioridad.toUpperCase();
                    document.getElementById('modalDescripcion').textContent = tarea.descripcion;
                    document.getElementById('modalCompletar').href = '/tarea/completar/' + tarea.id;
                    if (tarea.completada) {{
                        document.getElementById('modalCompletar').style.display = 'none';
                    }} else {{
                        document.getElementById('modalCompletar').style.display = 'inline-block';
                    }}
                    document.getElementById('modalTarea').style.display = 'flex';
                }}
            }}
            
            document.getElementById('modalTarea').addEventListener('click', function(e) {{
                if (e.target === this) this.style.display = 'none';
            }});
            </script>
        """
    
    return base_html(content, "Dashboard")


    return base_html(content, "Dashboard")


@app.route('/admin/empleados')
def admin_empleados():
    if 'user_id' not in session or session.get('rol') != 'admin':
        return redirect('/login')
    
    empleados = Usuario.query.all()
    
    filas = ""
    for e in empleados:
        tareas_totales = len(e.tareas)
        tareas_completadas = sum(1 for t in e.tareas if t.completada)
        productividad = int((tareas_completadas / tareas_totales * 100)) if tareas_totales > 0 else 0
        
        filas += f"""
        <tr>
            <td>{e.nombre_completo}</td>
            <td>{e.username}</td>
            <td>{e.departamento}</td>
            <td>{e.cargo}</td>
            <td><span class="badge {'badge-admin' if e.rol == 'admin' else 'badge-empleado'}">{e.rol.upper()}</span></td>
            <td>{tareas_totales} ({tareas_completadas} ✅)</td>
            <td>
                <div style="background: #ddd; height: 6px; width: 100px; border-radius: 3px;">
                    <div style="background: #27ae60; height: 6px; width: {productividad}px; border-radius: 3px;"></div>
                </div>
                {productividad}%
            </td>
            <td>
                <a href="/admin/empleado/editar/{e.id}" class="btn btn-warning btn-sm">✏️</a>
                {f'<a href="/admin/empleado/eliminar/{e.id}" class="btn btn-danger btn-sm" onclick="return confirm(&quot;¿Eliminar a {e.nombre_completo}?&quot;)">🗑️</a>' if e.rol != 'admin' else ''}
            </td>
        </tr>
        """
    
    content = f"""
        <h2>👥 Gestión de Empleados</h2>
        <p><a href="/registro" class="btn btn-success">➕ Nuevo Empleado</a></p>
        
        <table>
            <thead>
                <tr>
                    <th>Nombre</th>
                    <th>Usuario</th>
                    <th>Departamento</th>
                    <th>Cargo</th>
                    <th>Rol</th>
                    <th>Tareas</th>
                    <th>Productividad</th>
                    <th>Acciones</th>
                </tr>
            </thead>
            <tbody>
                {filas}
            </tbody>
        </table>
    """
    return base_html(content, "Empleados")

@app.route('/admin/empleado/editar/<int:id>', methods=['GET', 'POST'])
def editar_empleado(id):
    if 'user_id' not in session or session.get('rol') != 'admin':
        return redirect('/dashboard')
    
    empleado = Usuario.query.get_or_404(id)
    
    if request.method == 'POST':
        empleado.nombre_completo = request.form['nombre_completo']
        empleado.departamento = request.form['departamento']
        empleado.cargo = request.form['cargo']
        if request.form.get('password'):
            empleado.password_hash = generate_password_hash(request.form['password'], method='pbkdf2:sha256')
        db.session.commit()
        flash('✅ Empleado actualizado')
        return redirect('/admin/empleados')
    
    content = f"""
        <h2>✏️ Editar Empleado: {empleado.nombre_completo}</h2>
        <form method="POST" style="max-width: 500px;">
            <div class="form-group">
                <label>Nombre Completo</label>
                <input type="text" name="nombre_completo" class="form-control" value="{empleado.nombre_completo}" required>
            </div>
            <div class="form-group">
                <label>Departamento</label>
                <select name="departamento" class="form-control">
                    {''.join([f'<option value="{d}" {"selected" if d == empleado.departamento else ""}>{d}</option>' 
                              for d in ['Administración', 'RRHH', 'Contabilidad', 'Ventas', 'IT', 'General']])}
                </select>
            </div>
            <div class="form-group">
                <label>Cargo</label>
                <input type="text" name="cargo" class="form-control" value="{empleado.cargo}">
            </div>
            <div class="form-group">
                <label>Nueva Contraseña (dejar en blanco para no cambiar)</label>
                <input type="password" name="password" class="form-control">
            </div>
            <button type="submit" class="btn btn-primary">Guardar Cambios</button>
            <a href="/admin/empleados" class="btn" style="background: #95a5a6; color: white;">Cancelar</a>
        </form>
    """
    return base_html(content, "Editar Empleado")

@app.route('/admin/empleado/eliminar/<int:id>')
def eliminar_empleado(id):
    if 'user_id' not in session or session.get('rol') != 'admin':
        return redirect('/dashboard')
    
    empleado = Usuario.query.get_or_404(id)
    if empleado.rol != 'admin':
        Tarea.query.filter_by(usuario_id=id).delete()
        db.session.delete(empleado)
        db.session.commit()
        flash('✅ Empleado eliminado')
    else:
        flash('❌ No se puede eliminar al administrador')
    
    return redirect('/admin/empleados')

@app.route('/admin/asignar', methods=['GET', 'POST'])
def admin_asignar():
    if 'user_id' not in session or session.get('rol') != 'admin':
        return redirect('/dashboard')
    
    if request.method == 'POST':
        tarea = Tarea(
            titulo=request.form['titulo'],
            descripcion=request.form.get('descripcion', ''),
            fecha_limite=request.form.get('fecha_limite', ''),
            prioridad=request.form.get('prioridad', 'media'),
            usuario_id=request.form['usuario_id']
        )
        db.session.add(tarea)
        db.session.commit()
        flash('✅ Tarea asignada correctamente')
        return redirect('/admin/asignar')
    
    empleados = Usuario.query.all()
    opciones = "".join([f'<option value="{e.id}">{e.nombre_completo} ({e.departamento})</option>' for e in empleados])
    
    content = f"""
        <h2>📋 Asignar Nueva Tarea</h2>
        <form method="POST" style="max-width: 600px;">
            <div class="form-group">
                <label>Empleado *</label>
                <select name="usuario_id" class="form-control" required>
                    <option value="">Seleccionar empleado...</option>
                    {opciones}
                </select>
            </div>
            <div class="form-group">
                <label>Título de la tarea *</label>
                <input type="text" name="titulo" class="form-control" required>
            </div>
            <div class="form-group">
                <label>Descripción</label>
                <textarea name="descripcion" class="form-control" rows="3"></textarea>
            </div>
            <div class="form-group">
                <label>Fecha límite</label>
                <input type="date" name="fecha_limite" class="form-control">
            </div>
            <div class="form-group">
                <label>Prioridad</label>
                <select name="prioridad" class="form-control">
                    <option value="alta">🔴 Alta</option>
                    <option value="media" selected>🟡 Media</option>
                    <option value="baja">🟢 Baja</option>
                </select>
            </div>
            <button type="submit" class="btn btn-primary">Asignar Tarea</button>
        </form>
    """
    return base_html(content, "Asignar Tarea")

@app.route('/tarea/completar/<int:id>')
def completar_tarea(id):
    if 'user_id' not in session:
        return redirect('/login')
    tarea = Tarea.query.get_or_404(id)
    if tarea.usuario_id == session['user_id'] or session.get('rol') == 'admin':
        tarea.completada = True
        db.session.commit()
        flash('✅ Tarea completada')
    return redirect('/dashboard')

@app.route('/logout')
def logout():
    session.clear()
    flash('Has cerrado sesión')
    return redirect('/login')

# ========== CHAT INTERNO ==========
@app.route('/chat')
def chat():
    if 'user_id' not in session:
        return redirect('/login')
    
    user = Usuario.query.get(session.get('user_id'))
    if not user:
        session.clear()
        return redirect('/login')
    
    if session.get('rol') == 'admin':
        # Admin ve lista de empleados
        empleados = Usuario.query.filter(Usuario.id != user.id).all()
        
        empleados_html = ""
        for emp in empleados:
            # Contar mensajes no leídos de este empleado
            no_leidos = Mensaje.query.filter_by(
                emisor_id=emp.id, 
                receptor_id=user.id, 
                leido=False
            ).count()
            
            badge = f'<span class="badge-new">{no_leidos}</span>' if no_leidos > 0 else ''
            
            empleados_html += f"""
            <div class="chat-sidebar-item" onclick="window.location.href='/chat/{emp.id}'">
                👤 {emp.nombre_completo} {badge}
                <small style="display:block;color:#666;">{emp.departamento}</small>
            </div>
            """
        
        content = f"""
            <h2>💬 Chat con Empleados</h2>
            <div class="chat-container">
                <div class="chat-sidebar">
                    {empleados_html if empleados else '<p style="padding:15px;">No hay empleados</p>'}
                </div>
                <div class="chat-main">
                    <div class="chat-header">Selecciona un empleado para chatear</div>
                    <div class="chat-messages" style="display:flex;align-items:center;justify-content:center;color:#999;">
                        👈 Elige un contacto
                    </div>
                </div>
            </div>
        """
    else:
        # Empleado: chat directo con admin
        admin = Usuario.query.filter_by(rol='admin').first()
        return redirect(f'/chat/{admin.id}')
    
    return base_html(content, "Chat")

@app.route('/chat/<int:otro_id>', methods=['GET', 'POST'])
def chat_con(otro_id):
    if 'user_id' not in session:
        return redirect('/login')
    
    user = Usuario.query.get(session.get('user_id'))
    if not user:
        session.clear()
        return redirect('/login')
    
    otro = Usuario.query.get_or_404(otro_id)
    
    # Verificar permisos
    if session.get('rol') != 'admin' and otro.rol != 'admin':
        flash('❌ No tienes permiso')
        return redirect('/chat')
    
    if request.method == 'POST':
        contenido = request.form.get('mensaje')
        if contenido:
            msg = Mensaje(
                contenido=contenido,
                emisor_id=user.id,
                receptor_id=otro.id
            )
            db.session.add(msg)
            db.session.commit()
        return redirect(f'/chat/{otro_id}')
    
    # Marcar mensajes como leídos
    Mensaje.query.filter_by(
        emisor_id=otro.id,
        receptor_id=user.id,
        leido=False
    ).update({'leido': True})
    db.session.commit()
    
    # Obtener conversación
    mensajes = Mensaje.query.filter(
        ((Mensaje.emisor_id == user.id) & (Mensaje.receptor_id == otro.id)) |
        ((Mensaje.emisor_id == otro.id) & (Mensaje.receptor_id == user.id))
    ).order_by(Mensaje.id).all()
    
    mensajes_html = ""
    for m in mensajes:
        sent_class = 'sent' if m.emisor_id == user.id else ''
        mensajes_html += f"""
        <div class="chat-message {sent_class}">
            <div class="chat-message-bubble">
                {m.contenido}
                <div class="chat-message-time">{m.fecha}</div>
            </div>
        </div>
        """
    
    # Sidebar para admin
    sidebar_html = ""
    if session.get('rol') == 'admin':
        empleados = Usuario.query.filter(Usuario.id != user.id).all()
        for emp in empleados:
            no_leidos = Mensaje.query.filter_by(
                emisor_id=emp.id,
                receptor_id=user.id,
                leido=False
            ).count()
            badge = f'<span class="badge-new">{no_leidos}</span>' if no_leidos > 0 else ''
            active_class = 'active' if emp.id == otro_id else ''
            sidebar_html += f"""
            <div class="chat-sidebar-item {active_class}" onclick="window.location.href='/chat/{emp.id}'">
                👤 {emp.nombre_completo} {badge}
            </div>
            """
    
    content = f"""
        <h2>💬 Chat con {otro.nombre_completo}</h2>
        <div class="chat-container">
            {f'<div class="chat-sidebar">{sidebar_html}</div>' if session.get("rol") == "admin" else ''}
            <div class="chat-main">
                <div class="chat-header">👤 {otro.nombre_completo} ({otro.departamento})</div>
                <div class="chat-messages" id="chat-messages">
                    {mensajes_html if mensajes else '<p style="text-align:center;color:#999;padding:20px;">No hay mensajes. ¡Empieza la conversación!</p>'}
                </div>
                <form method="POST" class="chat-input">
                    <input type="text" name="mensaje" placeholder="Escribe un mensaje..." required autocomplete="off">
                    <button type="submit" class="btn btn-primary">📤 Enviar</button>
                </form>
            </div>
        </div>
        
        <script>
            // Scroll automático al final
            const messagesDiv = document.getElementById('chat-messages');
            if (messagesDiv) messagesDiv.scrollTop = messagesDiv.scrollHeight;
        </script>
    """
    return base_html(content, f"Chat - {otro.nombre_completo}")

# ========== GESTIÓN DE CLIENTES (CRM) ==========

@app.route('/clientes')
def clientes():
    if 'user_id' not in session:
        return redirect('/login')
    
    user = Usuario.query.get(session.get('user_id'))
    if not user:
        session.clear()
        return redirect('/login')
    
    # Admin ve todos los clientes, empleados solo los suyos
    if session.get('rol') == 'admin':
        clientes_list = Cliente.query.order_by(Cliente.nombre).all()
    else:
        clientes_list = Cliente.query.filter_by(usuario_id=user.id).order_by(Cliente.nombre).all()
    
    busqueda = request.args.get('busqueda', '')
    if busqueda:
        clientes_list = [c for c in clientes_list if 
                        busqueda.lower() in c.nombre.lower() or 
                        (c.empresa and busqueda.lower() in c.empresa.lower()) or
                        (c.email and busqueda.lower() in c.email.lower())]
    
    # Estadísticas rápidas
    total_clientes = len(clientes_list)
    clientes_con_email = sum(1 for c in clientes_list if c.email)
    clientes_con_telefono = sum(1 for c in clientes_list if c.telefono)
    
    # Agrupar por empresa
    empresas = {}
    for c in clientes_list:
        emp = c.empresa or 'Sin empresa'
        if emp not in empresas:
            empresas[emp] = []
        empresas[emp].append(c)
    
    clientes_html = ""
    for c in clientes_list[:20]:  # Mostrar primeros 20 para no sobrecargar
        inicial = c.nombre[0].upper() if c.nombre else '?'
        interacciones_count = len(c.interacciones)
        ultima_interaccion = c.interacciones[-1].fecha if c.interacciones else 'Sin actividad'
        
        # Color según actividad
        if interacciones_count > 5:
            border_color = '#27ae60'
        elif interacciones_count > 2:
            border_color = '#f39c12'
        else:
            border_color = '#e74c3c'
        
        clientes_html += f"""
        <div class="cliente-card" style="border-top: 4px solid {border_color};">
            <div class="cliente-header">
                <div class="cliente-avatar" style="background: {border_color};">{inicial}</div>
                <div class="cliente-info">
                    <h3>{c.nombre}</h3>
                    <p>{c.empresa or 'Sin empresa'}</p>
                </div>
            </div>
            <div class="cliente-body">
                <div class="cliente-contacto">
                    {f'<span>📧 {c.email}</span>' if c.email else ''}
                    {f'<span>📱 {c.telefono}</span>' if c.telefono else ''}
                    {f'<span>📍 {c.direccion[:30]}...</span>' if c.direccion else ''}
                </div>
                <div style="display: flex; justify-content: space-between; margin-top: 10px; font-size: 12px; color: #666;">
                    <span>📊 {interacciones_count} interacciones</span>
                    <span>🕐 {ultima_interaccion}</span>
                </div>
                {f'<p style="background:#f5f5f5; padding:8px; border-radius:6px; margin-top:10px; font-size:13px;">📝 {c.notas[:60]}...</p>' if c.notas else ''}
            </div>
            <div class="cliente-footer">
                <small>👤 {c.usuario.nombre_completo}</small>
                <div>
                    <a href="/cliente/{c.id}" class="btn btn-primary btn-sm">👁️ Ver</a>
                    <a href="/cliente/editar/{c.id}" class="btn btn-warning btn-sm">✏️</a>
                    <a href="/cliente/eliminar/{c.id}" class="btn btn-danger btn-sm" onclick="return confirm('¿Eliminar cliente?')">🗑️</a>
                </div>
            </div>
        </div>
        """
    
    content = f"""
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 25px;">
            <h2 style="margin:0;">👥 Gestión de Clientes</h2>
            <a href="/cliente/nuevo" class="btn btn-success">➕ Nuevo Cliente</a>
        </div>
        
        <!-- KPIs de clientes -->
        <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px;">
            <div style="background: linear-gradient(135deg, #667eea, #764ba2); color: white; padding: 15px; border-radius: 12px;">
                <div style="font-size: 14px; opacity: 0.9;">👥 Total Clientes</div>
                <div style="font-size: 28px; font-weight: bold;">{total_clientes}</div>
            </div>
            <div style="background: linear-gradient(135deg, #f093fb, #f5576c); color: white; padding: 15px; border-radius: 12px;">
                <div style="font-size: 14px; opacity: 0.9;">📧 Con Email</div>
                <div style="font-size: 28px; font-weight: bold;">{clientes_con_email}</div>
            </div>
            <div style="background: linear-gradient(135deg, #4facfe, #00f2fe); color: white; padding: 15px; border-radius: 12px;">
                <div style="font-size: 14px; opacity: 0.9;">📱 Con Teléfono</div>
                <div style="font-size: 28px; font-weight: bold;">{clientes_con_telefono}</div>
            </div>
            <div style="background: linear-gradient(135deg, #43e97b, #38f9d7); color: white; padding: 15px; border-radius: 12px;">
                <div style="font-size: 14px; opacity: 0.9;">🏢 Empresas</div>
                <div style="font-size: 28px; font-weight: bold;">{len(empresas)}</div>
            </div>
        </div>
        
        <!-- Buscador -->
        <div style="display: flex; gap: 10px; margin-bottom: 20px;">
            <form method="GET" style="flex: 1; display: flex; gap: 10px;">
                <input type="text" name="busqueda" placeholder="🔍 Buscar por nombre, email o empresa..." value="{busqueda}" 
                       style="flex:1; padding:12px; border:1px solid #ddd; border-radius:8px; font-size:14px;">
                <button type="submit" class="btn btn-primary">Buscar</button>
                {f'<a href="/clientes" class="btn" style="background:#95a5a6;color:white;">Limpiar</a>' if busqueda else ''}
            </form>
        </div>
        
        <!-- Lista de clientes -->
        <div class="clientes-grid">
            {clientes_html if clientes_list else '<div style="grid-column:1/-1; text-align:center; padding:40px; color:#999;"><p style="font-size:48px;">👥</p><p>No hay clientes. ¡Añade tu primer cliente!</p></div>'}
        </div>
        
        {f'<p style="text-align:center; margin-top:20px; color:#666;">Mostrando {min(20, len(clientes_list))} de {len(clientes_list)} clientes</p>' if len(clientes_list) > 20 else ''}
    """
    return base_html(content, "Clientes")

@app.route('/cliente/nuevo', methods=['GET', 'POST'])
def nuevo_cliente():
    if 'user_id' not in session:
        return redirect('/login')
    
    user = Usuario.query.get(session.get('user_id'))
    if not user:
        session.clear()
        return redirect('/login')
    
    if request.method == 'POST':
        cliente = Cliente(
            nombre=request.form['nombre'],
            email=request.form.get('email', ''),
            telefono=request.form.get('telefono', ''),
            empresa=request.form.get('empresa', ''),
            direccion=request.form.get('direccion', ''),
            notas=request.form.get('notas', ''),
            usuario_id=user.id
        )
        db.session.add(cliente)
        db.session.commit()
        flash('✅ Cliente creado correctamente')
        return redirect(f'/cliente/{cliente.id}')
    
    content = """
        <h2>➕ Nuevo Cliente</h2>
        <form method="POST" style="max-width: 600px;">
            <div class="form-group">
                <label>Nombre completo *</label>
                <input type="text" name="nombre" class="form-control" required>
            </div>
            <div class="form-group">
                <label>Email</label>
                <input type="email" name="email" class="form-control">
            </div>
            <div class="form-group">
                <label>Teléfono</label>
                <input type="tel" name="telefono" class="form-control">
            </div>
            <div class="form-group">
                <label>Empresa</label>
                <input type="text" name="empresa" class="form-control">
            </div>
            <div class="form-group">
                <label>Dirección</label>
                <input type="text" name="direccion" class="form-control">
            </div>
            <div class="form-group">
                <label>Notas</label>
                <textarea name="notas" class="form-control" rows="4"></textarea>
            </div>
            <button type="submit" class="btn btn-primary">Guardar Cliente</button>
            <a href="/clientes" class="btn" style="background:#95a5a6;color:white;">Cancelar</a>
        </form>
    """
    return base_html(content, "Nuevo Cliente")

@app.route('/cliente/<int:id>')
def ver_cliente(id):
    if 'user_id' not in session:
        return redirect('/login')
    
    cliente = Cliente.query.get_or_404(id)
    user = Usuario.query.get(session.get('user_id'))
    
    # Verificar permisos
    if session.get('rol') != 'admin' and cliente.usuario_id != user.id:
        flash('❌ No tienes permiso para ver este cliente')
        return redirect('/clientes')
    
    interacciones = Interaccion.query.filter_by(cliente_id=cliente.id).order_by(Interaccion.id.desc()).all()
    
    interacciones_html = ""
    for i in interacciones:
        icono = {'Llamada': '📞', 'Email': '📧', 'Reunión': '🤝', 'Nota': '📝'}.get(i.tipo, '📌')
        interacciones_html += f"""
        <div class="interaccion-item">
            <div class="interaccion-icon">{icono}</div>
            <div class="interaccion-content">
                <div class="interaccion-header">
                    <span class="interaccion-tipo">{i.tipo}</span>
                    <span class="interaccion-fecha">{i.fecha}</span>
                </div>
                <p>{i.descripcion}</p>
                <small>👤 {i.usuario.nombre_completo}</small>
            </div>
        </div>
        """
    
    content = f"""
        <h2>👤 {cliente.nombre}</h2>
        
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
            <div class="card">
                <h3>📋 Información</h3>
                <p><strong>Empresa:</strong> {cliente.empresa or '-'}</p>
                <p><strong>Email:</strong> {cliente.email or '-'}</p>
                <p><strong>Teléfono:</strong> {cliente.telefono or '-'}</p>
                <p><strong>Dirección:</strong> {cliente.direccion or '-'}</p>
                <p><strong>Notas:</strong><br>{cliente.notas or '-'}</p>
                <p><strong>Creado:</strong> {cliente.fecha_creacion} por {cliente.usuario.nombre_completo}</p>
                
                <div style="margin-top:20px;">
                    <a href="/cliente/editar/{cliente.id}" class="btn btn-warning">✏️ Editar</a>
                    <a href="/clientes" class="btn" style="background:#95a5a6;color:white;">← Volver</a>
                </div>
            </div>
            
            <div class="card">
                <h3>📝 Añadir Interacción</h3>
                <form action="/cliente/{cliente.id}/interaccion" method="POST">
                    <div class="form-group">
                        <label>Tipo</label>
                        <select name="tipo" class="form-control">
                            <option value="Llamada">📞 Llamada</option>
                            <option value="Email">📧 Email</option>
                            <option value="Reunión">🤝 Reunión</option>
                            <option value="Nota">📝 Nota</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Descripción</label>
                        <textarea name="descripcion" class="form-control" rows="3" required></textarea>
                    </div>
                    <button type="submit" class="btn btn-primary">💾 Guardar Interacción</button>
                </form>
            </div>
        </div>
        
        <div class="card" style="margin-top:20px;">
            <h3>📅 Historial de Interacciones</h3>
            <div class="interaccion-timeline">
                {interacciones_html if interacciones else '<p style="color:#999;">No hay interacciones registradas</p>'}
            </div>
        </div>
    """
    return base_html(content, f"Cliente: {cliente.nombre}")

@app.route('/cliente/<int:id>/interaccion', methods=['POST'])
def añadir_interaccion(id):
    if 'user_id' not in session:
        return redirect('/login')
    
    cliente = Cliente.query.get_or_404(id)
    user = Usuario.query.get(session.get('user_id'))
    
    interaccion = Interaccion(
        tipo=request.form['tipo'],
        descripcion=request.form['descripcion'],
        cliente_id=cliente.id,
        usuario_id=user.id
    )
    db.session.add(interaccion)
    db.session.commit()
    flash('✅ Interacción registrada')
    return redirect(f'/cliente/{cliente.id}')

@app.route('/cliente/editar/<int:id>', methods=['GET', 'POST'])
def editar_cliente(id):
    if 'user_id' not in session:
        return redirect('/login')
    
    cliente = Cliente.query.get_or_404(id)
    user = Usuario.query.get(session.get('user_id'))
    
    if session.get('rol') != 'admin' and cliente.usuario_id != user.id:
        flash('❌ No tienes permiso')
        return redirect('/clientes')
    
    if request.method == 'POST':
        cliente.nombre = request.form['nombre']
        cliente.email = request.form.get('email', '')
        cliente.telefono = request.form.get('telefono', '')
        cliente.empresa = request.form.get('empresa', '')
        cliente.direccion = request.form.get('direccion', '')
        cliente.notas = request.form.get('notas', '')
        db.session.commit()
        flash('✅ Cliente actualizado')
        return redirect(f'/cliente/{cliente.id}')
    
    content = f"""
        <h2>✏️ Editar Cliente</h2>
        <form method="POST" style="max-width: 600px;">
            <div class="form-group">
                <label>Nombre completo *</label>
                <input type="text" name="nombre" class="form-control" value="{cliente.nombre}" required>
            </div>
            <div class="form-group">
                <label>Email</label>
                <input type="email" name="email" class="form-control" value="{cliente.email or ''}">
            </div>
            <div class="form-group">
                <label>Teléfono</label>
                <input type="tel" name="telefono" class="form-control" value="{cliente.telefono or ''}">
            </div>
            <div class="form-group">
                <label>Empresa</label>
                <input type="text" name="empresa" class="form-control" value="{cliente.empresa or ''}">
            </div>
            <div class="form-group">
                <label>Dirección</label>
                <input type="text" name="direccion" class="form-control" value="{cliente.direccion or ''}">
            </div>
            <div class="form-group">
                <label>Notas</label>
                <textarea name="notas" class="form-control" rows="4">{cliente.notas or ''}</textarea>
            </div>
            <button type="submit" class="btn btn-primary">Actualizar</button>
            <a href="/cliente/{cliente.id}" class="btn" style="background:#95a5a6;color:white;">Cancelar</a>
        </form>
    """
    return base_html(content, "Editar Cliente")

@app.route('/cliente/eliminar/<int:id>')
def eliminar_cliente(id):
    if 'user_id' not in session:
        return redirect('/login')
    
    cliente = Cliente.query.get_or_404(id)
    user = Usuario.query.get(session.get('user_id'))
    
    if session.get('rol') != 'admin' and cliente.usuario_id != user.id:
        flash('❌ No tienes permiso')
        return redirect('/clientes')
    
    db.session.delete(cliente)
    db.session.commit()
    flash('✅ Cliente eliminado')
    return redirect('/clientes')
@app.route('/admin/exportar-todo')
def exportar_todo():
    if 'user_id' not in session or session.get('rol') != 'admin':
        return redirect('/dashboard')
    
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    import io
    
    wb = Workbook()
    
    # Hoja 1: Empleados
    ws1 = wb.active
    ws1.title = "Empleados"
    headers1 = ['ID', 'Usuario', 'Nombre', 'Departamento', 'Cargo', 'Rol', 'Tareas Totales', 'Completadas', 'Productividad']
    for col, h in enumerate(headers1, 1):
        cell = ws1.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="1a73e8", end_color="1a73e8", fill_type="solid")
        cell.font = Font(bold=True, color="FFFFFF")
    
    empleados = Usuario.query.all()
    for row, e in enumerate(empleados, 2):
        total = len(e.tareas)
        completadas = sum(1 for t in e.tareas if t.completada)
        prod = f"{(completadas/total*100):.1f}%" if total > 0 else "0%"
        ws1.cell(row=row, column=1, value=e.id)
        ws1.cell(row=row, column=2, value=e.username)
        ws1.cell(row=row, column=3, value=e.nombre_completo)
        ws1.cell(row=row, column=4, value=e.departamento)
        ws1.cell(row=row, column=5, value=e.cargo)
        ws1.cell(row=row, column=6, value=e.rol)
        ws1.cell(row=row, column=7, value=total)
        ws1.cell(row=row, column=8, value=completadas)
        ws1.cell(row=row, column=9, value=prod)
    
    # Hoja 2: Tareas
    ws2 = wb.create_sheet("Tareas")
    headers2 = ['ID', 'Título', 'Descripción', 'Asignado a', 'Departamento', 'Prioridad', 'Fecha Límite', 'Estado', 'Creada']
    for col, h in enumerate(headers2, 1):
        cell = ws2.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="1a73e8", end_color="1a73e8", fill_type="solid")
        cell.font = Font(bold=True, color="FFFFFF")
    
    tareas = Tarea.query.all()
    for row, t in enumerate(tareas, 2):
        ws2.cell(row=row, column=1, value=t.id)
        ws2.cell(row=row, column=2, value=t.titulo)
        ws2.cell(row=row, column=3, value=t.descripcion or '')
        ws2.cell(row=row, column=4, value=t.asignado.nombre_completo)
        ws2.cell(row=row, column=5, value=t.asignado.departamento)
        ws2.cell(row=row, column=6, value=t.prioridad)
        ws2.cell(row=row, column=7, value=t.fecha_limite or '')
        ws2.cell(row=row, column=8, value='Completada' if t.completada else 'Pendiente')
        ws2.cell(row=row, column=9, value=t.fecha_creacion)
    
    # Hoja 3: Clientes
    ws3 = wb.create_sheet("Clientes")
    headers3 = ['ID', 'Nombre', 'Email', 'Teléfono', 'Empresa', 'Dirección', 'Creado por', 'Fecha']
    for col, h in enumerate(headers3, 1):
        cell = ws3.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="1a73e8", end_color="1a73e8", fill_type="solid")
        cell.font = Font(bold=True, color="FFFFFF")
    
    clientes = Cliente.query.all()
    for row, c in enumerate(clientes, 2):
        ws3.cell(row=row, column=1, value=c.id)
        ws3.cell(row=row, column=2, value=c.nombre)
        ws3.cell(row=row, column=3, value=c.email or '')
        ws3.cell(row=row, column=4, value=c.telefono or '')
        ws3.cell(row=row, column=5, value=c.empresa or '')
        ws3.cell(row=row, column=6, value=c.direccion or '')
        ws3.cell(row=row, column=7, value=c.usuario.nombre_completo)
        ws3.cell(row=row, column=8, value=c.fecha_creacion)
    
    # Ajustar ancho de columnas
    for ws in [ws1, ws2, ws3]:
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[col_letter].width = min(max_len + 2, 40)
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'reporte_completo_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'
    )

# ========== CHAT GRUPAL POR DEPARTAMENTOS ==========
@app.route('/chat-grupal')
def chat_grupal():
    if 'user_id' not in session:
        return redirect('/login')
    
    user = Usuario.query.get(session.get('user_id'))
    if not user:
        session.clear()
        return redirect('/login')
    
    # Determinar departamentos disponibles
    if session.get('rol') == 'admin':
        departamentos = db.session.query(Usuario.departamento).distinct().all()
        deptos_disponibles = [d[0] for d in departamentos]
    else:
        # Si el empleado no tiene departamento, usar 'General'
        deptos_disponibles = [user.departamento] if user.departamento else ['General']
    
    # Sidebar con departamentos
    sidebar_html = ""
    for depto in deptos_disponibles:
        # Contar mensajes no leídos (simplificado - todos como no leídos por ahora)
        sidebar_html += f"""
        <div class="chat-grupal-item" onclick="window.location.href='/chat-grupal/{depto}'">
            🏢 {depto}
        </div>
        """
    
    content = f"""
        <h2>💬 Chat por Departamentos</h2>
        <div class="chat-grupal-container">
            <div class="chat-grupal-sidebar">
                {sidebar_html if sidebar_html else '<p style="padding:15px;">No hay departamentos</p>'}
            </div>
            <div class="chat-grupal-main">
                <div class="chat-grupal-header">👈 Selecciona un departamento</div>
                <div class="chat-grupal-messages" style="display:flex;align-items:center;justify-content:center;color:#999;">
                    Elige un departamento para ver la conversación
                </div>
            </div>
        </div>
    """
    return base_html(content, "Chat Grupal")

@app.route('/chat-grupal/<depto>', methods=['GET', 'POST'])
def chat_grupal_depto(depto):
    if 'user_id' not in session:
        return redirect('/login')
    
    user = Usuario.query.get(session.get('user_id'))
    if not user:
        session.clear()
        return redirect('/login')
    
    # Verificar permisos
    if session.get('rol') != 'admin' and user.departamento != depto:
        flash('❌ No tienes acceso a este departamento')
        return redirect('/chat-grupal')
    
    if request.method == 'POST':
        contenido = request.form.get('mensaje')
        if contenido:
            msg = MensajeGrupal(
                contenido=contenido,
                departamento=depto,
                usuario_id=user.id
            )
            db.session.add(msg)
            db.session.commit()
        return redirect(f'/chat-grupal/{depto}')
    
    # Obtener mensajes del departamento
    mensajes = MensajeGrupal.query.filter_by(departamento=depto).order_by(MensajeGrupal.id.desc()).limit(50).all()
    mensajes = list(reversed(mensajes))
    
    mensajes_html = ""
    for m in mensajes:
        es_mio = m.usuario_id == user.id
        mensajes_html += f"""
        <div class="chat-message {'sent' if es_mio else ''}">
            <div class="chat-message-bubble">
                <strong>{m.usuario.nombre_completo}</strong><br>
                {m.contenido}
                <div class="chat-message-time">{m.fecha}</div>
            </div>
        </div>
        """
    
    # Sidebar para admin (todos los deptos) o empleado (solo el suyo)
    if session.get('rol') == 'admin':
        departamentos = db.session.query(Usuario.departamento).distinct().all()
        deptos_disponibles = [d[0] for d in departamentos]
    else:
        deptos_disponibles = [user.departamento]
    
    sidebar_html = ""
    for d in deptos_disponibles:
        active_class = 'active' if d == depto else ''
        sidebar_html += f"""
        <div class="chat-grupal-item {active_class}" onclick="window.location.href='/chat-grupal/{d}'">
            🏢 {d}
        </div>
        """
    
    content = f"""
        <h2>💬 Chat de {depto}</h2>
        <div class="chat-grupal-container">
            <div class="chat-grupal-sidebar">
                {sidebar_html}
            </div>
            <div class="chat-grupal-main">
                <div class="chat-grupal-header">
                    🏢 {depto} | 👥 {Usuario.query.filter_by(departamento=depto).count()} miembros
                </div>
                <div class="chat-grupal-messages" id="chat-messages">
                    {mensajes_html if mensajes else '<p style="text-align:center;color:#999;padding:20px;">No hay mensajes. ¡Empieza la conversación!</p>'}
                </div>
                <form method="POST" class="chat-grupal-input">
                    <input type="text" name="mensaje" placeholder="Escribe un mensaje para {depto}..." required autocomplete="off" style="flex:1; padding:10px; border:1px solid #ddd; border-radius:20px;">
                    <button type="submit" class="btn btn-primary">📤 Enviar</button>
                </form>
            </div>
        </div>
        
        <script>
            const messagesDiv = document.getElementById('chat-messages');
            if (messagesDiv) messagesDiv.scrollTop = messagesDiv.scrollHeight;
        </script>
    """
    return base_html(content, f"Chat - {depto}")

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect('/login')
        user = db.session.get(Usuario, session.get('user_id'))
        if not user:
            session.clear()
            return redirect('/login')
        return f(*args, **kwargs)
    return decorated_function

# ========== INICIALIZACIÓN ==========
with app.app_context():
    db.create_all()
    if Usuario.query.count() == 0:
        admin = Usuario(
            username='admin',
            password_hash=generate_password_hash('admin123', method='pbkdf2:sha256'),
            nombre_completo='Administrador',
            departamento='Dirección',
            cargo='Administrador',
            rol='admin'
        )
        db.session.add(admin)
        db.session.commit()
        print("✅ Admin creado: admin / admin123")
        print("📝 La base de datos está vacía. Registra empleados desde la app.")

if __name__ == '__main__':
    app.run(debug=True)