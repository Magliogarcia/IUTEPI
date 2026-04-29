# Archivo: app.py
from flask import Flask, render_template, request, redirect, url_for, session, jsonify, Response, send_from_directory
import sqlite3
import os
import sys
import base64
import re
import threading
import webbrowser
import random
import io
import xlsxwriter
from datetime import datetime, timedelta, time

if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
    APPLICATION_PATH = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))
    APPLICATION_PATH = BASE_DIR

app = Flask(__name__, template_folder=os.path.join(BASE_DIR, 'templates'), static_folder=os.path.join(BASE_DIR, 'static'))
app.secret_key = 'iutepi_secreto_super_seguro' 
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(minutes=3)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024 

DB_PATH = os.path.join(APPLICATION_PATH, 'iutepi.db')
FOTOS_DIR = os.path.join(APPLICATION_PATH, 'fotos_asistencia')
os.makedirs(FOTOS_DIR, exist_ok=True)

MESES_ES = {"01":"Enero", "02":"Febrero", "03":"Marzo", "04":"Abril", "05":"Mayo", "06":"Junio", 
            "07":"Julio", "08":"Agosto", "09":"Septiembre", "10":"Octubre", "11":"Noviembre", "12":"Diciembre"}
DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

# ==========================================
# GENERADOR AUTOMÁTICO DE DATOS Y HORARIOS AVANZADOS
# ==========================================
def seed_database_for_presentation(conn):
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) as c FROM personal")
    if cursor.fetchone()['c'] > 0: return 

    print("Generando base de datos de prueba avanzada para la defensa...")
    
    empleados = [
        ('11111111', 'Carlos Mendoza', 'Docente', 'Profesor Titular'),     
        ('22222222', 'Maria Perez', 'Docente', 'Profesora Invitada'),      
        ('33333333', 'Luis Rodriguez', 'Docente', 'Profesor Auxiliar'),    
        ('44444444', 'Ana Gomez', 'Administrativo', 'Secretaria'),         
        ('55555555', 'Jose Silva', 'Ambiente', 'Mantenimiento'),           
        ('66666666', 'Elena Torres', 'Administrativo', 'Coordinadora'),
        ('77777777', 'Miguel Castro', 'Ambiente', 'Seguridad'), # Turno 12am - 12pm
        ('88888888', 'Patricia Diaz', 'Docente', 'Profesor Titular'),      
        ('99999999', 'Roberto Ruiz', 'Administrativo', 'Contador'),
        ('10101010', 'Sofia Vargas', 'Ambiente', 'Limpieza'),
        ('12121212', 'Fernando Gil', 'Docente', 'Profesor Auxiliar'),
        ('13131313', 'Lucía Rios', 'Administrativo', 'Asistente'),
        ('14141414', 'Diego Blanco', 'Ambiente', 'Seguridad'), # Turno 12pm - 12am
        ('15151515', 'Carmen Vega', 'Docente', 'Profesora Invitada'),
        ('16161616', 'Andres Peña', 'Administrativo', 'Director')
    ]
    for emp in empleados:
        cursor.execute("INSERT INTO personal (cedula, nombre, departamento, cargo) VALUES (?,?,?,?)", emp)

    # 1. HORARIOS PARA DOCENTES (INTERCALADOS Y VARIADOS)
    horarios_docentes = {
        '11111111': [('Matemática I', '07:40', '09:10'), ('Estadística', '12:15', '14:00')],
        '22222222': [('Programación Web', '08:30', '10:00'), ('Bases de Datos', '12:30', '14:30')],
        '33333333': [('Redes de Computadoras', '07:40', '09:40'), ('Seguridad Informática', '11:00', '12:30')],
        '88888888': [('Arquitectura del Computador', '09:00', '11:00'), ('Sistemas Operativos', '13:00', '14:30')],
        '12121212': [('Ingeniería de Software', '07:40', '09:40'), ('Gestión de Proyectos', '10:15', '11:45')],
        '15151515': [('Inteligencia Artificial', '08:00', '09:30'), ('Robótica Avanzada', '11:30', '13:30')]
    }
    for cedula, materias in horarios_docentes.items():
        for dia in range(5): # Lunes a Viernes
            for m in materias:
                cursor.execute("INSERT INTO horarios (cedula, materia, dia_semana, hora_entrada, hora_salida) VALUES (?, ?, ?, ?, ?)", (cedula, m[0], dia, m[1], m[2]))
        # Clase extra los Sábados para Carlos Mendoza
        if cedula == '11111111': cursor.execute("INSERT INTO horarios (cedula, materia, dia_semana, hora_entrada, hora_salida) VALUES (?, ?, ?, ?, ?)", (cedula, 'Seminario de Tesis', 5, '07:40', '11:00'))

    # 2. HORARIOS ADMINISTRATIVO Y AMBIENTE (CONTINUOS 7:00 a 2:00 PM)
    para_generales = ['44444444', '55555555', '66666666', '99999999', '10101010', '13131313', '16161616']
    for c in para_generales:
        for dia in range(5):
            cursor.execute("INSERT INTO horarios (cedula, materia, dia_semana, hora_entrada, hora_salida) VALUES (?, ?, ?, ?, ?)", (c, 'Jornada Laboral', dia, '07:00', '14:00'))

    # 3. HORARIOS DE SEGURIDAD (ROTATIVOS 24 HORAS LUN-DOM)
    for dia in range(7):
        cursor.execute("INSERT INTO horarios (cedula, materia, dia_semana, hora_entrada, hora_salida) VALUES (?, ?, ?, ?, ?)", ('77777777', 'Turno Madrugada/Mañana', dia, '00:00', '12:00'))
        cursor.execute("INSERT INTO horarios (cedula, materia, dia_semana, hora_entrada, hora_salida) VALUES (?, ?, ?, ?, ?)", ('14141414', 'Turno Tarde/Noche', dia, '12:01', '23:59'))

    # NOVEDADES
    cursor.execute("INSERT INTO novedades (cedula, tipo, fecha_inicio, fecha_fin, descripcion) VALUES (?,?,?,?,?)", ('44444444', 'VACACIONES', '2026-02-10', '2026-02-20', 'Vacaciones anuales'))
    cursor.execute("INSERT INTO novedades (cedula, tipo, fecha_inicio, fecha_fin, descripcion) VALUES (?,?,?,?,?)", ('55555555', 'REPOSO MÉDICO', '2026-02-05', '2026-02-12', 'Reposo por Gripe fuerte'))

    # ASISTENCIAS FICTICIAS FEBRERO
    for day in range(2, 28):
        fecha_str = f"2026-02-{day:02d}"
        dt = datetime.strptime(fecha_str, "%Y-%m-%d")
        dia_semana = dt.weekday()
        
        for emp in empleados:
            cedula, nombre, depto, cargo = emp
            if cedula == '44444444' and 10 <= day <= 20: continue
            if cedula == '55555555' and 5 <= day <= 12: continue

            # Consultar cual es el bloque real de este empleado este dia
            cursor.execute("SELECT MIN(hora_entrada) as ent, MAX(hora_salida) as sal FROM horarios WHERE cedula=? AND dia_semana=?", (cedula, dia_semana))
            h = cursor.fetchone()
            if not h or not h['ent']: continue # Si no tiene horario este dia, no asiste

            hora_meta_ent = h['ent']
            hora_meta_sal = h['sal']

            est_ent, est_sal, obs = "Puntual", "Correcta", ""
            
            # Algoritmo de simulación
            ent_dt = datetime.strptime(hora_meta_ent, "%H:%M")
            sal_dt = datetime.strptime(hora_meta_sal, "%H:%M")
            
            if cedula == '11111111': # Perfecto
                ent_real = (ent_dt - timedelta(minutes=10)).strftime("%H:%M:%S")
                sal_real = (sal_dt + timedelta(minutes=5)).strftime("%H:%M:%S")
            elif cedula == '22222222': # Tarde
                ent_real = (ent_dt + timedelta(minutes=25)).strftime("%H:%M:%S")
                sal_real = (sal_dt + timedelta(minutes=5)).strftime("%H:%M:%S")
                est_ent = "TARDÍA"
            elif cedula == '88888888': # Olvida Salida
                ent_real = (ent_dt - timedelta(minutes=5)).strftime("%H:%M:%S")
                sal_real = (sal_dt + timedelta(minutes=45)).strftime("%H:%M:%S")
                if random.random() > 0.5:
                    sal_real, est_sal, obs = sal_real, "NO MARCO", "Cierre Automático"
            else: # Regulares
                r = random.random()
                if r > 0.3:
                    ent_real = (ent_dt - timedelta(minutes=5)).strftime("%H:%M:%S")
                else:
                    ent_real = (ent_dt + timedelta(minutes=20)).strftime("%H:%M:%S")
                    est_ent = "TARDÍA"
                    
                if random.random() > 0.9:
                    sal_real = (sal_dt - timedelta(minutes=60)).strftime("%H:%M:%S")
                    est_sal, obs = "ANTICIPADA", "Permiso personal"
                else:
                    sal_real = (sal_dt + timedelta(minutes=5)).strftime("%H:%M:%S")

            cursor.execute("INSERT INTO asistencia (cedula, nombre, departamento, cargo, fecha, hora_entrada, estado_entrada, hora_salida, estado_salida, observacion) VALUES (?,?,?,?,?,?,?,?,?,?)",
                           (cedula, nombre, depto, cargo, fecha_str, ent_real, est_ent, sal_real, est_sal, obs))
    conn.commit()

# ==========================================
# CONEXIÓN E INICIALIZACIÓN
# ==========================================
def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row 
    return conn

try:
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS usuarios (id INTEGER PRIMARY KEY AUTOINCREMENT, usuario TEXT, password TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS personal (id INTEGER PRIMARY KEY AUTOINCREMENT, cedula TEXT, nombre TEXT, departamento TEXT, cargo TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS asistencia (id INTEGER PRIMARY KEY AUTOINCREMENT, cedula TEXT, nombre TEXT, departamento TEXT, cargo TEXT, fecha TEXT, hora_entrada TEXT, estado_entrada TEXT, hora_salida TEXT, estado_salida TEXT, observacion TEXT, foto_entrada TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS novedades (id INTEGER PRIMARY KEY AUTOINCREMENT, cedula TEXT, tipo TEXT, fecha_inicio TEXT, fecha_fin TEXT, descripcion TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS horarios (id INTEGER PRIMARY KEY AUTOINCREMENT, cedula TEXT, materia TEXT, dia_semana INTEGER, hora_entrada TEXT, hora_salida TEXT)''')
    
    try: cursor.execute("ALTER TABLE horarios ADD COLUMN materia TEXT NOT NULL DEFAULT 'Clase General'")
    except: pass
    try: cursor.execute("ALTER TABLE asistencia ADD COLUMN foto_salida TEXT DEFAULT ''")
    except: pass
    
    cursor.execute("SELECT COUNT(*) as count FROM usuarios")
    if cursor.fetchone()['count'] == 0:
        cursor.execute("INSERT INTO usuarios (usuario, password) VALUES (?, ?)", ('admin', 'Admin28*'))
    
    seed_database_for_presentation(conn)
    conn.commit()
    cursor.close()
    conn.close()
except Exception as e:
    print("Error inicializando base de datos:", e)

@app.before_request
def make_session_permanent():
    session.permanent = True

@app.route('/foto/<filename>')
def serve_foto(filename):
    return send_from_directory(FOTOS_DIR, filename)

def to_12h(val):
    if not val or val in ["0:00:00", "00:00:00", ""]: return None
    try: return datetime.strptime(val, "%H:%M:%S").strftime("%I:%M:%S %p")
    except: 
        try: return datetime.strptime(val, "%H:%M").strftime("%I:%M %p")
        except: return val

def parse_time(val):
    if isinstance(val, str): 
        try: return datetime.strptime(val, "%H:%M:%S").time()
        except: return datetime.strptime(val, "%H:%M").time()
    return val

def clave_segura(pwd):
    if len(pwd) < 8: return False, "Mínimo 8 caracteres."
    if not re.search(r"[A-Z]", pwd): return False, "Debe incluir una Mayúscula."
    if not re.search(r"[a-z]", pwd): return False, "Debe incluir una Minúscula."
    if not re.search(r"\d", pwd): return False, "Debe incluir un Número."
    if not re.search(r"[@$!%*?&#.\-_]", pwd): return False, "Debe incluir un símbolo."
    return True, ""

def auto_marcar_salidas(conn):
    cursor = conn.cursor()
    cursor.execute("SELECT id, cedula, fecha FROM asistencia WHERE hora_salida='' OR hora_salida='Pendiente' OR hora_salida IS NULL")
    pendientes = cursor.fetchall()
    ahora = datetime.now()
    
    for p in pendientes:
        fecha_str = p['fecha']
        cedula = p['cedula']
        try: fecha_dt = datetime.strptime(fecha_str, "%Y-%m-%d")
        except: continue
        
        dia_semana = fecha_dt.weekday()
        meta_salida = None
        
        cursor.execute("SELECT MAX(hora_salida) as sal FROM horarios WHERE cedula=? AND dia_semana=?", (cedula, dia_semana))
        h = cursor.fetchone()
        if h and h['sal']: 
            meta_salida = parse_time(h['sal'])
            dt_meta_salida = datetime.combine(fecha_dt.date(), meta_salida)
            dt_limite = dt_meta_salida + timedelta(minutes=30)
            
            if ahora >= dt_limite:
                cursor.execute("UPDATE asistencia SET hora_salida=?, estado_salida='NO MARCO', observacion='Cierre Automático (Sistema)' WHERE id=?", 
                               (meta_salida.strftime("%H:%M:%S"), p['id']))
    conn.commit()

# ==========================================
# RUTAS DE LOGIN Y GESTIÓN
# ==========================================
@app.route('/login', methods=['GET', 'POST'])
def login():
    msg = ""
    if request.method == 'POST':
        usuario = request.form['usuario']
        password = request.form['password']
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM usuarios WHERE usuario=? AND password=?", (usuario, password))
        user = cursor.fetchone()
        conn.close()
        if user:
            session['usuario'] = user['usuario']
            return redirect(url_for('admin_panel'))
        else:
            msg = "❌ Usuario o contraseña incorrectos"
    return render_template('login.html', msg=msg)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

@app.route('/crear_admin', methods=['GET', 'POST'])
def crear_admin():
    if 'usuario' not in session: return redirect(url_for('login'))
    mensaje = ""
    conn = get_db_connection()
    cursor = conn.cursor()
    if request.method == 'POST':
        nuevo_usuario, nuevo_pass = request.form['nuevo_usuario'], request.form['nuevo_pass']
        es_segura, error_msg = clave_segura(nuevo_pass)
        if not es_segura: mensaje = f"<div class='alert error'>❌ {error_msg}</div>"
        else:
            cursor.execute("SELECT id FROM usuarios WHERE usuario=?", (nuevo_usuario,))
            if cursor.fetchone(): mensaje = f"<div class='alert error'>❌ El usuario '{nuevo_usuario}' ya existe.</div>"
            else:
                cursor.execute("INSERT INTO usuarios (usuario, password) VALUES (?, ?)", (nuevo_usuario, nuevo_pass))
                conn.commit()
                mensaje = "<div class='alert success'>✅ Admin creado con éxito.</div>"
    
    cursor.execute("SELECT * FROM usuarios")
    admins = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return render_template('crear_admin.html', mensaje=mensaje, usuario_actual=session['usuario'], admins=admins)

@app.route('/eliminar_admin/<int:id>')
def eliminar_admin(id):
    if 'usuario' not in session: return redirect(url_for('login'))
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT usuario FROM usuarios WHERE id=?", (id,))
    user = cursor.fetchone()
    if user and user['usuario'] != session['usuario'] and user['usuario'].lower() != 'admin':
        cursor.execute("DELETE FROM usuarios WHERE id=?", (id,))
        conn.commit()
    conn.close()
    return redirect(url_for('crear_admin'))

# ==========================================
# RUTAS DE PERSONAL
# ==========================================
@app.route('/registrar_personal', methods=['GET', 'POST'])
def registrar_personal():
    if 'usuario' not in session: return redirect(url_for('login'))
    mensaje, tipo_mensaje = "", ""
    if request.args.get('msg') == 'editado': mensaje, tipo_mensaje = "Datos actualizados exitosamente.", "success"
    elif request.args.get('msg') == 'error_cedula': mensaje, tipo_mensaje = "La cédula ya pertenece a otro empleado.", "error"
    elif request.args.get('msg') == 'error_nombre': mensaje, tipo_mensaje = "Ese Nombre y Apellido exactos ya existen.", "error"
        
    conn = get_db_connection()
    cursor = conn.cursor()
    if request.method == 'POST':
        cedula, nombre, apellido = request.form['cedula'], request.form['nombre'].strip(), request.form['apellido'].strip()
        departamento, cargo = request.form['departamento'], request.form['cargo'].strip()
        nombre_completo = f"{nombre} {apellido}"
        cursor.execute("SELECT id FROM personal WHERE cedula=?", (cedula,))
        if cursor.fetchone():
            mensaje, tipo_mensaje = "La cédula ya está registrada.", "error"
        else:
            cursor.execute("SELECT id FROM personal WHERE nombre=?", (nombre_completo,))
            if cursor.fetchone():
                mensaje, tipo_mensaje = "El empleado ya se encuentra registrado.", "error"
            else:
                cursor.execute("INSERT INTO personal (cedula, nombre, departamento, cargo) VALUES (?, ?, ?, ?)", (cedula, nombre_completo, departamento, cargo))
                conn.commit()
                mensaje, tipo_mensaje = "Personal registrado exitosamente.", "success"

    cursor.execute("SELECT * FROM personal ORDER BY departamento ASC, nombre ASC")
    personal_list = [dict(r) for r in cursor.fetchall()]
    
    personal_agrupado = {}
    for p in personal_list:
        dep = p['departamento']
        if dep not in personal_agrupado: personal_agrupado[dep] = []
        personal_agrupado[dep].append(p)
        
    conn.close()
    return render_template('registrar_personal.html', mensaje=mensaje, tipo_mensaje=tipo_mensaje, usuario_actual=session['usuario'], personal_agrupado=personal_agrupado)

@app.route('/editar_personal/<int:id>', methods=['GET', 'POST'])
def editar_personal(id):
    if 'usuario' not in session: return redirect(url_for('login'))
    conn = get_db_connection()
    cursor = conn.cursor()
    if request.method == 'POST':
        cedula, nombre, apellido = request.form['cedula'], request.form['nombre'].strip(), request.form['apellido'].strip()
        departamento, cargo = request.form['departamento'], request.form['cargo'].strip()
        nombre_completo = f"{nombre} {apellido}"
        cursor.execute("SELECT id FROM personal WHERE cedula=? AND id!=?", (cedula, id))
        if cursor.fetchone(): return redirect(url_for('registrar_personal', msg='error_cedula'))
        cursor.execute("SELECT id FROM personal WHERE nombre=? AND id!=?", (nombre_completo, id))
        if cursor.fetchone(): return redirect(url_for('registrar_personal', msg='error_nombre'))
        cursor.execute("UPDATE personal SET cedula=?, nombre=?, departamento=?, cargo=? WHERE id=?", (cedula, nombre_completo, departamento, cargo, id))
        conn.commit()
        conn.close()
        return redirect(url_for('registrar_personal', msg='editado'))

    cursor.execute("SELECT * FROM personal WHERE id=?", (id,))
    empleado = cursor.fetchone()
    conn.close()
    nom, ape = "", ""
    if empleado:
        partes = empleado['nombre'].split(' ', 1)
        nom, ape = partes[0], partes[1] if len(partes) > 1 else ""
    return render_template('editar_personal.html', empleado=empleado, nom=nom, ape=ape, usuario_actual=session['usuario'])

@app.route('/eliminar_personal/<int:id>')
def eliminar_personal(id):
    if 'usuario' not in session: return redirect(url_for('login'))
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM personal WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return redirect(url_for('registrar_personal'))

# ==========================================
# RUTAS DE HORARIOS (NUEVA MATEMÁTICA VISUAL)
# ==========================================
@app.route('/gestionar_horarios', methods=['GET', 'POST'])
def gestionar_horarios():
    if 'usuario' not in session: return redirect(url_for('login'))
    mensaje, tipo_mensaje = "", ""
    cedula_seleccionada = request.args.get('cedula') or (request.form.get('cedula') if request.method == 'POST' else "")
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    if request.method == 'POST' and 'guardar_horario' in request.form:
        materia, dia_str, h_ent_nueva, h_sal_nueva = request.form['materia'].strip(), request.form['dia_semana'], request.form['hora_entrada'], request.form['hora_salida']
        
        if not h_ent_nueva or not h_sal_nueva: 
            mensaje, tipo_mensaje = "Debes seleccionar las horas correctamente.", "error"
        elif h_sal_nueva <= h_ent_nueva: 
            mensaje, tipo_mensaje = "La hora de salida no puede ser menor o igual a la de entrada.", "error"
        else:
            cursor.execute("SELECT materia, hora_entrada, hora_salida FROM horarios WHERE cedula=? AND dia_semana=?", (cedula_seleccionada, dia_str))
            existentes = cursor.fetchall()
            hay_choque = False
            materia_choque = ""
            for e in existentes:
                if (h_ent_nueva < e['hora_salida']) and (h_sal_nueva > e['hora_entrada']):
                    hay_choque, materia_choque = True, e['materia']
                    break 

            if hay_choque:
                mensaje, tipo_mensaje = f"⛔ Error: Este horario choca con <b>'{materia_choque}'</b> ({e['hora_entrada']} - {e['hora_salida']}).", "error"
            else:
                cursor.execute("INSERT INTO horarios (cedula, materia, dia_semana, hora_entrada, hora_salida) VALUES (?, ?, ?, ?, ?)", (cedula_seleccionada, materia, dia_str, h_ent_nueva, h_sal_nueva))
                conn.commit()
                mensaje, tipo_mensaje = "Horario asignado exitosamente.", "success"

    cursor.execute("SELECT cedula, nombre, departamento FROM personal ORDER BY departamento ASC, nombre ASC")
    empleados = [dict(r) for r in cursor.fetchall()]
    
    # MATEMÁTICA DEL CALENDARIO
    horarios_semana = {0: [], 1: [], 2: [], 3: [], 4: [], 5: [], 6: []}
    nombre_empleado = ""
    grid_start_hour = 6
    grid_end_hour = 17 # De 6 AM a 5 PM por defecto
    
    if cedula_seleccionada:
        cursor.execute("SELECT nombre, departamento, cargo FROM personal WHERE cedula=?", (cedula_seleccionada,))
        fila = cursor.fetchone()
        if fila: 
            nombre_empleado = fila['nombre']
            if fila['departamento'] == 'Ambiente' and 'Seguridad' in fila['cargo']:
                grid_start_hour = 0
                grid_end_hour = 24
            elif fila['departamento'] == 'Docente':
                grid_start_hour = 7
                grid_end_hour = 16
            else:
                grid_start_hour = 6
                grid_end_hour = 15

        cursor.execute("SELECT * FROM horarios WHERE cedula=? ORDER BY dia_semana ASC, hora_entrada ASC", (cedula_seleccionada,))
        total_grid_mins = (grid_end_hour - grid_start_hour) * 60
        
        for h in cursor.fetchall():
            hd = dict(h)
            hd['hora_in_fmt'] = to_12h(hd['hora_entrada'])
            hd['hora_out_fmt'] = to_12h(hd['hora_salida'])
            
            # Cálculo de posiciones absolutas en la grilla CSS
            h_ent = parse_time(hd['hora_entrada'])
            h_sal = parse_time(hd['hora_salida'])
            start_min = h_ent.hour * 60 + h_ent.minute
            end_min = h_sal.hour * 60 + h_sal.minute
            
            top_pct = ((start_min - (grid_start_hour * 60)) / total_grid_mins) * 100
            height_pct = ((end_min - start_min) / total_grid_mins) * 100
            
            # Limitar visualmente si el horario excede la tabla
            if top_pct < 0: top_pct = 0
            if top_pct + height_pct > 100: height_pct = 100 - top_pct
            
            hd['top'] = top_pct
            hd['height'] = height_pct
            
            if hd['dia_semana'] in horarios_semana:
                horarios_semana[hd['dia_semana']].append(hd)
                
    conn.close()
    return render_template('gestionar_horarios.html', mensaje=mensaje, tipo_mensaje=tipo_mensaje, usuario_actual=session['usuario'], empleados=empleados, cedula_seleccionada=cedula_seleccionada, nombre_empleado=nombre_empleado, horarios_semana=horarios_semana, dias=DIAS_SEMANA, start_h=grid_start_hour, end_h=grid_end_hour)

@app.route('/eliminar_horario/<int:id>/<cedula>')
def eliminar_horario(id, cedula):
    if 'usuario' not in session: return redirect(url_for('login'))
    conn = get_db_connection()
    conn.execute("DELETE FROM horarios WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return redirect(url_for('gestionar_horarios', cedula=cedula))

# ==========================================
# RUTA DE RENDIMIENTO
# ==========================================
@app.route('/rendimiento', methods=['GET', 'POST'])
def rendimiento():
    if 'usuario' not in session: return redirect(url_for('login'))
    conn = get_db_connection()
    auto_marcar_salidas(conn) 
    
    desde = request.form.get('desde') if request.method == 'POST' else request.args.get('desde', '')
    hasta = request.form.get('hasta') if request.method == 'POST' else request.args.get('hasta', '')
    
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM personal ORDER BY departamento ASC, nombre ASC")
    personal = cursor.fetchall()
    
    if desde and hasta:
        cursor.execute("SELECT cedula, estado_entrada, estado_salida FROM asistencia WHERE fecha BETWEEN ? AND ?", (desde, hasta))
    else:
        cursor.execute("SELECT cedula, estado_entrada, estado_salida FROM asistencia")
        
    asistencias = cursor.fetchall()
    conn.close()
    
    stats_personal = {}
    for p in personal:
        stats_personal[p['cedula']] = {
            'nombre': p['nombre'], 'departamento': p['departamento'], 'cargo': p['cargo'],
            'total': 0, 'puntual': 0, 'tarde': 0, 'correcta': 0, 'anticipada': 0, 'no_marco': 0
        }
        
    for a in asistencias:
        c = a['cedula']
        if c in stats_personal:
            stats_personal[c]['total'] += 1
            if a['estado_entrada'] == 'Puntual': stats_personal[c]['puntual'] += 1
            elif a['estado_entrada'] == 'TARDÍA': stats_personal[c]['tarde'] += 1
            if a['estado_salida'] == 'Correcta': stats_personal[c]['correcta'] += 1
            elif a['estado_salida'] == 'ANTICIPADA': stats_personal[c]['anticipada'] += 1
            elif a['estado_salida'] == 'NO MARCO': stats_personal[c]['no_marco'] += 1
            
    return render_template('rendimiento.html', usuario_actual=session['usuario'], stats=stats_personal, desde=desde, hasta=hasta)

# ==========================================
# RUTA PRINCIPAL E INTELIGENCIA DE ASISTENCIA (ESTRICTA)
# ==========================================
@app.route('/buscar_personal', methods=['POST'])
def buscar_personal():
    cedula = request.form.get('cedula')
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM personal WHERE cedula = ?", (cedula,))
    row = cursor.fetchone()
    
    if row:
        hoy_str = datetime.now().strftime("%Y-%m-%d")
        dia_hoy = datetime.now().weekday()
        
        # Verificar Inteligencia de Botones
        cursor.execute("SELECT hora_salida FROM asistencia WHERE cedula=? AND fecha=?", (cedula, hoy_str))
        asis_hoy = cursor.fetchone()
        ha_marcado_entrada = True if asis_hoy else False
        ha_marcado_salida = True if (asis_hoy and asis_hoy['hora_salida'] and asis_hoy['hora_salida'] not in ["", "Pendiente", "0:00:00", "00:00:00"]) else False
        
        # VALIDACIÓN ESTRICTA: SI NO TIENE HORARIO, BLOQUEA
        cursor.execute("SELECT MIN(hora_entrada) as ent, MAX(hora_salida) as sal FROM horarios WHERE cedula=? AND dia_semana=?", (cedula, dia_hoy))
        h = cursor.fetchone()
        
        if not h or not h['ent']:
            conn.close()
            return jsonify({'success': False, 'error_msg': 'No tienes turno asignado para hoy.'})
            
        target_out = parse_time(h['sal']).strftime("%H:%M")
        conn.close()
        
        return jsonify({'success': True, 'nombre': row['nombre'], 'departamento': row['departamento'], 'cargo': row['cargo'], 'hora_salida_esperada': target_out, 'ha_marcado_entrada': ha_marcado_entrada, 'ha_marcado_salida': ha_marcado_salida})
    
    conn.close()
    return jsonify({'success': False})

@app.route('/', methods=['GET', 'POST'])
def index():
    if 'usuario' in session: return redirect(url_for('admin_panel'))
    
    conn = get_db_connection()
    auto_marcar_salidas(conn) 
    
    mensaje, tipo_mensaje = "", ""
    hoy_dt = datetime.now()
    hoy_str = hoy_dt.strftime("%Y-%m-%d")
    hora_actual = hoy_dt.time()
    hora_str = hora_actual.strftime("%H:%M:%S")
    dia_semana_hoy = hoy_dt.weekday()

    cursor = conn.cursor()
    if request.method == 'POST':
        cedula, nombre, depto = request.form.get('cedula'), request.form.get('nombre'), request.form.get('departamento')
        cargo, observacion, accion, foto_base64 = request.form.get('cargo'), request.form.get('observacion', ''), request.form.get('accion'), request.form.get('foto_base64', '')

        if not nombre or not depto:
            mensaje, tipo_mensaje = "⛔ Error: Usuario no reconocido.", "error"
        else:
            cursor.execute("SELECT MIN(hora_entrada) as ent, MAX(hora_salida) as sal FROM horarios WHERE cedula=? AND dia_semana=?", (cedula, dia_semana_hoy))
            horario_hoy = cursor.fetchone()
            
            if not horario_hoy or not horario_hoy['ent']:
                mensaje, tipo_mensaje = "⛔ ACCESO DENEGADO: No tienes turno o clases asignadas para el día de hoy.", "error"
            else:
                meta_entrada = parse_time(horario_hoy['ent'])
                meta_salida = parse_time(horario_hoy['sal'])

                dt_meta_entrada = datetime.combine(hoy_dt.date(), meta_entrada)
                dt_limite_gracia = dt_meta_entrada + timedelta(minutes=15)
                dt_actual = datetime.combine(hoy_dt.date(), hora_actual)
                
                nombre_foto = ""
                if foto_base64 and ';base64,' in foto_base64:
                    try:
                        img_data = base64.b64decode(foto_base64.split(';base64,')[1])
                        tipo_nom = "ent" if accion == "entrada" else "sal"
                        nombre_foto = f"foto_{tipo_nom}_{cedula}_{hoy_dt.strftime('%H%M%S')}.jpg"
                        with open(os.path.join(FOTOS_DIR, nombre_foto), 'wb') as f: f.write(img_data)
                    except: pass

                if accion == "entrada":
                    cursor.execute("SELECT id FROM asistencia WHERE cedula=? AND fecha=?", (cedula, hoy_str))
                    if cursor.fetchone(): mensaje, tipo_mensaje = "⚠️ Ya marcó entrada hoy.", "warn"
                    else:
                        cursor.execute("SELECT * FROM novedades WHERE cedula=? AND ? BETWEEN fecha_inicio AND fecha_fin", (cedula, hoy_str))
                        novedad = cursor.fetchone()
                        if novedad:
                            mensaje, tipo_mensaje = f"⛔ ACCESO DENEGADO: Usted está de {novedad['tipo'].upper()}.", "error"
                        else:
                            estado = "TARDÍA" if dt_actual > dt_limite_gracia else "Puntual"
                            cursor.execute("INSERT INTO asistencia (cedula, nombre, departamento, cargo, fecha, hora_entrada, estado_entrada, hora_salida, estado_salida, observacion, foto_entrada, foto_salida) VALUES (?, ?, ?, ?, ?, ?, ?, '', 'Pendiente', '', ?, '')", 
                                           (cedula, nombre, depto, cargo, hoy_str, hora_str, estado, nombre_foto))
                            conn.commit()
                            mensaje, tipo_mensaje = f"📸 Entrada registrada: {nombre}", "success"

                elif accion == "salida":
                    cursor.execute("SELECT id, hora_salida FROM asistencia WHERE cedula=? AND fecha=?", (cedula, hoy_str))
                    registro = cursor.fetchone()
                    if not registro: mensaje, tipo_mensaje = "⛔ No tiene entrada hoy.", "error"
                    elif registro['hora_salida'] and registro['hora_salida'] != "0:00:00" and registro['hora_salida'] != "":
                        mensaje, tipo_mensaje = "⚠️ Ya marcó salida.", "warn"
                    else:
                        est_sal = "ANTICIPADA" if hora_actual < meta_salida else "Correcta"
                        cursor.execute("UPDATE asistencia SET hora_salida=?, estado_salida=?, observacion=?, foto_salida=? WHERE cedula=? AND fecha=?", (hora_str, est_sal, observacion, nombre_foto, cedula, hoy_str))
                        conn.commit()
                        mensaje, tipo_mensaje = "👋 Salida registrada.", "success"

    cursor.execute("SELECT * FROM asistencia WHERE fecha=? ORDER BY id DESC LIMIT 6", (hoy_str,))
    registros_hoy = []
    for r in cursor.fetchall():
        rd = dict(r)
        rd['hora_entrada_fmt'] = to_12h(rd['hora_entrada'])
        rd['hora_salida_fmt'] = to_12h(rd['hora_salida'])
        registros_hoy.append(rd)
    conn.close()
    return render_template('index.html', mensaje=mensaje, tipo_mensaje=tipo_mensaje, registros=registros_hoy)

@app.route('/borrar/<int:id>')
def borrar(id):
    if 'usuario' not in session: return redirect(url_for('login'))
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT foto_entrada, foto_salida FROM asistencia WHERE id=?", (id,))
    row = cursor.fetchone()
    if row:
        if row['foto_entrada']:
            try: os.remove(os.path.join(FOTOS_DIR, row['foto_entrada']))
            except: pass
        if row['foto_salida']:
            try: os.remove(os.path.join(FOTOS_DIR, row['foto_salida']))
            except: pass
            
    cursor.execute("DELETE FROM asistencia WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return redirect(url_for('admin_panel'))

@app.route('/admin_panel', methods=['GET', 'POST'])
def admin_panel():
    if 'usuario' not in session: return redirect(url_for('login'))
    conn = get_db_connection()
    auto_marcar_salidas(conn) 
    
    hoy_dt = datetime.now()
    hoy = hoy_dt.strftime("%Y-%m-%d")
    primer_dia_mes = hoy_dt.replace(day=1).strftime("%Y-%m-%d")
    
    desde = request.form.get('desde') if request.method == 'POST' else request.args.get('desde', primer_dia_mes)
    hasta = request.form.get('hasta') if request.method == 'POST' else request.args.get('hasta', hoy)
    
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) as total FROM asistencia WHERE fecha=?", (hoy,))
    total = cursor.fetchone()['total']
    cursor.execute("SELECT COUNT(*) as total FROM asistencia WHERE fecha=? AND estado_entrada='TARDÍA'", (hoy,))
    tarde = cursor.fetchone()['total']
    cursor.execute("SELECT COUNT(*) as total FROM asistencia WHERE fecha=? AND (hora_salida IS NULL OR hora_salida='0:00:00' OR hora_salida='')", (hoy,))
    activos = cursor.fetchone()['total']
    stats = {'total': total, 'tarde': tarde, 'activos': activos}

    cursor.execute("SELECT * FROM asistencia WHERE fecha BETWEEN ? AND ? ORDER BY fecha DESC, hora_entrada DESC", (desde, hasta))
    asistencias = []
    for a in cursor.fetchall():
        ad = dict(a)
        ad['hora_entrada_fmt'] = to_12h(ad['hora_entrada'])
        ad['hora_salida_fmt'] = to_12h(ad['hora_salida'])
        asistencias.append(ad)
    conn.close()
    
    return render_template('admin_panel.html', usuario_actual=session['usuario'], stats=stats, asistencias=asistencias, desde=desde, hasta=hasta)

@app.route('/registrar_novedad', methods=['GET', 'POST'])
def registrar_novedad():
    if 'usuario' not in session: return redirect(url_for('login'))
    mensaje, nombre_encontrado, cedula_buscada = "", "", ""
    conn = get_db_connection()
    cursor = conn.cursor()
    if request.method == 'POST':
        if 'buscar_cedula' in request.form:
            cedula_buscada = request.form['cedula']
            cursor.execute("SELECT nombre, cargo FROM personal WHERE cedula=?", (cedula_buscada,))
            fila = cursor.fetchone()
            if fila: nombre_encontrado = f"{fila['nombre']} ({fila['cargo']})"
            else: mensaje = "<div class='alert error'>❌ Cédula no encontrada.</div>"
        elif 'guardar_novedad' in request.form:
            cedula, tipo, inicio, fin = request.form['cedula_final'], request.form['tipo'], request.form['fecha_inicio'], request.form['fecha_fin']
            desc = request.form.get('descripcion', '') 
            if fin < inicio: mensaje = "<div class='alert error'>⛔ Error: La fecha final es anterior a la inicial.</div>"
            else:
                cursor.execute("SELECT tipo, fecha_inicio, fecha_fin FROM novedades WHERE cedula=? AND fecha_inicio<=? AND fecha_fin>=?", (cedula, fin, inicio))
                c = cursor.fetchone()
                if c: mensaje = f"<div class='alert error'>⛔ Error: Ya tiene <b>{c['tipo']}</b> registrado. No se pueden cruzar.</div>"
                else:
                    cursor.execute("INSERT INTO novedades (cedula, tipo, fecha_inicio, fecha_fin, descripcion) VALUES (?, ?, ?, ?, ?)", (cedula, tipo, inicio, fin, desc))
                    conn.commit()
                    mensaje = "<div class='alert success'>✅ Novedad registrada.</div>"
    
    cursor.execute("SELECT n.*, p.nombre FROM novedades n JOIN personal p ON n.cedula = p.cedula ORDER BY n.id DESC")
    novedades_list = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return render_template('registrar_novedad.html', mensaje=mensaje, nombre_encontrado=nombre_encontrado, cedula_buscada=cedula_buscada, hoy=datetime.now().strftime("%Y-%m-%d"), usuario_actual=session['usuario'], novedades_list=novedades_list)

@app.route('/eliminar_novedad/<int:id>')
def eliminar_novedad(id):
    if 'usuario' not in session: return redirect(url_for('login'))
    conn = get_db_connection()
    conn.execute("DELETE FROM novedades WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return redirect(url_for('registrar_novedad'))

@app.route('/exportar')
def exportar():
    if 'usuario' not in session: return redirect(url_for('login'))
    desde = request.args.get('desde')
    hasta = request.args.get('hasta')
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    if desde and hasta: 
        try:
            d_dt = datetime.strptime(desde, "%Y-%m-%d")
            h_dt = datetime.strptime(hasta, "%Y-%m-%d")
            if d_dt.month == h_dt.month and d_dt.year == h_dt.year:
                mes_nombre = MESES_ES.get(f"{d_dt.month:02d}", str(d_dt.month))
                rango = f"Reporte del Mes de {mes_nombre} del {d_dt.year}"
                nombre_archivo = f"Reporte_IUTEPI_{d_dt.month:02d}-{d_dt.year}.xls"
            else:
                rango = f"Desde: {desde} Hasta: {hasta}"
                nombre_archivo = f"Reporte_IUTEPI_{datetime.now().strftime('%d-%m-%Y')}.xls"
        except:
            rango = f"Desde: {desde} Hasta: {hasta}"
            nombre_archivo = f"Reporte_IUTEPI_Rango.xls"
            
        cursor.execute("SELECT * FROM asistencia WHERE fecha BETWEEN ? AND ? ORDER BY fecha DESC, hora_entrada DESC", (desde, hasta))
    else: 
        rango = "Histórico Completo"
        nombre_archivo = f"Reporte_IUTEPI_Completo.xls"
        cursor.execute("SELECT * FROM asistencia ORDER BY fecha DESC, hora_entrada DESC")
    
    html = f"""<meta charset="UTF-8"><table border="1"><tr style="background-color:#D4001F;color:white;"><th colspan="10">REPORTE OFICIAL DE ASISTENCIA - IUTEPI</th></tr><tr style="background-color:#333;color:white;"><th colspan="10">{rango}</th></tr><tr style="background-color:#eee;"><th>Fecha</th><th>Cédula</th><th>Nombre</th><th>Depto</th><th>Cargo</th><th>H. Ent</th><th>Est. Ent</th><th>H. Sal</th><th>Est. Sal</th><th>Observación</th></tr>"""
    for f in cursor.fetchall():
        html += f"<tr><td>{f['fecha']}</td><td>{f['cedula']}</td><td>{f['nombre']}</td><td>{f['departamento']}</td><td>{f['cargo']}</td><td>{to_12h(f['hora_entrada']) or '--'}</td><td>{f['estado_entrada']}</td><td>{to_12h(f['hora_salida']) or '--'}</td><td>{f['estado_salida']}</td><td>{f['observacion'] or '-'}</td></tr>"
    conn.close()
    html += "</table>"
    
    return Response(html, headers={"Content-Disposition": f"attachment; filename={nombre_archivo}", "Content-Type": "application/vnd.ms-excel"})

@app.route('/exportar_rendimiento')
def exportar_rendimiento():
    if 'usuario' not in session: return redirect(url_for('login'))
    desde = request.args.get('desde')
    hasta = request.args.get('hasta')
    
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM personal ORDER BY departamento ASC, nombre ASC")
    personal = cursor.fetchall()
    
    if desde and hasta: cursor.execute("SELECT cedula, estado_entrada, estado_salida FROM asistencia WHERE fecha BETWEEN ? AND ?", (desde, hasta))
    else: cursor.execute("SELECT cedula, estado_entrada, estado_salida FROM asistencia")
    asistencias = cursor.fetchall()
    conn.close()
    
    stats_personal = {}
    for p in personal:
        stats_personal[p['cedula']] = {'nombre': p['nombre'], 'departamento': p['departamento'], 'cargo': p['cargo'], 'total': 0, 'puntual': 0, 'tarde': 0, 'correcta': 0, 'anticipada': 0, 'no_marco': 0}
        
    for a in asistencias:
        c = a['cedula']
        if c in stats_personal:
            stats_personal[c]['total'] += 1
            if a['estado_entrada'] == 'Puntual': stats_personal[c]['puntual'] += 1
            elif a['estado_entrada'] == 'TARDÍA': stats_personal[c]['tarde'] += 1
            if a['estado_salida'] == 'Correcta': stats_personal[c]['correcta'] += 1
            elif a['estado_salida'] == 'ANTICIPADA': stats_personal[c]['anticipada'] += 1
            elif a['estado_salida'] == 'NO MARCO': stats_personal[c]['no_marco'] += 1
            
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    deptos_data = {}
    for c, data in stats_personal.items():
        d = data['departamento']
        if d not in deptos_data: deptos_data[d] = []
        deptos_data[d].append((c, data))
        
    for depto_nombre, emps in deptos_data.items():
        ws = workbook.add_worksheet(depto_nombre[:31])
        header_format = workbook.add_format({'bold': True, 'bg_color': '#2c3e50', 'font_color': 'white'})
        
        headers = ['Cédula', 'Nombre', 'Cargo', 'Total Asistencias', 'Puntual', 'Tarde', 'Salida Anticipada', 'No Marcó Salida']
        for col_num, data in enumerate(headers):
            ws.write(0, col_num, data, header_format)
            ws.set_column(col_num, col_num, 15)
        ws.set_column(1, 1, 25) 
            
        row = 1
        for cedula, data in emps:
            ws.write(row, 0, cedula)
            ws.write(row, 1, data['nombre'])
            ws.write(row, 2, data['cargo'])
            ws.write(row, 3, data['total'])
            ws.write(row, 4, data['puntual'])
            ws.write(row, 5, data['tarde'])
            ws.write(row, 6, data['anticipada'])
            ws.write(row, 7, data['no_marco'])
            
            if data['total'] > 0:
                chart = workbook.add_chart({'type': 'doughnut'})
                chart.add_series({
                    'name':       f"Rendimiento",
                    'categories': [depto_nombre[:31], 0, 4, 0, 7], 
                    'values':     [depto_nombre[:31], row, 4, row, 7],
                    'points': [{'fill': {'color': '#27ae60'}}, {'fill': {'color': '#f1c40f'}}, {'fill': {'color': '#e67e22'}}, {'fill': {'color': '#c0392b'}}],
                })
                chart.set_title({'name': data['nombre']})
                chart.set_size({'width': 350, 'height': 250})
                ws.insert_chart(row, 9, chart)
                row += 13
            else:
                row += 1

    workbook.close()
    output.seek(0)
    
    try:
        d_dt = datetime.strptime(desde, "%Y-%m-%d")
        mes_nombre = MESES_ES.get(f"{d_dt.month:02d}", str(d_dt.month))
        nombre_archivo = f"Rendimiento_{d_dt.month:02d}-{d_dt.year}.xlsx"
    except:
        nombre_archivo = f"Rendimiento_General.xlsx"

    return Response(output, headers={"Content-Disposition": f"attachment; filename={nombre_archivo}", "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"})

def open_browser():
    webbrowser.open_new('http://127.0.0.1:5000/')

if __name__ == '__main__':
    threading.Timer(1.2, open_browser).start()
    app.run(port=5000)