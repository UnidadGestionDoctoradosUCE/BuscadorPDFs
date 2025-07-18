import time
from datetime import datetime
import os
import fitz  
import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
from PIL import Image, ImageTk
import json
from datetime import datetime
import hashlib
import string
import sys

documentos_drive = []
ruta_por_iid = {} 
nombre_usuario_actual = ""

logo_img_global = None
login_logo_img_global = None

def encontrar_rutas_drive():
    ruta_doctorados = None
    ruta_oficios = None
    for letra in string.ascii_uppercase:
        posible_doctorados = f"{letra}:\\Mi unidad\\Doctorados\\DB_General\\Universidad"
        if not ruta_doctorados and os.path.exists(posible_doctorados):
            ruta_doctorados = posible_doctorados
        posible_oficios = f"{letra}:\\Mi unidad\\Oficios\\DB_GeneralOficios"
        if not ruta_oficios and os.path.exists(posible_oficios):
            ruta_oficios = posible_oficios
    return ruta_doctorados, ruta_oficios

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

ruta_doctorados, ruta_oficios = encontrar_rutas_drive()

items_clave = {
    "1. Hoja de Vida": ["h.vida"],
    "2. Contrato de Beca": ["c.beca"],
    "3. Aval de Facultad": ["avalfacultad"],
    "4. Acción Personal": ["acciónpersonal"],
    "5. Cédula y Papeleta de Votación": ["cédulaypapeleta"],
    "6. Pasaporte": ["pasaporte"],
    "7. Declaración Juramentada": ["declaraciónjuramentada"],
    "8. Solicitud de Pedido al Rectorado": ["solicitudpedidorectorado"],
    "9. Carta de Admisión de la Universidad": ["cartaadmisiónuniversidad"],
    "10. Certificado de Cuenta Bancaria": ["certificadobancario"],
    "11. Matrícula": ["matrícula"],
    "12. Certificado de Aprobación Tesis": ["certificadoaprobación"],
    "13. Reporte de Notas Semestrales": ["reportenotas"],
    "14. Línea de Investigación y Plan de Tesis": ["lineamientoinvestigacióntesis"],
    "15. Informe de Avance de Tesis Emitido por el Tutor": ["informeavance"],
    "16. Correspondencia": ["correspondencia"],
    "17. Registro de Títulos SENESCYT": ["registrosenescyt"],
    "18. Adenda":["adenda"],
    "19. Oficios":["oficio"]
}

def cargar_documentos(ruta):
    if not ruta:
        return []
    documentos = []
    for carpeta, _, archivos in os.walk(ruta):
        for archivo in archivos:
            if archivo.lower().endswith('.pdf'):
                ruta_completa = os.path.join(carpeta, archivo)
                nombre_archivo = archivo.replace('.pdf', '').replace('_', ' ').lower()

                partes_ruta = os.path.relpath(ruta_completa, ruta).split(os.sep)
                universidad = partes_ruta[0] if len(partes_ruta) > 0 else ''
                programa = partes_ruta[1] if len(partes_ruta) > 1 else ''
                estudiante = partes_ruta[2] if len(partes_ruta) > 2 else ''

                documentos.append({
                    'universidad': universidad,
                    'programa': programa,
                    'estudiante': estudiante,
                    'nombre': nombre_archivo,
                    'ruta': ruta_completa
                })
    return documentos

def obtener_todas_universidades(ruta):
    if not ruta or not os.path.exists(ruta):
        return []
    return sorted([nombre for nombre in os.listdir(ruta) if os.path.isdir(os.path.join(ruta, nombre))])

def actualizar_universidades():
    global universidades_lista
    universidades_lista = obtener_todas_universidades(ruta_doctorados)
    combo_universidad['values'] = ['(Todos)'] + universidades_lista
    combo_universidad.set('(Todos)')
    actualizar_programas()

def actualizar_programas():
    u = combo_universidad.get()
    if u == '(Todos)':
        programas = sorted(set(doc['programa'] for doc in documentos_drive))
    else:
        programas = sorted(set(doc['programa'] for doc in documentos_drive if doc['universidad'] == u))
    combo_programa['values'] = ['(Todos)'] + programas
    combo_programa.set('(Todos)')
    actualizar_estudiantes()

def actualizar_estudiantes():
    u = combo_universidad.get()
    p = combo_programa.get()
    estudiantes = set()
    for doc in documentos_drive:
        if (u == '(Todos)' or doc['universidad'] == u) and \
           (p == '(Todos)' or doc['programa'] == p):
            estudiantes.add(doc['estudiante'])
    combo_estudiante['values'] = ['(Todos)'] + sorted(estudiantes)
    combo_estudiante.set('(Todos)')

def buscar():
    resultados.delete(*resultados.get_children())
    ruta_por_iid.clear()
    filtro_u = combo_universidad.get()
    filtro_p = combo_programa.get()
    filtro_e = combo_estudiante.get()
    filtro_nombre = entrada_nombre.get().strip().lower()
    filtro_item = combo_item_clave.get()
    encontrados = []
    for d in documentos_drive:
        if (filtro_u == '(Todos)' or d['universidad'] == filtro_u) and \
           (filtro_p == '(Todos)' or d['programa'] == filtro_p) and \
           (filtro_e == '(Todos)' or d['estudiante'] == filtro_e) and \
           (filtro_nombre in d['nombre']):
            if filtro_item == '(Todos)':
                encontrados.append(d)
            else:
                alias_list = items_clave.get(filtro_item, [])
                if any(alias.lower() in d['nombre'] for alias in alias_list):
                    encontrados.append(d)
    if not encontrados:
        messagebox.showinfo("No encontrado", "No se encontró ningún documento con los filtros aplicados.")
        etiqueta_resumen.config(text="Mostrando 0 documentos.")
        limpiar_detalles()
        return
    for i, doc in enumerate(encontrados):
        tag = 'evenrow' if i % 2 == 0 else 'oddrow'
        iid = resultados.insert("", "end", values=(
            doc['universidad'].title(),
            doc['programa'].title(),
            doc['estudiante'].title(),
            doc['nombre'].title()
        ), tags=(tag,))
        ruta_por_iid[iid] = doc['ruta']
    etiqueta_resumen.config(text=f"Mostrando {len(encontrados)} documento(s).")
    limpiar_detalles()

def abrir_pdf():
    item = resultados.focus()
    if item and item in ruta_por_iid:
        ruta = ruta_por_iid[item]
        try:
            os.startfile(ruta)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")
    else:
        messagebox.showwarning("Selecciona algo", "Debes seleccionar un resultado.")

def abrir_carpeta():
    item = resultados.focus()
    if item and item in ruta_por_iid:
        ruta = ruta_por_iid[item]
        carpeta = os.path.dirname(ruta)
        try:
            if os.name == 'nt':  
                os.startfile(carpeta)
            elif os.name == 'posix':  
                subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', carpeta])
            else:
                messagebox.showwarning("No soportado", "No se puede abrir la carpeta en este sistema.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta:\n{e}")
    else:
        messagebox.showwarning("Selecciona algo", "Debes seleccionar un resultado.")

def mostrar_detalles(event):
    item = resultados.focus()
    if item and item in ruta_por_iid:
        ruta = ruta_por_iid[item]
        try:
            doc = fitz.open(ruta)
            num_paginas = doc.page_count
            metadata = doc.metadata
            doc.close()
            info = (
                f"Ruta: {ruta}\n"
                f"Páginas: {num_paginas}\n"
                f"Título: {metadata.get('title', 'N/A')}\n"
                f"Herramienta Usada: {metadata.get('author', 'N/A')}\n"
            )
            texto_detalles.config(state='normal')
            texto_detalles.delete('1.0', tk.END)
            texto_detalles.insert(tk.END, info)
            texto_detalles.config(state='disabled')
        except Exception as e:
            texto_detalles.config(state='normal')
            texto_detalles.delete('1.0', tk.END)
            texto_detalles.insert(tk.END, f"No se pudo cargar detalles:\n{e}")
            texto_detalles.config(state='disabled')
    else:
        limpiar_detalles()

def limpiar_detalles():
    texto_detalles.config(state='normal')
    texto_detalles.delete('1.0', tk.END)
    texto_detalles.config(state='disabled')

def abrir_pdf_event(event):
    abrir_pdf()

def sincronizar_drive(ventana):
    global documentos_drive
    if not ruta_doctorados:
        documentos_drive = []
        actualizar_universidades()
        buscar()
        messagebox.showinfo("Sincronización", "No se encontró la ruta del doctorados. No hay documentos cargados.")
        return
    documentos_drive = cargar_documentos(ruta_doctorados)
    actualizar_universidades()
    buscar()
    messagebox.showinfo("Sincronización", "¡Sincronización completada! Los documentos han sido actualizados.")

def crear_encabezado(ventana):
    frame = tk.Frame(ventana, bg="#e8f0f7")
    frame.pack(pady=10, fill="x")

    frame_izquierdo = tk.Frame(frame, bg="#e8f0f7")
    frame_izquierdo.pack(side="left", padx=10)

    try:
        ruta_imagen = resource_path("imagenes/logouce.png")
        imagen_logo = Image.open(ruta_imagen)
        imagen_logo = imagen_logo.resize((120, 120), Image.LANCZOS)
        ventana.logo_img = ImageTk.PhotoImage(imagen_logo)
        etiqueta_logo = tk.Label(frame_izquierdo, image=ventana.logo_img, bg="#e8f0f7")
        etiqueta_logo.pack(side="left", padx=(0,10))
    except Exception as e:
        print("No se pudo cargar el logo:", e)

    btn_sincronizar = ttk.Button(frame_izquierdo, text="Sincronizar", command=lambda: sincronizar_drive(ventana), style='Sincronizar.TButton')
    btn_sincronizar.pack(side="left", padx=(0,10), pady=10)

    frame_centro = tk.Frame(frame, bg="#e8f0f7")
    frame_centro.pack(side="left", expand=True)
    tk.Label(frame_centro, text="Buscador de Doctorados", font=("Segoe UI", 28, "bold"), bg="#e8f0f7").pack(expand=True)

    frame_derecho = tk.Frame(frame, bg="#e8f0f7")
    frame_derecho.pack(side="right", padx=20)
    btn_cerrar_sesion = ttk.Button(frame_derecho, text="Cerrar Sesión", command=lambda: cerrar_sesion(ventana), style='CerrarSesion.TButton')
    btn_cerrar_sesion.pack(pady=10)

    return frame

def crear_resumen(ventana):
    frame = tk.Frame(ventana, bg="#e8f0f7")
    frame.pack(fill="x", padx=10, pady=(0, 5))
    etiqueta = tk.Label(frame, text="Mostrando 0 documentos.", font=("Segoe UI", 12, 'bold'), bg="#e8f0f7")
    etiqueta.pack(side="right", anchor="e", padx=10)
    return etiqueta

def crear_filtros(ventana):
    frame = tk.Frame(ventana, bg="#e8f0f7")
    frame.pack(pady=10)
    
    def crear_etiqueta(texto, fila, columna):
        tk.Label(frame, text=texto, font=('Segoe UI', 11), bg="#e8f0f7").grid(row=fila, column=columna, padx=5, pady=3)

    crear_etiqueta("Universidad:", 0, 0)
    global combo_universidad
    combo_universidad = ttk.Combobox(frame, width=25, state="readonly")
    combo_universidad.grid(row=0, column=1)

    crear_etiqueta("Programa:", 0, 2)
    global combo_programa
    combo_programa = ttk.Combobox(frame, width=25, state="readonly")
    combo_programa.grid(row=0, column=3)

    crear_etiqueta("Estudiante:", 0, 4)
    global combo_estudiante
    combo_estudiante = ttk.Combobox(frame, width=25, state="readonly")
    combo_estudiante.grid(row=0, column=5)

    crear_etiqueta("Filtro ítem clave:", 0, 6)
    global combo_item_clave
    combo_item_clave = ttk.Combobox(frame, width=25, state="readonly")
    combo_item_clave.grid(row=0, column=7)

    return combo_universidad, combo_programa, combo_estudiante, combo_item_clave

def crear_entrada_nombre(ventana):
    global entrada_nombre
    tk.Label(ventana, text="Filtrar por nombre del docente (opcional):", font=("Segoe UI", 12), bg="#e8f0f7").pack(pady=(10, 0))
    entrada_nombre = tk.Entry(ventana, width=70, font=("Segoe UI", 12), relief="solid", bd=1)
    entrada_nombre.pack(pady=8)
    entrada_nombre.insert(0, "")

def crear_botones(ventana):
    frame = tk.Frame(ventana, bg="#e8f0f7")
    frame.pack(pady=10)

    btn_buscar = ttk.Button(frame, text="Buscar", command=buscar)
    btn_buscar.grid(row=0, column=0, padx=10)

    btn_abrir = ttk.Button(frame, text="Abrir PDF", command=abrir_pdf)
    btn_abrir.grid(row=0, column=1, padx=10)

    btn_abrir_carpeta = ttk.Button(frame, text="Abrir carpeta", command=abrir_carpeta)
    btn_abrir_carpeta.grid(row=0, column=2, padx=10)

def crear_resultados(ventana):
    global resultados
    columnas = ("Universidad", "Programa", "Estudiante", "Nombre")
    frame_resultados = tk.Frame(ventana)
    frame_resultados.pack(padx=20, pady=10, fill="both", expand=True)
    scrollbar = tk.Scrollbar(frame_resultados, orient="vertical")
    resultados = ttk.Treeview(frame_resultados, columns=columnas, show="headings", height=25, yscrollcommand=scrollbar.set)
    scrollbar.config(command=resultados.yview)
    scrollbar.pack(side="right", fill="y")
    resultados.pack(side="left", fill="both", expand=True)
    for col in columnas:
        resultados.heading(col, text=col)
        resultados.column(col, anchor="center", width=200)
    resultados.tag_configure('oddrow', background="white")
    resultados.tag_configure('evenrow', background="#d3d3d3")
    resultados.bind("<<TreeviewSelect>>", mostrar_detalles)
    resultados.bind("<Double-1>", abrir_pdf_event)

def crear_texto_detalles(ventana):
    global texto_detalles
    tk.Label(ventana, text="Detalles del archivo seleccionado:", font=("Segoe UI", 12), bg="#e8f0f7").pack()
    texto_detalles = tk.Text(ventana, height=4, width=100, font=("Segoe UI", 11), relief="solid", bd=1)
    texto_detalles.pack(padx=40, pady=(0, 25), fill='x')
    texto_detalles.config(state="disabled")

def verificar_credenciales(usuario, contrasena):
    ruta_json = resource_path("data/usuarios.json")
    try:
        with open(ruta_json, 'r', encoding='utf-8') as f:
            lista_usuarios = json.load(f)
            usuario_hash = hashlib.sha256(usuario.encode('utf-8')).hexdigest()
            contrasena_hash = hashlib.sha256(contrasena.encode('utf-8')).hexdigest()

            usuario_encontrado = None
            for u in lista_usuarios:
                if u['usuario'] == usuario_hash:
                    usuario_encontrado = u
                    break

            if usuario_encontrado:
                if usuario_encontrado['contrasena'] == contrasena_hash:
                    return True, usuario_encontrado.get('nombre', usuario)
                else:
                    messagebox.showwarning("Credenciales incorrectas", "Vuelve a intentarlo: usuario o contraseña incorrectos.")
                    return False, None
            else:
                messagebox.showerror("Usuario no autorizado", "Este usuario no está registrado en el sistema.")
                return False, None

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo verificar credenciales:\n{e}")
        return False, None

def limpiar_ventana(ventana):
    for widget in ventana.winfo_children():
        widget.destroy()

def mostrar_login(ventana):
    limpiar_ventana(ventana)
    ventana.title("Unidad Administrativa de Gestión de Doctorados")
    ventana.state('zoomed')
    ventana.configure(bg="#f2f2f2")

    # Estilos
    style = ttk.Style()
    style.theme_use('clam')
    style.configure('TLabel', font=('Segoe UI', 12), background="#f2f2f2")
    style.configure('TButton', font=('Segoe UI', 12), padding=6)
    style.configure('TEntry', padding=6)

    fondo = tk.Frame(ventana, bg="#f2f2f2")
    fondo.pack(fill='both', expand=True)

    frame = tk.Frame(fondo, padx=30, pady=30, bg="#f2f2f2")
    frame.place(relx=0.5, rely=0.5, anchor='center')

    try:
        ruta_imagen = resource_path("imagenes/logouce.png")
        imagen_logo = Image.open(ruta_imagen)
        imagen_logo = imagen_logo.resize((130, 130), Image.LANCZOS)
        ventana.logo_img = ImageTk.PhotoImage(imagen_logo)
        etiqueta_logo = tk.Label(frame, image=ventana.logo_img, bg="#f2f2f2")
        etiqueta_logo.pack(pady=(0, 10))
    except Exception as e:
        print("No se pudo cargar el logo:", e)

    tk.Label(frame, text="Unidad Administrativa de Gestión de Doctorados", font=("Segoe UI", 14, "bold"), bg="#f2f2f2", fg="#2e4a62").pack(pady=(0, 20))

    fila_usuario = tk.Frame(frame, bg="#f2f2f2")
    fila_usuario.pack(pady=10)

    tk.Label(fila_usuario, text="Usuario:", font=("Segoe UI", 12), bg="#f2f2f2", width=12, anchor='e').pack(side='left')
    global entrada_usuario
    entrada_usuario = tk.Entry(fila_usuario, font=("Segoe UI", 12), width=30, relief="solid", bd=1)
    entrada_usuario.pack(side='left')

    fila_contra = tk.Frame(frame, bg="#f2f2f2")
    fila_contra.pack(pady=10)

    tk.Label(fila_contra, text="Contraseña:", font=("Segoe UI", 12), bg="#f2f2f2", width=12, anchor='e').pack(side='left')

    subframe_contra = tk.Frame(fila_contra, bg="#f2f2f2")
    subframe_contra.pack(side='left')

    global entrada_contra
    entrada_contra = tk.Entry(subframe_contra, font=("Segoe UI", 12), show="*", width=27, relief="solid", bd=1)
    entrada_contra.pack(side="left")

    mostrar_contra = tk.BooleanVar(value=False)

    def toggle_contra():
        if mostrar_contra.get():
            entrada_contra.config(show="*")
            boton_ojo.config(text="👁")
            mostrar_contra.set(False)
        else:
            entrada_contra.config(show="")
            boton_ojo.config(text="🔒")
            mostrar_contra.set(True)

    boton_ojo = tk.Button(subframe_contra, text="👁", command=toggle_contra, relief="flat", bg="#f2f2f2", bd=0, font=("Segoe UI", 12))
    boton_ojo.pack(side="left", padx=(5, 0))

    btn_ingresar = ttk.Button(frame, text="Ingresar", command=lambda: intentar_login(ventana))
    btn_ingresar.pack(pady=25)

    ventana.bind('<Return>', lambda event: intentar_login(ventana))

def intentar_login(ventana):
    global nombre_usuario_actual
    rutas_info = []
    if ruta_doctorados:
        rutas_info.append(f"Ruta Doctorados encontrada:\n{ruta_doctorados}")
    else:
        rutas_info.append("No se encontró la ruta de doctorados.")
    if ruta_oficios:
        rutas_info.append(f"Ruta Oficios encontrada:\n{ruta_oficios}")
    else:
        rutas_info.append("No se encontró la ruta de oficios.")
    messagebox.showinfo("Información de rutas", "\n\n".join(rutas_info))
    usuario = entrada_usuario.get().strip()
    contrasena = entrada_contra.get().strip()
    if not usuario or not contrasena:
        messagebox.showwarning("Campos requeridos", "Debes ingresar usuario y contraseña.")
        return
    valido, nombre_real = verificar_credenciales(usuario, contrasena)
    if valido:
        nombre_usuario_actual = nombre_real  
        mostrar_carga_y_abrir_main(ventana, nombre_real)
    else:
        messagebox.showerror("Error de inicio de sesión", "Usuario o contraseña incorrectos.")

def mostrar_seleccion(ventana, nombre_usuario):
    limpiar_ventana(ventana)
    ventana.title("Seleccione una opción")
    ventana.state('zoomed')
    ventana.configure(bg="#e8f0f7")

    frame = tk.Frame(ventana, bg="#e8f0f7")
    frame.place(relx=0.5, rely=0.5, anchor='center')

    tk.Label(frame, text=f"Bienvenido, {nombre_usuario}", font=("Segoe UI", 18, "bold"), bg="#e8f0f7").pack(pady=(0, 20))
    tk.Label(frame, text="¿Qué desea consultar?", font=("Segoe UI", 16), bg="#e8f0f7").pack(pady=(0, 30))

    btn_docentes = ttk.Button(frame, text="Docentes", width=20, command=lambda: main_docentes(ventana))
    btn_docentes.pack(pady=10)
    btn_oficios = ttk.Button(frame, text="Oficios", width=20, command=lambda: main_oficios(ventana))
    btn_oficios.pack(pady=10)

def mostrar_carga_y_abrir_main(ventana, nombre_usuario):
    limpiar_ventana(ventana)
    ventana.title("Cargando Sistema...")
    ventana.state('zoomed')
    ventana.configure(bg="#f2f2f7")

    style = ttk.Style(ventana)
    style.theme_use('clam')
    style.configure('TLabel', font=('Segoe UI', 16), background="#f2f2f2", foreground="#2e4a62")
    style.configure('Titulo.TLabel', font=('Segoe UI', 36, 'bold'), background="#f2f2f2", foreground="#2e4a62")
    style.configure('Subtitulo.TLabel', font=('Segoe UI', 20), background="#f2f2f2", foreground="#2e4a62")
    style.configure('Usuario.TLabel', font=('Segoe UI', 14), background="#f2f2f2", foreground="#2e4a62")
    style.configure('TProgressbar', thickness=20)

    frame = tk.Frame(ventana, bg="#f2f2f2")
    frame.place(relx=0.5, rely=0.5, anchor='center')

    ttk.Label(frame, text="CARGANDO SISTEMA", style='Titulo.TLabel').pack(pady=(20, 10))
    ttk.Label(frame, text="Bienvenido al Buscador de Expedientes\nUnidad Administrativa de Gestión de Doctorados", style='Subtitulo.TLabel', justify='center').pack(pady=(0, 20))

    hora_actual = datetime.now().strftime("%d/%m/%Y - %H:%M:%S")
    ttk.Label(frame, text=f"Usuario: {nombre_usuario}", style='Usuario.TLabel').pack(pady=(0, 5))
    ttk.Label(frame, text=f"Hora de ingreso: {hora_actual}", style='Usuario.TLabel').pack(pady=(0, 15))

    barra = ttk.Progressbar(frame, mode='indeterminate', length=600)
    barra.pack(pady=20)
    barra.start(10)

    def finalizar_carga():
        barra.stop()
        frame.destroy()
        mostrar_seleccion(ventana, nombre_usuario)

    ventana.after(3000, finalizar_carga)

def main_docentes(ventana):
    limpiar_ventana(ventana)
    nav_frame = tk.Frame(ventana, bg="#e8f0f7")
    nav_frame.pack(fill='x', pady=(10, 0))
    btn_volver_seleccion = ttk.Button(nav_frame, text="Volver a selección", command=lambda: mostrar_seleccion(ventana, nombre_usuario_actual))
    btn_volver_seleccion.pack(side='left', padx=10)
    main(ventana, skip_clear=True)

def obtener_estructura_oficios(ruta):
    """
    Devuelve la estructura de carpetas de oficios:
    {
        '2016': {
            'Octubre-Diciembre': ['Recibidos', 'Enviados'],
            ...
        },
        ...
    }
    """
    estructura = {}
    if not ruta or not os.path.exists(ruta):
        return estructura
    for anio in os.listdir(ruta):
        ruta_anio = os.path.join(ruta, anio)
        if os.path.isdir(ruta_anio):
            estructura[anio] = {}
            for periodo in os.listdir(ruta_anio):
                ruta_periodo = os.path.join(ruta_anio, periodo)
                if os.path.isdir(ruta_periodo):
                    tipos = []
                    for tipo in os.listdir(ruta_periodo):
                        ruta_tipo = os.path.join(ruta_periodo, tipo)
                        if os.path.isdir(ruta_tipo):
                            tipos.append(tipo)
                    estructura[anio][periodo] = tipos
    return estructura

estructura_oficios = None

def main_oficios(ventana):
    global estructura_oficios
    limpiar_ventana(ventana)
    nav_frame = tk.Frame(ventana, bg="#e8f0f7")
    nav_frame.pack(fill='x', pady=(10, 0))
    btn_volver_seleccion = ttk.Button(nav_frame, text="Volver a selección", command=lambda: mostrar_seleccion(ventana, nombre_usuario_actual))
    btn_volver_seleccion.pack(side='left', padx=10)
    estructura_oficios = obtener_estructura_oficios(ruta_oficios)
    global documentos_oficios
    documentos_oficios = cargar_documentos_oficios(ruta_oficios)
    ventana.title("Buscador de Oficios")
    ventana.state('zoomed')
    ventana.configure(bg="#e8f0f7")
    configurar_estilos()
    crear_encabezado(ventana)
    global etiqueta_resumen_oficios
    etiqueta_resumen_oficios = crear_resumen(ventana)
    crear_filtros_oficios(ventana)
    crear_entrada_oficio(ventana)
    crear_botones_oficios(ventana)
    crear_texto_detalles_oficio(ventana)
    crear_resultados_oficios(ventana)
    actualizar_anios_oficios()
    combo_anio_oficio.bind("<<ComboboxSelected>>", lambda e: actualizar_tipo_oficios())
    combo_tipo_oficio.bind("<<ComboboxSelected>>", lambda e: None)
    etiqueta_resumen_oficios.config(text="Mostrando 0 oficios.")

def cargar_documentos_oficios(ruta):
    if not ruta:
        return []
    documentos = []
    for carpeta, _, archivos in os.walk(ruta):
        for archivo in archivos:
            if archivo.lower().endswith('.pdf'):
                ruta_completa = os.path.join(carpeta, archivo)
                partes_ruta = os.path.relpath(ruta_completa, ruta).split(os.sep)
                if len(partes_ruta) >= 4:
                    anio = partes_ruta[0]
                    periodo = partes_ruta[1]
                    tipo = partes_ruta[2]  
                else:
                    anio = periodo = tipo = ''
                nombre_archivo = archivo.replace('.pdf', '').replace('_', ' ').lower()
                documentos.append({
                    'anio': anio,
                    'periodo': periodo,
                    'tipo': tipo,
                    'nombre': nombre_archivo,
                    'ruta': ruta_completa
                })
    return documentos

def crear_filtros_oficios(ventana):
    frame = tk.Frame(ventana, bg="#e8f0f7")
    frame.pack(pady=10)
    def crear_etiqueta(texto, fila, columna):
        tk.Label(frame, text=texto, font=('Segoe UI', 11), bg="#e8f0f7").grid(row=fila, column=columna, padx=5, pady=3)
    crear_etiqueta("Año:", 0, 0)
    global combo_anio_oficio
    combo_anio_oficio = ttk.Combobox(frame, width=15, state="readonly")
    combo_anio_oficio.grid(row=0, column=1)
    crear_etiqueta("Tipo:", 0, 2)
    global combo_tipo_oficio
    combo_tipo_oficio = ttk.Combobox(frame, width=15, state="readonly")
    combo_tipo_oficio.grid(row=0, column=3)
    return combo_anio_oficio, combo_tipo_oficio

def crear_entrada_oficio(ventana):
    global entrada_nombre_oficio
    tk.Label(ventana, text="Buscar por nombre o número de oficio:", font=("Segoe UI", 12), bg="#e8f0f7").pack(pady=(10, 0))
    entrada_nombre_oficio = tk.Entry(ventana, width=30, font=("Segoe UI", 12), relief="solid", bd=1)
    entrada_nombre_oficio.pack(pady=4)

def crear_botones_oficios(ventana):
    frame = tk.Frame(ventana, bg="#e8f0f7")
    frame.pack(pady=10)
    btn_buscar = ttk.Button(frame, text="Buscar", command=buscar_oficios)
    btn_buscar.grid(row=0, column=0, padx=10)
    btn_abrir = ttk.Button(frame, text="Abrir PDF", command=abrir_pdf_oficio)
    btn_abrir.grid(row=0, column=1, padx=10)
    btn_carpeta = ttk.Button(frame, text="Abrir carpeta", command=abrir_carpeta_oficio)
    btn_carpeta.grid(row=0, column=2, padx=10)

def crear_resultados_oficios(ventana):
    global resultados_oficios, ruta_por_iid_oficios
    columnas = ("Año", "Tipo", "Nombre")
    resultados_oficios = ttk.Treeview(ventana, columns=columnas, show="headings", height=25)
    resultados_oficios.pack(padx=20, pady=10, fill="both", expand=True)
    for col in columnas:
        resultados_oficios.heading(col, text=col)
        resultados_oficios.column(col, anchor="center", width=200)
    resultados_oficios.tag_configure('oddrow', background="white")
    resultados_oficios.tag_configure('evenrow', background="#d3d3d3")
    resultados_oficios.bind("<<TreeviewSelect>>", mostrar_detalles_oficio)
    resultados_oficios.bind("<Double-1>", lambda e: abrir_pdf_oficio())
    ruta_por_iid_oficios = {}

def actualizar_anios_oficios():
    global estructura_oficios
    if estructura_oficios:
        anios = sorted(estructura_oficios.keys())
    else:
        anios = []
    combo_anio_oficio['values'] = ['(Todos)'] + anios
    combo_anio_oficio.set('(Todos)')
    actualizar_tipo_oficios()

def actualizar_tipo_oficios():
    global estructura_oficios
    anio = combo_anio_oficio.get()
    tipos = set()
    if estructura_oficios:
        if anio != '(Todos)':
            for tipos_periodo in estructura_oficios.get(anio, {}).values():
                tipos.update(tipos_periodo)
        else:
            for anio_val in estructura_oficios.values():
                for tipos_periodo in anio_val.values():
                    tipos.update(tipos_periodo)
    tipos = sorted(tipos)
    combo_tipo_oficio['values'] = ['(Todos)'] + tipos
    combo_tipo_oficio.set('(Todos)')

def buscar_oficios():
    resultados_oficios.delete(*resultados_oficios.get_children())
    ruta_por_iid_oficios.clear()
    filtro_anio = combo_anio_oficio.get()
    filtro_tipo = combo_tipo_oficio.get()
    filtro_nombre = entrada_nombre_oficio.get().strip().lower()
    docs = documentos_oficios
    if filtro_anio != '(Todos)':
        docs = [doc for doc in docs if doc['anio'] == filtro_anio]
    if filtro_tipo != '(Todos)':
        docs = [doc for doc in docs if doc['tipo'] == filtro_tipo]
    if filtro_nombre:
        docs = [doc for doc in docs if filtro_nombre in doc['nombre']]
    if not docs:
        messagebox.showinfo("No encontrado", "No se encontró ningún oficio con los filtros aplicados.")
        etiqueta_resumen_oficios.config(text="Mostrando 0 oficios.")
        limpiar_detalles_oficio()
        return
    for i, doc in enumerate(docs):
        tag = 'evenrow' if i % 2 == 0 else 'oddrow'
        iid = resultados_oficios.insert("", "end", values=(
            doc['anio'],
            doc['tipo'],
            doc['nombre'].title()
        ), tags=(tag,))
        ruta_por_iid_oficios[iid] = doc['ruta']
    etiqueta_resumen_oficios.config(text=f"Mostrando {len(docs)} oficio(s).")
    limpiar_detalles_oficio()

def mostrar_detalles_oficio(event):
    item = resultados_oficios.focus()
    if item and item in ruta_por_iid_oficios:
        doc = next((d for d in documentos_oficios if d['ruta'] == ruta_por_iid_oficios[item]), None)
        if doc:
            info = (
                f"Año: {doc['anio']}\n"
                f"Tipo: {doc['tipo']}\n"
                f"Nombre: {doc['nombre'].title()}\n"
                f"Ruta: {doc['ruta']}\n"
            )
            texto_detalles_oficio.config(state='normal')
            texto_detalles_oficio.delete('1.0', tk.END)
            texto_detalles_oficio.insert(tk.END, info)
            texto_detalles_oficio.config(state='disabled')
    else:
        limpiar_detalles_oficio()

def limpiar_detalles_oficio():
    texto_detalles_oficio.config(state='normal')
    texto_detalles_oficio.delete('1.0', tk.END)
    texto_detalles_oficio.config(state='disabled')

def abrir_pdf_oficio():
    item = resultados_oficios.focus()
    if item and item in ruta_por_iid_oficios:
        ruta = ruta_por_iid_oficios[item]
        try:
            os.startfile(ruta)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")
    else:
        messagebox.showwarning("Selecciona algo", "Debes seleccionar un resultado.")

def abrir_carpeta_oficio():
    item = resultados_oficios.focus()
    if item and item in ruta_por_iid_oficios:
        ruta = ruta_por_iid_oficios[item]
        carpeta = os.path.dirname(ruta)
        try:
            if os.name == 'nt': 
                os.startfile(carpeta)
            elif os.name == 'posix':
                subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', carpeta])
            else:
                messagebox.showwarning("No soportado", "No se puede abrir la carpeta en este sistema.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta:\n{e}")
    else:
        messagebox.showwarning("Selecciona algo", "Debes seleccionar un resultado.")

def cerrar_sesion(ventana):
    global nombre_usuario_actual
    if messagebox.askyesno("Cerrar Sesión", "¿Estás seguro de que quieres cerrar sesión?"):
        nombre_usuario_actual = "" 
        mostrar_login(ventana)

def main(ventana, skip_clear=False):
    if not skip_clear:
        limpiar_ventana(ventana)
    global documentos_drive
    documentos_drive = cargar_documentos(ruta_doctorados)
    ventana.title("Buscador de Doctorados")
    ventana.state('zoomed')
    ventana.configure(bg="#e8f0f7")
    configurar_estilos()
    crear_encabezado(ventana)
    global etiqueta_resumen
    etiqueta_resumen = crear_resumen(ventana)
    crear_filtros(ventana)
    crear_entrada_nombre(ventana)
    crear_botones(ventana)
    crear_texto_detalles(ventana)
    crear_resultados(ventana)
    actualizar_universidades()
    combo_programa.bind("<<ComboboxSelected>>", lambda e: actualizar_estudiantes())
    combo_estudiante.bind("<<ComboboxSelected>>", lambda e: None)
    combo_item_clave['values'] = ['(Todos)'] + list(items_clave.keys())
    combo_item_clave.set('(Todos)')
    if not ruta_doctorados:
        etiqueta_resumen.config(text="No se encontró la ruta de doctorados. No hay documentos cargados.")
    ventana.mainloop()

def configurar_estilos():
    style = ttk.Style()
    style.theme_use('clam')
    style.configure("Treeview",
                    background="#f9f9f9",
                    foreground="black",
                    rowheight=28,
                    fieldbackground="#f9f9f9",
                    font=('Segoe UI', 11))
    style.map('Treeview', background=[('selected', '#347083')], foreground=[('selected', 'white')])
    style.configure('TButton', font=('Segoe UI', 12), padding=6)
    
    style.configure('TCombobox', 
                    font=('Segoe UI', 11),
                    fieldbackground='white',
                    background='white',
                    selectbackground='#347083',
                    selectforeground='white')
    style.map('TCombobox',
              fieldbackground=[('readonly', 'white')],
              selectbackground=[('readonly', '#347083')],
              selectforeground=[('readonly', 'white')])
    
    style.configure('CerrarSesion.TButton', 
                    font=('Segoe UI', 11, 'bold'), 
                    padding=8,
                    background='#dc3545',
                    foreground='white')
    style.map('CerrarSesion.TButton',
              background=[('active', '#c82333'), ('pressed', '#bd2130')],
              foreground=[('active', 'white'), ('pressed', 'white')])
    style.configure('Sincronizar.TButton', 
                    font=('Segoe UI', 11, 'bold'), 
                    padding=8,
                    background='#007bff',
                    foreground='white')
    style.map('Sincronizar.TButton',
              background=[('active', '#0056b3'), ('pressed', '#004085')],
              foreground=[('active', 'white'), ('pressed', 'white')])

def crear_texto_detalles_oficio(ventana):
    global texto_detalles_oficio
    tk.Label(ventana, text="Detalles del oficio seleccionado:", font=("Segoe UI", 12), bg="#e8f0f7").pack()
    texto_detalles_oficio = tk.Text(ventana, height=4, width=100, font=("Segoe UI", 11), relief="solid", bd=1)
    texto_detalles_oficio.pack(padx=40, pady=(0, 25), fill='x')
    texto_detalles_oficio.config(state="disabled")

universidades_lista = []
def abrir_selector_universidad():
    global universidades_lista
    top = tk.Toplevel()
    top.title("Seleccionar Universidad")
    top.geometry("350x350")
    top.transient()
    top.grab_set()
    tk.Label(top, text="Selecciona una universidad:", font=("Segoe UI", 12, "bold")).pack(pady=8)
    frame_list = tk.Frame(top)
    frame_list.pack(fill="both", expand=True, padx=10, pady=10)
    scrollbar = tk.Scrollbar(frame_list)
    scrollbar.pack(side="right", fill="y")
    listbox = tk.Listbox(frame_list, font=("Segoe UI", 11), yscrollcommand=scrollbar.set, selectmode="single")
    for u in ["(Todos)"] + universidades_lista:
        listbox.insert("end", u)
    listbox.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=listbox.yview)
    # Selecciona el actual
    try:
        idx = ["(Todos)"] + universidades_lista.index(universidad_seleccionada.get())
        listbox.selection_set(idx)
        listbox.see(idx)
    except:
        pass
    def seleccionar():
        sel = listbox.curselection()
        if sel:
            universidad_seleccionada.set(listbox.get(sel[0]))
            top.destroy()
            actualizar_programas()
    btn_sel = tk.Button(top, text="Seleccionar", font=("Segoe UI", 11), command=seleccionar)
    btn_sel.pack(pady=8)
    listbox.bind('<Double-1>', lambda e: seleccionar())
    listbox.bind('<Return>', lambda e: seleccionar())

if __name__ == "__main__":
    root = tk.Tk()
    mostrar_login(root)
    root.mainloop()