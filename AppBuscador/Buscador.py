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

# Variables globales para mantener la referencia de las imágenes
logo_img_global = None
login_logo_img_global = None

def encontrar_ruta_drive():
    for letra in string.ascii_uppercase:
        posible_ruta = f"{letra}:\\Mi unidad\\Doctorados\\DB_General\\Universidad"
        if os.path.exists(posible_ruta):
            return posible_ruta
    return None

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

ruta_drive = encontrar_ruta_drive()

if ruta_drive is None:
    print("No se encontró la ruta de Doctorados.")
else:
    print(f"Ruta encontrada:\n{ruta_drive}")

items_clave = {
    "1. Hoja de Vida": ["h.vida"],
    "2. Contrato de Beca": ["c.beca"],
    "3. Aval de Facultad": ["avalfacultad"],
    "4. Acción Personal": ["acciónpersonal"],
    "5. Cédula y Papeleta de Votación": ["cédula"],
    "6. Pasaporte": ["pasaporte"],
    "7. Declaración Juramentada": ["declaraciónjuramentada"],
    "8. Solicitud de Pedido al Rectorado": ["solicitudpedidorectorado"],
    "9. Carta de Admisión de la Universidad": ["cartaadmisiónuniversidad"],
    "10. Certificado de Cuenta Bancaria": ["certificadobancario"],
    "11. Matrícula": ["matrícula"],
    "12. Certificado de Aprobación Tesis": ["certificadoaprobación"],
    "13. Reporte de Notas Semestrales": ["reportenotas"],
    "14. Línea de Investigación y Plan de Tesis": ["lineainvestigación"],
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

def actualizar_universidades():
    universidades = sorted(set(doc['universidad'] for doc in documentos_drive))
    combo_universidad['values'] = ['(Todos)'] + universidades
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
            
            # Filtrar por ítem clave
            if filtro_item == '(Todos)':
                encontrados.append(d)
            else:
                alias_list = items_clave.get(filtro_item, [])
                # Comprobar si el nombre contiene alguna de las palabras clave/alias
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
            if os.name == 'nt':  # Windows
                os.startfile(carpeta)
            elif os.name == 'posix':  # Mac/Linux
                import sys
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
    if not ruta_drive:
        documentos_drive = []
        actualizar_universidades()
        buscar()
        messagebox.showinfo("Sincronización", "No se encontró la ruta del drive. No hay documentos cargados.")
        return
    documentos_drive = cargar_documentos(ruta_drive)
    actualizar_universidades()
    buscar()
    messagebox.showinfo("Sincronización", "¡Sincronización completada! Los documentos han sido actualizados.")

def crear_encabezado(ventana):
    frame = tk.Frame(ventana, bg="#e8f0f7")
    frame.pack(pady=10, fill="x")

    # Frame izquierdo para logo y botón de sincronizar
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

    # Frame central para el nombre de la app
    frame_centro = tk.Frame(frame, bg="#e8f0f7")
    frame_centro.pack(side="left", expand=True)
    tk.Label(frame_centro, text="Buscador de Doctorados", font=("Segoe UI", 28, "bold"), bg="#e8f0f7").pack(expand=True)

    # Frame derecho para el botón de cerrar sesión
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
    resultados = ttk.Treeview(ventana, columns=columnas, show="headings", height=25)  # más alto
    resultados.pack(padx=20, pady=10, fill="both", expand=True)

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

    # Fondo para centrar el contenido
    fondo = tk.Frame(ventana, bg="#f2f2f2")
    fondo.pack(fill='both', expand=True)

    # Frame principal centrado
    frame = tk.Frame(fondo, padx=30, pady=30, bg="#f2f2f2")
    frame.place(relx=0.5, rely=0.5, anchor='center')

    # Logo superior
    try:
        ruta_imagen = resource_path("imagenes/logouce.png")
        imagen_logo = Image.open(ruta_imagen)
        imagen_logo = imagen_logo.resize((130, 130), Image.LANCZOS)
        ventana.logo_img = ImageTk.PhotoImage(imagen_logo)
        etiqueta_logo = tk.Label(frame, image=ventana.logo_img, bg="#f2f2f2")
        etiqueta_logo.pack(pady=(0, 10))
    except Exception as e:
        print("No se pudo cargar el logo:", e)

    # Título
    tk.Label(frame, text="Unidad Administrativa de Gestión de Doctorados", font=("Segoe UI", 14, "bold"), bg="#f2f2f2", fg="#2e4a62").pack(pady=(0, 20))

    # Usuario (en línea)
    fila_usuario = tk.Frame(frame, bg="#f2f2f2")
    fila_usuario.pack(pady=10)

    tk.Label(fila_usuario, text="Usuario:", font=("Segoe UI", 12), bg="#f2f2f2", width=12, anchor='e').pack(side='left')
    global entrada_usuario
    entrada_usuario = tk.Entry(fila_usuario, font=("Segoe UI", 12), width=30, relief="solid", bd=1)
    entrada_usuario.pack(side='left')

    # Contraseña (en línea)
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

    # Botón de login
    btn_ingresar = ttk.Button(frame, text="Ingresar", command=lambda: intentar_login(ventana))
    btn_ingresar.pack(pady=25)

    # Permitir ingresar con Enter
    ventana.bind('<Return>', lambda event: intentar_login(ventana))

def intentar_login(ventana):
    # Mostrar advertencia o información sobre la ruta del drive al intentar ingresar
    if ruta_drive is None:
        messagebox.showwarning("Advertencia", "No se encontró la ruta de Doctorados.")
    else:
        messagebox.showinfo("Información", f"Ruta encontrada:\n{ruta_drive}")
    usuario = entrada_usuario.get().strip()
    contrasena = entrada_contra.get().strip()
    valido, nombre = verificar_credenciales(usuario, contrasena)
    if valido:
        mostrar_carga_y_abrir_main(ventana, nombre)
    else:
        messagebox.showerror("Acceso denegado", "Usuario o contraseña incorrectos.")

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
    # Botón de navegación arriba
    nav_frame = tk.Frame(ventana, bg="#e8f0f7")
    nav_frame.pack(fill='x', pady=(10, 0))
    btn_volver_seleccion = ttk.Button(nav_frame, text="Volver a selección", command=lambda: mostrar_seleccion(ventana, ""))
    btn_volver_seleccion.pack(side='left', padx=10)
    # Resto del buscador
    main(ventana, skip_clear=True)

def main_oficios(ventana):
    limpiar_ventana(ventana)
    # Botón de navegación arriba
    nav_frame = tk.Frame(ventana, bg="#e8f0f7")
    nav_frame.pack(fill='x', pady=(10, 0))
    btn_volver_seleccion = ttk.Button(nav_frame, text="Volver a selección", command=lambda: mostrar_seleccion(ventana, ""))
    btn_volver_seleccion.pack(side='left', padx=10)
    global documentos_drive
    documentos_drive = cargar_documentos(ruta_drive)
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
    actualizar_universidades_oficios()
    combo_universidad_oficio.bind("<<ComboboxSelected>>", lambda e: actualizar_programas_oficios())
    combo_programa_oficio.bind("<<ComboboxSelected>>", lambda e: actualizar_estudiantes_oficios())
    combo_estudiante_oficio.bind("<<ComboboxSelected>>", lambda e: None)
    combo_item_clave_oficio['values'] = ['Oficios']
    combo_item_clave_oficio.set('Oficios')
    etiqueta_resumen_oficios.config(text="Mostrando 0 oficios.")

def crear_filtros_oficios(ventana):
    frame = tk.Frame(ventana, bg="#e8f0f7")
    frame.pack(pady=10)
    def crear_etiqueta(texto, fila, columna):
        tk.Label(frame, text=texto, font=('Segoe UI', 11), bg="#e8f0f7").grid(row=fila, column=columna, padx=5, pady=3)
    crear_etiqueta("Universidad:", 0, 0)
    global combo_universidad_oficio
    combo_universidad_oficio = ttk.Combobox(frame, width=25, state="readonly")
    combo_universidad_oficio.grid(row=0, column=1)
    crear_etiqueta("Programa:", 0, 2)
    global combo_programa_oficio
    combo_programa_oficio = ttk.Combobox(frame, width=25, state="readonly")
    combo_programa_oficio.grid(row=0, column=3)
    crear_etiqueta("Estudiante:", 0, 4)
    global combo_estudiante_oficio
    combo_estudiante_oficio = ttk.Combobox(frame, width=25, state="readonly")
    combo_estudiante_oficio.grid(row=0, column=5)
    crear_etiqueta("Filtro ítem clave:", 0, 6)
    global combo_item_clave_oficio
    combo_item_clave_oficio = ttk.Combobox(frame, width=25, state="readonly")
    combo_item_clave_oficio.grid(row=0, column=7)
    return combo_universidad_oficio, combo_programa_oficio, combo_estudiante_oficio, combo_item_clave_oficio

def crear_entrada_oficio(ventana):
    global entrada_num_oficio, entrada_fecha_oficio
    tk.Label(ventana, text="Buscar por número de oficio:", font=("Segoe UI", 12), bg="#e8f0f7").pack(pady=(10, 0))
    entrada_num_oficio = tk.Entry(ventana, width=30, font=("Segoe UI", 12), relief="solid", bd=1)
    entrada_num_oficio.pack(pady=4)
    tk.Label(ventana, text="Buscar por fecha (formato libre):", font=("Segoe UI", 12), bg="#e8f0f7").pack(pady=(10, 0))
    entrada_fecha_oficio = tk.Entry(ventana, width=30, font=("Segoe UI", 12), relief="solid", bd=1)
    entrada_fecha_oficio.pack(pady=4)

def crear_botones_oficios(ventana):
    frame = tk.Frame(ventana, bg="#e8f0f7")
    frame.pack(pady=10)
    btn_buscar = ttk.Button(frame, text="Buscar", command=buscar_oficios)
    btn_buscar.grid(row=0, column=0, padx=10)
    btn_volver = ttk.Button(frame, text="Volver", command=lambda: mostrar_seleccion(ventana, ""))
    btn_volver.grid(row=0, column=1, padx=10)

def crear_resultados_oficios(ventana):
    global resultados_oficios
    columnas = ("Universidad", "Programa", "Estudiante", "Nombre")
    resultados_oficios = ttk.Treeview(ventana, columns=columnas, show="headings", height=25)
    resultados_oficios.pack(padx=20, pady=10, fill="both", expand=True)
    for col in columnas:
        resultados_oficios.heading(col, text=col)
        resultados_oficios.column(col, anchor="center", width=200)
    resultados_oficios.tag_configure('oddrow', background="white")
    resultados_oficios.tag_configure('evenrow', background="#d3d3d3")
    resultados_oficios.bind("<<TreeviewSelect>>", mostrar_detalles_oficio)

def crear_texto_detalles_oficio(ventana):
    global texto_detalles_oficio
    tk.Label(ventana, text="Detalles del oficio seleccionado:", font=("Segoe UI", 12), bg="#e8f0f7").pack()
    texto_detalles_oficio = tk.Text(ventana, height=4, width=100, font=("Segoe UI", 11), relief="solid", bd=1)
    texto_detalles_oficio.pack(padx=40, pady=(0, 25), fill='x')
    texto_detalles_oficio.config(state="disabled")

def actualizar_universidades_oficios():
    universidades = sorted(set(doc['universidad'] for doc in documentos_drive if 'oficio' in doc['nombre']))
    combo_universidad_oficio['values'] = ['(Todos)'] + universidades
    combo_universidad_oficio.set('(Todos)')
    actualizar_programas_oficios()

def actualizar_programas_oficios():
    u = combo_universidad_oficio.get()
    docs = [doc for doc in documentos_drive if 'oficio' in doc['nombre']]
    if u == '(Todos)':
        programas = sorted(set(doc['programa'] for doc in docs))
    else:
        programas = sorted(set(doc['programa'] for doc in docs if doc['universidad'] == u))
    combo_programa_oficio['values'] = ['(Todos)'] + programas
    combo_programa_oficio.set('(Todos)')
    actualizar_estudiantes_oficios()

def actualizar_estudiantes_oficios():
    u = combo_universidad_oficio.get()
    p = combo_programa_oficio.get()
    docs = [doc for doc in documentos_drive if 'oficio' in doc['nombre']]
    estudiantes = set()
    for doc in docs:
        if (u == '(Todos)' or doc['universidad'] == u) and (p == '(Todos)' or doc['programa'] == p):
            estudiantes.add(doc['estudiante'])
    combo_estudiante_oficio['values'] = ['(Todos)'] + sorted(estudiantes)
    combo_estudiante_oficio.set('(Todos)')

def buscar_oficios():
    resultados_oficios.delete(*resultados_oficios.get_children())
    filtro_u = combo_universidad_oficio.get()
    filtro_p = combo_programa_oficio.get()
    filtro_e = combo_estudiante_oficio.get()
    filtro_num = entrada_num_oficio.get().strip().lower()
    filtro_fecha = entrada_fecha_oficio.get().strip().lower()
    docs = [doc for doc in documentos_drive if 'oficio' in doc['nombre']]
    encontrados = []
    for d in docs:
        if (filtro_u == '(Todos)' or d['universidad'] == filtro_u) and \
           (filtro_p == '(Todos)' or d['programa'] == filtro_p) and \
           (filtro_e == '(Todos)' or d['estudiante'] == filtro_e):
            # Filtrar por número de oficio y fecha (en el nombre)
            if (filtro_num in d['nombre']) and (filtro_fecha in d['nombre']):
                encontrados.append(d)
    if not encontrados:
        messagebox.showinfo("No encontrado", "No se encontró ningún oficio con los filtros aplicados.")
        etiqueta_resumen_oficios.config(text="Mostrando 0 oficios.")
        limpiar_detalles_oficio()
        return
    for i, doc in enumerate(encontrados):
        tag = 'evenrow' if i % 2 == 0 else 'oddrow'
        resultados_oficios.insert("", "end", values=(
            doc['universidad'].title(),
            doc['programa'].title(),
            doc['estudiante'].title(),
            doc['nombre'].title()
        ), tags=(tag,))
    etiqueta_resumen_oficios.config(text=f"Mostrando {len(encontrados)} oficio(s).")
    limpiar_detalles_oficio()

def mostrar_detalles_oficio(event):
    item = resultados_oficios.focus()
    if item:
        valores = resultados_oficios.item(item, 'values')
        info = (
            f"Universidad: {valores[0]}\n"
            f"Programa: {valores[1]}\n"
            f"Estudiante: {valores[2]}\n"
            f"Nombre: {valores[3]}\n"
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

def cerrar_sesion(ventana):
    if messagebox.askyesno("Cerrar Sesión", "¿Estás seguro de que quieres cerrar sesión?"):
        mostrar_login(ventana)

def main(ventana, skip_clear=False):
    if not skip_clear:
        limpiar_ventana(ventana)
    global documentos_drive
    documentos_drive = cargar_documentos(ruta_drive)
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
    combo_universidad.bind("<<ComboboxSelected>>", lambda e: actualizar_programas())
    combo_programa.bind("<<ComboboxSelected>>", lambda e: actualizar_estudiantes())
    combo_estudiante.bind("<<ComboboxSelected>>", lambda e: None)
    combo_item_clave['values'] = ['(Todos)'] + list(items_clave.keys())
    combo_item_clave.set('(Todos)')
    if not ruta_drive:
        etiqueta_resumen.config(text="No se encontró la ruta del drive. No hay documentos cargados.")

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
    
    # Estilo personalizado para el botón de cerrar sesión
    style.configure('CerrarSesion.TButton', 
                    font=('Segoe UI', 11, 'bold'), 
                    padding=8,
                    background='#dc3545',
                    foreground='white')
    style.map('CerrarSesion.TButton',
              background=[('active', '#c82333'), ('pressed', '#bd2130')],
              foreground=[('active', 'white'), ('pressed', 'white')])
    # Estilo para el botón de sincronización
    style.configure('Sincronizar.TButton', 
                    font=('Segoe UI', 11, 'bold'), 
                    padding=8,
                    background='#007bff',
                    foreground='white')
    style.map('Sincronizar.TButton',
              background=[('active', '#0056b3'), ('pressed', '#004085')],
              foreground=[('active', 'white'), ('pressed', 'white')])

if __name__ == "__main__":
    root = tk.Tk()
    mostrar_login(root)
    root.mainloop()
