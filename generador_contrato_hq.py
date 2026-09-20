import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from datetime import datetime
import os
import subprocess
import sys
import json
from tkcalendar import DateEntry
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Image as RLImage, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Soporte para Excel
try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from PIL import Image, ImageTk
    HAS_PILLOW = True
except ImportError:
    HAS_PILLOW = False

ARCHIVO_FOLIO = "ultimo_folio_num.txt"
ARCHIVO_FLOTA = "flota_vehiculos.json"
ARCHIVO_EXCEL = "control_rentas.xlsx"
ARCHIVO_LOGOS = "config_logos.json"

class GeneradorContrato:
    def __init__(self, root):
        self.root = root
        self.root.title("TRAVELLER CAR RENTAL - Panel de Administración")
        self.root.geometry("1020x950")
        
        self.num_folio_actual = 1
        self.total_global_antes_anticipo = 0.0
        self.flota = self.cargar_flota()
        self.logos_config = self.cargar_logos_config()
        self.logo_seleccionado_clave = None
        
        # Colores del Tema Traveller
        self.color_primario = "#0D47A1"   # Azul Marino Traveller
        self.color_secundario = "#1565C0" # Azul Accent
        self.color_fondo = "#F5F7FA"     # Gris Claro
        self.color_blanco = "#FFFFFF"
        self.color_texto = "#212121"

        self.root.configure(bg=self.color_fondo)

        # --- ENCABEZADO PRINCIPAL ---
        frame_header_app = tk.Frame(root, bg=self.color_primario, pady=10, padx=15)
        frame_header_app.pack(fill="x", side="top")
        
        self.lbl_logo_preview_header = tk.Label(frame_header_app, bg=self.color_primario)
        self.lbl_logo_preview_header.pack(side="left", padx=(5, 15))

        self.lbl_titulo_app = tk.Label(
            frame_header_app, 
            text="TRAVELLER CAR RENTAL\nPanel de Administración y Control de Contratos", 
            font=("Helvetica", 15, "bold"), 
            fg=self.color_blanco, 
            bg=self.color_primario, 
            justify="left"
        )
        self.lbl_titulo_app.pack(side="left", fill="y")

        # --- CONTENEDOR CON BARRA DE DESPLAZAMIENTO (SCROLLBAR) ---
        container = tk.Frame(root, bg=self.color_fondo)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, bg=self.color_fondo, highlightthickness=0)
        scrollbar_v = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        
        self.scrollable_frame = tk.Frame(canvas, bg=self.color_fondo)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar_v.set)

        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=5)
        scrollbar_v.pack(side="right", fill="y")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # Notebook principal
        style_notebook = ttk.Style()
        style_notebook.configure("TNotebook", background=self.color_fondo)
        style_notebook.configure("TNotebook.Tab", font=("Helvetica", 9, "bold"), padding=[10, 5])
        
        self.notebook = ttk.Notebook(self.scrollable_frame)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)
        
        tab_logos = ttk.Frame(self.notebook)
        tab_general = ttk.Frame(self.notebook)
        tab_vehiculo = ttk.Frame(self.notebook)
        tab_coberturas = ttk.Frame(self.notebook)
        tab_costos = ttk.Frame(self.notebook)
        tab_flota = ttk.Frame(self.notebook)
        tab_excel = ttk.Frame(self.notebook)
        
        self.notebook.add(tab_logos, text="🖼️ Logos y Empresa")
        self.notebook.add(tab_general, text="1. Datos Generales")
        self.notebook.add(tab_vehiculo, text="2. Vehículo e Inventario")
        self.notebook.add(tab_coberturas, text="3. Coberturas y Tarjeta")
        self.notebook.add(tab_costos, text="4. Costos y Garantías")
        self.notebook.add(tab_flota, text="🚗 Gestión Flota")
        self.notebook.add(tab_excel, text="📊 Registro Excel")

        horas_disponibles = [f"{i:02d}" for i in range(24)]
        minutos_disponibles = [f"{i:02d}" for i in range(0, 60, 5)]

        # ==========================================
        # PESTAÑA LOGOS Y EMPRESA
        # ==========================================
        frame_logo_mgr = tk.LabelFrame(tab_logos, text=" Selección y Carga de Logos para PDF ", font=("Helvetica", 10, "bold"), padx=10, pady=10)
        frame_logo_mgr.pack(fill="x", padx=15, pady=10)

        tk.Label(frame_logo_mgr, text="Logo Activo para Contratos:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.combo_logos = ttk.Combobox(frame_logo_mgr, state="readonly", width=30)
        self.combo_logos.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        self.combo_logos.bind("<<ComboboxSelected>>", self.al_seleccionar_logo)

        btn_cargar_logo = tk.Button(frame_logo_mgr, text="➕ Agregar Nuevo Logo", bg=self.color_secundario, fg="white", font=("Helvetica", 9, "bold"), command=self.agregar_nuevo_logo)
        btn_cargar_logo.grid(row=0, column=2, padx=10, pady=5)

        btn_eliminar_logo = tk.Button(frame_logo_mgr, text="🗑 Eliminar Logo Seleccionado", bg="#D32F2F", fg="white", font=("Helvetica", 9, "bold"), command=self.eliminar_logo_actual)
        btn_eliminar_logo.grid(row=0, column=3, padx=5, pady=5)

        frame_empresa_info = tk.LabelFrame(tab_logos, text=" Configuración de Datos de Empresa (Para PDF) ", font=("Helvetica", 10, "bold"), padx=10, pady=10)
        frame_empresa_info.pack(fill="x", padx=15, pady=10)

        tk.Label(frame_empresa_info, text="Nombre Comercial de Empresa:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.entry_empresa_nombre = tk.Entry(frame_empresa_info, width=40)
        self.entry_empresa_nombre.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        tk.Label(frame_empresa_info, text="Correo Electrónico de Empresa:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.entry_empresa_email = tk.Entry(frame_empresa_info, width=40)
        self.entry_empresa_email.grid(row=1, column=1, sticky="w", padx=5, pady=5)

        btn_guardar_emp_nombre = tk.Button(frame_empresa_info, text="💾 Guardar Datos de Empresa", bg="#4CAF50", fg="white", font=("Helvetica", 9, "bold"), command=self.guardar_nombre_empresa)
        btn_guardar_emp_nombre.grid(row=2, column=0, columnspan=2, pady=10)

        self.actualizar_combo_logos()

        # ==========================================
        # PESTAÑA 1: DATOS GENERALES
        # ==========================================
        tk.Label(tab_general, text="Folio No.:", font=("Helvetica", 10, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        frame_folio_control = tk.Frame(tab_general)
        frame_folio_control.grid(row=0, column=1, columnspan=2, sticky="w", padx=10, pady=5)
        
        self.entry_folio_prefijo = tk.Entry(frame_folio_control, width=5, font=("Helvetica", 10, "bold"), fg="blue", justify="center")
        self.entry_folio_prefijo.insert(0, "TCR-")
        self.entry_folio_prefijo.pack(side="left", padx=(0, 5))

        tk.Button(frame_folio_control, text="◀", command=self.retroceder_folio, bg="#757575", fg="white", font=("Helvetica", 8, "bold"), width=3).pack(side="left", padx=2)
        self.entry_folio = tk.Entry(frame_folio_control, width=10, font=("Helvetica", 10, "bold"), fg="red", justify="center")
        self.entry_folio.pack(side="left", padx=2)
        tk.Button(frame_folio_control, text="▶", command=self.avanzar_folio, bg=self.color_primario, fg="white", font=("Helvetica", 8, "bold"), width=3).pack(side="left", padx=2)
        self.inicializar_folio()

        tk.Label(tab_general, text="Nombre/Razón Social:", font=("Helvetica", 9, "bold")).grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.entry_arrendador_nom = tk.Entry(tab_general, width=35)
        self.entry_arrendador_nom.insert(0, "RICARDO MEDINA GOMEZ")
        self.entry_arrendador_nom.grid(row=1, column=1, sticky="w", padx=10, pady=5)
        
        tk.Label(tab_general, text="Tel Celular:").grid(row=1, column=2, sticky="w", padx=5, pady=5)
        self.entry_arrendador_tel = tk.Entry(tab_general, width=20)
        self.entry_arrendador_tel.insert(0, "9991227269")
        self.entry_arrendador_tel.grid(row=1, column=3, sticky="w", padx=5, pady=5)

        tk.Label(tab_general, text="Nombre Cliente:", font=("Helvetica", 9, "bold")).grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.entry_nombre = tk.Entry(tab_general, width=45)
        self.entry_nombre.grid(row=2, column=1, columnspan=2, sticky="w", padx=10, pady=5)
        
        tk.Label(tab_general, text="Licencia No.:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        self.entry_licencia = tk.Entry(tab_general, width=20)
        self.entry_licencia.grid(row=3, column=1, sticky="w", padx=10, pady=5)
        
        tk.Label(tab_general, text="Vencimiento Lic.:").grid(row=3, column=2, sticky="w", padx=5, pady=5)
        self.cal_venc_lic = DateEntry(tab_general, width=12, background=self.color_primario, date_pattern='dd/mm/yyyy', state="readonly")
        self.cal_venc_lic.grid(row=3, column=3, sticky="w", padx=5, pady=5)
        
        tk.Label(tab_general, text="Domicilio/Hotel/Airbnb:").grid(row=4, column=0, sticky="w", padx=10, pady=5)
        self.entry_domicilio = tk.Entry(tab_general, width=45)
        self.entry_domicilio.grid(row=4, column=1, columnspan=2, sticky="w", padx=10, pady=5)
        
        tk.Label(tab_general, text="RFC Arrendatario:").grid(row=5, column=0, sticky="w", padx=10, pady=5)
        self.entry_rfc = tk.Entry(tab_general, width=20)
        self.entry_rfc.grid(row=5, column=1, sticky="w", padx=10, pady=5)
        
        tk.Label(tab_general, text="Correo Electrónico:").grid(row=5, column=2, sticky="w", padx=5, pady=5)
        self.entry_email = tk.Entry(tab_general, width=20)
        self.entry_email.grid(row=5, column=3, sticky="w", padx=5, pady=5)
        
        tk.Label(tab_general, text="Teléfono 1:").grid(row=6, column=0, sticky="w", padx=10, pady=5)
        self.entry_tel1 = tk.Entry(tab_general, width=20)
        self.entry_tel1.grid(row=6, column=1, sticky="w", padx=10, pady=5)
        
        tk.Label(tab_general, text="Teléfono 2:").grid(row=6, column=2, sticky="w", padx=5, pady=5)
        self.entry_tel2 = tk.Entry(tab_general, width=20)
        self.entry_tel2.grid(row=6, column=3, sticky="w", padx=5, pady=5)
        
        tk.Label(tab_general, text="Cond. Adicional 1:").grid(row=7, column=0, sticky="w", padx=10, pady=5)
        self.entry_cond1 = tk.Entry(tab_general, width=25)
        self.entry_cond1.grid(row=7, column=1, sticky="w", padx=10, pady=5)
        tk.Label(tab_general, text="Lic / Venc:").grid(row=7, column=2, sticky="w", padx=5, pady=5)
        frame_lic1 = tk.Frame(tab_general)
        frame_lic1.grid(row=7, column=3, sticky="w", padx=5, pady=5)
        self.entry_lic_cond1 = tk.Entry(frame_lic1, width=12)
        self.entry_lic_cond1.pack(side="left", padx=2)
        self.cal_venc_cond1 = DateEntry(frame_lic1, width=10, background=self.color_primario, date_pattern='dd/mm/yyyy', state="readonly")
        self.cal_venc_cond1.pack(side="left", padx=2)

        tk.Label(tab_general, text="Cond. Adicional 2:").grid(row=8, column=0, sticky="w", padx=10, pady=5)
        self.entry_cond2 = tk.Entry(tab_general, width=25)
        self.entry_cond2.grid(row=8, column=1, sticky="w", padx=10, pady=5)
        tk.Label(tab_general, text="Lic / Venc:").grid(row=8, column=2, sticky="w", padx=5, pady=5)
        frame_lic2 = tk.Frame(tab_general)
        frame_lic2.grid(row=8, column=3, sticky="w", padx=5, pady=5)
        self.entry_lic_cond2 = tk.Entry(frame_lic2, width=12)
        self.entry_lic_cond2.pack(side="left", padx=2)
        self.cal_venc_cond2 = DateEntry(frame_lic2, width=10, background=self.color_primario, date_pattern='dd/mm/yyyy', state="readonly")
        self.cal_venc_cond2.pack(side="left", padx=2)

        tk.Label(tab_general, text="Lugar de Entrega:").grid(row=9, column=0, sticky="w", padx=10, pady=5)
        self.entry_lugar = tk.Entry(tab_general, width=45)
        self.entry_lugar.insert(0, "AEROPUERTO")
        self.entry_lugar.grid(row=9, column=1, columnspan=2, sticky="w", padx=10, pady=5)

        tk.Label(tab_general, text="Lugar Devolución:").grid(row=10, column=0, sticky="w", padx=10, pady=5)
        self.entry_lugar_dev = tk.Entry(tab_general, width=45)
        self.entry_lugar_dev.insert(0, "AEROPUERTO")
        self.entry_lugar_dev.grid(row=10, column=1, columnspan=2, sticky="w", padx=10, pady=5)
        
        tk.Label(tab_general, text="Formato Renta:", font=("Helvetica", 9, "bold")).grid(row=11, column=0, sticky="w", padx=10, pady=5)
        self.combo_formato_tiempo = ttk.Combobox(tab_general, values=["Por ciclo de 24 Horas", "Por Día Calendario (Diario)"], width=25, state="readonly")
        self.combo_formato_tiempo.set("Por ciclo de 24 Horas")
        self.combo_formato_tiempo.grid(row=11, column=1, sticky="w", padx=10, pady=5)
        self.combo_formato_tiempo.bind("<<ComboboxSelected>>", lambda e: self.recalcular_tiempos_automatico())

        tk.Label(tab_general, text="Fecha Salida:").grid(row=12, column=0, sticky="w", padx=10, pady=5)
        self.cal_salida = DateEntry(tab_general, width=12, background=self.color_primario, date_pattern='dd/mm/yyyy', state="readonly")
        self.cal_salida.grid(row=12, column=1, sticky="w", padx=10, pady=5)
        self.cal_salida.bind("<<DateEntrySelected>>", lambda e: self.recalcular_tiempos_automatico())
        
        frame_h_salida = tk.Frame(tab_general)
        frame_h_salida.grid(row=12, column=3, sticky="w", padx=5, pady=5)
        tk.Label(tab_general, text="Hora Salida:").grid(row=12, column=2, sticky="w", padx=5, pady=5)
        self.combo_h_salida = ttk.Combobox(frame_h_salida, values=horas_disponibles, width=3, state="readonly")
        self.combo_h_salida.set("12")
        self.combo_h_salida.pack(side="left")
        self.combo_h_salida.bind("<<ComboboxSelected>>", lambda e: self.recalcular_tiempos_automatico())
        self.combo_m_salida = ttk.Combobox(frame_h_salida, values=minutos_disponibles, width=3, state="readonly")
        self.combo_m_salida.set("00")
        self.combo_m_salida.pack(side="left")
        self.combo_m_salida.bind("<<ComboboxSelected>>", lambda e: self.recalcular_tiempos_automatico())
        
        tk.Label(tab_general, text="Fecha Devolución:").grid(row=13, column=0, sticky="w", padx=10, pady=5)
        self.cal_devolucion = DateEntry(tab_general, width=12, background=self.color_primario, date_pattern='dd/mm/yyyy', state="readonly")
        self.cal_devolucion.grid(row=13, column=1, sticky="w", padx=10, pady=5)
        self.cal_devolucion.bind("<<DateEntrySelected>>", lambda e: self.recalcular_tiempos_automatico())
        
        frame_h_dev = tk.Frame(tab_general)
        frame_h_dev.grid(row=13, column=3, sticky="w", padx=5, pady=5)
        tk.Label(tab_general, text="Hora Devolución:").grid(row=13, column=2, sticky="w", padx=5, pady=5)
        self.combo_h_dev = ttk.Combobox(frame_h_dev, values=horas_disponibles, width=3, state="readonly")
        self.combo_h_dev.set("12")
        self.combo_h_dev.pack(side="left")
        self.combo_h_dev.bind("<<ComboboxSelected>>", lambda e: self.recalcular_tiempos_automatico())
        self.combo_m_dev = ttk.Combobox(frame_h_dev, values=minutos_disponibles, width=3, state="readonly")
        self.combo_m_dev.set("00")
        self.combo_m_dev.pack(side="left")
        self.combo_m_dev.bind("<<ComboboxSelected>>", lambda e: self.recalcular_tiempos_automatico())

        # ==========================================
        # PESTAÑA 2: VEHÍCULO E INVENTARIO
        # ==========================================
        frame_busqueda = tk.LabelFrame(tab_vehiculo, text=" 1. Registro de Flota (Buscador/Selección Automática) ")
        frame_busqueda.grid(row=0, column=0, columnspan=4, sticky="ew", padx=10, pady=5)
        
        tk.Label(frame_busqueda, text="Seleccionar de Flota Registrada:").pack(side="left", padx=5, pady=5)
        self.combo_seleccionar_placa = ttk.Combobox(frame_busqueda, state="readonly", width=25)
        self.combo_seleccionar_placa.pack(side="left", padx=5, pady=5)
        self.combo_seleccionar_placa.bind("<<ComboboxSelected>>", self.cargar_datos_auto_seleccionado)

        btn_eliminar_auto_p2 = tk.Button(frame_busqueda, text="🗑 Eliminar Auto", bg="#D32F2F", fg="white", font=("Helvetica", 8, "bold"), command=self.eliminar_auto_desde_pestana_vehiculo)
        btn_eliminar_auto_p2.pack(side="left", padx=10, pady=5)

        tk.Label(tab_vehiculo, text="Marca y Modelo:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        frame_marca_mod = tk.Frame(tab_vehiculo)
        frame_marca_mod.grid(row=1, column=1, sticky="w", padx=10, pady=5)
        self.entry_marca = tk.Entry(frame_marca_mod, width=15)
        self.entry_marca.pack(side="left", padx=(0,2))
        self.entry_modelo = tk.Entry(frame_marca_mod, width=15)
        self.entry_modelo.pack(side="left")

        tk.Label(tab_vehiculo, text="Año / Color:").grid(row=1, column=2, sticky="w", padx=5, pady=5)
        self.entry_anio_color = tk.Entry(tab_vehiculo, width=18)
        self.entry_anio_color.grid(row=1, column=3, sticky="w", padx=5, pady=5)
        
        tk.Label(tab_vehiculo, text="Número de Serie (VIN):").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.entry_serie = tk.Entry(tab_vehiculo, width=25)
        self.entry_serie.grid(row=2, column=1, sticky="w", padx=10, pady=5)

        tk.Label(tab_vehiculo, text="Placas:").grid(row=2, column=2, sticky="w", padx=5, pady=5)
        self.entry_placas = tk.Entry(tab_vehiculo, width=18)
        self.entry_placas.grid(row=2, column=3, sticky="w", padx=5, pady=5)

        tk.Label(tab_vehiculo, text="Kilometraje Actual:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        self.entry_km_salida = tk.Entry(tab_vehiculo, width=25)
        self.entry_km_salida.grid(row=3, column=1, sticky="w", padx=10, pady=5)

        # VEHÍCULO DE REPUESTO / ASIGNACIÓN
        frame_repuesto_box = tk.LabelFrame(tab_vehiculo, text=" 5. Vehículo de Repuesto / Asignación ", font=("Helvetica", 9, "bold"))
        frame_repuesto_box.grid(row=4, column=0, columnspan=4, sticky="ew", padx=10, pady=5)
        
        tk.Label(frame_repuesto_box, text="¿Activar Repuesto?:").grid(row=0, column=0, sticky="w", padx=5, pady=3)
        self.combo_repuesto_sn = ttk.Combobox(frame_repuesto_box, values=["no", "si"], width=5, state="readonly")
        self.combo_repuesto_sn.set("no")
        self.combo_repuesto_sn.grid(row=0, column=1, sticky="w", padx=5, pady=3)

        tk.Label(frame_repuesto_box, text="Seleccionar de Flota Registrada:").grid(row=0, column=2, sticky="w", padx=5, pady=3)
        self.combo_repuesto_flota = ttk.Combobox(frame_repuesto_box, state="readonly", width=25)
        self.combo_repuesto_flota.grid(row=0, column=3, sticky="w", padx=5, pady=3)
        self.combo_repuesto_flota.bind("<<ComboboxSelected>>", self.cargar_datos_auto_repuesto_seleccionado)
        
        tk.Label(frame_repuesto_box, text="Marca/Modelo:").grid(row=1, column=0, sticky="w", padx=5, pady=3)
        self.entry_rep_marca_modelo = tk.Entry(frame_repuesto_box, width=20)
        self.entry_rep_marca_modelo.grid(row=1, column=1, sticky="w", padx=5, pady=3)

        tk.Label(frame_repuesto_box, text="Año / Color:").grid(row=1, column=2, sticky="w", padx=5, pady=3)
        self.entry_rep_anio_color = tk.Entry(frame_repuesto_box, width=20)
        self.entry_rep_anio_color.grid(row=1, column=3, sticky="w", padx=5, pady=3)

        tk.Label(frame_repuesto_box, text="Placas:").grid(row=2, column=0, sticky="w", padx=5, pady=3)
        self.entry_rep_placas = tk.Entry(frame_repuesto_box, width=20)
        self.entry_rep_placas.grid(row=2, column=1, sticky="w", padx=5, pady=3)

        tk.Label(frame_repuesto_box, text="Número de Serie:").grid(row=2, column=2, sticky="w", padx=5, pady=3)
        self.entry_rep_serie = tk.Entry(frame_repuesto_box, width=20)
        self.entry_rep_serie.grid(row=2, column=3, sticky="w", padx=5, pady=3)

        tk.Label(frame_repuesto_box, text="Kilometraje al Entregar:").grid(row=3, column=0, sticky="w", padx=5, pady=3)
        self.entry_rep_km = tk.Entry(frame_repuesto_box, width=20)
        self.entry_rep_km.grid(row=3, column=1, sticky="w", padx=5, pady=3)

        tk.Label(frame_repuesto_box, text="Gasolina Repuesto:").grid(row=3, column=2, sticky="w", padx=5, pady=3)
        self.combo_rep_gasolina = ttk.Combobox(frame_repuesto_box, values=["4/4 (Full)", "3/4 Tanque", "1/2 Tanque", "1/4 Tanque", "Reserva", "Vacío"], width=18, state="readonly")
        self.combo_rep_gasolina.set("4/4 (Full)")
        self.combo_rep_gasolina.grid(row=3, column=3, sticky="w", padx=5, pady=3)

        tk.Label(tab_vehiculo, text="Aspectos Mecánicos:").grid(row=5, column=0, sticky="w", padx=10, pady=5)
        self.entry_asp_mecanicos = tk.Entry(tab_vehiculo, width=25)
        self.entry_asp_mecanicos.insert(0, "Excelentes")
        self.entry_asp_mecanicos.grid(row=5, column=1, sticky="w", padx=10, pady=5)
        tk.Label(tab_vehiculo, text="Aspectos Carrocería:").grid(row=5, column=2, sticky="w", padx=5, pady=5)
        self.entry_asp_carroceria = tk.Entry(tab_vehiculo, width=18)
        self.entry_asp_carroceria.insert(0, "Excelentes")
        self.entry_asp_carroceria.grid(row=5, column=3, sticky="w", padx=5, pady=5)

        tk.Label(tab_vehiculo, text="Niveles de Aceite:").grid(row=6, column=0, sticky="w", padx=10, pady=5)
        self.combo_aceite = ttk.Combobox(tab_vehiculo, values=["Lleno", "Vacío"], width=10, state="readonly")
        self.combo_aceite.set("Lleno")
        self.combo_aceite.grid(row=6, column=1, sticky="w", padx=10, pady=5)

        frame_gas = tk.LabelFrame(tab_vehiculo, text=" Tablero de Nivel de Gasolina ")
        frame_gas.grid(row=7, column=0, columnspan=4, sticky="ew", padx=10, pady=5)
        self.slider_gasolina = tk.Scale(frame_gas, from_=0, to=5, orient="horizontal", length=400, showvalue=False, command=self.actualizar_gasolina)
        self.slider_gasolina.set(5)
        self.slider_gasolina.pack(pady=2)
        
        frame_lbls = tk.Frame(frame_gas)
        frame_lbls.pack(fill="x")
        for txt in ["Vacío", "Reserva", "1/4", "1/2", "3/4", "4/4 (Full)"]:
            tk.Label(frame_lbls, text=txt, width=8, anchor="center").pack(side="left", expand=True)
        self.lbl_gas_txt = tk.Label(frame_gas, text="Nivel Seleccionado: 4/4 (Full)", font=("Helvetica", 10, "bold"), fg="green")
        self.lbl_gas_txt.pack()
        self.gas_actual_str = "4/4 (Full)"

        # ==========================================
        # PESTAÑA 3: COBERTURAS Y INFORMACIÓN BANCARIA
        # ==========================================
        tk.Label(tab_coberturas, text="Compañía Aseguradora:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.entry_aseguradora = tk.Entry(tab_coberturas, width=25)
        self.entry_aseguradora.insert(0, "ANA SEGUROS -- CHUBB")
        self.entry_aseguradora.grid(row=0, column=1, sticky="w", padx=10, pady=5)

        tk.Label(tab_coberturas, text="Tipo de Cobertura Global:").grid(row=0, column=2, sticky="w", padx=10, pady=5)
        self.entry_tipo_cobertura = tk.Entry(tab_coberturas, width=25)
        self.entry_tipo_cobertura.insert(0, "COBERTURA AMPLIA")
        self.entry_tipo_cobertura.grid(row=0, column=3, sticky="w", padx=10, pady=5)

        frame_deducibles = tk.LabelFrame(tab_coberturas, text=" 3. Coberturas y Seguros (CDW) - Opciones de Deducible ", font=("Helvetica", 9, "bold"))
        frame_deducibles.grid(row=1, column=0, columnspan=4, sticky="ew", padx=10, pady=10)

        self.var_deducible_opcion = tk.StringVar(value="15")
        
        rb1 = tk.Radiobutton(frame_deducibles, text="20% Deducible — $60,000.00", variable=self.var_deducible_opcion, value="15", font=("Helvetica", 9))
        rb1.pack(anchor="w", padx=15, pady=3)
        rb2 = tk.Radiobutton(frame_deducibles, text="10% Deducible — $40,000.00", variable=self.var_deducible_opcion, value="10", font=("Helvetica", 9))
        rb2.pack(anchor="w", padx=15, pady=3)
        rb3 = tk.Radiobutton(frame_deducibles, text="5% Deducible — $20,000.00", variable=self.var_deducible_opcion, value="5", font=("Helvetica", 9))
        rb3.pack(anchor="w", padx=15, pady=3)

        frame_tarjeta = tk.LabelFrame(tab_coberturas, text=" 4. Información del Titular y Tarjeta ", font=("Helvetica", 9, "bold"))
        frame_tarjeta.grid(row=2, column=0, columnspan=4, sticky="ew", padx=10, pady=10)

        tk.Label(frame_tarjeta, text="Titular de la Tarjeta:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.entry_tarjeta_titular = tk.Entry(frame_tarjeta, width=40)
        self.entry_tarjeta_titular.grid(row=0, column=1, columnspan=2, sticky="w", padx=10, pady=5)

        tk.Label(frame_tarjeta, text="Número de la Tarjeta:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.entry_tarjeta_num = tk.Entry(frame_tarjeta, width=25)
        self.entry_tarjeta_num.grid(row=1, column=1, sticky="w", padx=10, pady=5)

        tk.Label(frame_tarjeta, text="Vencimiento (EXP):").grid(row=1, column=2, sticky="w", padx=5, pady=5)
        self.entry_tarjeta_exp = tk.Entry(frame_tarjeta, width=12)
        self.entry_tarjeta_exp.grid(row=1, column=3, sticky="w", padx=5, pady=5)

        frame_obs = tk.LabelFrame(tab_coberturas, text=" 6. Observaciones / Notas Adicionales ", font=("Helvetica", 9, "bold"))
        frame_obs.grid(row=3, column=0, columnspan=4, sticky="ew", padx=10, pady=10)

        self.txt_observaciones = tk.Text(frame_obs, height=4, width=65)
        self.txt_observaciones.pack(padx=10, pady=5, fill="x")

        # ==========================================
        # PESTAÑA 4: COSTOS, CARGOS Y GARANTÍAS
        # ==========================================
        frame_cargos_sec = tk.LabelFrame(tab_costos, text=" 2. Costos, Cargos Adicionales y Garantías ", font=("Helvetica", 9, "bold"))
        frame_cargos_sec.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_cargos_sec, text="Días Vigencia Contrato:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.entry_dias = tk.Entry(frame_cargos_sec, width=12, font=("Helvetica", 9, "bold"), bg="#E0F7FA")
        self.entry_dias.grid(row=0, column=1, sticky="w", padx=10, pady=5)

        tk.Label(frame_cargos_sec, text="Renta Base ($ por día):").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        self.entry_tarifa = tk.Entry(frame_cargos_sec, width=12)
        self.entry_tarifa.grid(row=0, column=3, sticky="w", padx=5, pady=5)

        tk.Label(frame_cargos_sec, text="Garantía / Depósito ($):").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.entry_deposito = tk.Entry(frame_cargos_sec, width=12)
        self.entry_deposito.insert(0, "3000.00")
        self.entry_deposito.grid(row=1, column=1, sticky="w", padx=10, pady=5)

        tk.Label(frame_cargos_sec, text="Método Garantía:").grid(row=1, column=2, sticky="w", padx=5, pady=5)
        self.combo_metodo_garantia = ttk.Combobox(frame_cargos_sec, values=["Efectivo", "Tarjeta de Crédito", "Tarjeta de Débito", "Transferencia", "Voucher firmado"], width=18, state="readonly")
        self.combo_metodo_garantia.set("Tarjeta de Crédito")
        self.combo_metodo_garantia.grid(row=1, column=3, sticky="w", padx=5, pady=5)

        tk.Label(frame_cargos_sec, text="Cargo por Daño ($):").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.entry_cargo_dano = tk.Entry(frame_cargos_sec, width=12)
        self.entry_cargo_dano.insert(0, "0")
        self.entry_cargo_dano.grid(row=2, column=1, sticky="w", padx=10, pady=5)

        tk.Label(frame_cargos_sec, text="Cargo Traslado (Drop-Off) ($):").grid(row=2, column=2, sticky="w", padx=5, pady=5)
        self.entry_delivery = tk.Entry(frame_cargos_sec, width=12)
        self.entry_delivery.insert(0, "0")
        self.entry_delivery.grid(row=2, column=3, sticky="w", padx=5, pady=5)

        tk.Label(frame_cargos_sec, text="Cuota Protección (Seguro) ($):").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        self.entry_cuota_proteccion = tk.Entry(frame_cargos_sec, width=12)
        self.entry_cuota_proteccion.insert(0, "0")
        self.entry_cuota_proteccion.grid(row=3, column=1, sticky="w", padx=10, pady=5)

        tk.Label(frame_cargos_sec, text="Cobro de Gasolina ($):").grid(row=3, column=2, sticky="w", padx=5, pady=5)
        self.entry_cobro_gasolina = tk.Entry(frame_cargos_sec, width=12)
        self.entry_cobro_gasolina.insert(0, "0")
        self.entry_cobro_gasolina.grid(row=3, column=3, sticky="w", padx=5, pady=5)

        tk.Label(frame_cargos_sec, text="Faltante de Gasolina ($):").grid(row=4, column=0, sticky="w", padx=10, pady=5)
        self.entry_faltante_gasolina = tk.Entry(frame_cargos_sec, width=12)
        self.entry_faltante_gasolina.insert(0, "0")
        self.entry_faltante_gasolina.grid(row=4, column=1, sticky="w", padx=10, pady=5)

        tk.Label(frame_cargos_sec, text="Costo de Accesorios ($):").grid(row=4, column=2, sticky="w", padx=5, pady=5)
        self.entry_costo_accesorios = tk.Entry(frame_cargos_sec, width=12)
        self.entry_costo_accesorios.insert(0, "0")
        self.entry_costo_accesorios.grid(row=4, column=3, sticky="w", padx=5, pady=5)

        tk.Label(frame_cargos_sec, text="Horas Extra Calculadas:").grid(row=5, column=0, sticky="w", padx=10, pady=5)
        self.entry_cortesia_hrs = tk.Entry(frame_cargos_sec, width=12, bg="#E0F7FA")
        self.entry_cortesia_hrs.insert(0, "0")
        self.entry_cortesia_hrs.grid(row=5, column=1, sticky="w", padx=10, pady=5)

        tk.Label(frame_cargos_sec, text="Costo Total Horas Extra ($):").grid(row=5, column=2, sticky="w", padx=5, pady=5)
        self.entry_hora_extra_costo = tk.Entry(frame_cargos_sec, width=12)
        self.entry_hora_extra_costo.insert(0, "0")
        self.entry_hora_extra_costo.grid(row=5, column=3, sticky="w", padx=5, pady=5)

        tk.Label(frame_cargos_sec, text="IVA (%):").grid(row=6, column=0, sticky="w", padx=10, pady=5)
        self.entry_iva_porcentaje = tk.Entry(frame_cargos_sec, width=12)
        self.entry_iva_porcentaje.insert(0, "0")
        self.entry_iva_porcentaje.grid(row=6, column=1, sticky="w", padx=10, pady=5)

        tk.Label(frame_cargos_sec, text="Monto Anticipo / Reserva ($):").grid(row=6, column=2, sticky="w", padx=5, pady=5)
        self.entry_anticipo = tk.Entry(frame_cargos_sec, width=12)
        self.entry_anticipo.insert(0, "0")
        self.entry_anticipo.grid(row=6, column=3, sticky="w", padx=5, pady=5)

        frame_tot = tk.Frame(tab_costos)
        frame_tot.pack(pady=10)
        tk.Label(frame_tot, text="NETO A PAGAR ($):", font=("Helvetica", 11, "bold"), fg=self.color_primario).pack(side="left", padx=5)
        self.entry_total = tk.Entry(frame_tot, width=15, font=("Helvetica", 11, "bold"))
        self.entry_total.pack(side="left", padx=5)
        
        btn_calc = tk.Button(frame_tot, text="Calcular Total", command=self.calcular_todo, bg="#4CAF50", fg="white", font=("Helvetica", 9, "bold"))
        btn_calc.pack(side="left", padx=5)

        btn_generar = tk.Button(tab_costos, text="📄 GENERAR CONTRATO, REGISTRAR EXCEL E IMPRIMIR", font=("Helvetica", 11, "bold"), bg="#FF5722", fg="white", height=2, command=self.generar_pdf)
        btn_generar.pack(fill="x", padx=15, pady=15)

        # ==========================================
        # PESTAÑA 5: GESTIÓN DE FLOTA
        # ==========================================
        frame_flota_inputs = tk.LabelFrame(tab_flota, text=" Agregar / Editar Vehículo a la Flota (Estructura Completa) ", font=("Helvetica", 9, "bold"))
        frame_flota_inputs.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_flota_inputs, text="Placa:").grid(row=0, column=0, padx=5, pady=5)
        self.flota_placa = tk.Entry(frame_flota_inputs, width=10)
        self.flota_placa.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(frame_flota_inputs, text="Marca:").grid(row=0, column=2, padx=5, pady=5)
        self.flota_marca = tk.Entry(frame_flota_inputs, width=12)
        self.flota_marca.grid(row=0, column=3, padx=5, pady=5)

        tk.Label(frame_flota_inputs, text="Modelo:").grid(row=0, column=4, padx=5, pady=5)
        self.flota_modelo = tk.Entry(frame_flota_inputs, width=12)
        self.flota_modelo.grid(row=0, column=5, padx=5, pady=5)

        tk.Label(frame_flota_inputs, text="Año:").grid(row=1, column=0, padx=5, pady=5)
        self.flota_anio = tk.Entry(frame_flota_inputs, width=10)
        self.flota_anio.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(frame_flota_inputs, text="Color:").grid(row=1, column=2, padx=5, pady=5)
        self.flota_color = tk.Entry(frame_flota_inputs, width=12)
        self.flota_color.grid(row=1, column=3, padx=5, pady=5)

        tk.Label(frame_flota_inputs, text="VIN/Serie:").grid(row=1, column=4, padx=5, pady=5)
        self.flota_vin = tk.Entry(frame_flota_inputs, width=12)
        self.flota_vin.grid(row=1, column=5, padx=5, pady=5)

        tk.Label(frame_flota_inputs, text="Km Actual:").grid(row=2, column=0, padx=5, pady=5)
        self.flota_km = tk.Entry(frame_flota_inputs, width=10)
        self.flota_km.grid(row=2, column=1, padx=5, pady=5)

        btn_guardar_auto = tk.Button(frame_flota_inputs, text="➕ Guardar Vehículo a la Flota", bg="#4CAF50", fg="white", font=("Helvetica", 9, "bold"), command=self.agregar_auto_flota)
        btn_guardar_auto.grid(row=2, column=2, columnspan=4, pady=8)

        frame_tabla = tk.Frame(tab_flota)
        frame_tabla.pack(fill="both", expand=True, padx=10, pady=5)

        cols = ("Placa", "Marca", "Modelo", "Año", "Color", "Serie", "KM")
        self.tree_flota = ttk.Treeview(frame_tabla, columns=cols, show="headings", height=8)
        for c in cols:
            self.tree_flota.heading(c, text=c)
            self.tree_flota.column(c, width=90, anchor="center")
        self.tree_flota.pack(side="left", fill="both", expand=True)

        btn_borrar_auto = tk.Button(tab_flota, text="🗑 Eliminar Vehículo Seleccionado de la Flota", bg="#D32F2F", fg="white", font=("Helvetica", 9, "bold"), command=self.eliminar_auto_flota)
        btn_borrar_auto.pack(pady=5)

        self.actualizar_tabla_flota()
        self.actualizar_combo_placas()

        # ==========================================
        # PESTAÑA 6: REGISTRO Y CONTROL EXCEL
        # ==========================================
        frame_excel_mgr = tk.LabelFrame(tab_excel, text=" Administración del Archivo de Control de Rentas (control_rentas.xlsx) ", font=("Helvetica", 10, "bold"), padx=10, pady=10)
        frame_excel_mgr.pack(fill="both", expand=True, padx=15, pady=10)

        cols_excel = ("Folio", "Cliente", "Teléfono", "Auto / Modelo", "Placa", "Fecha Salida", "Hora Salida", "Fecha Devolución", "Hora Devolución", "Total Renta ($)")
        self.tree_excel = ttk.Treeview(frame_excel_mgr, columns=cols_excel, show="headings", height=10)
        for c in cols_excel:
            self.tree_excel.heading(c, text=c)
            self.tree_excel.column(c, width=100, anchor="center")
        self.tree_excel.pack(side="top", fill="both", expand=True, pady=5)

        frame_btn_excel = tk.Frame(frame_excel_mgr)
        frame_btn_excel.pack(fill="x", pady=5)

        btn_refrescar_excel = tk.Button(frame_btn_excel, text="🔄 Actualizar Vista Tabla", bg=self.color_secundario, fg="white", font=("Helvetica", 9, "bold"), command=self.cargar_datos_excel_tabla)
        btn_refrescar_excel.pack(side="left", padx=10)

        btn_borrar_excel = tk.Button(frame_btn_excel, text="🗑 Borrar/Vaciar Registro de Rentas (Excel)", bg="#D32F2F", fg="white", font=("Helvetica", 9, "bold"), command=self.vaciar_archivo_excel)
        btn_borrar_excel.pack(side="right", padx=10)

        self.cargar_datos_excel_tabla()

        # BARRA DE NAVEGACIÓN
        frame_navegacion = tk.Frame(root, bg=self.color_primario, height=45)
        frame_navegacion.pack(fill="x", side="bottom", ipady=3)
        
        self.btn_anterior = tk.Button(frame_navegacion, text="◀ Anterior", font=("Helvetica", 9, "bold"), bg="#757575", fg="white", width=12, command=self.pestaña_anterior)
        self.btn_anterior.pack(side="left", padx=25, pady=5)
        
        self.btn_siguiente = tk.Button(frame_navegacion, text="Siguiente ▶", font=("Helvetica", 9, "bold"), bg=self.color_secundario, fg="white", width=12, command=self.pestaña_siguiente)
        self.btn_siguiente.pack(side="right", padx=25, pady=5)
        
        self.notebook.bind("<<NotebookTabChanged>>", lambda e: self.actualizar_botones_navegacion())
        self.recalcular_tiempos_automatico()
        self.actualizar_botones_navegacion()

    # --- FUNCIONES DE LOGOS MULTI-EMPRESA Y EMAIL ---
    def cargar_logos_config(self):
        if os.path.exists(ARCHIVO_LOGOS):
            try:
                with open(ARCHIVO_LOGOS, "r") as f:
                    return json.load(f)
            except Exception:
                pass
        return {
            "TRAVELLER CAR RENTAL": {
                "ruta": "logo.jpg",
                "empresa_nombre": "TRAVELLER CAR RENTAL",
                "empresa_email": "TRAVELLERCARRENTALMID@GMAIL.COM"
            }
        }

    def guardar_logos_config(self):
        with open(ARCHIVO_LOGOS, "w") as f:
            json.dump(self.logos_config, f, indent=4)

    def actualizar_combo_logos(self):
        claves = list(self.logos_config.keys())
        self.combo_logos['values'] = claves
        if claves:
            if not self.logo_seleccionado_clave or self.logo_seleccionado_clave not in self.logos_config:
                self.logo_seleccionado_clave = claves[0]
            self.combo_logos.set(self.logo_seleccionado_clave)
            self.al_seleccionar_logo()

    def al_seleccionar_logo(self, event=None):
        clave = self.combo_logos.get()
        if clave in self.logos_config:
            self.logo_seleccionado_clave = clave
            cfg = self.logos_config[clave]
            
            self.entry_empresa_nombre.delete(0, tk.END)
            self.entry_empresa_nombre.insert(0, cfg.get("empresa_nombre", clave))
            
            self.entry_empresa_email.delete(0, tk.END)
            self.entry_empresa_email.insert(0, cfg.get("empresa_email", "TRAVELLERCARRENTALMID@GMAIL.COM"))
            
            ruta_img = cfg.get("ruta", "")
            if os.path.exists(ruta_img) and HAS_PILLOW:
                try:
                    img_raw = Image.open(ruta_img)
                    if img_raw.mode != 'RGB': img_raw = img_raw.convert('RGB')
                    img_resized = img_raw.resize((80, 50), Image.Resampling.LANCZOS)
                    self.logo_img_app = ImageTk.PhotoImage(img_resized)
                    self.lbl_logo_preview_header.config(image=self.logo_img_app, text="")
                except Exception:
                    self.lbl_logo_preview_header.config(image="", text="[Error Imagen]", fg="yellow", bg=self.color_primario)
            else:
                self.lbl_logo_preview_header.config(image="", text="[Sin Logo]", fg="white", bg=self.color_primario)

    def agregar_nuevo_logo(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar Logo",
            filetypes=[("Archivos de Imagen", "*.jpg *.jpeg *.png")]
        )
        if file_path:
            nombre_clave = os.path.splitext(os.path.basename(file_path))[0].upper()
            self.logos_config[nombre_clave] = {
                "ruta": file_path,
                "empresa_nombre": nombre_clave,
                "empresa_email": "TRAVELLERCARRENTALMID@GMAIL.COM"
            }
            self.guardar_logos_config()
            self.logo_seleccionado_clave = nombre_clave
            self.actualizar_combo_logos()
            messagebox.showinfo("Éxito", f"Logo '{nombre_clave}' agregado correctamente.")

    def eliminar_logo_actual(self):
        clave = self.combo_logos.get()
        if not clave:
            return
        if len(self.logos_config) <= 1:
            messagebox.showwarning("Atención", "Debe existir al menos un logo registrado.")
            return
        if messagebox.askyesno("Confirmar", f"¿Desea eliminar el logo '{clave}'?"):
            del self.logos_config[clave]
            self.logo_seleccionado_clave = None
            self.guardar_logos_config()
            self.actualizar_combo_logos()

    def guardar_nombre_empresa(self):
        clave = self.combo_logos.get()
        if clave in self.logos_config:
            nuevo_nombre = self.entry_empresa_nombre.get().strip()
            nuevo_email = self.entry_empresa_email.get().strip()
            self.logos_config[clave]["empresa_nombre"] = nuevo_nombre
            self.logos_config[clave]["empresa_email"] = nuevo_email
            self.guardar_logos_config()
            messagebox.showinfo("Éxito", "Datos de la empresa actualizados correctamente.")

    # --- FUNCIONES DE FLOTA COMPLETA ---
    def cargar_flota(self):
        if os.path.exists(ARCHIVO_FLOTA):
            try:
                with open(ARCHIVO_FLOTA, "r") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def guardar_flota(self):
        with open(ARCHIVO_FLOTA, "w") as f:
            json.dump(self.flota, f, indent=4)

    def agregar_auto_flota(self):
        placa = self.flota_placa.get().strip().upper()
        if not placa:
            messagebox.showerror("Error", "Ingrese la Placa del vehículo.")
            return

        self.flota[placa] = {
            "marca": self.flota_marca.get().strip(),
            "modelo": self.flota_modelo.get().strip(),
            "anio": self.flota_anio.get().strip(),
            "color": self.flota_color.get().strip(),
            "vin": self.flota_vin.get().strip(),
            "km": self.flota_km.get().strip()
        }
        self.guardar_flota()
        self.actualizar_tabla_flota()
        self.actualizar_combo_placas()
        messagebox.showinfo("Éxito", f"Vehículo {placa} guardado.")

    def eliminar_auto_flota(self):
        selected = self.tree_flota.selection()
        if not selected:
            return
        item = self.tree_flota.item(selected[0])
        placa = item['values'][0]
        if messagebox.askyesno("Confirmar", f"¿Eliminar vehículo {placa}?"):
            if str(placa) in self.flota:
                del self.flota[str(placa)]
                self.guardar_flota()
                self.actualizar_tabla_flota()
                self.actualizar_combo_placas()

    def eliminar_auto_desde_pestana_vehiculo(self):
        placa = self.combo_seleccionar_placa.get().strip()
        if not placa:
            messagebox.showwarning("Atención", "Seleccione un auto de la lista para eliminar.")
            return
        if messagebox.askyesno("Confirmar", f"¿Desea eliminar el vehículo con placas {placa} de la flota?"):
            if placa in self.flota:
                del self.flota[placa]
                self.guardar_flota()
                self.actualizar_tabla_flota()
                self.actualizar_combo_placas()
                self.combo_seleccionar_placa.set("")
                messagebox.showinfo("Éxito", f"Vehículo {placa} eliminado de la flota.")

    def actualizar_tabla_flota(self):
        for i in self.tree_flota.get_children():
            self.tree_flota.delete(i)
        for placa, d in self.flota.items():
            self.tree_flota.insert("", "end", values=(
                placa, 
                d.get("marca", ""), 
                d.get("modelo", ""),
                d.get("anio", ""),
                d.get("color", ""),
                d.get("vin", ""),
                d.get("km", "")
            ))

    def actualizar_combo_placas(self):
        placas = list(self.flota.keys())
        self.combo_seleccionar_placa['values'] = placas
        self.combo_repuesto_flota['values'] = placas

    def cargar_datos_auto_seleccionado(self, event=None):
        placa = self.combo_seleccionar_placa.get()
        if placa in self.flota:
            d = self.flota[placa]
            self.entry_placas.delete(0, tk.END); self.entry_placas.insert(0, placa)
            self.entry_marca.delete(0, tk.END); self.entry_marca.insert(0, d.get("marca", ""))
            self.entry_modelo.delete(0, tk.END); self.entry_modelo.insert(0, d.get("modelo", ""))
            
            anio_col = f"{d.get('anio', '')} / {d.get('color', '')}".strip(" /")
            self.entry_anio_color.delete(0, tk.END); self.entry_anio_color.insert(0, anio_col)
            self.entry_serie.delete(0, tk.END); self.entry_serie.insert(0, d.get("vin", ""))
            self.entry_km_salida.delete(0, tk.END); self.entry_km_salida.insert(0, d.get("km", ""))

    def cargar_datos_auto_repuesto_seleccionado(self, event=None):
        placa = self.combo_repuesto_flota.get()
        if placa in self.flota:
            d = self.flota[placa]
            self.entry_rep_placas.delete(0, tk.END); self.entry_rep_placas.insert(0, placa)
            
            marca_mod = f"{d.get('marca', '')} {d.get('modelo', '')}".strip()
            self.entry_rep_marca_modelo.delete(0, tk.END); self.entry_rep_marca_modelo.insert(0, marca_mod)
            
            anio_col = f"{d.get('anio', '')} / {d.get('color', '')}".strip(" /")
            self.entry_rep_anio_color.delete(0, tk.END); self.entry_rep_anio_color.insert(0, anio_col)
            
            self.entry_rep_serie.delete(0, tk.END); self.entry_rep_serie.insert(0, d.get("vin", ""))
            self.entry_rep_km.delete(0, tk.END); self.entry_rep_km.insert(0, d.get("km", ""))

    # --- REGISTRO Y CONTROL EN EXCEL ---
    def registrar_en_excel(self):
        if not HAS_OPENPYXL: return
        encabezados = ["Folio", "Cliente", "Teléfono", "Auto / Modelo", "Placa", "Fecha Salida", "Hora Salida", "Fecha Devolución", "Hora Devolución", "Total Renta ($)"]
        try:
            if not os.path.exists(ARCHIVO_EXCEL):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Entregas_Devoluciones"
                ws.append(encabezados)
            else:
                wb = openpyxl.load_workbook(ARCHIVO_EXCEL)
                ws = wb.active

            f_salida = self.cal_salida.get_date().strftime('%d/%m/%Y')
            f_dev = self.cal_devolucion.get_date().strftime('%d/%m/%Y')
            h_salida = f"{self.combo_h_salida.get()}:{self.combo_m_salida.get()}"
            h_dev = f"{self.combo_h_dev.get()}:{self.combo_m_dev.get()}"

            folio_completo = f"{self.entry_folio_prefijo.get().strip()}{self.entry_folio.get()}"

            fila = [
                folio_completo,
                self.entry_nombre.get(),
                self.entry_tel1.get(),
                f"{self.entry_marca.get()} {self.entry_modelo.get()}",
                self.entry_placas.get(),
                f_salida,
                h_salida,
                f_dev,
                h_dev,
                self.entry_total.get()
            ]
            ws.append(fila)
            wb.save(ARCHIVO_EXCEL)
            self.cargar_datos_excel_tabla()
        except Exception as e:
            messagebox.showerror("Error Excel", f"Error al guardar Excel: {str(e)}")

    def cargar_datos_excel_tabla(self):
        for i in self.tree_excel.get_children():
            self.tree_excel.delete(i)
        if os.path.exists(ARCHIVO_EXCEL) and HAS_OPENPYXL:
            try:
                wb = openpyxl.load_workbook(ARCHIVO_EXCEL)
                ws = wb.active
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if any(row):
                        self.tree_excel.insert("", "end", values=row)
            except Exception:
                pass

    def vaciar_archivo_excel(self):
        if not messagebox.askyesno("Confirmar Borrado", "¿Está seguro de que desea BORRAR/VACIAR todo el reporte de Excel?\nEsta acción creará un nuevo control de rentas desde cero."):
            return
        encabezados = ["Folio", "Cliente", "Teléfono", "Auto / Modelo", "Placa", "Fecha Salida", "Hora Salida", "Fecha Devolución", "Hora Devolución", "Total Renta ($)"]
        try:
            if HAS_OPENPYXL:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Entregas_Devoluciones"
                ws.append(encabezados)
                wb.save(ARCHIVO_EXCEL)
            elif os.path.exists(ARCHIVO_EXCEL):
                os.remove(ARCHIVO_EXCEL)
            self.cargar_datos_excel_tabla()
            messagebox.showinfo("Éxito", "El archivo de reporte de rentas ha sido vaciado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo limpiar el archivo Excel: {str(e)}")

    # --- NAVEGACIÓN Y CÁLCULOS ---
    def pestaña_siguiente(self):
        index = self.notebook.index(self.notebook.select())
        if index < 6: self.notebook.select(index + 1)

    def pestaña_anterior(self):
        index = self.notebook.index(self.notebook.select())
        if index > 0: self.notebook.select(index - 1)

    def actualizar_botones_navegacion(self):
        index = self.notebook.index(self.notebook.select())
        self.btn_anterior.config(state="disabled" if index == 0 else "normal")
        self.btn_siguiente.config(state="disabled" if index == 6 else "normal")

    def actualizar_gasolina(self, val):
        opciones = {"0": "Vacío", "1": "Reserva", "2": "1/4 Tanque", "3": "1/2 Tanque", "4": "3/4 Tanque", "5": "4/4 (Full)"}
        self.gas_actual_str = opciones.get(str(val), "4/4 (Full)")
        self.lbl_gas_txt.config(text=f"Nivel Seleccionado: {self.gas_actual_str}")

    def recalcular_tiempos_automatico(self):
        try:
            f_salida_dt = self.cal_salida.get_date()
            f_dev_dt = self.cal_devolucion.get_date()
            dt_inicio = datetime(f_salida_dt.year, f_salida_dt.month, f_salida_dt.day, int(self.combo_h_salida.get()), int(self.combo_m_salida.get()))
            dt_fin = datetime(f_dev_dt.year, f_dev_dt.month, f_dev_dt.day, int(self.combo_h_dev.get()), int(self.combo_m_dev.get()))
            
            if dt_fin <= dt_inicio:
                self.entry_dias.delete(0, tk.END); self.entry_dias.insert(0, "0")
                return

            diferencia = dt_fin - dt_inicio
            dias = diferencia.days
            horas_restantes = diferencia.seconds // 3600
            if (diferencia.seconds % 3600) // 60 >= 15: horas_restantes += 1
            if dias == 0 and horas_restantes > 0: dias = 1; horas_restantes = 0

            self.entry_dias.delete(0, tk.END); self.entry_dias.insert(0, str(dias))
            self.entry_cortesia_hrs.delete(0, tk.END); self.entry_cortesia_hrs.insert(0, str(horas_restantes))
        except Exception: pass

    def inicializar_folio(self):
        if os.path.exists(ARCHIVO_FOLIO):
            try:
                with open(ARCHIVO_FOLIO, "r") as f: self.num_folio_actual = int(f.read().strip())
            except ValueError: self.num_folio_actual = 1
        self.actualizar_pantalla_folio()

    def actualizar_pantalla_folio(self):
        self.entry_folio.delete(0, tk.END)
        self.entry_folio.insert(0, f"{self.num_folio_actual:04d}")

    def avanzar_folio(self):
        try: self.num_folio_actual = int(self.entry_folio.get())
        except ValueError: pass
        self.num_folio_actual += 1
        self.actualizar_pantalla_folio()

    def retroceder_folio(self):
        try: self.num_folio_actual = int(self.entry_folio.get())
        except ValueError: pass
        self.num_folio_actual = max(1, self.num_folio_actual - 1)
        self.actualizar_pantalla_folio()

    def guardar_ultimo_folio(self):
        with open(ARCHIVO_FOLIO, "w") as f: f.write(str(self.num_folio_actual))

    def calcular_todo(self):
        try:
            self.recalcular_tiempos_automatico()
            dias = int(self.entry_dias.get() or 0)
            tarifa = float(self.entry_tarifa.get() or 0)
            h_extra = float(self.entry_hora_extra_costo.get() or 0)
            delivery = float(self.entry_delivery.get() or 0)
            cargo_dano = float(self.entry_cargo_dano.get() or 0)
            cuota_prot = float(self.entry_cuota_proteccion.get() or 0)
            cobro_gas = float(self.entry_cobro_gasolina.get() or 0)
            falt_gas = float(self.entry_faltante_gasolina.get() or 0)
            costo_acc = float(self.entry_costo_accesorios.get() or 0)
            anticipo = float(self.entry_anticipo.get() or 0)

            subtotal = (dias * tarifa) + h_extra + delivery + cargo_dano + cuota_prot + cobro_gas + falt_gas + costo_acc
            iva = subtotal * (float(self.entry_iva_porcentaje.get() or 0) / 100)
            
            self.total_global_antes_anticipo = subtotal + iva
            total_final_neto = self.total_global_antes_anticipo - anticipo
            
            self.entry_total.delete(0, tk.END)
            self.entry_total.insert(0, f"{total_final_neto:.2f}")
        except ValueError:
            messagebox.showerror("Error", "Revise los valores numéricos ingresados.")

    # --- GENERACIÓN DEL PDF CON ESTILO EXACTO TRAVELLER ---
    def generar_pdf(self):
        try:
            self.calcular_todo()
            folio_num_str = self.entry_folio.get() or "0001"
            prefijo = self.entry_folio_prefijo.get().strip()
            folio_str = f"{prefijo}{folio_num_str}"
            nombre_archivo = f"Contrato_{folio_num_str}.pdf"
            
            # Márgenes reducidos para optimizar espacio
            doc = SimpleDocTemplate(nombre_archivo, pagesize=letter, rightMargin=12, leftMargin=12, topMargin=8, bottomMargin=8)
            story = []
            styles = getSampleStyleSheet()
            
            # Tipografías aumentadas para mejor lectura
            style_brand = ParagraphStyle('Brand', parent=styles['Heading1'], fontSize=16, alignment=0, textColor=colors.HexColor("#0D47A1"), leading=18, fontName="Helvetica-Bold")
            style_sub_brand = ParagraphStyle('SubBrand', parent=styles['Normal'], fontSize=8.5, leading=10.5, textColor=colors.black, fontName="Helvetica-Bold")
            style_head_right = ParagraphStyle('HeadRight', parent=styles['Normal'], fontSize=8, leading=10, alignment=2)
            style_folio = ParagraphStyle('Folio', parent=styles['Normal'], fontSize=11, leading=13, textColor=colors.red, alignment=2, fontName="Helvetica-Bold")
            
            style_section_head = ParagraphStyle('SecHead', parent=styles['Normal'], fontSize=8.5, leading=10, textColor=colors.white, fontName="Helvetica-Bold", alignment=1)
            style_cell = ParagraphStyle('Cell', parent=styles['Normal'], fontSize=7.5, leading=9.5, fontName="Helvetica")
            style_cell_bold = ParagraphStyle('CellB', parent=styles['Normal'], fontSize=7.5, leading=9.5, fontName="Helvetica-Bold")
            style_cell_center = ParagraphStyle('CellC', parent=styles['Normal'], fontSize=7.5, leading=9.5, alignment=1, fontName="Helvetica")
            
            style_legal = ParagraphStyle('Legal', parent=styles['Normal'], fontSize=6.5, leading=8.5, alignment=4, fontName="Helvetica")
            style_pagare = ParagraphStyle('Pagare', parent=styles['Normal'], fontSize=6.5, leading=8.5, alignment=4, fontName="Helvetica-Bold")

            clave_logo = self.combo_logos.get()
            cfg_logo = self.logos_config.get(clave_logo, {})
            nombre_empresa_pdf = cfg_logo.get("empresa_nombre", "TRAVELLER CAR RENTAL")
            email_empresa_pdf = cfg_logo.get("empresa_email", "TRAVELLERCARRENTALMID@GMAIL.COM")
            ruta_logo_pdf = cfg_logo.get("ruta", "")

            empresa_info = [
                Paragraph(f"<b>{nombre_empresa_pdf}</b>", style_brand),
                Paragraph("<b>ARRENDAMIENTO DE VEHICULOS</b>", style_sub_brand),
                Paragraph("<b>ARRENDADOR:</b> " + self.entry_arrendador_nom.get(), style_cell),
                Paragraph("<b>DOMICILIO:</b> CALLE 28 No. 315/23 y 25, COL. MANUEL CRESCENCIO REJÓN, MÉRIDA, YUC. C.P. 97255", style_cell),
                Paragraph("<b>R.F.C.:</b> MEGR9210230ZA", style_cell)
            ]
            
            header_der = [
                Paragraph(f"<b>FOLIO: {folio_str}</b>", style_folio),
                Paragraph("<b>HORARIO DE ATENCION DE LUNES A DOMINGO DE 7:00 A 21:00HS</b>", style_head_right),
                Paragraph(f"<b>TELEFONO OFICINA:</b> {self.entry_arrendador_tel.get()}", style_head_right),
                Paragraph("<b>TELEFONO EMERGENCIAS:</b> (999 436 7345 - 999 907 8749 24 HRS/7)", style_head_right),
                Paragraph(f"<b>CORREO ELECTRONICO:</b> {email_empresa_pdf}", style_head_right),
                Paragraph("<font color='red'><b>EL SEGURO SOLO CUBRE YUCATAN Y QUINTANA ROO.</b></font>", style_head_right)
            ]

            img_logo = RLImage(ruta_logo_pdf, width=75, height=60) if os.path.exists(ruta_logo_pdf) else Paragraph(f"<b>{nombre_empresa_pdf}</b>", style_brand)
            
            t_header = Table([[img_logo, empresa_info, header_der]], colWidths=[80, 255, 253])
            t_header.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('PADDING', (0,0), (-1,-1), 1)]))
            story.append(t_header)
            story.append(Spacer(1, 2))

            # DATOS DEL ARRENDATARIO
            story.append(Table([[Paragraph("<b>DATOS DEL ARRENDATARIO</b>", style_section_head)]], colWidths=[588], style=[('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#0D47A1")), ('PADDING', (0,0), (-1,-1), 2)]))
            
            venc_titular = self.cal_venc_lic.get_date().strftime('%d/%m/%Y')
            f_salida_str = f"{self.cal_salida.get_date().strftime('%d/%m/%Y')}"
            h_salida_str = f"{self.combo_h_salida.get()}:{self.combo_m_salida.get()} HRS"
            f_dev_str = f"{self.cal_devolucion.get_date().strftime('%d/%m/%Y')}"
            h_dev_str = f"{self.combo_h_dev.get()}:{self.combo_m_dev.get()} HRS"

            t_arr_data = [
                [Paragraph(f"<b>NOMBRE:</b> {self.entry_nombre.get()}", style_cell), Paragraph(f"<b>TELEFONO:</b> {self.entry_tel1.get()}", style_cell)],
                [Paragraph(f"<b>LUGAR DE HOSPEDAJE:</b> {self.entry_domicilio.get()}", style_cell), Paragraph(f"<b>TELEFONO 2:</b> {self.entry_tel2.get()}", style_cell)],
                [Paragraph(f"<b>CORREO ELECTRONICO:</b> {self.entry_email.get()}", style_cell), Paragraph(f"<b>RFC:</b> {self.entry_rfc.get()}", style_cell)],
                [Paragraph(f"<b>FECHA DE SALIDA:</b> {f_salida_str}", style_cell), Paragraph(f"<b>HORA:</b> {h_salida_str}", style_cell)],
                [Paragraph(f"<b>FECHA DE ENTREGA:</b> {f_dev_str}", style_cell), Paragraph(f"<b>HORA:</b> {h_dev_str}", style_cell)],
                [Paragraph(f"<b>DURACION DE ARRENDAMIENTO:</b> {self.entry_dias.get()} DIAS", style_cell), Paragraph(f"<b>LUGAR DE ENTREGA:</b> {self.entry_lugar.get()}", style_cell)],
                [Paragraph("<b>RENTA POR KILOMETRO:</b> ILIMITADO EN LA PENINSULA", style_cell), Paragraph("<b>GASOLINA:</b> NO ESTA INCLUIDA, NO REEMBOLSABLE", style_cell)]
            ]
            t_arr = Table(t_arr_data, colWidths=[294, 294])
            t_arr.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('GRID', (0,0), (-1,-1), 0.3, colors.lightgrey), ('PADDING', (0,0), (-1,-1), 1.5)]))
            story.append(t_arr)
            story.append(Spacer(1, 2))

            # CONDUCTORES Y LICENCIAS
            venc_c1 = self.cal_venc_cond1.get_date().strftime('%d/%m/%Y') if self.entry_cond1.get() else ""
            venc_c2 = self.cal_venc_cond2.get_date().strftime('%d/%m/%Y') if self.entry_cond2.get() else ""
            
            t_cond_data = [
                [Paragraph("<b>NOMBRE DE LOS CONDUCTORES:</b>", style_cell_bold), Paragraph("<b>LIC. DE CONDUCIR Y VENCIMIENTO:</b>", style_cell_bold)],
                [Paragraph(f"1.- {self.entry_nombre.get()}", style_cell), Paragraph(f"LIC: {self.entry_licencia.get()} | VENC: {venc_titular}", style_cell)],
                [Paragraph(f"2.- {self.entry_cond1.get() or 'N/A'}", style_cell), Paragraph(f"LIC: {self.entry_lic_cond1.get() or 'N/A'} | VENC: {venc_c1 or 'N/A'}", style_cell)],
                [Paragraph(f"3.- {self.entry_cond2.get() or 'N/A'}", style_cell), Paragraph(f"LIC: {self.entry_lic_cond2.get() or 'N/A'} | VENC: {venc_c2 or 'N/A'}", style_cell)]
            ]
            t_cond = Table(t_cond_data, colWidths=[294, 294])
            t_cond.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('GRID', (0,0), (-1,-1), 0.3, colors.lightgrey), ('PADDING', (0,0), (-1,-1), 1.5)]))
            story.append(t_cond)
            story.append(Spacer(1, 2))

            # VEHÍCULO PRINCIPAL / VEHÍCULO DE REPUESTO
            datos_auto_txt = (
                f"<b>MARCA / MODELO:</b> {self.entry_marca.get()} {self.entry_modelo.get()}<br/>"
                f"<b>AÑO / COLOR:</b> {self.entry_anio_color.get()} &nbsp;&nbsp;|&nbsp;&nbsp; <b>SERIE (VIN):</b> {self.entry_serie.get()}<br/>"
                f"<b>PLACAS:</b> {self.entry_placas.get()} &nbsp;&nbsp;|&nbsp;&nbsp; <b>KM SALIDA:</b> {self.entry_km_salida.get()} km<br/>"
                f"<b>NIVEL GASOLINA:</b> {self.gas_actual_str} &nbsp;&nbsp;|&nbsp;&nbsp; <b>ACEITE:</b> {self.combo_aceite.get()}"
            )

            if self.combo_repuesto_sn.get() == "si":
                repuesto_txt = (
                    f"<b>VEHÍCULO DE REPUESTO</b><br/>"
                    f"<b>MARCA/MOD:</b> {self.entry_rep_marca_modelo.get()}<br/>"
                    f"<b>PLACAS:</b> {self.entry_rep_placas.get()}<br/>"
                    f"<b>SERIE:</b> {self.entry_rep_serie.get()}<br/>"
                    f"<b>AÑO/COLOR:</b> {self.entry_rep_anio_color.get()}<br/>"
                    f"<b>KM ENTREGADO:</b> {self.entry_rep_km.get()} km<br/>"
                    f"<b>GASOLINA REP:</b> {self.combo_rep_gasolina.get()}"
                )
                col_derecha_contenido = Paragraph(repuesto_txt, style_cell)
            else:
                estado_vehiculo_txt = (
                    f"<b>ESTADO GENERAL DEL VEHÍCULO</b><br/>"
                    f"<b>Aspectos Mecánicos:</b> {self.entry_asp_mecanicos.get()}<br/>"
                    f"<b>Aspectos Carrocería:</b> {self.entry_asp_carroceria.get()}<br/>"
                    f"<b>Nivel de Aceite:</b> {self.combo_aceite.get()}<br/>"
                    f"<b>Gasolina:</b> {self.gas_actual_str}"
                )
                col_derecha_contenido = Paragraph(estado_vehiculo_txt, style_cell)

            t_vehiculo = Table([[Paragraph(datos_auto_txt, style_cell), col_derecha_contenido]], colWidths=[340, 248])
            t_vehiculo.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('GRID', (0,0), (-1,-1), 0.3, colors.lightgrey), ('PADDING', (0,0), (-1,-1), 2)]))
            story.append(t_vehiculo)
            story.append(Spacer(1, 2))

            # COBERTURAS Y INFORMACIÓN BANCARIA
            ded_opc = self.var_deducible_opcion.get()
            chk15 = "[X]" if ded_opc == "15" else "[  ]"
            chk10 = "[X]" if ded_opc == "10" else "[  ]"
            chk5  = "[X]" if ded_opc == "5"  else "[  ]"

            titular_tarjeta = self.entry_tarjeta_titular.get() or self.entry_nombre.get()
            num_tarjeta = self.entry_tarjeta_num.get() or "**** **** **** ****"
            exp_tarjeta = self.entry_tarjeta_exp.get() or "N/A"

            t_cob_data = [
                [Paragraph(f"<b>COBERTURA (CDW):</b> {chk15} 15% ($60k) | {chk10} 10% ($40k) | {chk5} 5% ($20k)", style_cell_bold), Paragraph(f"<b>TIPO:</b> {self.entry_tipo_cobertura.get()}", style_cell_bold)],
                [Paragraph(f"<b>TITULAR DE LA TARJETA:</b> {titular_tarjeta}", style_cell), Paragraph(f"<b>GARANTÍA / MÉTODO:</b> {self.combo_metodo_garantia.get()}", style_cell)],
                [Paragraph(f"<b>NÚMERO DE TARJETA:</b> {num_tarjeta}", style_cell), Paragraph(f"<b>VENCIMIENTO (EXP):</b> {exp_tarjeta}", style_cell)]
            ]
            t_cob = Table(t_cob_data, colWidths=[344, 244])
            t_cob.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('GRID', (0,0), (-1,-1), 0.3, colors.lightgrey), ('PADDING', (0,0), (-1,-1), 1.5)]))
            story.append(t_cob)
            story.append(Spacer(1, 2))

            # TABLA DE COSTOS, CARGOS ADICIONALES Y GARANTÍAS
            dias_val = self.entry_dias.get() or "1"
            tarifa_val = self.entry_tarifa.get() or "0"
            tot_renta = float(dias_val) * float(tarifa_val)
            
            t_servicios_data = [
                [Paragraph("<b>SERVICIOS Y CARGOS ADICIONALES</b>", style_cell_bold), Paragraph("<b>IMPORTE</b>", style_cell_bold)],
                [Paragraph(f"RENTA BASE (${tarifa_val} M.N. X {dias_val} DIAS)", style_cell), Paragraph(f"${tot_renta:,.2f} M.N.", style_cell_center)],
                [Paragraph(f"HORAS EXTRAS (${self.entry_hora_extra_costo.get()} M.N.)", style_cell), Paragraph(f"${float(self.entry_hora_extra_costo.get() or 0):,.2f} M.N.", style_cell_center)],
                [Paragraph(f"CARGO POR TRASLADO DE VEHÍCULO (DROP-OFF / DELIVERY)", style_cell), Paragraph(f"${float(self.entry_delivery.get() or 0):,.2f} M.N.", style_cell_center)],
                [Paragraph(f"CARGO POR DAÑO", style_cell), Paragraph(f"${float(self.entry_cargo_dano.get() or 0):,.2f} M.N.", style_cell_center)],
                [Paragraph(f"CUOTA DE PROTECCIÓN (SEGURO)", style_cell), Paragraph(f"${float(self.entry_cuota_proteccion.get() or 0):,.2f} M.N.", style_cell_center)],
                [Paragraph(f"COBRO DE GASOLINA / FALTANTE DE GASOLINA", style_cell), Paragraph(f"${(float(self.entry_cobro_gasolina.get() or 0) + float(self.entry_faltante_gasolina.get() or 0)):,.2f} M.N.", style_cell_center)],
                [Paragraph(f"COSTO DE ACCESORIOS", style_cell), Paragraph(f"${float(self.entry_costo_accesorios.get() or 0):,.2f} M.N.", style_cell_center)],
                [Paragraph(f"ANTICIPO DE RENTA (${self.entry_anticipo.get()}) | GARANTÍA RETENIDA: ${self.entry_deposito.get()}", style_cell), Paragraph(f"-${float(self.entry_anticipo.get() or 0):,.2f} M.N.", style_cell_center)],
                [Paragraph("<b>SUBTOTAL</b>", style_cell_bold), Paragraph(f"<b>${self.total_global_antes_anticipo:,.2f} M.N.</b>", style_cell_center)],
                [Paragraph("<b>IVA</b>", style_cell_bold), Paragraph(f"<b>${(self.total_global_antes_anticipo * (float(self.entry_iva_porcentaje.get() or 0)/100)):,.2f} M.N.</b>", style_cell_center)],
                [Paragraph("<b>MONTO A PAGAR NETO</b>", style_cell_bold), Paragraph(f"<b>${self.entry_total.get()} M.N.</b>", style_cell_center)]
            ]
            t_servicios = Table(t_servicios_data, colWidths=[448, 140])
            t_servicios.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('GRID', (0,0), (-1,-1), 0.3, colors.lightgrey),
                ('BACKGROUND', (0,0), (1,0), colors.HexColor("#E0E0E0")),
                ('BACKGROUND', (0,-1), (1,-1), colors.HexColor("#E3F2FD")),
                ('PADDING', (0,0), (-1,-1), 1.5)
            ]))
            story.append(t_servicios)
            story.append(Spacer(1, 2))

            # OBSERVACIONES EN EL PDF
            obs_texto = self.txt_observaciones.get("1.0", tk.END).strip()
            if obs_texto:
                t_obs = Table([[Paragraph(f"<b>OBSERVACIONES / NOTAS ADICIONALES:</b> {obs_texto}", style_cell)]], colWidths=[588])
                t_obs.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 0.5, colors.grey), ('PADDING', (0,0), (-1,-1), 2)]))
                story.append(t_obs)
                story.append(Spacer(1, 2))

            # TEXTO LEGAL Y CLÁUSULA DE RENTA MÍNIMA
            texto_renta_minima = (
                "<b>RENTA MINIMA:</b> Su renta mínima es de 24 hrs a partir de la hora en que se firma el contrato, "
                "teniendo 1 hora de tolerancia para la devolución de su vehículo. Si la devolución es posterior a lo "
                "mencionado, se le cobrará un día de renta a la tarifa contratada y sus adicionales del contrato firmado. "
                "<b>El vehículo no debe salir por ningún motivo fuera de la península de Yucatán.</b>"
            )
            story.append(Paragraph(texto_renta_minima, style_legal))
            story.append(Spacer(1, 2))

            # PAGARÉ INTEGRADO
            monto_total_str = self.entry_total.get() or "0.00"
            nombre_cliente = self.entry_nombre.get() or "____________________"
            
            texto_pagare = (
                f"<b>DEBE(MOS) Y PAGARE(MOS) INCONDICIONALMENTE POR ESTE PAGARÉ A LA ORDEN DE: RICARDO MEDINA GOMEZ EN MERIDA, YUCATAN LA CANTIDAD DE ${monto_total_str} MXN.</b><br/>"
                "ESTE PAGARÉ FORMARA PARTE DE LA SERIE NUMERADA DEL 1 AL... Y TODOS ESTAN SUJETOS A LAS CONDICIONES DE QUE AL NO PAGARSE CUALQUIERA DE ELLOS A SU VENCIMIENTO SERAN EXIGIBLES TODOS "
                "LOS QUE LE SIGAN EN NUMERO ADEMAS DE LOS YA VENCIDOS DESDE LA FECHA DE VENCIMIENTO DE ESTE DOCUMENTO HASTA EL DIA DE SU LIQUIDACION CAUSARA INTERESES MORATORIOS AL TIPO DE 5% "
                "MENSUAL PAGADERO EN ESTA CIUDAD JUNTAMENTE CON EL PRINCIPAL."
            )
            story.append(Paragraph(texto_pagare, style_pagare))
            story.append(Spacer(1, 2))

            # BLOQUE DE FIRMAS Y DATOS DEL DEUDOR
            t_firmas_data = [
                [
                    Paragraph(f"<b>NOMBRE DEUDOR:</b> {nombre_cliente}<br/><b>POBLACION / RFC:</b> MERIDA, YUC. / {self.entry_rfc.get()}", style_cell),
                    Paragraph(f"<b>BUENO POR:</b> ${monto_total_str} MXN<br/><b>ACEPTO (AMOS) FIRMA(S):</b> _______________________________________", style_cell)
                ],
                [
                    Paragraph("<br/><br/>________________________________________<br/><b>FIRMA O RUBRICA DE RESPONSIVA DEL ARRENDATARIO</b>", style_cell_center),
                    Paragraph("<br/><br/>________________________________________<br/><b>FIRMA ARRENDADOR / ENTREGADO POR</b>", style_cell_center)
                ]
            ]
            t_firmas = Table(t_firmas_data, colWidths=[294, 294])
            t_firmas.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('PADDING', (0,0), (-1,-1), 2)]))
            story.append(t_firmas)

            doc.build(story)

            # Registro en Excel e incremento del folio
            self.registrar_en_excel()
            self.guardar_ultimo_folio()

            messagebox.showinfo("Éxito", f"Contrato {nombre_archivo} generado correctamente e información guardada en Excel.")

            # Abrir archivo PDF generado
            if sys.platform.startswith('win'):
                os.startfile(nombre_archivo)
            elif sys.platform.startswith('darwin'):
                subprocess.call(('open', nombre_archivo))
            else:
                subprocess.call(('xdg-open', nombre_archivo))

        except Exception as e:
            messagebox.showerror("Error al generar PDF", f"Ocurrió un detalle al generar el archivo: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = GeneradorContrato(root)
    root.mainloop()