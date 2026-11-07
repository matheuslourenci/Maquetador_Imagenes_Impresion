import os
from tkinter import filedialog, messagebox, simpledialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image
import docx
from docx.shared import Cm, Mm

class OrganizadorImpresionApp(ttk.Window):
    def __init__(self):
        super().__init__(themename="yeti")
        self.title("Maquetador_A4 (Word)")
        self.resizable(True, True)
        
        try:
            self.iconbitmap("logo.ico")
        except Exception:
            pass
        
        # Centralizar pantalla
        ancho = 1000
        alto = 700
        ancho_pantalla = self.winfo_screenwidth()
        alto_pantalla = self.winfo_screenheight()
        x = (ancho_pantalla // 2) - (ancho // 2)
        y = (alto_pantalla // 2) - (alto // 2) 
        self.geometry(f"{ancho}x{alto}+{x}+{y}")
        
        self.lista_elementos = []
        self.crear_interfaz()

    def crear_interfaz(self):
        titulo = ttk.Label(self, text=" Maquetación de hojas A4 para impresión en Word (.docx)", font=("Helvetica", 16, "bold"))
        titulo.pack(pady=15)

        # 1. Cargar Imágenes y Asignar Alturas
        frame_archivos = ttk.LabelFrame(self, text=" 1. Añadir Imágenes y sus Alturas ")
        frame_archivos.pack(fill="x", padx=20, pady=5, ipady=10, ipadx=10)

        sub_botones = ttk.Frame(frame_archivos)
        sub_botones.pack(fill="x", padx=5, pady=5)

        btn_cargar = ttk.Button(sub_botones, text="Añadir Imágenes...", command=self.cargar_imagenes_con_dialogo, bootstyle="primary")
        btn_cargar.pack(side="left", padx=5, pady=5)

        btn_eliminar = ttk.Button(sub_botones, text="Eliminar Seleccionada", command=self.eliminar_imagen_seleccionada, bootstyle="warning")
        btn_eliminar.pack(side="left", padx=5, pady=5)

        btn_limpiar = ttk.Button(sub_botones, text="Limpiar Lista", command=self.limpiar_lista, bootstyle="danger-outline")
        btn_limpiar.pack(side="left", padx=5, pady=5)

        self.lbl_contador = ttk.Label(sub_botones, text="0 imágenes en la lista", font=("Helvetica", 10, "bold"))
        self.lbl_contador.pack(side="left", padx=15, pady=5)

        # Listbox para ver las imágenes añadidas
        frame_lista = ttk.Frame(frame_archivos)
        frame_lista.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.listbox = ttk.Treeview(frame_lista, columns=("Imagen", "Altura"), show="headings", height=6)
        self.listbox.heading("Imagen", text="Archivo de Imagen")
        self.listbox.heading("Altura", text="Altura Asignada (cm)")
        self.listbox.column("Imagen", width=450, anchor="w")
        self.listbox.column("Altura", width=180, anchor="center")
        
        scrollbar = ttk.Scrollbar(frame_lista, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 2. Márgenes de Impresión
        frame_margenes = ttk.LabelFrame(self, text=" 2. Márgenes de Hoja Word (mm) ")
        frame_margenes.pack(fill="x", padx=20, pady=10, ipady=10, ipadx=10)

        ttk.Label(frame_margenes, text="Margen Horizontal (mm):").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.entry_margen_h = ttk.Entry(frame_margenes, width=8)
        self.entry_margen_h.insert(0, "10")
        self.entry_margen_h.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(frame_margenes, text="Margen Vertical (mm):").grid(row=0, column=2, sticky="w", padx=15, pady=5)
        self.entry_margen_v = ttk.Entry(frame_margenes, width=8)
        self.entry_margen_v.insert(0, "10")
        self.entry_margen_v.grid(row=0, column=3, sticky="w", padx=5, pady=5)

        # 3. Botón de Generar Word
        self.btn_generar = ttk.Button(self, text="Generar Documento Word (.docx)", command=self.procesar_y_generar_word, bootstyle="success", width=35)
        self.btn_generar.pack(pady=10)

    def cargar_imagenes_con_dialogo(self):
        archivos = filedialog.askopenfilenames(
            title="Seleccionar imágenes",
            filetypes=[("Archivos de imagen", "*.jpg *.jpeg *.png *.bmp *.webp")]
        )
        if not archivos:
            return

        for archivo in archivos:
            nombre_corto = os.path.basename(archivo)
            
            while True:
                respuesta = simpledialog.askstring(
                    "Altura requerida", 
                    f"Introduce la altura en cm para:\n{nombre_corto}",
                    initialvalue="4.0"
                )
                
                if respuesta is None:
                    break
                
                try:
                    altura_cm = float(respuesta)
                    if altura_cm <= 0:
                        raise ValueError
                    
                    self.lista_elementos.append((archivo, altura_cm))
                    self.listbox.insert("", "end", values=(nombre_corto, f"{altura_cm} cm"))
                    break
                except ValueError:
                    messagebox.showerror("Valor inválido", "Por favor, introduce un número válido mayor a 0.")

        self.lbl_contador.config(text=f"{len(self.lista_elementos)} imágenes en la lista")

    def eliminar_imagen_seleccionada(self):
        seleccionados = self.listbox.selection()
        if not seleccionados:
            messagebox.showwarning("Advertencia", "Por favor, selecciona una imagen de la lista para eliminar.")
            return

        for item in seleccionados:
            indice = self.listbox.index(item)
            if 0 <= indice < len(self.lista_elementos):
                self.lista_elementos.pop(indice)
            self.listbox.delete(item)

        self.lbl_contador.config(text=f"{len(self.lista_elementos)} imágenes en la lista")

    def limpiar_lista(self):
        self.lista_elementos = []
        for item in self.listbox.get_children():
            self.listbox.delete(item)
        self.lbl_contador.config(text="0 imágenes en la lista")

    def procesar_y_generar_word(self):
        if not self.lista_elementos:
            messagebox.showwarning("Advertencia", "Por favor, añade al menos una imagen con su altura.")
            return

        try:
            margen_h_mm = float(self.entry_margen_h.get())
            margen_v_mm = float(self.entry_margen_v.get())
        except ValueError:
            messagebox.showerror("Error", "Introduce valores numéricos válidos en los márgenes.")
            return

        # Crear documento Word
        doc = docx.Document()
        
        # Configurar formato de página a A4 y establecer los márgenes
        seccion = doc.sections[0]
        seccion.page_width = Mm(210)
        seccion.page_height = Mm(297)
        seccion.top_margin = Mm(margen_v_mm)
        seccion.bottom_margin = Mm(margen_v_mm)
        seccion.left_margin = Mm(margen_h_mm)
        seccion.right_margin = Mm(margen_h_mm)

        # Definimos una tabla invisible de 3 columnas para organizar las imágenes de forma independiente
        columnas = 3
        filas = (len(self.lista_elementos) + columnas - 1) // columnas
        
        tabla = doc.add_table(rows=filas, cols=columnas)
        tabla.autofit = False

        # Rellenar la tabla celda por celda de forma independiente
        for idx, (ruta, altura_cm) in enumerate(self.lista_elementos):
            fila_actual = idx // columnas
            col_actual = idx % columnas
            celda = tabla.cell(fila_actual, col_actual)
            
            try:
                # Cada celda tiene su propio párrafo independiente
                parrafo = celda.paragraphs[0]
                run = parrafo.add_run()
                # Inserta la imagen respetando la altura elegida y calculando el ancho proporcional
                run.add_picture(ruta, height=Cm(altura_cm))
            except Exception as e:
                continue

        # Guardar archivo Word resultante
        ruta_salida = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Documento de Word", "*.docx")],
            title="Guardar documento Word para impresión"
        )
        
        if ruta_salida:
            doc.save(ruta_salida)
            messagebox.showinfo("¡Éxito!", f"Documento Word generado correctamente en:\n{ruta_salida}")

if __name__ == "__main__":
    app = OrganizadorImpresionApp()
    app.mainloop()