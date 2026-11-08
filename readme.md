# 📄 Maquetador A4 (Word) — Organizador de Imágenes para Impresión

Una aplicación de escritorio gráfica desarrollada en Python para automatizar el maquetado e inserción de múltiples imágenes en documentos de Microsoft Word (`.docx`), respetando dimensiones personalizadas y ajustando los márgenes de impresión en formato A4.

## Contexto y Caso de Uso

Al preparar archivos visuales o etiquetas para impresión en hojas A4, ajustar manualmente la altura y posición de cada imagen dentro de Word resulta un proceso lento y propenso a desalineaciones.

Esta herramienta permite a los usuarios seleccionar imágenes, definir individualmente la altura exacta en centímetros que debe tener cada una en el layout y generar automáticamente una cuadrícula organizada en Word con márgenes configurables.

El conjunto de imágenes generado en la hoja A4 es fácilmente ajustable por su estructuración en columnas preestablecidas.

---

## Características Principales

* **Ajuste de Altura Personalizado:** Permite asignar la altura exacta en centímetros (`cm`) a cada imagen al momento de cargarla.
* **Márgenes de Hoja Configurables:** Personalización de márgenes horizontales y verticales en milímetros (`mm`).
* **Organización en Cuadrícula:** Distribución automática de imágenes en 3 columnas dentro de una estructura invisible para optimizar el espacio A4.
* **Interfaz de Usuario Moderna:** Diseñada con `ttkbootstrap` (con soporte para modo claro *yeti* y modo oscuro  *darkly* ) para una experiencia más usable y accesible.
* **Vista Previa de Lista:** Gestión de elementos añadidos con opción de eliminar selecciones o limpiar la lista completa.

---

## Tecnologías Utilizadas

* **Lenguaje:** Python 3.x
* **Interfaz Gráfica (GUI):** `tkinter` / `ttkbootstrap`
* **Manipulación de Documentos:** `python-docx`
* **Procesamiento de Imágenes:** `Pillow` (PIL)

---

## Requisitos Previos e Instalación

1. **Clonar el repositorio:**

   ```bash
   git clone https://github.com/matheuslourenci/Maquetador_A4_Word.git
   cd Maquetador_A4_Word
   ```
2. **Instalar las dependencias:**

   ```bash
   pip install -r requirements.txt
   ```
3. **Ejecutar la aplicación:**

   ```bash
   python app.py
   ```
