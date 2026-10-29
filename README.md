# Convertidor de Facturas PDF a Excel - Papelería Acuarela

Aplicación web local desarrollada con **Flask** y **Python** para procesar y extraer automáticamente datos de facturas en formato PDF (mediante extracción de texto nativo y OCR) y exportarlos a hojas de cálculo de Excel (`.xlsx`) con un formato estructurado.

---

## 🚀 Requisitos

- **Python 3.10+**
- **pip** (administrador de paquetes de Python)

---

## 📦 Instalación

1. Clona este repositorio o descárgalo en el nuevo dispositivo:
   ```bash
   git clone <URL_DEL_REPOSITORIO>
   cd pdf_to_excel_app
   ```

2. (Opcional pero recomendado) Crea y activa un entorno virtual:
   ```bash
   python -m venv venv
   # En Windows:
   .\venv\Scripts\activate
   ```

3. Instala las dependencias necesarias:
   ```bash
   pip install -r requirements.txt
   ```

---

## 💻 Uso

### Opción 1: Acceso rápido (Windows)
Haz doble clic sobre el archivo ejecutable por lotes:
```text
iniciar_servidor.bat
```
Esto abrirá automáticamente tu navegador web predeterminado en `http://localhost:5000` y mantendrá el servidor Flask en ejecución.

### Opción 2: Desde terminal / consola
```bash
python app.py
```
Luego abre en tu navegador: [http://localhost:5000](http://localhost:5000)

---

## 📁 Estructura del Proyecto

- `app.py`: Servidor web Flask y rutas de la aplicación.
- `parser.py`: Motor de extracción de datos del PDF (OCR, tablas, expresiones regulares).
- `formato_de_tabla_excel.xlsx`: Plantilla base requerida para la generación del Excel final.
- `templates/index.html`: Interfaz web interactiva de usuario.
- `iniciar_servidor.bat`: Script de inicio rápido en Windows.
- `requirements.txt`: Lista de librerías Python requeridas.
