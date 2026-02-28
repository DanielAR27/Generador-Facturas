# Generador de Facturas de Evaluación 📝

Esta es una aplicación web interactiva desarrollada en Python y Streamlit para la automatización y generación masiva de rúbricas de evaluación en formato Excel (`.xlsx`). Diseñado para facilitar la labor de asistencia y revisión académica.

## Características
- **Lectura dinámica de datos:** Importa listas de cursos y estudiantes desde archivos `.csv`.
- **Gestión de Equipos:** Permite la evaluación individual o la creación de múltiples equipos de trabajo, evitando la duplicidad de estudiantes.
- **Configuración de Rúbricas:** Asignación dinámica de la cantidad de partes a evaluar y sus respectivos pesos porcentuales.
- **Inyección en Excel:** Modifica una plantilla base (`plantilla.xlsx`), clonando formatos, reestructurando filas y reescribiendo fórmulas de manera automatizada utilizando `openpyxl`.

## 🛠️ Requisitos Previos
Asegúrese de tener instalado Python 3 y las siguientes librerías:
- `streamlit`
- `pandas`
- `openpyxl`

Puede instalarlas ejecutando:
```bash
pip install streamlit pandas openpyxl
```

## 📁 Estructura del Proyecto requerida
Para que el programa funcione correctamente, el directorio debe contener:

- **facturas.py:** El código fuente principal.

- **cursos.csv:** Archivo separado por punto y coma (;) con las columnas Siglas y Curso.

- **members_*.csv:** Archivos de estudiantes con las columnas Apellidos, Nombre y Rol (debe existir al menos un rol student).

- **plantilla.xlsx:** El archivo Excel base que el programa utilizará como molde.

## ▶️ Ejecución
Gracias a la configuración local del proyecto, solo necesita abrir una terminal en la carpeta del proyecto y ejecutar:

```bash
streamlit run facturas.py
```

El servidor local se levantará automáticamente en http://localhost:8080 con la función de recarga automática (hot-reloading) activada al guardar cambios en el código.

---
Desarrollado para uso académico.

