# 📊 Procesador de Matrices

Aplicación web para crear matrices de intersección desde archivos Excel y CSV.

## Requisitos

- **Python 3.7+** (descargar desde [python.org](https://python.org/))

## Instalación

1. Descarga o clona este repositorio
2. Asegúrate de tener Python instalado

## Uso

### Opción 1: Doble clic (recomendado)
Simplemente haz **doble clic en `START.bat`**. La aplicación:
- Instalará automáticamente las dependencias necesarias (pandas, openpyxl)
- Abrirá tu navegador en `http://localhost:8080`

### Opción 2: Línea de comandos
```bash
python app.py
```

## Funcionalidades

### Paso 1: Cargar Archivos
- Arrastra y suelta archivos Excel (.xlsx, .xls) o CSV
- Carga múltiples archivos a la vez

### Paso 2: Seleccionar Hojas
- Elige qué hojas de cada archivo procesar
- Las hojas se auto-seleccionan para archivos CSV

### Paso 3: Definir Ejes
- **Eje X (Filas)**: Selecciona múltiples columnas que formarán las filas de la matriz
- **Eje Y (Columnas)**: Selecciona la columna que formará las columnas de la matriz
- Usa el botón "Aplicar selección a todos" para copiar la configuración a archivos con columnas similares
- Reordena las columnas de filas usando los botones ↑ ↓

### Paso 4: Filtrar (Opcional)
- Carga un archivo índice para filtrar las filas
- Útil para mantener solo empleados activos, por ejemplo

### Paso 5: Configurar Matrices
- Nombra cada matriz
- Combina múltiples fuentes en una sola matriz si es necesario

### Paso 6: Exportar
- Descarga un archivo Excel con:
  - **Hoja "Consulta"**: Búsqueda interactiva de permisos por usuario
  - **Hojas de matrices**: Una hoja por cada matriz generada

## Estructura de Archivos

```
Matriz/
├── app.py          # Servidor Python (backend)
├── index.html      # Interfaz web (frontend)
├── START.bat       # Ejecutable para Windows
├── README.md       # Este archivo
└── .gitignore      # Archivos ignorados por git
```

## Solución de Problemas

### La aplicación no abre
1. Verifica que Python esté instalado: `python --version`
2. Instala Python desde [python.org](https://python.org/)

### El navegador muestra una versión antigua
1. Cierra todas las pestañas de `localhost:8080`
2. Presiona `Ctrl+Shift+R` para forzar recarga sin caché

### Error al procesar archivos
1. Verifica que los archivos no estén corruptos
2. Asegúrate de que las columnas seleccionadas existan en los datos

## Licencia

MIT License
