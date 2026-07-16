# Excel Generator - STARK096 ALDEA

Módulo Python para generación de ficheros Excel parametrizados.

## Estructura

```
excelGenerator/
├── __init__.py                    # Package initialization
├── requirements.txt               # Python dependencies
├── utils.py                       # Common utilities
│
├── Scripts principales
├── excel_template_generator.py   # Generador de plantillas
├── batch_processor.py             # Procesador por lotes
├── config_manager.py              # Gestor de configuración
├── file_validator.py              # Validador de ficheros
│
├── config/                        # Configuración
│   ├── README.md
│   └── settings.yaml              # Configuración principal
│
├── input/                         # Datos de entrada
│   ├── README.md
│   ├── example_data.yaml          # Ejemplo YAML
│   └── example_data.csv           # Ejemplo CSV
│
├── output/                        # Ficheros generados
│   └── README.md
│
└── templates/                     # Plantillas base
    └── README.md
```

## Scripts Disponibles

### excel_template_generator.py
Genera plantillas Excel base con formato parametrizado.

**Uso:**
```bash
python3 excel_template_generator.py
```

**Output:**
- `output/template_viviendas_YYYYMMDD_HHMMSS.xlsx`

---

### batch_processor.py
Procesa múltiples ficheros de datos y genera Excel.

**Uso:**
```bash
python3 batch_processor.py
```

**Input:**
- Ficheros en `input/` (.yaml, .csv, .json)

**Output:**
- Múltiples ficheros Excel en `output/`

---

### config_manager.py
Gestiona la configuración del sistema.

**Uso:**
```bash
python3 config_manager.py
```

**Output:**
- `config/settings.yaml` actualizado

---

### file_validator.py
Valida ficheros Excel generados.

**Uso:**
```bash
python3 file_validator.py
```

**Verifica:**
- Estructura de ficheros
- Integridad de datos
- Formato correcto

---

## Utilidades (utils.py)

### Funciones de Configuración
- `load_config()` - Carga configuración desde YAML
- `create_output_dir()` - Crea directorio de salida
- `get_timestamp()` - Genera timestamp actual

### Funciones de Estilo Excel
- `get_header_style()` - Obtiene estilo de cabecera
- `get_cell_border()` - Obtiene borde de celda
- `apply_header_style()` - Aplica estilo a cabecera
- `apply_cell_border()` - Aplica borde a celda

### Funciones de Formateo
- `format_currency()` - Formatea valores monetarios

### Funciones de Validación
- `validate_housing_data()` - Valida datos de vivienda

### Funciones de Mensajes
- `print_success()` - Mensaje de éxito ✅
- `print_error()` - Mensaje de error ❌
- `print_warning()` - Mensaje de advertencia ⚠️
- `print_info()` - Mensaje informativo ℹ️

---

## Configuración (settings.yaml)

### Secciones Principales

#### company
Información de la empresa constructora:
- name, address, city, postal_code
- phone, email, website
- logo_path

#### formats
Formatos de datos:
- currency, currency_symbol
- date_format
- decimal_separator, thousands_separator

#### paths
Rutas del sistema:
- input, output, templates, config, temp

#### excel
Configuración de estilos Excel:
- default_font, default_font_size
- header_color, header_font_color
- auto_filter, freeze_panes

#### housing_types
Tipos de viviendas disponibles:
- code, description
- min_surface, max_surface

#### validation
Reglas de validación:
- required_fields
- check_duplicates
- min_price, max_price

---

## Formatos de Entrada

### YAML
```yaml
viviendas:
  - id: "VIV-001"
    bloque: "A"
    piso: 1
    puerta: "A"
    tipo: "2D"
    superficie: 85.5
    precio: 185000
    estado: "Disponible"
```

### CSV
```csv
id,bloque,piso,puerta,tipo,superficie,precio,estado
VIV-001,A,1,A,2D,85.5,185000,Disponible
VIV-002,A,1,B,3D,105.2,235000,Reservada
```

### JSON
```json
{
  "viviendas": [
    {
      "id": "VIV-001",
      "bloque": "A",
      "piso": 1,
      "puerta": "A",
      "tipo": "2D",
      "superficie": 85.5,
      "precio": 185000,
      "estado": "Disponible"
    }
  ]
}
```

---

## Dependencias

Ver `requirements.txt` para lista completa.

**Principales:**
- `openpyxl` - Lectura/escritura de Excel
- `pandas` - Análisis de datos
- `xlsxwriter` - Creación avanzada de Excel
- `numpy` - Operaciones numéricas
- `pyyaml` - Parsing de YAML
- `tqdm` - Barras de progreso

---

## Desarrollo

### Añadir Nuevo Script

1. Crear archivo `.py` en este directorio
2. Importar utilities:
   ```python
   from utils import (
       load_config,
       create_output_dir,
       print_success
   )
   ```
3. Seguir estructura estándar:
   ```python
   def main():
       try:
           # Lógica aquí
           return 0
       except Exception as e:
           print_error(f"Error: {e}")
           return 1
   
   if __name__ == "__main__":
       exit(main())
   ```

### Testing

```bash
# Ejecutar script individual
python3 -u mi_script.py

# Ejecutar todos los tests
python3 -m pytest tests/ -v
```

### Debugging

```python
# Modo verbose
import logging
logging.basicConfig(level=logging.DEBUG)
```

---

## Ejemplos de Uso

### Generar Plantilla Simple

```python
from utils import *
import openpyxl

config = load_config()
wb = openpyxl.Workbook()
ws = wb.active

# Cabecera
headers = ["ID", "Nombre", "Precio"]
for col, header in enumerate(headers, start=1):
    cell = ws.cell(row=1, column=col)
    cell.value = header
    apply_header_style(cell, config)

# Guardar
output_dir = create_output_dir()
wb.save(f"{output_dir}/mi_plantilla_{get_timestamp()}.xlsx")
```

### Validar Datos

```python
from utils import validate_housing_data, load_config

config = load_config()
data = {'id': 'VIV-001', 'precio': 185000}

errors = validate_housing_data(data, config)
if errors:
    for error in errors:
        print(f"Error: {error}")
```

---

## Notas

- Todos los scripts se ejecutan desde el contenedor Docker
- Los ficheros generados se guardan en `output/`
- La configuración se gestiona mediante `settings.yaml`
- Usa `utils.py` para funcionalidad común

---

Para más información:
- [QUICKSTART.md](../QUICKSTART.md) - Guía rápida
- [Dosieres/DESARROLLO.md](../Dosieres/DESARROLLO.md) - Guía de desarrollo
- [README.md](../README.md) - Documentación principal
