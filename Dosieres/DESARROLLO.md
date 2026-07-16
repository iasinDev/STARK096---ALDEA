# Guía de Desarrollo - STARK096 ALDEA

## Añadir Nuevas Funcionalidades

Esta guía explica cómo extender el sistema con nuevas características.

---

## 1. Añadir un Nuevo Script Python

### Paso 1: Crear el Script

Crea un nuevo archivo en `excelGenerator/`:

```python
#!/usr/bin/env python3
"""
STARK096 — ALDEA: [Nombre del Script]
[Descripción breve de lo que hace]
"""

from utils import (
    load_config,
    create_output_dir,
    get_timestamp,
    print_success,
    print_error,
    print_info
)


def main():
    """Función principal"""
    print_info("Starting [nombre del proceso]...")
    
    # Cargar configuración
    config = load_config()
    
    try:
        # Tu lógica aquí
        print_info("Processing...")
        
        # Ejemplo: crear directorio de salida
        output_dir = create_output_dir()
        
        # ... tu código ...
        
        print_success("Process completed successfully!")
        return 0
        
    except Exception as e:
        print_error(f"Process failed: {e}")
        return 1


if __name__ == "__main__":
    try:
        exit(main())
    except Exception as e:
        print_error(f"Fatal error: {e}")
        exit(1)
```

### Paso 2: Añadir Dependencias

Si necesitas librerías adicionales, añádelas a `requirements.txt`:

```bash
# Tu nueva librería
mi-libreria==1.0.0
```

### Paso 3: Actualizar menu.sh

Edita `menu.sh` y añade tu opción en el menú correspondiente:

```bash
show_excel_menu() {
    while true; do
        show_header
        echo -e "${CYAN}================================================================${NC}"
        echo -e "${CYAN}                      ${WHITE}EXCEL GENERATOR${CYAN}                      ${NC}"
        echo -e "${CYAN}================================================================${NC}"
        echo ""
        echo -e "${WHITE}${EXCEL} Available Options:${NC}"
        echo ""
        # ... opciones existentes ...
        
        # TU NUEVA OPCIÓN
        echo -e "${CYAN} ${GREEN}5.${NC} 🆕 Tu Nueva Funcionalidad                     ${NC}"
        echo -e "      ↳ Descripción de lo que hace                      ${NC}"
        echo ""
        
        echo -e "${CYAN} ${GREEN}0.${NC} ${EXIT} Back to Main Menu                               ${NC}"
        echo -e "${CYAN}================================================================${NC}"
        echo ""
        read -rp "" opc
        case "$opc" in
            # ... casos existentes ...
            
            5)
                show_header
                echo -e "${GREEN}🆕 Running tu nueva funcionalidad...${NC}"
                echo ""
                safe_run_python tu_nuevo_script.py
                exit_code=$?
                echo ""
                if [ $exit_code -eq 0 ]; then
                    echo -e "${GREEN}${CHECK} Completed successfully!${NC}"
                else
                    echo -e "${RED}${EXIT} Failed (exit code: $exit_code)${NC}"
                fi
                pause
                ;;
                
            0) return ;;
            *) echo -e "${RED}${EXIT} Invalid option.${NC}"; pause ;;
        esac
    done
}
```

### Paso 4: Reconstruir la Imagen

```bash
source menu.sh
# Opción 1: Container Management
# Opción 1: Build Image
```

---

## 2. Añadir Nuevos Tipos de Plantillas

### Crear una Nueva Plantilla

```python
#!/usr/bin/env python3
"""
Generador de plantilla personalizada
"""

import openpyxl
from utils import (
    load_config,
    create_output_dir,
    get_timestamp,
    apply_header_style,
    apply_cell_border
)


def create_custom_template():
    """Crea una plantilla personalizada"""
    
    config = load_config()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mi Plantilla"
    
    # Definir cabeceras personalizadas
    headers = ["Campo1", "Campo2", "Campo3", "Campo4"]
    
    # Aplicar estilos
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col)
        cell.value = header
        apply_header_style(cell, config)
    
    # Datos de ejemplo
    example_data = [
        ["Valor1", "Valor2", "Valor3", "Valor4"],
        ["Valor5", "Valor6", "Valor7", "Valor8"],
    ]
    
    for row_idx, data in enumerate(example_data, start=2):
        for col_idx, value in enumerate(data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = value
            apply_cell_border(cell)
    
    # Guardar
    output_dir = create_output_dir()
    output_file = f"{output_dir}/mi_plantilla_{get_timestamp()}.xlsx"
    wb.save(output_file)
    
    return output_file
```

---

## 3. Configuración Personalizada

### Añadir Nuevos Parámetros

Edita `excelGenerator/config/settings.yaml`:

```yaml
# Tu nueva sección
mi_seccion:
  parametro1: "valor1"
  parametro2: 123
  parametro3:
    - "opcion1"
    - "opcion2"
```

### Usar en tu Script

```python
config = load_config()

# Acceder a tus parámetros
mis_params = config.get('mi_seccion', {})
param1 = mis_params.get('parametro1', 'default')
```

---

## 4. Validación de Datos

### Usar el Validador

```python
from utils import validate_housing_data, print_error

data = {
    'id': 'VIV-001',
    'bloque': 'A',
    'tipo': '2D',
    'precio': 185000
}

config = load_config()
errors = validate_housing_data(data, config)

if errors:
    for error in errors:
        print_error(error)
    return 1
```

### Añadir Reglas de Validación

Edita `utils.py` y modifica `validate_housing_data()`:

```python
def validate_housing_data(data, config=None):
    """Validate housing data against configuration rules"""
    errors = []
    
    # ... validaciones existentes ...
    
    # TU NUEVA VALIDACIÓN
    if 'superficie' in data:
        if data['superficie'] < 30:
            errors.append("Superficie mínima no alcanzada")
    
    return errors
```

---

## 5. Procesamiento por Lotes

### Estructura Básica

```python
import os
import yaml
from tqdm import tqdm  # Barra de progreso

def process_batch():
    """Procesa múltiples ficheros"""
    
    # Listar ficheros de entrada
    input_dir = "input"
    input_files = [f for f in os.listdir(input_dir) 
                   if f.endswith(('.yaml', '.yml', '.csv'))]
    
    if not input_files:
        print_error("No input files found")
        return 1
    
    # Procesar cada fichero
    results = []
    for filename in tqdm(input_files, desc="Processing"):
        try:
            # Cargar datos
            if filename.endswith(('.yaml', '.yml')):
                with open(f"{input_dir}/{filename}", 'r') as f:
                    data = yaml.safe_load(f)
            
            # Procesar
            result = process_single_file(data)
            results.append((filename, result, "OK"))
            
        except Exception as e:
            results.append((filename, None, f"Error: {e}"))
    
    # Mostrar resumen
    print("\n" + "="*60)
    print("BATCH PROCESSING SUMMARY")
    print("="*60)
    for filename, result, status in results:
        print(f"  {filename}: {status}")
    
    return 0
```

---

## 6. Testing

### Crear Tests

Crea `excelGenerator/tests/test_template.py`:

```python
#!/usr/bin/env python3
"""
Tests for template generator
"""

import pytest
import sys
sys.path.insert(0, '..')

from excel_template_generator import create_template


def test_template_creation():
    """Test that template is created successfully"""
    result = create_template()
    assert result == 0


def test_template_content():
    """Test template has expected content"""
    # Implementar tests específicos
    pass
```

### Ejecutar Tests

```bash
# Dentro del contenedor
cd /app
python3 -m pytest tests/ -v
```

---

## 7. Utilidades Comunes

### Funciones Disponibles en utils.py

```python
from utils import (
    # Configuración
    load_config,
    create_output_dir,
    get_timestamp,
    
    # Estilos Excel
    get_header_style,
    get_cell_border,
    apply_header_style,
    apply_cell_border,
    
    # Formateo
    format_currency,
    
    # Validación
    validate_housing_data,
    
    # Mensajes
    print_success,
    print_error,
    print_warning,
    print_info
)
```

### Ejemplo de Uso Completo

```python
#!/usr/bin/env python3
"""
Ejemplo completo de uso de utilidades
"""

from utils import *
import openpyxl


def mi_generador():
    """Genera un reporte personalizado"""
    
    # Cargar configuración
    config = load_config()
    print_info("Configuration loaded")
    
    # Crear workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    
    # Cabecera
    headers = ["ID", "Nombre", "Precio"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col)
        cell.value = header
        apply_header_style(cell, config)
    
    # Datos con validación
    data_items = [
        {'id': '001', 'precio': 100000},
        {'id': '002', 'precio': 200000},
    ]
    
    for item in data_items:
        errors = validate_housing_data(item, config)
        if errors:
            for error in errors:
                print_warning(error)
    
    # Guardar
    output_dir = create_output_dir()
    filename = f"{output_dir}/reporte_{get_timestamp()}.xlsx"
    wb.save(filename)
    
    print_success(f"Report generated: {filename}")
    return 0


if __name__ == "__main__":
    try:
        exit(mi_generador())
    except Exception as e:
        print_error(f"Error: {e}")
        exit(1)
```

---

## 8. Mejores Prácticas

### Código Python

1. **Usa docstrings** para documentar funciones
2. **Maneja excepciones** apropiadamente
3. **Valida entradas** antes de procesarlas
4. **Usa las utilidades comunes** de `utils.py`
5. **Retorna códigos de salida** (0=éxito, 1=error)

### Menu.sh

1. **Mantén consistencia** con el formato existente
2. **Usa emojis** para mejorar la UX
3. **Proporciona feedback** claro al usuario
4. **Verifica el contenedor** antes de ejecutar scripts

### Configuración

1. **Documenta** nuevos parámetros en `settings.yaml`
2. **Usa valores por defecto** en el código
3. **No hardcodees** valores configurables
4. **Valida** la configuración al cargarla

---

## 9. Debugging

### Ver Logs del Contenedor

```bash
docker logs stark096_aldea_dev
```

### Ejecutar Script Manualmente

```bash
source menu.sh
# Opción 1 → Opción 3 (Enter Container)

# Dentro del contenedor
cd /app
python3 -u excel_template_generator.py
```

### Verificar Dependencias

```bash
# Dentro del contenedor
pip list
pip show openpyxl
```

---

## 10. Deployment

### Crear Versión Release

1. Actualiza `__init__.py` con nueva versión
2. Documenta cambios en `CHANGELOG.md`
3. Commit y tag en Git:

```bash
git add .
git commit -m "Release v1.1.0: New features"
git tag -a v1.1.0 -m "Version 1.1.0"
git push origin main --tags
```

---

## Recursos Adicionales

- [Documentación openpyxl](https://openpyxl.readthedocs.io/)
- [Documentación pandas](https://pandas.pydata.org/docs/)
- [Docker best practices](https://docs.docker.com/develop/dev-best-practices/)

---

**¿Dudas?** Consulta los archivos existentes como ejemplos de referencia.
