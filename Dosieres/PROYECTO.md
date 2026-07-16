# Documentación del Proyecto STARK096 — ALDEA

## Índice

1. [Visión General](#visión-general)
2. [Arquitectura del Sistema](#arquitectura-del-sistema)
3. [Guías de Uso](#guías-de-uso)
4. [Desarrollo](#desarrollo)

## Visión General

STARK096—ALDEA es un sistema automatizado para la generación de ficheros Excel parametrizados dirigido a empresas constructoras. El objetivo principal es facilitar la creación, gestión y validación de documentos relacionados con proyectos de construcción de viviendas.

## Arquitectura del Sistema

### Componentes Principales

```
┌─────────────────────────────────────────────┐
│          menu.sh (Interfaz)                 │
├─────────────────────────────────────────────┤
│         Docker Container                     │
│  ┌───────────────────────────────────────┐  │
│  │     Python 3.11                       │  │
│  │  ┌──────────────────────────────┐     │  │
│  │  │  Excel Generator Scripts     │     │  │
│  │  │  - template_generator.py     │     │  │
│  │  │  - batch_processor.py        │     │  │
│  │  │  - config_manager.py         │     │  │
│  │  │  - file_validator.py         │     │  │
│  │  └──────────────────────────────┘     │  │
│  └───────────────────────────────────────┘  │
└─────────────────────────────────────────────┘
```

### Flujo de Trabajo

1. **Configuración**: Define parámetros en `config/settings.yaml`
2. **Entrada de Datos**: Coloca ficheros en `input/`
3. **Procesamiento**: Ejecuta scripts desde el menú
4. **Generación**: Ficheros Excel creados en `output/`
5. **Validación**: Verifica integridad de ficheros generados

## Guías de Uso

### Primer Uso

1. Abre Git Bash en el directorio del proyecto
2. Ejecuta `source menu.sh`
3. Selecciona "Container Management" → "Build Image"
4. Espera a que se construya la imagen Docker
5. Vuelve al menú principal y selecciona "Excel Generator"

### Generación de Plantillas

1. Ejecuta `source menu.sh`
2. Opción 2: "Excel Generator"
3. Opción 1: "Generate Excel Template"
4. Los ficheros se generan en `excelGenerator/output/`

### Procesamiento por Lotes

1. Coloca tus datos en `excelGenerator/input/`
2. Configura parámetros si es necesario (opción 3)
3. Ejecuta "Process Batch Data" (opción 2)
4. Valida los resultados (opción 4)

## Desarrollo

### Añadir Nuevas Funcionalidades

1. Crea un nuevo script Python en `excelGenerator/`
2. Añade la dependencia en `requirements.txt` si es necesario
3. Actualiza `menu.sh` con la nueva opción del menú
4. Reconstruye la imagen Docker

### Estructura de Scripts Python

Todos los scripts deben seguir esta estructura:

```python
#!/usr/bin/env python3
"""
STARK096 — ALDEA: [Nombre del Script]
[Descripción breve]
"""

def main():
    """Función principal"""
    print("🏗️  [Descripción de la acción]")
    # Lógica del script
    return 0  # 0 = éxito, 1 = error

if __name__ == "__main__":
    try:
        exit(main())
    except Exception as e:
        print(f"❌ Error: {e}")
        exit(1)
```

### Dependencias

Ver `excelGenerator/requirements.txt` para lista completa.

Principales:
- `openpyxl` - Manipulación de Excel
- `pandas` - Procesamiento de datos
- `pyyaml` - Configuración

### Testing

```bash
# Entrar al contenedor
source menu.sh
# Opción 1 → Opción 3

# Dentro del contenedor
python3 -m pytest tests/
```

## Próximos Pasos

- [ ] Implementar procesamiento por lotes completo
- [ ] Añadir más tipos de plantillas
- [ ] Integración con bases de datos
- [ ] Sistema de reportes avanzado
- [ ] Interfaz web (opcional)

---

Para más información, consulta los archivos README.md en cada subdirectorio.
