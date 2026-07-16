# Estructura del Proyecto STARK096—ALDEA

```
STARK096---ALDEA/
│
├── 📄 README.md                         # Documentación principal del proyecto
├── 📄 QUICKSTART.md                     # Guía de inicio rápido
├── 📄 .gitignore                        # Archivos ignorados por Git
├── 🐳 Dockerfile                        # Definición del contenedor Docker
├── 🔧 menu.sh                           # Menú interactivo principal (source menu.sh)
│
├── 📁 Dosieres/                         # Documentación del proyecto
│   ├── PROYECTO.md                      # Visión general y arquitectura
│   └── DESARROLLO.md                    # Guía completa de desarrollo
│
└── 📁 excelGenerator/                   # Módulo Python principal
    │
    ├── 📄 README.md                     # Documentación del módulo
    ├── 📄 requirements.txt              # Dependencias Python
    ├── 📄 __init__.py                   # Inicialización del paquete
    ├── 🔧 utils.py                      # Utilidades comunes
    │
    ├── 📊 Scripts principales
    ├── excel_template_generator.py      # Generador de plantillas Excel
    ├── batch_processor.py               # Procesador de datos por lotes
    ├── config_manager.py                # Gestor de configuración
    └── file_validator.py                # Validador de ficheros Excel
    │
    ├── 📁 config/                       # Configuración del sistema
    │   ├── README.md                    # Documentación de configuración
    │   └── settings.yaml                # Configuración principal
    │
    ├── 📁 input/                        # Datos de entrada
    │   ├── README.md                    # Formatos soportados
    │   ├── example_data.yaml            # Ejemplo de datos YAML
    │   └── example_data.csv             # Ejemplo de datos CSV
    │
    ├── 📁 output/                       # Ficheros Excel generados
    │   └── README.md                    # Estructura de salida
    │
    └── 📁 templates/                    # Plantillas base
        └── README.md                    # Tipos de plantillas
```

---

## Descripción de Componentes

### 🔧 Archivos de Configuración

| Archivo | Descripción |
|---------|-------------|
| `menu.sh` | Menú interactivo bash para gestión del proyecto |
| `Dockerfile` | Configuración del contenedor Python 3.11 |
| `.gitignore` | Exclusión de archivos temporales y output |
| `requirements.txt` | Dependencias Python (openpyxl, pandas, etc.) |

### 📊 Scripts Python

| Script | Propósito |
|--------|-----------|
| `excel_template_generator.py` | Genera plantillas Excel con formato parametrizado |
| `batch_processor.py` | Procesa múltiples ficheros de entrada |
| `config_manager.py` | Gestiona parámetros del sistema |
| `file_validator.py` | Valida integridad de Excel generados |
| `utils.py` | Funciones comunes (estilos, validación, mensajes) |

### 📁 Directorios

| Directorio | Contenido |
|------------|-----------|
| `Dosieres/` | Documentación técnica y de desarrollo |
| `config/` | Archivos de configuración YAML |
| `input/` | Datos de entrada (CSV, YAML, JSON) |
| `output/` | Ficheros Excel generados (no versionado) |
| `templates/` | Plantillas base reutilizables |

---

## Flujo de Trabajo

```
┌────────────────────────────────────────────────────────────┐
│                    Usuario ejecuta:                        │
│                   source menu.sh                           │
└────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌────────────────────────────────────────────────────────────┐
│                    Menú Principal                          │
│  1. Container Management                                   │
│  2. Excel Generator                                        │
│  0. Exit                                                   │
└────────────────────────────────────────────────────────────┘
                            │
        ┌───────────────────┴───────────────────┐
        │                                       │
        ▼                                       ▼
┌──────────────────┐                  ┌──────────────────┐
│   Container      │                  │  Excel Generator │
│   Management     │                  │      Menu        │
└──────────────────┘                  └──────────────────┘
        │                                       │
        │ Build/Start/Stop                     │ Generate/Process
        │ Docker Container                     │ Excel Files
        │                                       │
        ▼                                       ▼
┌────────────────────────────────────────────────────────────┐
│          Docker Container: stark096_aldea_dev              │
│  ┌──────────────────────────────────────────────────────┐ │
│  │           Python 3.11 Environment                    │ │
│  │  ┌────────────────────────────────────────────────┐  │ │
│  │  │  excelGenerator/                               │  │ │
│  │  │  - Ejecuta scripts Python                      │  │ │
│  │  │  - Lee config/settings.yaml                    │  │ │
│  │  │  - Lee datos de input/                         │  │ │
│  │  │  - Genera Excel en output/                     │  │ │
│  │  └────────────────────────────────────────────────┘  │ │
│  └──────────────────────────────────────────────────────┘ │
└────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌────────────────────────────────────────────────────────────┐
│            excelGenerator/output/                          │
│  template_viviendas_20260716_123456.xlsx                   │
│  reporte_batch_20260716_124500.xlsx                        │
│  ...                                                       │
└────────────────────────────────────────────────────────────┘
```

---

## Comandos Principales

### Inicio del Sistema
```bash
cd "C:\Users\xete\Documents\LIFE_7.0 - TITAN\GitHub\STARK096---ALDEA"
source menu.sh
```

### Construcción de Imagen
```bash
# En el menú: 1 → 1
# Construye la imagen Docker con todas las dependencias
```

### Generación de Plantilla
```bash
# En el menú: 2 → 1
# Genera: output/template_viviendas_YYYYMMDD_HHMMSS.xlsx
```

### Procesamiento por Lotes
```bash
# En el menú: 2 → 2
# Procesa ficheros de input/ y genera Excel en output/
```

### Validación
```bash
# En el menú: 2 → 4
# Valida todos los ficheros .xlsx en output/
```

---

## Archivos Generados

### Durante la Ejecución

```
output/
├── template_viviendas_20260716_123456.xlsx    # Plantillas generadas
├── reporte_batch_20260716_124500.xlsx         # Reportes procesados
└── batch_output_20260716_130000/              # Carpetas de lotes
    ├── viviendas_bloque_A.xlsx
    ├── viviendas_bloque_B.xlsx
    └── resumen.xlsx
```

### Configuración Persistente

```
config/
├── settings.yaml           # Configuración principal (versionado)
├── secrets.yaml            # Credenciales (NO versionado)
└── parameters.yaml         # Parámetros custom (NO versionado)
```

---

## Tecnologías Utilizadas

| Tecnología | Versión | Uso |
|------------|---------|-----|
| Python | 3.11 | Lenguaje principal |
| Docker | Latest | Contenedorización |
| Bash | 4.x+ | Sistema de menús |
| openpyxl | 3.1.2 | Manipulación de Excel |
| pandas | 2.2.0 | Procesamiento de datos |
| PyYAML | 6.0.1 | Configuración |

---

## Estado de Implementación

| Funcionalidad | Estado | Notas |
|---------------|--------|-------|
| ✅ Estructura base | Completo | Menu + Docker + Scripts |
| ✅ Generador de plantillas | Completo | Plantilla básica funcional |
| ⏳ Procesamiento por lotes | Pendiente | Estructura creada |
| ⏳ Gestor de configuración | Pendiente | settings.yaml creado |
| ✅ Validador de ficheros | Completo | Validación básica |
| ⏳ Sistema de reportes | Pendiente | Por implementar |
| ⏳ Integración BD | Futuro | Por definir |

Leyenda: ✅ Completo | ⏳ En desarrollo | 🔄 Planificado

---

## Próximos Pasos

1. **Implementar batch_processor.py**
   - Leer múltiples ficheros de input/
   - Generar Excel parametrizados
   - Crear resumen de procesamiento

2. **Ampliar config_manager.py**
   - Interfaz interactiva para configuración
   - Validación de parámetros
   - Gestión de múltiples configuraciones

3. **Añadir más tipos de plantillas**
   - Plantilla de resumen ejecutivo
   - Plantilla de comparativa de precios
   - Plantilla de estado de obra

4. **Sistema de reportes**
   - Gráficos y visualizaciones
   - Exportación a PDF
   - Dashboard HTML

---

## Recursos

### Documentación
- [README.md](README.md) - Documentación principal
- [QUICKSTART.md](QUICKSTART.md) - Inicio rápido
- [Dosieres/PROYECTO.md](Dosieres/PROYECTO.md) - Arquitectura
- [Dosieres/DESARROLLO.md](Dosieres/DESARROLLO.md) - Guía desarrollo
- [excelGenerator/README.md](excelGenerator/README.md) - Módulo Python

### Enlaces Útiles
- [Documentación openpyxl](https://openpyxl.readthedocs.io/)
- [Documentación pandas](https://pandas.pydata.org/)
- [Docker Documentation](https://docs.docker.com/)

---

**STARK096 — ALDEA** | Generador de Excel para Constructoras
