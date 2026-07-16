# STARK096 — ALDEA 🏘️

**Generador de Ficheros Excel Parametrizados para Constructora**

Sistema de generación automatizada de documentos Excel para la gestión de proyectos de construcción de viviendas.

---

## 📋 Descripción

STARK096—ALDEA es una herramienta diseñada para automatizar la creación y gestión de ficheros Excel parametrizados para empresas constructoras. El sistema permite generar plantillas, procesar datos en lote, y validar los documentos generados, todo desde un contenedor Docker con Python.

## 🚀 Características

- **Generación de Plantillas Excel**: Crea plantillas parametrizadas con formato profesional
- **Procesamiento en Lote**: Genera múltiples ficheros desde datos de entrada
- **Gestión de Configuración**: Parámetros personalizables para cada proyecto
- **Validación Automática**: Verifica la integridad de los ficheros generados
- **Entorno Dockerizado**: Ejecución aislada y reproducible
- **Interfaz de Menú**: Sistema de menús interactivo con `menu.sh`

## 🛠️ Estructura del Proyecto

```
STARK096---ALDEA/
├── menu.sh                      # Menú principal (source menu.sh)
├── Dockerfile                   # Definición del contenedor Python
├── .gitignore                   # Archivos ignorados por Git
├── README.md                    # Este archivo
├── Dosieres/                    # Documentación del proyecto
└── excelGenerator/              # Código Python
    ├── requirements.txt         # Dependencias Python
    ├── excel_template_generator.py   # Generador de plantillas
    ├── batch_processor.py       # Procesador de lotes
    ├── config_manager.py        # Gestor de configuración
    ├── file_validator.py        # Validador de ficheros
    ├── config/                  # Archivos de configuración
    ├── input/                   # Datos de entrada
    ├── output/                  # Ficheros generados
    └── templates/               # Plantillas base
```

## 📦 Instalación y Uso

### Requisitos Previos

- Docker instalado y en ejecución
- Git Bash o terminal compatible con bash
- Windows / Linux / macOS

### Inicio Rápido

1. **Clonar el repositorio** (si es necesario)
   ```bash
   cd "C:\Users\xete\Documents\LIFE_7.0 - TITAN\GitHub\STARK096---ALDEA"
   ```

2. **Cargar el menú**
   ```bash
   source menu.sh
   ```

3. **Construir la imagen Docker**
   - Selecciona opción `1` (Container Management)
   - Luego opción `1` (Build Image)

4. **Usar el generador**
   - Vuelve al menú principal
   - Selecciona opción `2` (Excel Generator)
   - Elige la operación deseada

## 📊 Funcionalidades

### 1. Container Management
- **Build Image**: Construye la imagen Docker con todas las dependencias
- **Start Container**: Inicia el contenedor de desarrollo
- **Enter Container**: Accede a la shell del contenedor
- **Stop/Remove Container**: Detiene y elimina el contenedor
- **Container Status**: Muestra el estado del contenedor e imagen

### 2. Excel Generator
- **Generate Excel Template**: Crea plantillas base parametrizadas
- **Process Batch Data**: Procesa múltiples registros y genera ficheros
- **Configure Parameters**: Gestiona la configuración del sistema
- **Validate Generated Files**: Verifica la integridad de los ficheros Excel

## 🔧 Dependencias Python

- `openpyxl`: Manipulación de archivos Excel (lectura/escritura)
- `pandas`: Procesamiento de datos tabulares
- `xlsxwriter`: Creación avanzada de Excel
- `numpy`: Operaciones numéricas
- `pyyaml`: Gestión de configuración
- `tqdm`: Barras de progreso

## 🎨 Próximas Funcionalidades

Las siguientes opciones se implementarán según requisitos específicos:

- [ ] Generación de informes personalizados
- [ ] Integración con bases de datos
- [ ] Exportación a múltiples formatos
- [ ] Dashboard de visualización
- [ ] API REST para integración
- [ ] Sistema de plantillas avanzado

## 📝 Notas de Desarrollo

- El sistema usa `source menu.sh` para mantener el entorno en la sesión actual
- Los scripts Python se ejecutan dentro del contenedor Docker
- Los ficheros generados se almacenan en `excelGenerator/output/`
- La configuración se gestiona mediante archivos YAML

## 🤝 Contribución

Este es un proyecto personalizado para gestión de construcción de viviendas. Las características se irán añadiendo según necesidades específicas del proyecto.

## 📄 Licencia

Proyecto privado - STARK096

---

**STARK096 — ALDEA** | Excel Generator for Construction Management