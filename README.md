# STARK096 — ALDEA 🏘️

**Generador de Ficheros Excel Parametrizados para Constructora**

Sistema de generación automatizada de documentos Excel con 91 viviendas y compradores con nombres de perritos.

---

## 📋 Descripción

STARK096—ALDEA genera plantillas Excel con 91 viviendas distribuidas en 3 escaleras, con compradores asignados que tienen nombres graciosos basados en nombres de perritos.

## 🚀 Características

- **91 Viviendas**: Escalera 1 (28), Escalera 2 (28), Escalera 3 (35)
- **Nombres Graciosos**: Compradores con nombres de perritos (Firulais Ladrón Guardián, etc.)
- **2 Compradores**: ~40 viviendas tienen 2 compradores
- **Excel Final**: Genera un Excel con una pestaña por vivienda
- **Entorno Dockerizado**: Ejecución aislada y reproducible
- **Interfaz de Menú**: Sistema de menús interactivo con `menu.sh`

## 🛠️ Estructura del Proyecto

```
STARK096---ALDEA/
├── menu.sh                      # Menú principal (source menu.sh)
├── Dockerfile                   # Definición del contenedor Python
├── README.md                    # Este archivo
└── excelGenerator/              # Código Python
    ├── requirements.txt         # Dependencias Python
    ├── excel_template_generator.py   # Generador de plantillas
    ├── generate_final_sheets.py      # Generador de Excel final
    ├── utils.py                 # Utilidades comunes
    ├── output/                  # Plantillas base generadas
    └── output_final/            # Excel final con pestañas
```

## 📦 Instalación y Uso

### Requisitos Previos

- Docker instalado y en ejecución
- Git Bash o terminal compatible con bash

### Inicio Rápido

1. **Cargar el menú**
   ```bash
   cd "C:\Users\xete\Documents\LIFE_7.0 - TITAN\GitHub\STARK096---ALDEA"
   source menu.sh
   ```

2. **Construir la imagen Docker** (primera vez)
   - Selecciona: 1 (Container Management) → 1 (Build Image)

3. **Generar plantilla Excel**
   - Generar Excel final con pestañas**
   - Selecciona: 2 (Excel Generator) → 2 (Generate Final Sheets)

5. **Los ficheros se generan en**:
   - Plantilla base: `excelGenerator/output/`
   - Excel final: `excelGenerator/output_finalel Template)

4. **El fichero se genera en**: `excelGenerator/output/`

## 📊 Distribución de Viv

## 📋 Ficheros Generados

### 1. Plantilla Base (output/)
Excel con todas las viviendas y compradores en una sola hoja:
- Columnas: Identificador, Piso, Comprador 1 (Nombre, Apellido1, Apellido2), Comprador 2 (Nombre, Apellido1, Apellido2)
- 91 filas de datos + 1 cabecera

### 2. Excel Final (output_final/)
Excel con pestañas individuales:
- **Pestaña "Resumen"**: Estadísticas generales del proyecto
- **91 Pestañas**: Una por vivienda, nombre = columna "Piso" del original
- **Contenido de cada pestaña**:
  - Cabecera con nombre(s) del/los comprador(es)
  - Identificación de vivienda
  - Campo "Total Gastado: ____ € + IVA"iendas

- **Escalera 1**: Plantas 1-7, Puertas A-D (28 viviendas)
- **Escalera 2**: Plantas 1-7, Puertas A-D (28 viviendas)  
- **Escalera 3**: Plantas 1-7, Puertas A-E (35 viviendas)
- **Total**: 91 viviendas

## 🐕 Nombres de Compradores

Los compradores tienen nombres graciosos combinando:
- **Nombres**: Firulais, Scooby, Pluto, Luna, Rocky, Bella, etc.
- **Apellido 1**: Ladrón, Mordedor, Saltarín, Gruñón, Olfateador, etc.
- **Apellido 2**: Guardián, Valiente, Peludo, Leal, Nervioso, etc.

Ejemplo: "Firulais Ladrón Guardián", "Scooby Escarbabasura Nervioso"

## 🔧 Dependencias Python

- `openpyxl`: Manipulación de archivos Excel
- `tqdm`: Barras de progreso

## 📝 Notas de Desarrollo

- El sistema usa `source menu.sh` para mantener el entorno en la sesión actual
- Los scripts Python se ejecutan dentro del contenedor Docker
- Los ficheros generados se almacenan en `excelGenerator/output/`

---

**STARK096 — ALDEA** | Excel Generator for Construction Management