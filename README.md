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
- **Excel Final con Mejoras**: Sistema completo de personalización de viviendas
  - Hoja Resumen con todas las viviendas
  - Hoja Resumen Constructora para pedidos
  - 91 hojas individuales con tabla de conceptos
  - Desplegables de validación de datos
  - Fórmulas automáticas de cálculo
  - Validaciones según tipo de vivienda
- **Catálogo de Mejoras**: 11 categorías con +50 opciones del PDF oficial
- **Entorno Dockerizado**: Ejecución aislada y reproducible
- **Interfaz de Menú**: Sistema de menús interactivo con `menu.sh`

## 🛠️ Estructura del Proyecto

```
STARK096---ALDEA/
├── menu.sh                      # Menú principal (source menu.sh)
├── Dockerfile                   # Definición del contenedor Python
├── README.md                    # Este archivo
├── dosier/                      # Documentos fuente
│   └── 6_LISTADO MEJORAS Neptuno C63.pdf
└── excelGenerator/              # Código Python
    ├── requirements.txt         # Dependencias (openpyxl, pypdf2, tqdm)
    ├── excel_template_generator.py   # Generador de plantillas
    ├── generate_final_sheets.py      # Generador de Excel final ⭐
    ├── read_pdf_mejoras.py           # Lector de catálogo PDF
    ├── utils.py                 # Utilidades comunes
    ├── mejoras_catalog.json     # Catálogo estructurado de mejoras
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
Excel con todas las viviendas y co ⭐ COMPLETO

**Hoja "Resumen":**
- Listado de todas las viviendas
- Columnas: Comprador 1 (3 cols), Comprador 2 (3 cols), Piso, Total Gastado, Max Módulo, Sobrante
- Fórmulas automáticas que referencian las hojas individuales
- Cálculo de Sobrante = Max Módulo - Total Gastado

**Hoja "Resumen Constructora":**
- Listado de obra para pedidos a constructora
- Columnas: Código, Concepto, Cantidad Total
- Agregación de cantidades de todas las viviendas

**91 Hojas Individuales (una por vivienda):**
Cada hoja incluye:

1. **Header "FICHA DE CLIENTE"**:
   - Nombre y apellidos del/los comprador(es)
   - Identificación de vivienda (formato E1 4A)
   - Total Gastado con fórmula automática

2. **Tabla de Conceptos**:
   - Concepto | Selección | Precio (€) | Cantidad | Seleccionado (Sí/No) | Importe (€)
   - **Desplegables de validación** en columna Selección
   - **Validaciones condicionales**: Opciones de cocina según tipo de vivienda
   - **Fórmulas automáticas**: Importe = IF(Seleccionado="Sí", Precio × Cantidad, 0)
   - Fila TOTAL con suma de todos los importes

3. **Conceptos Incluidos**:
   - Pasos básicos: Carpintería Interior, Suelo Laminado, Baño General, Baño Suite, Cocina
   - Mejoras: Armarios empotrados, Muebles baño, Mamparas, Griferías, Iluminación, Suelos, Ventanas, Refuerzos, Cierre cocina
   - Total: ~12 conceptos configurables por vivienda

## 🎯 Catálogo de Mejoras

El sistema lee automáticamente el catálogo desde `dosier/6_LISTADO MEJORAS Neptuno C63.pdf`:

- **11 categorías** de mejoras
- **+50 conceptos** con precios
- **Validaciones por tipo de vivienda**: Las opciones de cocina se adaptan automáticamente según el tipo (1AD, 2AD, 3AE, 1BC, 2BC, 3BD, 3C)
- **Precios actualizados** desde el PDF oficial (rangos desde 0€ hasta 7.889€)mbre = columna "Piso" del original
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