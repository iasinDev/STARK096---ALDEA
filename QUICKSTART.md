# 🚀 Guía de Inicio Rápido - STARK096 ALDEA

## Instalación Inicial

### 1. Verificar Docker
```bash
docker --version
```
Si no tienes Docker, instálalo desde: https://www.docker.com/get-started

### 2. Navegar al Proyecto
```bash
cd "C:\Users\xete\Documents\LIFE_7.0 - TITAN\GitHub\STARK096---ALDEA"
```

### 3. Cargar el Menú
```bash
source menu.sh
```

### 4. Construir la Imagen Docker
En el menú:
- Opción `1` (Container Management)
- Opción `1` (Build Image)
- Espera unos minutos mientras se construye

## Uso Básico

### Generar una Plantilla Excel
```bash
source menu.sh
# Opción 2: Excel Generator
# Opción 1: Generate Excel Template
```

El fichero se generará en: `excelGenerator/output/`

### Ver el Fichero Generado
```bash
# En Windows
start excelGenerator/output/template_viviendas_*.xlsx

# En Linux/Mac
xdg-open excelGenerator/output/template_viviendas_*.xlsx
```

## Comandos Útiles

### Ver Estado del Contenedor
```bash
source menu.sh
# Opción 1: Container Management
# Opción 5: Container Status
```

### Acceder al Contenedor (Shell)
```bash
source menu.sh
# Opción 1: Container Management
# Opción 3: Enter Container Shell
```

Dentro del contenedor puedes:
```bash
# Ver archivos
ls -la

# Ejecutar scripts manualmente
python3 excel_template_generator.py

# Instalar dependencias adicionales
pip install nombre_paquete

# Salir
exit
```

### Detener el Contenedor
```bash
source menu.sh
# Opción 1: Container Management
# Opción 4: Stop / Remove Container
```

## Estructura de Archivos

```
STARK096---ALDEA/
├── menu.sh                  ← Carga esto con "source menu.sh"
├── Dockerfile               ← Configuración del contenedor
├── excelGenerator/
│   ├── config/              ← Configuraciones (settings.yaml)
│   ├── input/               ← Coloca aquí tus datos
│   ├── output/              ← Aquí se generan los Excel
│   ├── templates/           ← Plantillas base
│   └── *.py                 ← Scripts Python
```

## Flujo de Trabajo Típico

1. **Configurar** (primera vez)
   ```bash
   source menu.sh
   # Opción 2 → Opción 3 (Configure Parameters)
   ```

2. **Preparar Datos**
   - Coloca archivos CSV/JSON en `excelGenerator/input/`

3. **Procesar**
   ```bash
   source menu.sh
   # Opción 2 → Opción 2 (Process Batch Data)
   ```

4. **Validar**
   ```bash
   source menu.sh
   # Opción 2 → Opción 4 (Validate Generated Files)
   ```

5. **Revisar Resultados**
   - Los ficheros están en `excelGenerator/output/`

## Solución de Problemas

### Error: "Docker no encontrado"
- Instala Docker Desktop
- Asegúrate de que Docker está en ejecución

### Error: "Permission denied"
En Git Bash:
```bash
chmod +x menu.sh
source menu.sh
```

### Error: "Container not running"
El sistema lo iniciará automáticamente, pero si falla:
```bash
source menu.sh
# Opción 1 → Opción 2 (Start Container)
```

### Error al construir la imagen
```bash
# Limpiar todo y reconstruir
docker system prune -a
source menu.sh
# Opción 1 → Opción 1 (Build Image)
```

## Personalización

### Editar Configuración
```bash
# Con tu editor favorito
code excelGenerator/config/settings.yaml
# o
notepad excelGenerator/config/settings.yaml
```

### Añadir Scripts Propios
1. Crea `excelGenerator/mi_script.py`
2. Edita `menu.sh` para añadir la opción
3. Reconstruye la imagen (Opción 1 → 1)

## Más Información

- [README.md](README.md) - Documentación completa
- [Dosieres/PROYECTO.md](Dosieres/PROYECTO.md) - Documentación técnica
- [excelGenerator/config/README.md](excelGenerator/config/README.md) - Configuración

## Ayuda Rápida

```bash
# Recargar el menú (si cierras la terminal)
source menu.sh

# Ver archivos generados
ls -la excelGenerator/output/

# Ver logs (si algo falla)
docker logs stark096_aldea_dev

# Limpiar outputs
rm -rf excelGenerator/output/*.xlsx
```

---

¿Necesitas ayuda? Revisa la documentación en `Dosieres/` o consulta los README.md en cada carpeta.
