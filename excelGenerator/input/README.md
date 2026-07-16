# Directorio de Datos de Entrada

Coloca aquí los archivos de datos que serán procesados para generar los ficheros Excel.

## Formatos Soportados

- CSV (`.csv`)
- JSON (`.json`)
- YAML (`.yaml`, `.yml`)
- Excel (`.xlsx`, `.xls`)

## Estructura Esperada

Los archivos de entrada deben contener los siguientes campos (ejemplo para viviendas):

```yaml
viviendas:
  - id: VIV-001
    bloque: A
    piso: 1
    puerta: A
    tipo: 2D
    superficie: 85.5
    precio: 185000
    estado: Disponible
    fecha_entrega: 2026-12-31
```

## Ejemplos

Consulta los archivos de ejemplo en `templates/` para ver la estructura de datos esperada.
