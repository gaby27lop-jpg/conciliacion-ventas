# Conciliación automática de ventas contra catálogos maestros

> **Proyecto de portafolio con datos simulados.** Los productos, precios, clientes y ventas son inventados y no corresponden a ninguna empresa real. Los errores de captura se introdujeron a propósito para reproducir problemas comunes en datos operativos.

Automatización en Python que consolida reportes de ventas de varias regiones, limpia los datos, los valida contra la lista de precios y el catálogo de clientes, y genera un reporte de excepciones en Excel listo para compartir.

## El problema

En muchas áreas comerciales y de operaciones, la conciliación de ventas se hace a mano:

1. Copiar y pegar los reportes de cada región en un solo archivo.
2. Corregir espacios, mayúsculas y formatos.
3. Cruzar con `BUSCARV` contra la lista de precios y el catálogo de clientes.
4. Revisar fila por fila los precios distintos a lista, los códigos inexistentes y los importes mal capturados.

Es un proceso repetitivo, lento y propenso a errores, que se repite cada día, semana o mes.

## La solución

Un script que ejecuta todo el proceso en segundos y siempre con las mismas reglas:

| Paso | Qué hace |
|---|---|
| **Consolidación** | Lee todos los `.xlsx` de una carpeta, verifica que tengan las columnas esperadas y conserva el archivo de origen de cada registro. |
| **Limpieza** | Normaliza llaves (espacios, mayúsculas), convierte tipos de dato y elimina duplicados exactos. |
| **Validación** | Aplica 6 reglas de calidad a cada registro y documenta el motivo de cada excepción. |
| **Reporte** | Genera un Excel con resumen de KPIs, excepciones por motivo, calidad por archivo de origen y ventas validadas. |

### Reglas de validación

- Dato faltante en cualquier campo obligatorio
- SKU que no existe en la lista de precios
- Cliente que no existe en el catálogo
- Cantidad no válida (cero o negativa)
- Precio unitario fuera de lista (tolerancia configurable, 1% por defecto)
- Importe capturado que no cuadra con cantidad × precio

Un registro puede tener varias excepciones a la vez; todas quedan documentadas.

## Resultados con los datos de ejemplo

| Indicador | Valor |
|---|---|
| Registros leídos (3 regiones) | 1,224 |
| Duplicados eliminados | 22 |
| Registros validados | 1,071 |
| Registros con excepción | 131 |
| Calidad de datos | 89.1% |
| Venta en excepción | $799,702 |

Principal hallazgo: **54 registros con precio fuera de lista**, el motivo de excepción más frecuente. Esto apunta a descuentos no autorizados o errores de captura de precio que impactan directamente el margen.

## Reporte generado

El archivo `reportes/reporte_conciliacion_AAAAMMDD_HHMM.xlsx` contiene tres hojas:

- **Resumen:** indicadores generales, excepciones por motivo y porcentaje de excepción por archivo de origen.
- **Excepciones:** cada registro con problema, el precio de lista, la diferencia de precio, el importe recalculado y el motivo. Con filtros activados.
- **Ventas_Validadas:** registros limpios, enriquecidos con producto, categoría, cliente y canal, listos para cargar a un dashboard en Power BI o Tableau.

## Cómo ejecutarlo

```bash
pip install -r requirements.txt

# 1. Generar los datos de ejemplo (solo la primera vez)
python generar_datos.py

# 2. Ejecutar la conciliación
python conciliar_ventas.py
```

Parámetros opcionales:

```bash
python conciliar_ventas.py --ventas ruta/a/ventas --catalogos ruta/a/catalogos --salida ruta/a/reportes --tolerancia 0.02
```

Para usarlo con datos reales basta con colocar los archivos de ventas en la carpeta `data/ventas/` con las columnas `Folio, Fecha, ID_Cliente, SKU, Cantidad, Precio_Unitario, Importe`, y los catálogos en `data/catalogos/`.

## Estructura del proyecto

```
conciliacion-ventas/
├── conciliar_ventas.py      # Script principal de automatización
├── generar_datos.py         # Generador de datos de ejemplo con errores típicos
├── requirements.txt
├── data/
│   ├── catalogos/           # Lista de precios y catálogo de clientes
│   └── ventas/              # Reportes de ventas por región
└── reportes/                # Reportes de conciliación generados
```

## Habilidades aplicadas

Python (pandas, openpyxl) · Automatización de procesos · Limpieza y validación de datos · Conciliación de información de múltiples fuentes · Calidad de datos · Reglas de negocio · Reportes automatizados en Excel · Documentación de soluciones

## Nota

Los datos son simulados y se generan con semilla fija para que los resultados sean reproducibles. Los errores se introdujeron a propósito para replicar problemas reales de captura.
