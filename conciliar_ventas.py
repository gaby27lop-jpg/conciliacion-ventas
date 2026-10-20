"""
Conciliación automática de reportes de ventas contra catálogos maestros.

Qué hace:
1. Consolida todos los archivos de ventas (.xlsx) de una carpeta en una sola tabla.
2. Limpia los datos: espacios, mayúsculas, tipos de dato y duplicados.
3. Valida cada registro contra la lista de precios y el catálogo de clientes.
4. Genera un reporte en Excel con resumen de KPIs, excepciones y ventas validadas.

Uso:
    python conciliar_ventas.py
    python conciliar_ventas.py --ventas data/ventas --catalogos data/catalogos --salida reportes
    python conciliar_ventas.py --tolerancia 0.02

Reemplaza un proceso manual de copiar, pegar y comparar con BUSCARV entre varios archivos.
"""

import argparse
import logging
from datetime import datetime
from pathlib import Path

import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")
log = logging.getLogger("conciliacion")

COLUMNAS_REQUERIDAS = ["Folio", "Fecha", "ID_Cliente", "SKU", "Cantidad", "Precio_Unitario", "Importe"]


# ---------------------------------------------------------------- 1. Carga
def consolidar_ventas(carpeta: Path) -> pd.DataFrame:
    """Lee todos los .xlsx de la carpeta y los une, conservando el archivo de origen."""
    archivos = sorted(carpeta.glob("*.xlsx"))
    if not archivos:
        raise FileNotFoundError(f"No se encontraron archivos .xlsx en {carpeta}")

    tablas = []
    for archivo in archivos:
        df = pd.read_excel(archivo)
        faltantes = set(COLUMNAS_REQUERIDAS) - set(df.columns)
        if faltantes:
            log.warning("Se omite %s: faltan columnas %s", archivo.name, sorted(faltantes))
            continue
        df["Archivo_Origen"] = archivo.name
        tablas.append(df)
        log.info("Leído %s (%d registros)", archivo.name, len(df))

    if not tablas:
        raise ValueError("Ningún archivo tiene la estructura esperada.")
    return pd.concat(tablas, ignore_index=True)


# ---------------------------------------------------------------- 2. Limpieza
def limpiar(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Normaliza llaves y tipos de dato; elimina duplicados exactos."""
    df = df.copy()
    for col in ["Folio", "ID_Cliente", "SKU"]:
        df[col] = df[col].astype("string").str.strip().str.upper()

    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    for col in ["Cantidad", "Precio_Unitario", "Importe"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    antes = len(df)
    df = df.drop_duplicates(subset=COLUMNAS_REQUERIDAS).reset_index(drop=True)
    return df, antes - len(df)


# ---------------------------------------------------------------- 3. Validación
def validar(df: pd.DataFrame, precios: pd.DataFrame, clientes: pd.DataFrame,
            tolerancia: float) -> pd.DataFrame:
    """Agrega una columna por regla de validación y una columna con el motivo de excepción."""
    df = df.merge(precios[["SKU", "Producto", "Categoria", "Precio_Lista"]], on="SKU", how="left")
    df = df.merge(clientes[["ID_Cliente", "Nombre_Cliente", "Canal"]], on="ID_Cliente", how="left")

    df["Importe_Calculado"] = (df["Cantidad"] * df["Precio_Unitario"]).round(2)
    df["Dif_Precio_%"] = ((df["Precio_Unitario"] / df["Precio_Lista"]) - 1).round(4)

    reglas = {
        "Dato faltante": df[COLUMNAS_REQUERIDAS].isna().any(axis=1),
        "SKU no existe en catálogo": df["SKU"].notna() & df["Precio_Lista"].isna(),
        "Cliente no existe en catálogo": df["ID_Cliente"].notna() & df["Nombre_Cliente"].isna(),
        "Cantidad no válida (<= 0)": df["Cantidad"] <= 0,
        "Precio fuera de lista": df["Dif_Precio_%"].abs() > tolerancia,
        "Importe no cuadra": (df["Importe"] - df["Importe_Calculado"]).abs() > 0.01,
    }

    motivos = pd.Series([[] for _ in range(len(df))], index=df.index)
    for nombre, mascara in reglas.items():
        mascara = mascara.fillna(False)
        motivos[mascara] = motivos[mascara].apply(lambda m, n=nombre: m + [n])

    df["Motivo_Excepcion"] = motivos.apply(lambda m: "; ".join(m))
    df["Estatus"] = df["Motivo_Excepcion"].map(lambda m: "Excepción" if m else "Validado")
    return df


# ---------------------------------------------------------------- 4. Reporte
def construir_resumen(df: pd.DataFrame, duplicados: int) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    validos = df[df["Estatus"] == "Validado"]
    excepciones = df[df["Estatus"] == "Excepción"]

    kpis = pd.DataFrame({
        "Indicador": [
            "Registros leídos",
            "Duplicados eliminados",
            "Registros validados",
            "Registros con excepción",
            "% de calidad de datos",
            "Venta validada ($)",
            "Venta en excepción ($)",
        ],
        "Valor": [
            len(df) + duplicados,
            duplicados,
            len(validos),
            len(excepciones),
            round(len(validos) / len(df) * 100, 1),
            round(validos["Importe"].sum(), 2),
            round(excepciones["Importe"].sum(skipna=True), 2),
        ],
    })

    por_motivo = (
        excepciones["Motivo_Excepcion"].str.split("; ").explode()
        .value_counts().rename_axis("Motivo").reset_index(name="Registros")
    )

    por_origen = (
        df.groupby(["Archivo_Origen", "Estatus"]).size().unstack(fill_value=0).reset_index()
    )
    if "Excepción" in por_origen and "Validado" in por_origen:
        total = por_origen["Excepción"] + por_origen["Validado"]
        por_origen["% Excepción"] = (por_origen["Excepción"] / total * 100).round(1)

    return kpis, por_motivo, por_origen


def exportar(df: pd.DataFrame, kpis: pd.DataFrame, por_motivo: pd.DataFrame,
             por_origen: pd.DataFrame, carpeta_salida: Path) -> Path:
    carpeta_salida.mkdir(parents=True, exist_ok=True)
    archivo = carpeta_salida / f"reporte_conciliacion_{datetime.now():%Y%m%d_%H%M}.xlsx"

    columnas_salida = [
        "Archivo_Origen", "Folio", "Fecha", "ID_Cliente", "Nombre_Cliente", "Canal",
        "SKU", "Producto", "Categoria", "Cantidad", "Precio_Unitario", "Precio_Lista",
        "Dif_Precio_%", "Importe", "Importe_Calculado", "Estatus", "Motivo_Excepcion",
    ]

    with pd.ExcelWriter(archivo, engine="openpyxl") as writer:
        kpis.to_excel(writer, sheet_name="Resumen", index=False, startrow=1)
        por_motivo.to_excel(writer, sheet_name="Resumen", index=False, startrow=len(kpis) + 4)
        por_origen.to_excel(writer, sheet_name="Resumen", index=False,
                            startrow=len(kpis) + len(por_motivo) + 7)
        hoja = writer.sheets["Resumen"]
        hoja["A1"] = "Indicadores generales"
        hoja["D1"] = "Nota: datos simulados con fines de demostración; no corresponden a ninguna empresa real."
        hoja.cell(row=len(kpis) + 4, column=1, value="Excepciones por motivo")
        hoja.cell(row=len(kpis) + len(por_motivo) + 7, column=1, value="Calidad por archivo de origen")

        df.loc[df["Estatus"] == "Excepción", columnas_salida].to_excel(
            writer, sheet_name="Excepciones", index=False)
        df.loc[df["Estatus"] == "Validado", columnas_salida].to_excel(
            writer, sheet_name="Ventas_Validadas", index=False)

        for ws in writer.sheets.values():
            ws.freeze_panes = "A2" if ws.title != "Resumen" else None
            for columna in ws.columns:
                ancho = max(len(str(c.value)) if c.value is not None else 0 for c in columna)
                ws.column_dimensions[columna[0].column_letter].width = min(ancho + 2, 45)
            if ws.title != "Resumen":
                ws.auto_filter.ref = ws.dimensions

    return archivo


# ---------------------------------------------------------------- Ejecución
def main() -> None:
    base = Path(__file__).parent
    parser = argparse.ArgumentParser(description="Concilia reportes de ventas contra catálogos maestros.")
    parser.add_argument("--ventas", type=Path, default=base / "data" / "ventas")
    parser.add_argument("--catalogos", type=Path, default=base / "data" / "catalogos")
    parser.add_argument("--salida", type=Path, default=base / "reportes")
    parser.add_argument("--tolerancia", type=float, default=0.01,
                        help="Diferencia de precio permitida vs. lista (0.01 = 1%%)")
    args = parser.parse_args()

    ventas = consolidar_ventas(args.ventas)
    precios = pd.read_excel(args.catalogos / "lista_precios.xlsx")
    clientes = pd.read_excel(args.catalogos / "clientes.xlsx")

    ventas, duplicados = limpiar(ventas)
    log.info("Duplicados eliminados: %d", duplicados)

    resultado = validar(ventas, precios, clientes, args.tolerancia)
    kpis, por_motivo, por_origen = construir_resumen(resultado, duplicados)
    archivo = exportar(resultado, kpis, por_motivo, por_origen, args.salida)

    log.info("Reporte generado: %s", archivo)
    print("\n" + kpis.to_string(index=False))


if __name__ == "__main__":
    main()
