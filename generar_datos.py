"""
Genera datos SIMULADOS para el proyecto de conciliación de ventas.

Todos los datos (productos, precios, clientes y ventas) son inventados con fines de
demostración y no corresponden a ninguna empresa real.

Crea:
- data/catalogos/lista_precios.xlsx   -> catálogo maestro de productos y precios
- data/catalogos/clientes.xlsx        -> catálogo maestro de clientes
- data/ventas/ventas_<region>.xlsx    -> reportes de ventas por región, con errores
                                         típicos de captura incluidos a propósito

Los datos son simulados. La semilla fija permite reproducir siempre el mismo resultado.
"""

from pathlib import Path

import numpy as np
import pandas as pd

SEMILLA = 42
BASE = Path(__file__).parent / "data"
REGIONES = ["Norte", "Centro", "Sur"]
REGISTROS_POR_REGION = 400


def crear_catalogos(rng: np.random.Generator) -> tuple[pd.DataFrame, pd.DataFrame]:
    # Catálogo genérico e inventado: no corresponde a ninguna empresa real
    categorias = {
        "Abarrotes": ["Arroz 1 kg", "Frijol 1 kg", "Aceite vegetal 1 L", "Azúcar 1 kg"],
        "Bebidas": ["Agua natural 1.5 L", "Jugo de naranja 1 L", "Café molido 400 g"],
        "Limpieza": ["Detergente en polvo 1 kg", "Limpiador multiusos 1 L", "Jabón de barra 3 pzas"],
        "Cuidado personal": ["Shampoo 750 ml", "Pasta dental 100 ml", "Papel higiénico 4 rollos"],
    }
    productos = []
    n = 1
    for categoria, nombres in categorias.items():
        for nombre in nombres:
            productos.append({
                "SKU": f"SKU-{n:04d}",
                "Producto": nombre,
                "Categoria": categoria,
                "Precio_Lista": round(float(rng.uniform(8, 120)), 2),
            })
            n += 1
    precios = pd.DataFrame(productos)

    cadenas = ["Autoservicio", "Mayorista", "Papelería", "Farmacia", "Conveniencia"]
    clientes = pd.DataFrame({
        "ID_Cliente": [f"CLI-{i:03d}" for i in range(1, 41)],
        "Nombre_Cliente": [f"Cliente {i:03d}" for i in range(1, 41)],
        "Canal": rng.choice(cadenas, size=40),
        "Region": rng.choice(REGIONES, size=40),
    })
    return precios, clientes


def crear_ventas(region: str, precios: pd.DataFrame, clientes: pd.DataFrame,
                 rng: np.random.Generator) -> pd.DataFrame:
    n = REGISTROS_POR_REGION
    clientes_region = clientes.loc[clientes["Region"] == region, "ID_Cliente"].to_numpy()
    idx_prod = rng.integers(0, len(precios), size=n)

    df = pd.DataFrame({
        "Folio": [f"{region[:1]}-{i:05d}" for i in range(1, n + 1)],
        "Fecha": pd.to_datetime("2026-07-01") + pd.to_timedelta(rng.integers(0, 92, size=n), unit="D"),
        "ID_Cliente": rng.choice(clientes_region, size=n),
        "SKU": precios["SKU"].to_numpy()[idx_prod],
        "Cantidad": rng.integers(1, 200, size=n),
        "Precio_Unitario": precios["Precio_Lista"].to_numpy()[idx_prod],
    })

    # --- Errores típicos de captura, introducidos a propósito ---
    def filas(k):
        return rng.choice(df.index, size=k, replace=False)

    # 1) Espacios y minúsculas en llaves
    f = filas(15)
    df.loc[f, "SKU"] = " " + df.loc[f, "SKU"].str.lower() + " "
    # 2) Precio distinto a lista (descuentos no autorizados o errores)
    f = filas(18)
    df.loc[f, "Precio_Unitario"] = (df.loc[f, "Precio_Unitario"] * rng.uniform(0.7, 0.95, size=len(f))).round(2)
    # 3) SKU que no existe en el catálogo
    f = filas(6)
    df.loc[f, "SKU"] = "SKU-9999"
    # 4) Cliente que no existe en el catálogo
    f = filas(5)
    df.loc[f, "ID_Cliente"] = "CLI-999"
    # 5) Cantidades no válidas
    f = filas(4)
    df.loc[f, "Cantidad"] = [0, -5, 0, -1]
    # 6) Valores faltantes
    df.loc[filas(4), "Precio_Unitario"] = np.nan
    df.loc[filas(3), "ID_Cliente"] = np.nan
    # 7) Registros duplicados (mismo folio capturado dos veces)
    duplicados = df.loc[filas(8)]
    df = pd.concat([df, duplicados], ignore_index=True)

    df["Importe"] = (df["Cantidad"] * df["Precio_Unitario"]).round(2)
    # 8) Importes capturados a mano con error
    f = filas(5)
    df.loc[f, "Importe"] = (df.loc[f, "Importe"] * 1.1).round(2)

    return df.sample(frac=1, random_state=SEMILLA).reset_index(drop=True)


def main() -> None:
    rng = np.random.default_rng(SEMILLA)
    (BASE / "catalogos").mkdir(parents=True, exist_ok=True)
    (BASE / "ventas").mkdir(parents=True, exist_ok=True)

    precios, clientes = crear_catalogos(rng)
    precios.to_excel(BASE / "catalogos" / "lista_precios.xlsx", index=False)
    clientes.to_excel(BASE / "catalogos" / "clientes.xlsx", index=False)

    for region in REGIONES:
        ventas = crear_ventas(region, precios, clientes, rng)
        ventas.to_excel(BASE / "ventas" / f"ventas_{region.lower()}.xlsx", index=False)
        print(f"Generado ventas_{region.lower()}.xlsx ({len(ventas)} registros)")

    print("Catálogos generados en data/catalogos/")


if __name__ == "__main__":
    main()
