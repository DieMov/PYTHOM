# -*- coding: utf-8 -*-
"""
Análisis de sucursales con gráficos en navegador (Matplotlib -> HTML/PNG; Plotly -> browser)
Autor: tú :)
"""

# ========================== CONFIG ==========================
EXCEL_FILE   = "Limpia_250811_master_reto_sucursales (version 1).xlsx"
EXPORT_EXCEL = "resultadoS.xlsx"

# ======================== IMPORTS ===========================
import os
import tempfile
import webbrowser
from pathlib import Path

import numpy as np
import pandas as pd

# Matplotlib
import matplotlib
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401 (activa proyección 3D)

# Plotly
import plotly.express as px
import plotly.io as pio
pio.renderers.default = "browser"  # abre gráficos interactivos en el navegador

# Intentar mpld3 (interactivo HTML para Matplotlib); si falla, haremos fallback a PNG
try:
    import mpld3
    _HAS_MPLD3 = True
except Exception:
    _HAS_MPLD3 = False


# ===================== UTILIDADES UI ========================
def show_in_browser(fig: plt.Figure | None = None, title_prefix: str = "fig"):
    """
    Abre una figura de Matplotlib en el navegador.
    1) Intenta exportar a HTML con mpld3 si está disponible.
    2) Si mpld3 no está o falla, guarda un PNG temporal y lo abre.
    """
    if fig is None:
        fig = plt.gcf()

    # Intento HTML con mpld3
    if _HAS_MPLD3:
        try:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".html")
            html = mpld3.fig_to_html(fig)
            with open(tmp.name, "w", encoding="utf-8") as f:
                f.write(html)
            webbrowser.open("file://" + os.path.realpath(tmp.name))
            return
        except Exception as e:
            print(f"[Aviso] mpld3 falló, uso PNG estático. Detalle: {e}")

    # Fallback: PNG
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    fig.savefig(tmp.name, dpi=140, bbox_inches="tight")
    webbrowser.open("file://" + os.path.realpath(tmp.name))


# ===================== FUNCIONES CÁLCULO ====================
def safe_div(num, den):
    """División segura (evita división por cero)."""
    num = pd.to_numeric(num, errors="coerce")
    den = pd.to_numeric(den, errors="coerce").replace(0, np.nan)
    return num / den


# ========================= MAIN =============================
def main():
    # ---------- Cargar Excel ----------
    if not Path(EXCEL_FILE).exists():
        raise FileNotFoundError(
            f"No encontré el archivo '{EXCEL_FILE}'. "
            "Pon este script en la MISMA carpeta que el Excel o ajusta EXCEL_FILE."
        )

    df = pd.read_excel(EXCEL_FILE)
    print(f"✅ Excel cargado: {EXCEL_FILE} | Filas: {len(df):,}")

    # ---------- Filtrado inicial (todas estas columnas en 0 o NaN) ----------
    saldo_cols = [
        "Saldo Insoluto Actual",
        "Saldo Insoluto T-01",
        "Saldo Insoluto T-02",
        "Saldo Insoluto T-03",
        "Saldo Insoluto T-04",
        "Saldo Insoluto T-05",
        "Saldo Insoluto T-06",
        "Saldo Insoluto T-12",
    ]
    for c in saldo_cols:
        if c not in df.columns:
            raise KeyError(f"Falta la columna requerida en el Excel: '{c}'")

    filtro = np.logical_and.reduce([(df[c].isna() | (df[c] == 0)) for c in saldo_cols])

    print("\nVendedores con todas las columnas en 0 o nulas:")
    if "Vendedor" in df.columns:
        print(df.loc[filtro, "Vendedor"])
    else:
        print("(No existe columna 'Vendedor' en tu archivo.)")

    df_filtrado = df.loc[~filtro].copy()
    print(f"🔧 Se limpiaron {int(filtro.sum())} filas con todos los saldos 0/NaN.")

    # ---------- Reglas Capital–FPD ----------
    capital_cols = [
        "Capital Dispersado Actual", "Capital Dispersado T-01", "Capital Dispersado T-02",
        "Capital Dispersado T-03", "Capital Dispersado T-04", "Capital Dispersado T-05",
        "Capital Dispersado T-06", "Capital Dispersado T-07", "Capital Dispersado T-08",
        "Capital Dispersado T-09", "Capital Dispersado T-10", "Capital Dispersado T-11",
        "Capital Dispersado T-12"
    ]
    fpd_cols = [
        "% FPD Actual", "% FPD T-01", "% FPD T-02",
        "% FPD T-03", "% FPD T-04", "% FPD T-05",
        "% FPD T-06", "% FPD T-07", "% FPD T-08",
        "% FPD T-09", "% FPD T-10", "% FPD T-11",
        "% FPD T-12"
    ]

    for col in fpd_cols:
        if col in df_filtrado.columns:
            df_filtrado[col] = pd.to_numeric(df_filtrado[col], errors="coerce")

    for cap_col, fpd_col in zip(capital_cols, fpd_cols):
        if cap_col not in df_filtrado.columns or fpd_col not in df_filtrado.columns:
            continue
        mask_fill_zero = (df_filtrado[cap_col].fillna(0) != 0) & (df_filtrado[fpd_col].isna())
        df_filtrado.loc[mask_fill_zero, fpd_col] = 0

        mask_zero_to_nan = (df_filtrado[cap_col].fillna(0) == 0) & (df_filtrado[fpd_col].fillna(0) == 0)
        df_filtrado.loc[mask_zero_to_nan, fpd_col] = np.nan

    # ---------- Mapeo Región–Zona–Sucursal ----------
    datos = [
        # Brokers
        ("Brokers", "Centro Metrópolis", "Centro Metrópolis"),
        ("Brokers", "Conexión Magna", "Conexión Magna"),
        ("Brokers", "Enlace Regio", "Enlace Regio"),
        ("Brokers", "Puerto Magna", "Puerto Magna"),
        ("Brokers", "Brokers", "Brokers"),

        # Núcleo Uno
        ("Núcleo Uno", "División. Red Mexiquense", "Ciudad Pirámide"),
        ("Núcleo Uno", "División. Red Mexiquense", "Valle Verde"),
        ("Núcleo Uno", "División. Red Mexiquense", "Río Blanco"),
        ("Núcleo Uno", "División. Red Mexiquense", "Colina del Sol"),
        ("Núcleo Uno", "División. Red Mexiquense", "Colina del Sol BIS"),
        ("Núcleo Uno", "División. Red Mexiquense", "Parque Jurica"),
        ("Núcleo Uno", "División. Red Mexiquense", "Colina Plateada"),
        ("Núcleo Uno", "División. Red Mexiquense", "Altos de Querétaro"),
        ("Núcleo Uno", "División. Red Mexiquense", "Sol y Campo"),
        ("Núcleo Uno", "Conexión Naucalpan", "Satélite 1"),
        ("Núcleo Uno", "Conexión Naucalpan", "Satélite 2"),
        ("Núcleo Uno", "Conexión Naucalpan", "Satélite 3"),
        ("Núcleo Uno", "Zona Sur Central", "Bahía Dorada"),
        ("Núcleo Uno", "Zona Sur Central", "Costa Marquesa"),
        ("Núcleo Uno", "Zona Sur Central", "Bahía Dorada BIS"),
        ("Núcleo Uno", "Zona Sur Central", "Alto de Chilpan"),
        ("Núcleo Uno", "Zona Sur Central", "Cuautla Vista"),
        ("Núcleo Uno", "Zona Sur Central", "Jardines del Valle"),
        ("Núcleo Uno", "Zona Sur Central", "Llanos de Igualdad"),
        ("Núcleo Uno", "Zona Sur Central", "Parque Jojutla"),
        ("Núcleo Uno", "Zona Oriente Valle", "Valle Chalco"),
        ("Núcleo Uno", "Zona Oriente Valle", "Montaña Azul"),
        ("Núcleo Uno", "Zona Oriente Valle", "Reyes Paz A"),
        ("Núcleo Uno", "Zona Oriente Valle", "Reyes Paz B"),
        ("Núcleo Uno", "Zona Oriente Valle", "Bosques Neza"),
        ("Núcleo Uno", "Zona Oriente Valle", "Cumbre Neza"),
        ("Núcleo Uno", "Zona Oriente Valle", "Cumbre Neza BIS"),
        ("Núcleo Uno", "Zona Oriente Valle", "Riberas Texcoco"),
        ("Núcleo Uno", "Zona Norte Valle", "Pinar del Valle"),
        ("Núcleo Uno", "Zona Norte Valle", "Cielos de Metepec"),
        ("Núcleo Uno", "Zona Norte Valle", "Lomas de Naucalpan"),
        ("Núcleo Uno", "Zona Norte Valle", "Puente de Tlalne"),
        ("Núcleo Uno", "Zona Norte Valle", "Puente de Tlalne II"),
        ("Núcleo Uno", "Zona Norte Valle", "Valles Toluca"),
        ("Núcleo Uno", "Zona Norte Valle", "Cumbre Toluca"),
        ("Núcleo Uno", "Zona Norte Valle", "Bosques Tultitlán"),

        # Núcleo Dos
        ("Núcleo Dos", "División. Distrito Central", "Jardín Aragón A"),
        ("Núcleo Dos", "División. Distrito Central", "Pilares del Norte"),
        ("Núcleo Dos", "División. Distrito Central", "Pilares del Norte BIS"),
        ("Núcleo Dos", "División. Distrito Central", "Residencia A"),
        ("Núcleo Dos", "División. Distrito Central", "Residencia B"),
        ("Núcleo Dos", "División. Distrito Central", "Colinas GAM"),
        ("Núcleo Dos", "División. Distrito Central", "Plaza Central"),
        ("Núcleo Dos", "División. Distrito Central", "Los Arcos"),
        ("Núcleo Dos", "División. Distrito Central", "Los Arcos BIS"),
        ("Núcleo Dos", "División. Distrito Central", "Campo Zaragoza"),
        ("Núcleo Dos", "División. Distrito Central", "Lomas Zaragoza"),
        ("Núcleo Dos", "División. Distrito Central", "Campo Zaragoza BIS"),
        ("Núcleo Dos", "Núcleo Avance", "Avance 1"),
        ("Núcleo Dos", "Núcleo Avance", "Avance 2"),
        ("Núcleo Dos", "Núcleo Avance", "Avance 3"),
        ("Núcleo Dos", "Núcleo Avance", "Avance 4"),
        ("Núcleo Dos", "Zona Núcleo CDMX", "Parque Obregón"),
        ("Núcleo Dos", "Zona Núcleo CDMX", "Centro Viejo"),
        ("Núcleo Dos", "Zona Núcleo CDMX", "Mirador Tlalpan A"),
        ("Núcleo Dos", "Zona Núcleo CDMX", "Mirador Tlalpan A BIS"),
        ("Núcleo Dos", "Zona Núcleo CDMX", "Mirador Tlalpan B"),
        ("Núcleo Dos", "Zona Núcleo CDMX", "Lagunas de Xochimilco"),
        ("Núcleo Dos", "Zona Núcleo CDMX", "Plaza Zapata"),
        ("Núcleo Dos", "Zona Oriente Conexión", "Robledal A"),
        ("Núcleo Dos", "Zona Oriente Conexión", "Robledal B"),
        ("Núcleo Dos", "Zona Oriente Conexión", "Campo Florido A"),
        ("Núcleo Dos", "Zona Oriente Conexión", "Campo Florido B"),
        ("Núcleo Dos", "Zona Oriente Conexión", "Campo Florido C"),
        ("Núcleo Dos", "Zona Oriente Conexión", "Campo Florido D"),
        ("Núcleo Dos", "Zona Oriente Conexión", "Riberas del Sur"),
        ("Núcleo Dos", "Zona Cordillera Puebla", "Sierra Cordobesa"),
        ("Núcleo Dos", "Zona Cordillera Puebla", "Valles de Orizaba"),
        ("Núcleo Dos", "Zona Cordillera Puebla", "Alturas de Puebla"),
        ("Núcleo Dos", "Zona Cordillera Puebla", "Jardines Manuel"),
        ("Núcleo Dos", "Zona Cordillera Puebla", "Lomas Santiago"),
        ("Núcleo Dos", "Zona Cordillera Puebla", "Alturas de Puebla BIS"),
        ("Núcleo Dos", "Zona Cordillera Puebla", "Llanos Tehuacán"),
        ("Núcleo Dos", "Zona Cordillera Puebla", "Bosques Tlaxcala"),
        ("Núcleo Dos", "Zona Bahía Veracruz", "Colinas Mirón"),
        ("Núcleo Dos", "Zona Bahía Veracruz", "Valle Rica"),
        ("Núcleo Dos", "Zona Bahía Veracruz", "Puerto Bravo"),
        ("Núcleo Dos", "Zona Bahía Veracruz", "Puerta Cuauhtémoc"),
        ("Núcleo Dos", "Zona Bahía Veracruz", "Puerto Bravo BIS"),
        ("Núcleo Dos", "Zona Bahía Veracruz", "Cumbres Xalapa"),
        ("Núcleo Dos", "Zona Bahía Veracruz", "Lomas Xalapa"),

        # Red Norteña
        ("Red Norteña","División Red Norteña","Paso del Norte"),
        ("Red Norteña","División Red Norteña","Río Bravo"),
        ("Red Norteña","División Red Norteña","Aceros del Norte"),
        ("Red Norteña","División Red Norteña","Aceros del Norte BIS"),
        ("Red Norteña","División Red Norteña","Paso Nuevo"),
        ("Red Norteña","División Red Norteña","Paso Nuevo BIS"),
        ("Red Norteña","División Red Norteña","Piedras Altas"),
        ("Red Norteña","División Red Norteña","Piedras Altas BIS"),
        ("Red Norteña","División Red Norteña","Valles del Norte"),
        ("Red Norteña","División Red Norteña","Laguna Norte"),
        ("Red Norteña","División Red Norteña","Sabinas Sierra"),
        ("Red Norteña","División Red Norteña","Campos Saltillo"),
        ("Red Norteña","División Red Norteña","Centro Saltillo"),
        ("Red Norteña","División Red Norteña","Centro Saltillo BIS"),
        ("Red Norteña","División Red Norteña","Campos Saltillo BIS"),

        ("Red Norteña", "Zona Sierra Norte", "Lomas de Álamos"),
        ("Red Norteña", "Zona Sierra Norte", "Lomas de Álamos BIS"),
        ("Red Norteña", "Zona Sierra Norte", "Valle Apodaca"),
        ("Red Norteña", "Zona Sierra Norte", "Valle Apodaca BIS"),
        ("Red Norteña", "Zona Sierra Norte", "Puente Lincoln"),
        ("Red Norteña", "Zona Sierra Norte", "Cumbres Regias"),
        ("Red Norteña", "Zona Sierra Norte", "Centro Regio"),
        ("Red Norteña", "Zona Sierra Norte", "Bulevar Regio"),
        ("Red Norteña", "Zona Sierra Norte", "San Nicolás Valle"),
        ("Red Norteña", "Zona Sierra Norte", "San Nicolás Valle BIS"),
        ("Red Norteña", "Zona Sierra Norte", "Sierra Santa"),

        ("Red Norteña", "Zona Red frontera este", "Bosque Verde"),
        ("Red Norteña", "Zona Red frontera este", "Palacio del Norte"),
        ("Red Norteña", "Zona Red frontera este", "Palacio del Norte BIS"),
        ("Red Norteña", "Zona Red frontera este", "Valle de Guadalupe"),
        ("Red Norteña", "Zona Red frontera este", "Parque Madero"),
        ("Red Norteña", "Zona Red frontera este", "Parque Madero BIS"),
        ("Red Norteña", "Zona Red frontera este", "Expo Regia"),
        ("Red Norteña", "Zona Red frontera este", "Desierto Norte"),
        ("Red Norteña", "Zona Red frontera este", "Desierto Bravo"),
        ("Red Norteña", "Zona Red frontera este", "Río Revolución"),
        ("Red Norteña", "Zona Red frontera este", "Desierto Norte BIS"),

        ("Red Norteña", "Zona Bahía del Sol", "Valle Real"),
        ("Red Norteña", "Zona Bahía del Sol", "Victoria Alta"),
        ("Red Norteña", "Zona Bahía del Sol", "Victoria Alta BIS"),
        ("Red Norteña", "Zona Bahía del Sol", "Bahía Aeropuerto"),
        ("Red Norteña", "Zona Bahía del Sol", "Plaza Tampico"),
        ("Red Norteña", "Zona Bahía del Sol", "Colinas Tampico"),
        ("Red Norteña", "Zona Bahía del Sol", "Colinas Tampico BIS"),
        ("Red Norteña", "Zona Bahía del Sol", "Río Madero"),

        # Red Noroeste
        ("Red Noroeste","División Sierra del Desierto","Sierra Chihuahua"),
        ("Red Noroeste","División Sierra del Desierto","Campus Sierra"),
        ("Red Noroeste","División Sierra del Desierto","Victoria Sierra"),
        ("Red Noroeste","División Sierra del Desierto","Victoria Sierra BIS"),
        ("Red Noroeste","División Sierra del Desierto","Plaza Cuauhtémoc"),
        ("Red Noroeste","División Sierra del Desierto","Juárez Norte"),
        ("Red Noroeste","División Sierra del Desierto","Jardines del Norte"),
        ("Red Noroeste","División Sierra del Desierto","Americas Plaza"),
        ("Red Noroeste","División Sierra del Desierto","Americas Plaza BIS"),
        ("Red Noroeste","División Sierra del Desierto","Patio Grande"),
        ("Red Noroeste","División Sierra del Desierto","Colinas Jilotepec"),
        ("Red Noroeste","División Sierra del Desierto","Parral Viejo"),
        ("Red Noroeste","División Sierra del Desierto","División Sierra del Desierto"),
        ("Red Noroeste","Zona Costa del Pacífico","Bahía Azul"),
        ("Red Noroeste","Zona Costa del Pacífico","Bahía Azul BIS"),
        ("Red Noroeste","Zona Costa del Pacífico","Plaza Pacifico"),
        ("Red Noroeste","Zona Costa del Pacífico","Plaza Pacifico BIS"),
        ("Red Noroeste","Zona Costa del Pacífico","Cabo Fuerte"),
        ("Red Noroeste","Zona Costa del Pacífico","Valle Mexicali"),
        ("Red Noroeste","Zona Costa del Pacífico","Norte Mexicali"),
        ("Red Noroeste","Zona Costa del Pacífico","Valle Mexicali BIS"),
        ("Red Noroeste","Zona Costa del Pacífico","Frontera Oeste"),
        ("Red Noroeste","Zona Costa del Pacífico","Frontera Bravo"),
        ("Red Noroeste","Zona Costa del Pacífico","Frontera Bravo BIS"),
        ("Red Noroeste","Zona Costa del Pacífico","Zona Costa del Pacífico"),
        ("Red Noroeste","Zona Valle Dorado","Valles de Culiacán"),
        ("Red Noroeste","Zona Valle Dorado","Culiacán Norte"),
        ("Red Noroeste","Zona Valle Dorado","Valles de Culiacán BIS"),
        ("Red Noroeste","Zona Valle Dorado","Sierra Durango"),
        ("Red Noroeste","Zona Valle Dorado","Durango Norte"),
        ("Red Noroeste","Zona Valle Dorado","Valle del Río"),
        ("Red Noroeste","Zona Valle Dorado","Plaza Mochis"),
        ("Red Noroeste","Zona Valle Dorado","Plaza Mochis BIS"),
        ("Red Noroeste","Zona Valle Dorado","Bahía Dorada"),
        ("Red Noroeste","Zona Valle Dorado","Norte Dorado"),
        ("Red Noroeste","Zona Valle Dorado","Zona Valle Dorado"),
        ("Red Noroeste","Zona Desierto del Sol","Obregón Central"),
        ("Red Noroeste","Zona Desierto del Sol","Obregón Norte"),
        ("Red Noroeste","Zona Desierto del Sol","Obregón Central BIS"),
        ("Red Noroeste","Zona Desierto del Sol","Sierra Hermosillo"),
        ("Red Noroeste","Zona Desierto del Sol","Hermosillo Norte"),
        ("Red Noroeste","Zona Desierto del Sol","Sierra Hermosillo BIS"),
        ("Red Noroeste","Zona Desierto del Sol","Valle de Navojoa"),
        ("Red Noroeste","Zona Desierto del Sol","Frontera Nogales"),
        ("Red Noroeste","Zona Desierto del Sol","Zona Desierto del Sol"),

        # Occidente Conexión
        ("Occidente Conexión","Conexión GDL","Guadalajara Uno"),
        ("Occidente Conexión","Conexión GDL","Guadalajara Dos"),
        ("Occidente Conexión","Conexión GDL","Guadalajara Tres"),
        ("Occidente Conexión","Conexión GDL","Conexión GDL"),
        ("Occidente Conexión","Zona Corazón de la Sierra","Aguas Central"),
        ("Occidente Conexión","Zona Corazón de la Sierra","Aguas Norte"),
        ("Occidente Conexión","Zona Corazón de la Sierra","Aguas Central BIS"),
        ("Occidente Conexión","Zona Corazón de la Sierra","Sierra Colima"),
        ("Occidente Conexión","Zona Corazón de la Sierra","Río Fresnillo"),
        ("Occidente Conexión","Zona Corazón de la Sierra","Bahía Manzanillo"),
        ("Occidente Conexión","Zona Corazón de la Sierra","San Luis Norte"),
        ("Occidente Conexión","Zona Corazón de la Sierra","San Luis Alturas"),
        ("Occidente Conexión","Zona Corazón de la Sierra","Cumbres Zacatecas"),
        ("Occidente Conexión","Zona Corazón de la Sierra","Zona Corazón de la Sierra"),
        ("Occidente Conexión","Zona Valles Centrales","Plaza Celaya"),
        ("Occidente Conexión","Zona Valles Centrales","Hidalgo Valle"),
        ("Occidente Conexión","Zona Valles Centrales","Jardines Irapuato"),
        ("Occidente Conexión","Zona Valles Centrales","Jardines Irapuato BIS"),
        ("Occidente Conexión","Zona Valles Centrales","Cañadas León"),
        ("Occidente Conexión","Zona Valles Centrales","Norte León"),
        ("Occidente Conexión","Zona Valles Centrales","Cañadas León BIS"),
        ("Occidente Conexión","Zona Valles Centrales","Zona Valles Centrales"),
        ("Occidente Conexión","Zona Tierra de lagos","Valle Piedad"),
        ("Occidente Conexión","Zona Tierra de lagos","Bahía Lázaro"),
        ("Occidente Conexión","Zona Tierra de lagos","Colinas Morelia"),
        ("Occidente Conexión","Zona Tierra de lagos","Morelia Norte"),
        ("Occidente Conexión","Zona Tierra de lagos","Morelia Norte BIS"),
        ("Occidente Conexión","Zona Tierra de lagos","Camelinas Plaza"),
        ("Occidente Conexión","Zona Tierra de lagos","Jardines Uruapan"),
        ("Occidente Conexión","Zona Tierra de lagos","Valle Zamora"),
        ("Occidente Conexión","Zona Tierra de lagos","Valle Zamora BIS"),
        ("Occidente Conexión","Zona Tierra de lagos","Riviera Zihua"),
        ("Occidente Conexión","Zona Tierra de lagos","Montes Zitácuaro"),
        ("Occidente Conexión","Zona Tierra de lagos","Zona Tierra de lagos"),
        ("Occidente Conexión","Zona Cumbres del Pacífico","GDL Central"),
        ("Occidente Conexión","Zona Cumbres del Pacífico","Norte GDL"),
        ("Occidente Conexión","Zona Cumbres del Pacífico","Riviera Vallarta"),
        ("Occidente Conexión","Zona Cumbres del Pacífico","Riviera Vallarta BIS"),
        ("Occidente Conexión","Zona Cumbres del Pacífico","Valle Tepic"),
        ("Occidente Conexión","Zona Cumbres del Pacífico","Norte Tepic"),
        ("Occidente Conexión","Zona Cumbres del Pacífico","Jardines Tlaque"),
        ("Occidente Conexión","Zona Cumbres del Pacífico","Montes Tonalá"),
        ("Occidente Conexión","Zona Cumbres del Pacífico","Zona Cumbres del Pacífico"),
        ("Occidente Conexión","Zona Valles del Pacífico","Guzmán Valle"),
        ("Occidente Conexión","Zona Valles del Pacífico","IMSS GDL"),
        ("Occidente Conexión","Zona Valles del Pacífico","Oblatos Plaza"),
        ("Occidente Conexión","Zona Valles del Pacífico","Las Águilas"),
        ("Occidente Conexión","Zona Valles del Pacífico","Zapopan Plaza"),
        ("Occidente Conexión","Zona Valles del Pacífico","Zona Valles del Pacífico"),

        # Red Sureste
        ("Red Sureste","Zona Selva Alta","Río Coatzacoalcos"),
        ("Red Sureste","Zona Selva Alta","Valle Comitán"),
        ("Red Sureste","Zona Selva Alta","Selva Tapachula"),
        ("Red Sureste","Zona Selva Alta","Tuxtla Norte"),
        ("Red Sureste","Zona Selva Alta","Tuxtla Central"),
        ("Red Sureste","Zona Selva Alta","Villa Central"),
        ("Red Sureste","Zona Selva Alta","Villa Norte"),
        ("Red Sureste","Zona Selva Alta","Villa Norte BIS"),
        ("Red Sureste","Zona Selva Alta","Villa Alturas"),
        ("Red Sureste","Zona Selva Alta","Zona Selva Alta"),
        ("Red Sureste","Zona Sierra Escondida","Sierra Oaxaca"),
        ("Red Sureste","Zona Sierra Escondida","Riviera Escondida"),
        ("Red Sureste","Zona Sierra Escondida","Bahía Cruz"),
        ("Red Sureste","Zona Sierra Escondida","Bahía Cruz BIS"),
        ("Red Sureste","Zona Sierra Escondida","Río Tuxtepec"),
        ("Red Sureste","Zona Sierra Escondida","Zona Sierra Escondida"),
        ("Red Sureste","Zona Riviera del Caribe","Bahía Campeche"),
        ("Red Sureste","Zona Riviera del Caribe","Riviera Cancún"),
        ("Red Sureste","Zona Riviera del Caribe","Bahía Chetumal"),
        ("Red Sureste","Zona Riviera del Caribe","Isla del Carmen"),
        ("Red Sureste","Zona Riviera del Caribe","Sierra Mérida"),
        ("Red Sureste","Zona Riviera del Caribe","Mérida Norte"),
        ("Red Sureste","Zona Riviera del Caribe","Alturas Mérida"),
        ("Red Sureste","Zona Riviera del Caribe","Sierra Mérida BIS"),
        ("Red Sureste","Zona Riviera del Caribe","Riviera Playa"),
        ("Red Sureste","Zona Riviera del Caribe","Zona Riviera del Caribe"),
    ]

    df_mapa = pd.DataFrame(datos, columns=["Región", "Zona", "Sucursal"])
    df_final = df_filtrado.merge(df_mapa, on="Sucursal", how="left")
    cols = ["Región", "Zona", "Sucursal"] + [c for c in df_final.columns if c not in ["Región","Zona","Sucursal"]]
    df_final = df_final[cols]

    # ---------- Columnas calculadas ----------
    tasainteresanual = 0.65 / 12
    tasacostefondeo  = 0.11 / 12

    if {"Saldo Insoluto Actual", "Saldo Insoluto Vencido Actual"}.issubset(df_final.columns):
        df_final["SaldoInsolutoVigente"] = df_final["Saldo Insoluto Actual"] - df_final["Saldo Insoluto Vencido Actual"]
        df_final["InteresGenerado"]      = df_final["SaldoInsolutoVigente"] * tasainteresanual
        df_final["ServiciodeDeuda"]      = df_final["Saldo Insoluto Actual"] * tasacostefondeo

    # ---------- Agrupar por Región/Zona/Sucursal ----------
    cols_sumar = (
        ["Saldo Insoluto Actual", "Saldo Insoluto Vencido Actual"]
        + [f"Saldo Insoluto T-{i:02d}" for i in range(1,13)]
        + [f"Saldo Insoluto Vencido T-{i:02d}" for i in range(1,13)]
    )
    cols_sumar = [c for c in cols_sumar if c in df_final.columns]

    df_sucursal = df_final.groupby(["Región", "Zona", "Sucursal"], as_index=False)[cols_sumar].sum()

    print("\nVista rápida de df_sucursal:")
    print(df_sucursal.head())

    # ---------- ICV (manejo división por cero) ----------
    if {"Saldo Insoluto Vencido Actual", "Saldo Insoluto Actual"}.issubset(df_sucursal.columns):
        df_sucursal["ICV"] = safe_div(df_sucursal["Saldo Insoluto Vencido Actual"],
                                      df_sucursal["Saldo Insoluto Actual"])

    for i in range(1, 13):
        v = f"Saldo Insoluto Vencido T-{i:02d}"
        s = f"Saldo Insoluto T-{i:02d}"
        if {v, s}.issubset(df_sucursal.columns):
            df_sucursal[f"ICV T-{i:02d}"] = safe_div(df_sucursal[v], df_sucursal[s])

    # ---------- Análisis preliminar ----------
    if "ServiciodeDeuda" in df_final.columns:
        top15 = df_final.sort_values("ServiciodeDeuda", ascending=False).head(15)
        print("\nTOP 15 por Servicio de Deuda:")
        print(top15[["Sucursal","Región","Zona","ServiciodeDeuda"]])
    else:
        print("\n(No existe columna 'ServiciodeDeuda')")

    if "Saldo Insoluto Actual" in df_final.columns:
        suma_insoluto_stat = pd.to_numeric(df_final["Saldo Insoluto Actual"], errors="coerce").sum()
        print(f"\nSuma total 'Saldo Insoluto Actual': ${suma_insoluto_stat:,.2f}\n")

    # ---------- Gráfica tablero (Matplotlib -> navegador) ----------
    if "ServiciodeDeuda" in df_final.columns:
        aux = df_final.sort_values("ServiciodeDeuda", ascending=False).head(15)
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.bar(aux["Sucursal"].astype(str), aux["ServiciodeDeuda"])
        ax.set_title("Top 15 Sucursales por Servicio de Deuda")
        ax.tick_params(axis="x", labelrotation=90, labelsize=8)
        plt.tight_layout()
        show_in_browser(fig)

    # ---------- Boxplot ICV por Zona (Matplotlib -> navegador) ----------
    if {"ICV", "Zona"}.issubset(df_sucursal.columns):
        icv_series = pd.to_numeric(df_sucursal["ICV"], errors="coerce").replace([np.inf, -np.inf], np.nan)
        p01, p99 = icv_series.quantile([0.01, 0.99])
        df_clip = df_sucursal[(icv_series >= p01) & (icv_series <= p99)].copy()

        fig, ax = plt.subplots(figsize=(12, 5))
        df_clip.boxplot(column="ICV", by="Zona", showfliers=False, ax=ax)
        ax.set_title("ICV por Zona (recortado p1–p99)")
        plt.suptitle("")
        ax.set_ylabel("ICV")
        plt.xticks(rotation=45, ha="right")
        plt.tight_layout()
        show_in_browser(fig)

    # ---------- Boxplot ICV por Región (Matplotlib -> navegador) ----------
    if {"ICV", "Región"}.issubset(df_sucursal.columns):
        icv_series = pd.to_numeric(df_sucursal["ICV"], errors="coerce").replace([np.inf, -np.inf], np.nan)
        p01, p99 = icv_series.quantile([0.01, 0.99])
        df_clip = df_sucursal[(icv_series >= p01) & (icv_series <= p99)].copy()

        fig, ax = plt.subplots(figsize=(12, 5))
        df_clip.boxplot(column="ICV", by="Región", showfliers=False, ax=ax)
        ax.set_title("ICV por Región (recortado p1–p99)")
        plt.suptitle("")
        ax.set_ylabel("ICV")
        plt.xticks(rotation=45, ha="right")
        plt.tight_layout()
        show_in_browser(fig)

    # ---------- Scatter 3D interactivo (Plotly -> navegador) ----------
    y_candidates = ["%FPD Actual", "% FPD Actual"]
    y_col = next((c for c in y_candidates if c in df_final.columns), None)
    x_col = "Capital Dispersado Actual" if "Capital Dispersado Actual" in df_final.columns else None
    z_col = "Saldo Insoluto Actual" if "Saldo Insoluto Actual" in df_final.columns else None

    if all([x_col, y_col, z_col]) and {"Sucursal","Región","Zona"}.issubset(df_final.columns):
        _df3d = df_final[[x_col, y_col, z_col, "Sucursal", "Región", "Zona"]].copy()
        _df3d[x_col] = pd.to_numeric(_df3d[x_col], errors="coerce")
        _df3d[y_col] = pd.to_numeric(_df3d[y_col], errors="coerce")
        _df3d[z_col] = pd.to_numeric(_df3d[z_col], errors="coerce")
        _df3d = _df3d.dropna(subset=[x_col, y_col, z_col])

        fig_plotly = px.scatter_3d(
            _df3d,
            x=x_col, y=y_col, z=z_col,
            color="Región",
            hover_data=["Sucursal","Zona","Región"],
            size=x_col,
            opacity=0.7,
            title=f"Scatter 3D interactivo: {y_col} vs {x_col} vs {z_col}"
        )
        # abre en navegador automáticamente por pio.renderers
        fig_plotly.show()
    else:
        print("⛔ No se generó el scatter 3D interactivo (faltan columnas x/y/z o Sucursal/Región/Zona).")

    # ---------- Scatter 3D Matplotlib y versión recortada (-> navegador) ----------
    if all([x_col, y_col, z_col]) and {"Sucursal","Región","Zona"}.issubset(df_final.columns):
        OUT = Path("figuras_dimex"); OUT.mkdir(exist_ok=True)

        X = pd.to_numeric(df_final[x_col], errors="coerce")
        Y = pd.to_numeric(df_final[y_col], errors="coerce")
        Z = pd.to_numeric(df_final[z_col], errors="coerce")
        mask = X.notna() & Y.notna() & Z.notna()
        X, Y, Z = X[mask], Y[mask], Z[mask]

        # Completo
        fig = plt.figure()
        ax = fig.add_subplot(111, projection="3d")
        ax.scatter(X, Y, Z, s=8, alpha=0.7)
        ax.set_xlabel(x_col); ax.set_ylabel(y_col); ax.set_zlabel(z_col)
        ax.set_title(f"3D: {y_col} vs {x_col} vs {z_col}")
        plt.tight_layout()
        fig.savefig(OUT / "H_scatter3D_fdp_capital_saldo.png", dpi=140)
        show_in_browser(fig)

        # Recortado p99
        xq, yq, zq = X.quantile(0.99), Y.quantile(0.99), Z.quantile(0.99)
        m_clip = (X <= xq) & (Y <= yq) & (Z <= zq)
        fig2 = plt.figure()
        ax2 = fig2.add_subplot(111, projection="3d")
        ax2.scatter(X[m_clip], Y[m_clip], Z[m_clip], s=8, alpha=0.7)
        ax2.set_xlabel(x_col); ax2.set_ylabel(y_col); ax2.set_zlabel(z_col)
        ax2.set_title(f"3D (p99 clip): {y_col} vs {x_col} vs {z_col}")
        plt.tight_layout()
        fig2.savefig(OUT / "H_scatter3D_fdp_capital_saldo_p99.png", dpi=140)
        show_in_browser(fig2)

        print("✅ Scatter 3D (Matplotlib) generado y abierto en navegador. PNG guardados en ./figuras_dimex/")
    else:
        print("⛔ No se generaron PNG 3D (faltan columnas x/y/z o Sucursal/Región/Zona).")

    # ---------- Exportar Excel final ----------
    df_final.to_excel(EXPORT_EXCEL, index=False)
    print(f"\n✅ Exportado: {EXPORT_EXCEL} (aparece en tu carpeta de trabajo)\n")


# ===================== ENTRY POINT ==========================
if __name__ == "__main__":
    # En Windows, asegura backend compatible (opcional)
    try:
        # Backend por defecto suele funcionar; puedes forzar Agg si renders locales dieran problemas:
        # matplotlib.use("Agg")
        pass
    except Exception as e:
        print(f"[Aviso backend Matplotlib] {e}")

    main()
a