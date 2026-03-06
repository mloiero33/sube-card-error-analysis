import re
import folium
from streamlit_folium import st_folium
import io
import os
import pandas as pd
import streamlit as st
import openpyxl

# NUEVAS IMPORTACIONES PARA PDF
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, LEGAL
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend

# =========================
# Configuración General de la Página
# =========================
st.set_page_config(page_title="Tarjetas Rechazadas", layout="wide")
st.title("Análisis de Tarjetas Rechazadas vs. Detalle de Ventas")
st.caption("SUBE>VENTAS>Detalle Ventas x Coche>detalle.xlsx")
st.caption("SUBE MOTOSIERRA>CONTROL GERENCIAL>TARJETAS RECHAZADAS>error.xlsx")


# =========================
# FUNCIÓN PARA GENERAR PDF
# =========================
def generar_reporte_pdf(no_match_df_filtrado, carpeta_sel, tolerancia_min, params, filtros_activos):
    """
    Genera un reporte PDF completo con análisis de tarjetas rechazadas
    Ahora usa el DataFrame ya filtrado en lugar del resultado completo
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=LEGAL,
                            rightMargin=72, leftMargin=72,
                            topMargin=72, bottomMargin=18)

    # Estilos
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        spaceAfter=30,
        alignment=TA_CENTER,
        textColor=colors.darkblue
    )

    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=14,
        spaceAfter=12,
        textColor=colors.darkblue
    )

    date_style = ParagraphStyle(
        'DateStyle',
        parent=styles['Normal'],
        fontSize=12,
        alignment=TA_CENTER,
        spaceAfter=20
    )

    # Contenido del reporte
    story = []

    # Título y fecha
    fecha_actual = datetime.now().strftime("%d-%m-%Y")
    story.append(Paragraph(f"<u>DÍA {fecha_actual}</u>", date_style))
    story.append(Spacer(1, 20))

    # Información general
    story.append(Paragraph("INFORMACIÓN GENERAL", heading_style))
    info_data = [
        ['Carpeta de trabajo:', carpeta_sel],
        ['Tolerancia configurada:', f"{tolerancia_min} minuto(s)"]
    ]

    # Agregar información sobre filtros activos
    if filtros_activos.get('coches'):
        info_data.append(
            ['Filtro Coches:', ', '.join(filtros_activos['coches'])])
    if filtros_activos.get('tarjetas'):
        info_data.append(
            ['Filtro Tarjetas:', ', '.join(filtros_activos['tarjetas'])])
    if filtros_activos.get('descripciones'):
        desc_text = ', '.join(filtros_activos['descripciones'])
        if len(desc_text) > 100:
            desc_text = desc_text[:100] + '...'
        info_data.append(['Filtro Descripciones:', desc_text])

    # NUEVO: filtro horario por rangos
    if filtros_activos.get('rangos_labels'):
        info_data.append([
            'Filtro Horario (2h):',
            ' | '.join(filtros_activos['rangos_labels'])
        ])

    info_data.append(['Registros en reporte:', str(len(no_match_df_filtrado))])

    info_table = Table(info_data, colWidths=[2*inch, 3*inch])
    info_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ('BACKGROUND', (1, 0), (1, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    story.append(info_table)
    story.append(Spacer(1, 30))

    # Análisis de NO MATCH (ahora usa el DataFrame filtrado)
    if len(no_match_df_filtrado) > 0:
        story.append(
            Paragraph("ANÁLISIS DE REGISTROS SIN COINCIDENCIA", heading_style))

        # Top 10 Descripciones
        story.append(
            Paragraph("Top 10 Descripciones con Mayor Frecuencia", heading_style))
        top_descr = no_match_df_filtrado['Descripcion'].value_counts().nlargest(
            10).reset_index()

        descr_data = [['Descripción', 'Cantidad']]
        for _, row in top_descr.iterrows():
            # Truncar descripción si es muy larga
            desc = row['Descripcion'][:50] + \
                '...' if len(str(row['Descripcion'])) > 50 else str(
                    row['Descripcion'])
            descr_data.append([desc, str(row['count'])])

        descr_table = Table(descr_data, colWidths=[4*inch, 1*inch])
        descr_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('ALIGN', (1, 0), (1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        story.append(descr_table)
        story.append(Spacer(1, 20))

        # Top Coches y Tarjetas
        story.append(
            Paragraph("Análisis por Coches y Tarjetas", heading_style))

        # Top Coches
        coche_series = no_match_df_filtrado["Coche"]
        if params.get('excluir_coche_vacios', True):
            coche_series = coche_series.replace(
                ["", "NAN", "NONE"], pd.NA).dropna()
        top_coche = coche_series.value_counts().nlargest(10).reset_index()

        # Top Tarjetas
        top_tarjetas = no_match_df_filtrado['Tarjeta Ext'].value_counts().nlargest(
            10).reset_index()

        # Crear tabla combinada
        max_rows = max(len(top_coche), len(top_tarjetas))
        combined_data = [['Top Coches', 'Cant.', '', 'Top Tarjetas', 'Cant.']]

        for i in range(max_rows):
            row = ['', '', '', '', '']
            if i < len(top_coche):
                row[0] = str(top_coche.iloc[i]['Coche'])
                row[1] = str(top_coche.iloc[i]['count'])
            if i < len(top_tarjetas):
                row[3] = str(top_tarjetas.iloc[i]['Tarjeta Ext'])
                row[4] = str(top_tarjetas.iloc[i]['count'])
            combined_data.append(row)

        combined_table = Table(combined_data, colWidths=[
                               1.5*inch, 0.7*inch, 0.3*inch, 1.5*inch, 0.7*inch])
        combined_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (4, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (4, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('BACKGROUND', (0, 1), (1, -1), colors.lightgrey),
            ('BACKGROUND', (3, 1), (4, -1), colors.lightgrey),
            ('GRID', (0, 0), (1, -1), 1, colors.black),
            ('GRID', (3, 0), (4, -1), 1, colors.black),
            ('BACKGROUND', (2, 0), (2, -1), colors.white),
        ]))
        story.append(combined_table)
        story.append(PageBreak())

        # Gráfico de distribución horaria
        story.append(
            Paragraph("DISTRIBUCIÓN HORARIA DE ERRORES", heading_style))

        # Crear gráfico de distribución horaria (ahora usa datos filtrados)
        horas = pd.to_datetime(
            no_match_df_filtrado["Hora_error"], format="%H:%M:%S", errors="coerce").dt.hour
        if horas.notna().any():
            fig, ax = plt.subplots(figsize=(10, 6))
            distribucion = horas.value_counts().sort_index()

            ax.bar(distribucion.index, distribucion.values,
                   color='steelblue', alpha=0.8)
            ax.set_xlabel('Hora del Día', fontsize=12)
            ax.set_ylabel('Cantidad de Errores', fontsize=12)
            ax.set_title(
                'Distribución Horaria de Registros NO MATCH', fontsize=14, pad=20)
            ax.grid(axis='y', alpha=0.3)
            ax.set_xticks(range(24))

            # Añadir valores encima de las barras
            for i, v in zip(distribucion.index, distribucion.values):
                ax.text(i, v + max(distribucion.values)*0.01, str(v),
                        ha='center', va='bottom', fontsize=10)

            plt.tight_layout()

            # Guardar gráfico
            img_buffer = io.BytesIO()
            plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
            img_buffer.seek(0)
            plt.close()

            # Añadir imagen al reporte
            img = Image(img_buffer, width=6*inch, height=3.6*inch)
            story.append(img)
            story.append(Spacer(1, 20))
            # --- Debajo del gráfico de distribución horaria ---
            story.append(
                Paragraph("MAPA ESTÁTICO DE ERRORES (NO MATCH)", heading_style))

            map_buf = generar_png_mapa_estatico(no_match_df_filtrado)
            if map_buf is not None:
                img_mapa = Image(map_buf, width=6*inch, height=3.6*inch)
                story.append(img_mapa)
                story.append(Spacer(1, 20))
            else:
                story.append(Paragraph(
                    "No hay coordenadas válidas para mostrar el mapa.", styles['Normal']))
                story.append(Spacer(1, 12))

    else:
        story.append(Paragraph(
            "NO HAY REGISTROS SIN COINCIDENCIA CON LOS FILTROS APLICADOS", heading_style))

    # Construir PDF
    doc.build(story)

    # Obtener bytes del PDF
    pdf_bytes = buffer.getvalue()
    buffer.close()

    return pdf_bytes


def generar_png_mapa_estatico(no_match_df_filtrado):
    """
    Genera un PNG (BytesIO) con los errores sobre un mapa de fondo (si hay contextily/pyproj).
    Si faltan dependencias, cae a un mapa sin fondo (hexbin + scatter).
    """
    # 1) Extraer coordenadas
    coords_series = no_match_df_filtrado.get('Posicion_url')
    if coords_series is None:
        return None

    coords = coords_series.dropna().apply(extraer_coordenadas)
    coords = [c for c in coords if isinstance(
        c, tuple) and c[0] is not None and c[1] is not None]
    if not coords:
        return None

    lats, lons = zip(*coords)  # lon eje X, lat eje Y

    # 2) Intentar basemap con contextily (EPSG:3857)
    try:
        import contextily as cx
        from pyproj import Transformer

        # Proyección a Web Mercator
        transformer = Transformer.from_crs(
            "EPSG:4326", "EPSG:3857", always_xy=True)
        xs, ys = transformer.transform(lons, lats)

        fig, ax = plt.subplots(figsize=(10, 6), dpi=150)

        # Densidad (hexbin) en coordenadas proyectadas + puntos
        hb = ax.hexbin(xs, ys, gridsize=50, mincnt=1, alpha=0.85)
        ax.scatter(xs, ys, s=8, alpha=0.6)

        # Extensión con pequeño padding (metros)
        pad_x = max(200, 0.03 * (max(xs) - min(xs) if len(xs) > 1 else 1000))
        pad_y = max(200, 0.03 * (max(ys) - min(ys) if len(ys) > 1 else 1000))
        ax.set_xlim(min(xs) - pad_x, max(xs) + pad_x)
        ax.set_ylim(min(ys) - pad_y, max(ys) + pad_y)

        # Basemap (ligero y claro)
        cx.add_basemap(ax, crs="EPSG:3857",
                       source=cx.providers.CartoDB.Positron)

        ax.set_title("Mapa estático de errores (NO MATCH)")
        ax.set_xlabel("")  # en Web Mercator no tiene sentido mostrar lat/lon
        ax.set_ylabel("")
        ax.grid(False)

        cbar = plt.colorbar(hb, ax=ax)
        cbar.set_label('Densidad de puntos')

        buf = io.BytesIO()
        plt.tight_layout()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        plt.close(fig)
        return buf

    except Exception:
        # 3) Fallback: mismo gráfico que ya tenías (sin basemap)
        fig, ax = plt.subplots(figsize=(10, 6), dpi=150)
        hb = ax.hexbin(lons, lats, gridsize=40, mincnt=1, alpha=0.9)
        ax.scatter(lons, lats, s=5, alpha=0.4)
        ax.set_xlabel('Longitud')
        ax.set_ylabel('Latitud')
        ax.set_title('Mapa estático de errores (NO MATCH)')
        ax.grid(alpha=0.3)
        cbar = plt.colorbar(hb, ax=ax)
        cbar.set_label('Densidad de puntos')

        buf = io.BytesIO()
        plt.tight_layout()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        plt.close(fig)
        return buf


# =========================
# Funciones de Utilidad
# =========================
def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    """Convierte un DataFrame de pandas a bytes en formato Excel para descarga."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output.getvalue()


def _normalizar_desc(s: pd.Series) -> pd.Series:
    """
    Normaliza descripciones para agrupar/comparar de forma robusta:
    - Quita tildes (NFKD)
    - Convierte a MAYÚSCULAS
    - Colapsa espacios múltiples
    - Quita punto(s) final(es)
    - Recorta espacios en blanco
    """
    return (
        s.astype(str)
        .str.normalize("NFKD")
        .str.encode("ascii", errors="ignore")
        .str.decode("utf-8")
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
        .str.replace(r"[.]+$", "", regex=True)
        .str.strip()
        .replace({"NAN": "SIN DESCRIPCION"})
    )


# =========================
# Carga de Datos (con caché)
# =========================
@st.cache_data(show_spinner="Leyendo archivo de errores...")
def read_error(path: str, header_row: int, mtime: float) -> pd.DataFrame:
    """Lee y procesa el archivo error.xlsx."""
    df = pd.read_excel(path, header=header_row)
    df = df.rename(columns=lambda x: str(x).strip())

    needed = ["Tarjeta Ext", "Descripcion", "Hora", "Coche", "Posicion"]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(
            f"Faltan columnas requeridas en error.xlsx: {missing}")

    df["Tarjeta Ext"] = df["Tarjeta Ext"].astype(str).str.extract(r"(\d+)")[0]
    df["Descripcion"] = _normalizar_desc(df["Descripcion"])
    df["Hora_error"] = df["Hora"].astype(str).str.strip()
    df["Hora_error_dt"] = pd.to_datetime(
        df["Hora_error"], errors="coerce", format="%H:%M:%S")
    df["Coche"] = df["Coche"].astype(str).str.strip()

    fecha_col = next(
        (c for c in df.columns if "FECHA" in str(c).upper()), df.columns[0])
    df["Fecha_dt"] = pd.to_datetime(df[fecha_col], errors="coerce")
    df["Fecha"] = df["Fecha_dt"].dt.strftime("%d-%m-%Y")

    # Extraer hipervínculos de la columna "Posicion" (columna J, índice 9 en 0-based)
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet = wb.active
    hyperlinks = []
    start_row = header_row + 2
    for row_idx in range(start_row, start_row + len(df)):
        cell = sheet.cell(row=row_idx, column=10)  # Columna J es 10
        if cell.hyperlink:
            hyperlinks.append(cell.hyperlink.target)
        else:
            hyperlinks.append(None)
    df["Posicion_url"] = hyperlinks

    return df[["Tarjeta Ext", "Descripcion", "Hora_error", "Hora_error_dt", "Coche", "Fecha_dt", "Fecha", "Posicion_url"]].copy()


@st.cache_data(show_spinner="Leyendo archivo de detalle de ventas...")
def read_detalle(path: str, header_row: int, mtime: float) -> pd.DataFrame:
    """Lee y procesa el archivo detalle.xlsx, buscando la tarjeta y la hora."""
    df = pd.read_excel(path, header=header_row)
    df = df.rename(columns=lambda x: str(x).strip())

    tarjeta_col = next(
        (c for c in df.columns if "TARJETA" in str(c).upper()), None)
    if not tarjeta_col:
        raise ValueError(
            "No se encontró una columna con 'Tarjeta' en detalle.xlsx")
    df["N° Tarjeta"] = df[tarjeta_col].astype(str).str.extract(r"(\d+)")[0]

    hora_col = next((c for c in df.columns if "HORA" in str(c).upper()), None)
    hora_dt = pd.Series(pd.NaT, index=df.index)

    if hora_col:
        hora_dt = pd.to_datetime(df[hora_col], errors="coerce")
    else:
        # Alternativa: Buscar en columnas de fecha con componente de tiempo
        for col in [c for c in df.columns if any(k in str(c).upper() for k in ["FECHA", "DATE"])]:
            temp_dt = pd.to_datetime(df[col], errors="coerce")
            if temp_dt.notna().any() and (temp_dt.dt.hour.sum() > 0 or temp_dt.dt.minute.sum() > 0):
                hora_dt = temp_dt
                break

    df["Hora_detalle_dt"] = hora_dt
    df["Hora_detalle"] = hora_dt.dt.strftime("%H:%M:%S").fillna("")

    # Consolidar por tarjeta, priorizando la primera hora no nula
    first = (
        df[["N° Tarjeta", "Hora_detalle", "Hora_detalle_dt"]]
        .sort_values(by=["N° Tarjeta", "Hora_detalle_dt"])
        .drop_duplicates(subset=["N° Tarjeta"], keep="first")
    )
    return first.reset_index(drop=True)


# =========================
# Lógica de negocio
# =========================
def deduplicar_error(error_df: pd.DataFrame, ignorar_coche: bool) -> tuple[pd.DataFrame, int, list]:
    """
    Elimina duplicados de error_df según criterio y registros con Tarjeta Ext = 0.
    """
    error_df = error_df[error_df["Tarjeta Ext"].astype(str) != "0"]

    subset = ["Tarjeta Ext", "Descripcion", "Fecha_dt", "Hora_error"]
    if not ignorar_coche:
        subset.append("Coche")

    dedup_df = error_df.drop_duplicates(subset=subset, keep="first")
    eliminados = len(error_df) - len(dedup_df)
    return dedup_df, eliminados, subset


@st.cache_data(show_spinner="Cruzando datos...")
def cruzar_datos(error_df: pd.DataFrame, detalle_df: pd.DataFrame, tolerancia_min: int) -> pd.DataFrame:
    """Realiza el cruce entre los dataframes de error y detalle."""
    res = pd.merge(
        error_df,
        detalle_df,
        left_on="Tarjeta Ext",
        right_on="N° Tarjeta",
        how="left",
    )

    fecha_base_str = res["Fecha_dt"].dt.strftime("%Y-%m-%d")
    err_alineada = pd.to_datetime(
        fecha_base_str + " " + res["Hora_error_dt"].dt.strftime("%H:%M:%S"), errors="coerce")
    det_alineada = pd.to_datetime(
        fecha_base_str + " " + res["Hora_detalle_dt"].dt.strftime("%H:%M:%S"), errors="coerce")

    diff_seconds = (det_alineada - err_alineada).dt.total_seconds()
    res["DiffMin"] = diff_seconds / 60.0

    cond_ok = (
        res["N° Tarjeta"].notna()
        & res["DiffMin"].notna()
        & (res["DiffMin"] >= 0)
        & (res["DiffMin"] <= float(tolerancia_min))
    )
    res["Estado"] = "NO MATCH"
    res.loc[cond_ok, "Estado"] = "OK"

    return res.drop(columns=["N° Tarjeta"])


def configurar_sidebar():
    """Crea y gestiona todos los controles de la barra lateral."""
    st.sidebar.header("Parámetros de Análisis")

    tolerancia_min = st.sidebar.slider(
        "Tolerancia de coincidencia (min)", 0, 60, 5, 1)

    st.sidebar.divider()
    st.sidebar.subheader("Limpieza (error.xlsx)")
    dedup_enable = st.sidebar.checkbox(
        "Eliminar duplicados exactos", value=True)
    dedup_ignorar_coche = st.sidebar.checkbox(
        "Ignorar 'Coche' al deduplicar", value=False)
    st.sidebar.caption("Criterio: Tarjeta + Descripción + Fecha + Hora.")

    st.sidebar.divider()
    st.sidebar.subheader("Visualización")
    show_ok = st.sidebar.checkbox("Mostrar coincidencias (OK)", value=False)
    sort_no_match_by_hora = st.sidebar.checkbox(
        "Ordenar NO MATCH por Hora", value=False)
    sort_no_match_by_coche = st.sidebar.checkbox(
        "Ordenar NO MATCH por Coche", value=False)

    st.sidebar.divider()
    st.sidebar.subheader("Análisis NO MATCH")
    top_n_coche = st.sidebar.slider("Top N coches", 5, 50, 10, 1)
    excluir_coche_vacios = st.sidebar.checkbox(
        "Excluir coches vacíos/NaN", value=True)

    st.sidebar.divider()
    st.sidebar.subheader("Filtro NO MATCH")
    filtro_placeholder = st.sidebar.empty()

    # NUEVO: Filtro horario en rangos fijos de 2 horas
    st.sidebar.subheader("Filtro horario (rangos de 2 horas)")
    rangos_2h = [(0, 2), (2, 4), (4, 6), (6, 8),
                 (8, 10), (10, 12), (12, 14), (14, 16),
                 (16, 18), (18, 20), (20, 22), (22, 24)]
    etiquetas_rangos = [
        f"{ini:02d}:00 - {fin:02d}:00" for ini, fin in rangos_2h]

    rangos_sel_labels = st.sidebar.multiselect(
        "Seleccionar rango(s) horario(s):",
        etiquetas_rangos
    )

    rangos_sel = [
        rangos_2h[etiquetas_rangos.index(lbl)]
        for lbl in rangos_sel_labels
    ]

    return {
        "tolerancia_min": tolerancia_min,
        "dedup_enable": dedup_enable,
        "dedup_ignorar_coche": dedup_ignorar_coche,
        "show_ok": show_ok,
        "sort_no_match_by_hora": sort_no_match_by_hora,
        "sort_no_match_by_coche": sort_no_match_by_coche,
        "top_n_coche": top_n_coche,
        "excluir_coche_vacios": excluir_coche_vacios,
        "filtro_placeholder": filtro_placeholder,
        # NUEVO
        "rangos_horarios": rangos_sel,
        "rangos_horarios_labels": rangos_sel_labels,
    }


def mostrar_metricas(resultado: pd.DataFrame, tolerancia_min: int):
    """Muestra las métricas principales del resultado del cruce."""
    total = len(resultado)
    ok = (resultado["Estado"] == "OK").sum()
    no_match = total - ok

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Registros (error.xlsx)", f"{total:,}".replace(",", "."))
    col2.metric("Coincidencias (OK)", f"{ok:,}".replace(",", "."))
    col3.metric("Sin Coincidencia (NO MATCH)",
                f"{no_match:,}".replace(",", "."))
    st.caption(f"Tolerancia configurada: **{tolerancia_min} minuto(s)**.")


# =========================
# Flujo Principal de la App
# =========================
def main():
    """Función principal que ejecuta la aplicación Streamlit."""

    # --- Selección de Carpeta ---
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        fechas_dir = os.path.join(base_dir, "Fechas")
        subcarpetas = sorted([d for d in os.listdir(
            fechas_dir) if os.path.isdir(os.path.join(fechas_dir, d))])
    except FileNotFoundError:
        st.error(f"No se encontró la carpeta 'Fechas' en el directorio del script.")
        st.stop()

    if not subcarpetas:
        st.error("La carpeta 'Fechas' no contiene subcarpetas.")
        st.stop()

    st.sidebar.title("📁 Selección de Carpeta")
    carpeta_sel = st.sidebar.selectbox(
        "Elige la carpeta de trabajo:", subcarpetas)
    carpeta_activa = os.path.join(fechas_dir, carpeta_sel)

    archivo_error = os.path.join(carpeta_activa, "error.xlsx")
    archivo_detalle = os.path.join(carpeta_activa, "detalle.xlsx")

    if not os.path.exists(archivo_error) or not os.path.exists(archivo_detalle):
        st.error(
            f"En la carpeta '{carpeta_sel}' falta uno o ambos archivos: 'error.xlsx', 'detalle.xlsx'.")
        st.stop()

    st.info(f"Carpeta activa: **{carpeta_sel}**")

    # --- Configuración y Parámetros ---
    params = configurar_sidebar()
    hdr_error, hdr_detalle = 3, 17  # Fila 4 y 18 en Excel

    # --- Carga y Procesamiento de Datos ---
    try:
        error_df = read_error(archivo_error, hdr_error,
                              os.path.getmtime(archivo_error))
        detalle_df = read_detalle(
            archivo_detalle, hdr_detalle, os.path.getmtime(archivo_detalle))
    except Exception as e:
        st.error(f"Error al leer los archivos: {e}")
        st.exception(e)
        st.stop()

    if params["dedup_enable"]:
        error_df, eliminados, criterio = deduplicar_error(
            error_df, params["dedup_ignorar_coche"])
        st.success(
            f"Deduplicación activa: se eliminaron {eliminados} registros. Criterio: {', '.join(criterio)}")

    resultado = cruzar_datos(error_df, detalle_df, params["tolerancia_min"])

    # --- Visualización de Resultados ---
    mostrar_metricas(resultado, params["tolerancia_min"])

    # --- Análisis de NO MATCH ---
    no_match_df = resultado[resultado["Estado"] == "NO MATCH"].copy()

    # Aplicar filtro horario por rangos de 2h si se seleccionaron
    if params["rangos_horarios"]:
        horas = no_match_df["Hora_error_dt"].dt.hour
        mascara_total = pd.Series(False, index=no_match_df.index)

        for inicio, fin in params["rangos_horarios"]:
            if fin == 24:
                # Último rango 22-24 incluye horas 22 y 23
                mascara = (horas >= inicio) & (horas <= 23)
            else:
                mascara = (horas >= inicio) & (horas < fin)
            mascara_total |= mascara

        no_match_df = no_match_df[mascara_total]

    st.divider()
    st.header("Análisis de Registros Sin Coincidencia (NO MATCH)")

    # --- Filtros dinámicos en sidebar para NO MATCH ---
    coches_disponibles = sorted(no_match_df["Coche"].dropna().unique())
    coches_sel = params["filtro_placeholder"].multiselect(
        "Filtrar por Coche(s):", coches_disponibles)

    # Filtro por Tarjeta Ext
    tarjetas_disponibles = sorted(no_match_df["Tarjeta Ext"].dropna().unique())
    tarjetas_sel = st.sidebar.multiselect(
        "Filtrar por Tarjeta(s):", tarjetas_disponibles)

    # Filtro por Descripción
    descripciones_disponibles = sorted(
        no_match_df["Descripcion"].dropna().unique())
    descripciones_sel = st.sidebar.multiselect(
        "Filtrar por Descripción(es):", descripciones_disponibles)

    # Aplica los filtros
    if coches_sel:
        no_match_df = no_match_df[no_match_df["Coche"].isin(coches_sel)]
    if tarjetas_sel:
        no_match_df = no_match_df[no_match_df["Tarjeta Ext"].isin(
            tarjetas_sel)]
    if descripciones_sel:
        no_match_df = no_match_df[no_match_df["Descripcion"].isin(
            descripciones_sel)]

    # --- Grillas de Top 5 Descripciones, Top N Coches y Top 10 Tarjetas ---
    col_a, col_b, col_c = st.columns([3, 1, 1])

    with col_a:
        st.subheader("Top Descripciones")
        top_descr = no_match_df['Descripcion'].value_counts().nlargest(
            10).reset_index()
        top_descr.columns = ["Descripción", "Cantidad"]
        st.dataframe(top_descr, use_container_width=True, hide_index=True)

    with col_b:
        st.subheader(f"Top Coches")
        coche_series = no_match_df["Coche"]
        if params['excluir_coche_vacios']:
            coche_series = coche_series.replace(
                ["", "NAN", "NONE"], pd.NA).dropna()
        top_coche = coche_series.value_counts().nlargest(
            params['top_n_coche']).reset_index()
        top_coche.columns = ["Coche", "Cantidad"]
        st.dataframe(
            top_coche,
            hide_index=True,
            column_config={
                "Coche": st.column_config.TextColumn(width="small"),
                "Cantidad": st.column_config.NumberColumn(width="small"),
            }
        )

    with col_c:
        st.subheader("Top Tarjetas")
        top_tarjetas = no_match_df['Tarjeta Ext'].value_counts().nlargest(
            10).reset_index()
        top_tarjetas.columns = ["Tarjeta Ext", "Cantidad"]
        st.dataframe(
            top_tarjetas,
            hide_index=True,
            column_config={
                "Tarjeta Ext": st.column_config.TextColumn(width="small"),
                "Cantidad": st.column_config.NumberColumn(width="small"),
            }
        )

    # --- Listado NO MATCH ---
    st.subheader("Listado Detallado de NO MATCH")
    nm_listado = no_match_df[["Tarjeta Ext",
                              "Descripcion", "Fecha", "Hora_error", "Coche", "Posicion_url"]]

    if params["sort_no_match_by_hora"]:
        nm_listado = nm_listado.sort_values(by="Hora_error")
    if params["sort_no_match_by_coche"]:
        nm_listado = nm_listado.sort_values(by="Coche")

    st.dataframe(
        nm_listado,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Posicion_url": st.column_config.LinkColumn(
                label="Posicion",
                display_text="Link",
                validate="^https?://",
                max_chars=100,
            )
        }
    )
    st.download_button(
        "Descargar Listado NO MATCH (xlsx)",
        data=to_excel_bytes(nm_listado, "NO_MATCH"),
        file_name=f"no_match_{carpeta_sel}.xlsx"
    )

    # --- Distribución horaria ---
    horas = pd.to_datetime(
        no_match_df["Hora_error"], format="%H:%M:%S", errors="coerce").dt.hour
    if horas.notna().any():
        st.subheader("Distribución Horaria de NO MATCH")

        # Crear dataframe con horas y cantidades
        distribucion = horas.value_counts().sort_index().reset_index()
        distribucion.columns = ["Hora del Día", "Cantidad"]

        # Mostrar tabla con horas y cantidades
        st.write(distribucion)

        # Mostrar gráfico de barras
        st.bar_chart(distribucion.set_index("Hora del Día"))

    # --- Visualización de OK (opcional) ---
    if params["show_ok"]:
        st.divider()
        st.header("Detalle de Coincidencias (OK)")
        ok_df = resultado[resultado["Estado"] == "OK"][[
            "Tarjeta Ext", "Descripcion", "Fecha", "Coche",
            "Hora_error", "Hora_detalle", "DiffMin"
        ]].rename(columns={
            "Hora_error": "Hora (Error)",
            "Hora_detalle": "Hora (Detalle)",
            "DiffMin": "Diferencia (min)"
        })
        st.dataframe(ok_df, use_container_width=True, hide_index=True)
        st.download_button(
            "Descargar Listado OK (xlsx)",
            data=to_excel_bytes(ok_df, "OK"),
            file_name=f"ok_{carpeta_sel}.xlsx"
        )

    # --- Mapa de Zonas Calientes ---
    filtros_aplicados = any([coches_sel, tarjetas_sel, descripciones_sel])

    if filtros_aplicados:
        mostrar_mapa_zonas(no_match_df)
    else:
        st.info(
            "Aplicá al menos un filtro (Coche, Tarjeta o Descripción) para ver el mapa de zonas calientes.")

    # --- SECCIÓN: Generación de Reporte PDF ---
    st.divider()

    col1, col2, col3 = st.columns([2, 2, 2])

    with col2:
        # Mostrar información sobre filtros activos
        filtros_info = []
        if coches_sel:
            filtros_info.append(f"Coches: {len(coches_sel)} seleccionados")
        if tarjetas_sel:
            filtros_info.append(f"Tarjetas: {len(tarjetas_sel)} seleccionadas")
        if descripciones_sel:
            filtros_info.append(
                f"Descripciones: {len(descripciones_sel)} seleccionadas")
        if params["rangos_horarios_labels"]:
            filtros_info.append(
                "Horario(s): " + " | ".join(params["rangos_horarios_labels"])
            )

        if filtros_info:
            st.caption(f"**Filtros activos:** {' | '.join(filtros_info)}")

        if st.button("Generar Reporte", type="primary", use_container_width=True):
            try:
                with st.spinner("Generando reporte PDF..."):
                    # Preparar información de filtros para el PDF
                    filtros_activos = {
                        'coches': coches_sel,
                        'tarjetas': tarjetas_sel,
                        'descripciones': descripciones_sel,
                        'rangos_labels': params["rangos_horarios_labels"],
                    }

                    # Pasar el DataFrame filtrado en lugar del resultado completo
                    pdf_bytes = generar_reporte_pdf(
                        no_match_df, carpeta_sel, params["tolerancia_min"], params, filtros_activos)

                    fecha_actual = datetime.now().strftime("%Y%m%d_%H%M%S")
                    nombre_archivo = f"reporte_tarjetas_{carpeta_sel}_{fecha_actual}.pdf"

                st.success("✅ ¡Reporte generado exitosamente!")

                st.download_button(
                    label="Descargar Reporte PDF",
                    data=pdf_bytes,
                    file_name=nombre_archivo,
                    mime="application/pdf",
                    type="secondary",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"❌ Error al generar el reporte: {str(e)}")
                st.exception(e)


def extraer_coordenadas(url):
    """
    Extrae latitud y longitud de un link de Google Maps.
    Devuelve (lat, lon) como float o (None, None) si falla.
    """
    if not isinstance(url, str):
        return None, None
    match = re.search(r'@?(-?\d+\.\d+),\s*(-?\d+\.\d+)', url)
    if match:
        return float(match.group(1)), float(match.group(2))
    return None, None


def mostrar_mapa_zonas(no_match_df):
    """
    Muestra mapa de calor con ubicaciones de errores (NO MATCH).
    """
    st.subheader("🗺️ Mapa de Zonas Calientes (NO MATCH)")

    coords = no_match_df['Posicion_url'].dropna().apply(extraer_coordenadas)
    coords = coords.dropna().tolist()
    coords = [(lat, lon) for lat, lon in coords if lat and lon]

    if not coords:
        st.info("No hay coordenadas válidas en los registros actuales.")
        return

    # Punto central
    centro = coords[0]

    m = folium.Map(location=centro, zoom_start=13)

    # Capa de calor
    from folium.plugins import HeatMap
    HeatMap(coords, radius=20).add_to(m)

    # Círculos de 300m (opcional)
    for lat, lon in coords:
        folium.Circle(
            location=(lat, lon),
            radius=300,
            color="blue",
            fill=True,
            fill_opacity=0.1
        ).add_to(m)

    st_data = st_folium(m, width=1300, height=600)


if __name__ == "__main__":
    main()
