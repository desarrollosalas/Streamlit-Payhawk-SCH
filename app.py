import datetime
import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
import os
import time
import numpy as np
from openpyxl import Workbook

# =========================================================
# CONFIGURACIÓN DE PÁGINA (CENTRADA)
# =========================================================

st.set_page_config(
    page_title="Generador Prinex - Payhawk",
    layout="centered"
)

st.markdown(
    """
    <style>
        .block-container {
            max-width: 900px;
            margin: auto;
            padding-top: 2rem;
        }

        h1, h2, h3 {
            text-align: center;
        }

        div.stButton > button {
            width: 100%;
            font-size: 1.1rem;
        }
    </style>
    """,
    unsafe_allow_html=True
)

# =========================================================
# PLANTILLA PRINEX FIJA
# =========================================================

COLUMNAS_PLANTILLA_PRINEX = [
    "SOCIEDAD","ORDEN","CIF","CODIGO","NUM.FRA","FECHA.FRA","FECHA.CONTABLE",
    "DIARIO_CONTB","IMP.BRUTO","TOTAL","OP.ALQ","D347","TIPO.FRA","SUJ_RECC",
    "DELEGACION","BASE_RETENCION","PORCENTAJE_RETENCION","IMPORTE_RETENCION",
    "APLICAR_RETENCION","BASE_IRPF","PORCENTAJE_IRPF","IMPORTE_IRPF",
    "CLAVE_IRPF","SUBCLAVE_IRPF","CEUTA","CONCEPTO","CTA_ACREEDORA",
    "SCTA_ACREEDORA","CTA_GARANTIA","SCTA_GARANTIA","CTA_IRPF","SCTA_IRPF",
    "CTA_IVAD","SCTA_IVAD","CONDICIONES","PAGADA","CTA_BANCO","SCTA_BANCO",
    "APUNTE","AUTOREPE_INVE_SUJE_PASI","SERIE_AUTOREPE","DIARIO_AUTOREPE",
    "TIPO_FRA_SII","CLAVE_RE","CLAVE_RE_AD1","CLAVE_RE_AD2","TIPO_OP_INTRA",
    "DESC_BIENES","DESCRIPCION_OP","SIMPLIFICADA","FRA_SIMPLI_IDEN",
    "BIEN_ART25","DOCU_ART25","PROT_ART25","NOTA_ART25",
    "DIARIO1","BASE1","IVA1","CUOTA1",
    "DIARIO2","BASE2","IVA2","CUOTA2",
    "DIARIO3","BASE3","IVA3","CUOTA3",
    "DIARIO4","BASE4","IVA4","CUOTA4",
    "DIARIO5","BASE5","IVA5","CUOTA5",
    "PROYECTO","TIPO_INMUEBLE","CLAVE1","CLAVE2","CLAVE3","CLAVE4",
    "IMPORTE_GASTO","TIPO_PARTIDA","APARTADO","CAPITULO","PARTIDA",
    "CTA_GASTO","SCTA_GASTO","COD_COEF","NOMBRE","CARACTERISTICA","RUTA","ETAPA"
]

def crear_plantilla_prinex_vacia():
    return pd.DataFrame(columns=COLUMNAS_PLANTILLA_PRINEX)

# =========================================================
# FUNCIONES AUXILIARES
# =========================================================

def convertir_df_a_excel(df):
    """
    Convierte un DataFrame a Excel plano, sin negrita ni bordes ni colores.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Plantilla Prinex"

    # Escribir encabezados (solo texto)
    ws.append(list(df.columns))

    # Escribir filas de datos
    for row in df.itertuples(index=False):
        ws.append(list(row))

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

def validar_archivos_cargados(zip_bytes):
    errores = []
    csv_ok = False
    pdf_ok = False

    with zipfile.ZipFile(BytesIO(zip_bytes)) as zip_ref:
        for nombre in zip_ref.namelist():
            if nombre.lower().endswith(".csv"):
                csv_ok = True
            elif nombre.lower().endswith(".pdf"):
                pdf_ok = True

    if not csv_ok:
        errores.append("El ZIP no contiene el archivo CSV de Payhawk.")
    if not pdf_ok:
        errores.append("El ZIP no contiene facturas en PDF.")

    return errores

# =========================================================
# PROCESAMIENTO PRINCIPAL
# =========================================================

def procesar_zip_payhawk(zip_bytes_payhawk, fecha_elegida):
    st.write("### 1️⃣ Descomprimiendo ZIP de Payhawk")

    df_payhawk = None
    archivos_pdf = {}

    with zipfile.ZipFile(BytesIO(zip_bytes_payhawk)) as zip_ref:
        for nombre in zip_ref.namelist():
            if nombre.lower().endswith(".csv"):
                with zip_ref.open(nombre) as f:
                    df_payhawk = pd.read_csv(f)
            elif nombre.lower().endswith(".pdf"):
                archivos_pdf[os.path.basename(nombre)] = zip_ref.read(nombre)

    st.success("ZIP descomprimido correctamente")

    st.write("### 2️⃣ Mapeando datos a Prinex")

    df_payhawk.columns = df_payhawk.columns.str.strip()
    df_prinex = crear_plantilla_prinex_vacia()
    df_prinex = pd.DataFrame(index=range(len(df_payhawk)), columns=df_prinex.columns)

    # -----------------------------------------------------
    # VALORES FIJOS
    # -----------------------------------------------------
    df_prinex["SOCIEDAD"] = 666
    df_prinex["DIARIO_CONTB"] = 1
    df_prinex["OP.ALQ"] = "N"
    df_prinex["D347"] = "N"
    df_prinex["DIARIO1"] = 1

    df_prinex["PAGADA"] = "S"
    df_prinex["CTA_BANCO"] = "5720"
    df_prinex["SCTA_BANCO"] = "001"
    df_prinex["APUNTE"] = "S"

    df_prinex["CARACTERISTICA"] = "PAYHAWK"
    df_prinex["CONDICIONES"] = "COMPTAT"
    df_prinex["RUTA"] = 9
    df_prinex["ETAPA"] = "CARGA PAYHAWK"

    df_prinex["CODIGO"] = 4444

    # -----------------------------------------------------
    # TIPO.FRA y CODIGO condicional según Document Type
    # -----------------------------------------------------
    df_prinex["TIPO.FRA"] = np.where(
        df_payhawk["Document Type"] == "Receipt",
        "C",
        np.where(df_payhawk["Document Type"] == "Invoice", "F", "")
    )

    df_prinex["TIPO.FRA"] = np.where(
        df_payhawk["Payment Type"] == "mileage",
        "C",
        df_prinex["TIPO.FRA"]
    )


    # -----------------------------------------------------
    # MAPEOS PAYHAWK → PRINEX
    # -----------------------------------------------------
    column_map = {
        "ORDEN": "Expense ID",
        "NUM.FRA": "Document Number",
        "IMP.BRUTO": "Net Amount (EUR)",
        "TOTAL": "Total Amount (EUR)",
        "BASE1": "Net Amount (EUR)",
        "IVA1": "Tax Rate %",
        "CUOTA1": "Tax Amount (EUR)",
        "PROYECTO": "Promoción External ID",
        "IMPORTE_GASTO": "Net Amount (EUR)",
        "NOMBRE": "File Name 1"
    }

    for prinex_col, payhawk_col in column_map.items():
        if payhawk_col not in df_payhawk.columns:
            continue

        if prinex_col == "NUM.FRA":
            df_prinex[prinex_col] = np.where(
                df_payhawk["Payment Type"] == "mileage",
                "KM-" + df_payhawk["Expense ID"].astype(str),
                df_payhawk[payhawk_col]
            )
        else:
            df_prinex[prinex_col] = df_payhawk[payhawk_col]

    mask_c = df_prinex["TIPO.FRA"] == "C"

    df_prinex.loc[mask_c, "IMP.BRUTO"] = df_payhawk.loc[mask_c, "Total Amount (EUR)"]
    df_prinex.loc[mask_c, "BASE1"] = df_payhawk.loc[mask_c, "Total Amount (EUR)"]
    df_prinex.loc[mask_c, "IMPORTE_GASTO"] = df_payhawk.loc[mask_c, "Total Amount (EUR)"]

    df_prinex.loc[mask_c, ["IVA1", "CUOTA1"]] = 0


    # -----------------------------------------------------
    # FECHAS
    # -----------------------------------------------------
    if "Document Date" in df_payhawk.columns:
        df_prinex["FECHA.FRA"] = pd.to_datetime(
            df_payhawk["Document Date"], errors="coerce"
        ).dt.strftime("%d/%m/%Y")

    fecha_formateada = fecha_elegida.strftime("%d/%m/%Y")
    df_prinex["FECHA.CONTABLE"] = fecha_formateada
    
    # -----------------------------------------------------
    # CTA / SCTA GASTO
    # -----------------------------------------------------
    if "Account Code" in df_payhawk.columns:
        split = df_payhawk["Account Code"].astype(str).str.split("-", n=1, expand=True)
        df_prinex["CTA_GASTO"] = split[0]
        df_prinex["SCTA_GASTO"] = split[1].fillna("") if 1 in split.columns else ""

    df_prinex = df_prinex.fillna("")
    st.success("Mapeo completado correctamente")

    return df_prinex, archivos_pdf

# =========================================================
# INTERFAZ STREAMLIT
# =========================================================

st.title("🚀 Generador de Carga Masiva Prinex desde Payhawk")

st.info(
    "Sube el archivo ZIP descargado desde Payhawk. "
    "La plantilla Prinex se genera automáticamente."
)

if "procesado" not in st.session_state:
    st.session_state.procesado = False
    st.session_state.zip_final = None
    st.session_state.df_preview = None

st.header("📦 Cargar ZIP de Payhawk")
archivo_zip = st.file_uploader("Selecciona el archivo ZIP", type=["zip"])

fecha_usuario = datetime.date.today()

if archivo_zip is not None:
    st.write("---")
    col_fecha, col_vacia = st.columns([1, 2])
    with col_fecha:
        fecha_usuario = st.date_input(
            "📅 Selecciona la Fecha Contable para este archivo:",
            value=datetime.date.today()
        )

st.divider()

if st.button("✨ Generar archivo de carga para Prinex", type="primary"):
    if archivo_zip is None:
        st.warning("Debes subir el ZIP de Payhawk.")
    else:
        try:
            zip_bytes = archivo_zip.getvalue()
            errores = validar_archivos_cargados(zip_bytes)

            if errores:
                st.error("**Errores encontrados:**\n" + "\n".join(f"- {e}" for e in errores))
            else:
                inicio = time.time()
                with st.spinner("Procesando archivos..."):
                    df_final, pdfs = procesar_zip_payhawk(zip_bytes, fecha_usuario)
                    
                    excel_bytes = convertir_df_a_excel(df_final)

                    zip_salida = BytesIO()
                    with zipfile.ZipFile(zip_salida, "w", zipfile.ZIP_DEFLATED) as z:
                        z.writestr("plantilla_prinex.xlsx", excel_bytes)
                        for nombre, contenido in pdfs.items():
                            z.writestr(f"facturas/{nombre}", contenido)

                st.session_state.procesado = True
                st.session_state.zip_final = zip_salida.getvalue()
                st.session_state.df_preview = df_final.head()

                st.success(f"Proceso completado en {time.time() - inicio:.2f} segundos")

        except Exception as e:
            st.error(f"Error inesperado: {e}")

if st.session_state.procesado:
    st.divider()
    st.header("📥 Descargar resultados")

    st.subheader("Vista previa (5 primeras filas)")
    st.dataframe(st.session_state.df_preview)

    st.download_button(
        label="Descargar ZIP final",
        data=st.session_state.zip_final,
        file_name="carga_prinex_con_facturas.zip",
        mime="application/zip",
        type="primary"
    )