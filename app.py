import streamlit as st
import pandas as pd
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP
from docx import Document
import io

# ======================
#  UTILIDADES DE TEXTO
# ======================

UNIDADES = (
    "cero", "uno", "dos", "tres", "cuatro", "cinco", "seis",
    "siete", "ocho", "nueve", "diez", "once", "dos­ce", "trece",
    "catorce", "quince", "dieciséis", "diecisiete", "dieciocho",
    "diecinueve", "veinte"
)

DECENAS = (
    "cero", "diez", "veinte", "treinta", "cuarenta", "cincuenta",
    "sesenta", "setenta", "ochenta", "noventa"
)

CENTENAS = (
    "cero", "cien", "doscientos", "trescientos", "cuatrocientos",
    "quinientos", "seiscientos", "setecientos",
    "ochocientos", "novecientos"
)


def numero_a_letras_menor_1000(n: int) -> str:
    n = int(n)
    if n < 21:
        return UNIDADES[n]
    if n < 100:
        d, u = divmod(n, 10)
        if u == 0:
            return DECENAS[d]
        if d == 2:
            return "veinti" + UNIDADES[u]
        return DECENAS[d] + " y " + UNIDADES[u]
    c, r = divmod(n, 100)
    if n == 100:
        return "cien"
    pref = CENTENAS[c]
    if r == 0:
        return pref
    return pref + " " + numero_a_letras_menor_1000(r)


def numero_a_letras_centavos(n: int) -> str:
    return numero_a_letras_menor_1000(n)


def numero_a_letras_pesos(valor: float) -> str:
    v = Decimal(str(valor)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    entero = int(v)
    centavos = int((v - Decimal(entero)) * 100)

    millones, resto = divmod(entero, 1_000_000)
    miles, unidades = divmod(resto, 1_000)

    partes = []
    if millones > 0:
        if millones == 1:
            partes.append("un millón")
        else:
            partes.append(numero_a_letras_menor_1000(millones) + " millones")

    if miles > 0:
        if miles == 1:
            partes.append("mil")
        else:
            partes.append(numero_a_letras_menor_1000(miles) + " mil")

    if unidades > 0 or entero == 0:
        partes.append(numero_a_letras_menor_1000(unidades))

    texto_entero = " ".join(partes)

    if centavos == 0:
        return f"{texto_entero} pesos"
    else:
        return f"{texto_entero} pesos con {numero_a_letras_centavos(centavos)} centavos"


# ======================
#  LECTURA DE USURA
# ======================

def cargar_usura(path: str):
    df = pd.read_excel(path)
    df["Fecha desde"] = pd.to_datetime(df["Fecha desde"]).dt.date
    df = df.sort_values("Fecha desde")
    return df


def obtener_tasa_ea(df_usura, fecha):
    filtro = df_usura[df_usura["Fecha desde"] <= fecha]
    if filtro.empty:
        st.error(f"No hay tasa de usura para la fecha {fecha}")
        st.stop()
    return Decimal(str(filtro.iloc[-1]["Tasa EA"]))


# ======================
#  MOTOR DE LIQUIDACIÓN
# ======================

def liquidar_obligacion(fila, df_usura, fecha_liquidacion):

    capital = Decimal(str(fila["CAPITAL"]))

    fecha_venc = pd.to_datetime(fila["FECHA VENCIMIENTO PAGARÉ"]).date()
    fecha_intereses = fecha_venc + timedelta(days=1)

    fecha_actual = fecha_intereses
    interes_acum = Decimal("0")
    filas = []

    while fecha_actual <= fecha_liquidacion:

        fin_mes = (fecha_actual.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)
        fecha_hasta = min(fin_mes, fecha_liquidacion)

        dias = (fecha_hasta - fecha_actual).days + 1

        tasa_ea = obtener_tasa_ea(df_usura, fecha_actual)

        factor_dia = ((Decimal("1") + tasa_ea) ** (Decimal("1") / Decimal("365"))) - Decimal("1")

        interes_periodo = (capital * factor_dia * dias).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        interes_acum = (interes_acum + interes_periodo).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

        filas.append({
            "fecha_desde": fecha_actual,
            "fecha_hasta": fecha_hasta,
            "tasa_ea": float(tasa_ea),
            "factor_dia": float(factor_dia),
            "dias": dias,
            "interes_periodo": float(interes_periodo),
            "interes_acumulado": float(interes_acum)
        })

        fecha_actual = fecha_hasta + timedelta(days=1)

    df_detalle = pd.DataFrame(filas)

    resumen = {
        "nombre": fila["NOMBRE"],
        "cedula": fila["CEDULA"],
        "pagaré": fila["No. PAGARÉ"],
        "juzgado": fila["JUZGADO"],
        "correo_juzgado": fila["CORREO JUZGADO"],
        "radicado": fila["RADICADO"],
        "capital": float(capital),
        "total_mora": float(interes_acum),
        "saldo_total": float(capital + interes_acum),
        "fecha_intereses": fecha_intereses,
        "fecha_liquidacion": fecha_liquidacion
    }

    return df_detalle, resumen


# ======================
#  GENERADOR MEMORIAL
# ======================

def reemplazar(doc, placeholder, valor):
    for p in doc.paragraphs:
        if placeholder in p.text:
            p.text = p.text.replace(placeholder, valor)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if placeholder in cell.text:
                    cell.text = cell.text.replace(placeholder, valor)


def generar_memorial(resumen, df_detalle):
    doc = Document("FORMATO MEMORIAL APORTA LIQUIDACIÓN DE CRÉDITO.docx")

    reemplazar(doc, "{{JUZGADO}}", resumen["juzgado"])
    reemplazar(doc, "{{CORREO_JUZGADO}}", resumen["correo_juzgado"])
    reemplazar(doc, "{{RADICADO}}", str(resumen["radicado"]))
    reemplazar(doc, "{{NOMBRE}}", resumen["nombre"])
    reemplazar(doc, "{{CEDULA}}", str(resumen["cedula"]))
    reemplazar(doc, "{{PAGARE}}", str(resumen["pagaré"]))

    reemplazar(doc, "{{FECHA_INTERESES}}", resumen["fecha_intereses"].strftime("%d/%m/%Y"))
    reemplazar(doc, "{{FECHA_LIQUIDACION}}", resumen["fecha_liquidacion"].strftime("%d/%m/%Y"))

    reemplazar(doc, "{{CAPITAL}}", f"${resumen['capital']:,.2f}")
    reemplazar(doc, "{{TOTAL_MORA}}", f"${resumen['total_mora']:,.2f}")
    reemplazar(doc, "{{SALDO_TOTAL}}", f"${resumen['saldo_total']:,.2f}")

    valor_letras = numero_a_letras_pesos(resumen["saldo_total"])
    reemplazar(doc, "{{VALOR_LETRAS}}", valor_letras)
    reemplazar(doc, "{{VALOR_NUM}}", f"${resumen['saldo_total']:,.2f}")

    # SEGUNDA HOJA CON LA TABLA
    doc.add_page_break()
    tabla = doc.add_table(rows=1, cols=7)
    tabla.rows[0].cells[0].text = "Desde"
    tabla.rows[0].cells[1].text = "Hasta"
    tabla.rows[0].cells[2].text = "EA"
    tabla.rows[0].cells[3].text = "Factor día"
    tabla.rows[0].cells[4].text = "Días"
    tabla.rows[0].cells[5].text = "Interés periodo"
    tabla.rows[0].cells[6].text = "Acumulado"

    for _, r in df_detalle.iterrows():
        row = tabla.add_row().cells
        row[0].text = r["fecha_desde"].strftime("%d/%m/%Y")
        row[1].text = r["fecha_hasta"].strftime("%d/%m/%Y")
        row[2].text = f"{r['tasa_ea']*100:.2f}%"
        row[3].text = f"{r['factor_dia']*100:.5f}%"
        row[4].text = str(r["dias"])
        row[5].text = f"${r['interes_periodo']:,.2f}"
        row[6].text = f"${r['interes_acumulado']:,.2f}"

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


# ======================
#  STREAMLIT UI
# ======================

st.title("💼 Liquidador Judicial Masivo – Banco GNB Sudameris")

st.subheader("1️⃣ Cargar base de obligaciones")
archivo_base = st.file_uploader("Sube el archivo Excel", type=["xlsx"])

if archivo_base:
    df_base = pd.read_excel(archivo_base)

    st.success(f"Base cargada con {len(df_base)} registros.")

    st.subheader("2️⃣ Seleccionar fecha de liquidación")
    fecha_liquidacion = st.date_input("Fecha de liquidación", value=date.today())

    st.subheader("3️⃣ Cargar histórico de tasas de usura")
    df_usura = cargar_usura("TASAS_DE_USURA.xlsx")
    st.success("Tasas de usura cargadas.")

    st.subheader("4️⃣ Previsualizar obligación")
    lista_pagare = df_base["No. PAGARÉ"].astype(str).tolist()
    pagare_sel = st.selectbox("Selecciona la obligación a revisar:", lista_pagare)

    fila_sel = df_base[df_base["No. PAGARÉ"].astype(str) == pagare_sel].iloc[0]

    df_detalle, resumen = liquidar_obligacion(fila_sel, df_usura, fecha_liquidacion)

    st.write("### 🔍 Resumen de liquidación")
    st.json(resumen)

    st.write("### 📊 Detalle mensual")
    st.dataframe(df_detalle)

    st.subheader("5️⃣ Generar memorial")
    if st.button("Generar memorial de ESTA obligación"):
        archivo = generar_memorial(resumen, df_detalle)
        st.download_button(
            "Descargar memorial",
            archivo,
            file_name=f"MEMORIAL_{resumen['pagaré']}.docx"
        )

    st.subheader("6️⃣ Generar memoriales masivos")
    if st.button("Generar TODOS los memoriales"):
        mem_zip = io.BytesIO()
        import zipfile
        with zipfile.ZipFile(mem_zip, "w") as z:
            for _, fila in df_base.iterrows():
                df_d, res = liquidar_obligacion(fila, df_usura, fecha_liquidacion)
                arc = generar_memorial(res, df_d)
                z.writestr(f"MEMORIAL_{res['pagaré']}.docx", arc)

        st.download_button(
            "Descargar ZIP de memoriales",
            mem_zip.getvalue(),
            file_name="MEMORIALES_GNB.zip"
        )
