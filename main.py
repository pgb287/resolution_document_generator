import os

import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert

# ============================================================
# CONFIGURACIÓN
# ============================================================

EXCEL = "listado_rectificar.xlsx"
PLANTILLA = "res.docx"

CARPETA_DOCX = "temp_docx"
CARPETA_PDF = "resoluciones_pdf"

# Último número de resolución asignado
CONTADOR_RES_INICIAL = 2005

# ============================================================
# CREAR CARPETAS
# ============================================================

os.makedirs(CARPETA_DOCX, exist_ok=True)
os.makedirs(CARPETA_PDF, exist_ok=True)


# ============================================================
# FUNCIONES
# ============================================================

def normalizar_padron(valor):
    """Convierte el padrón a texto eliminando decimales."""
    return str(int(float(valor)))


def normalizar_numero(valor):
    """Convierte un número a texto eliminando decimales."""
    return str(int(float(valor)))


def formatear_valor(valor):
    """
    Normaliza, redondea y formatea valores monetarios.

    Ejemplos:
        1234567.89       -> 1.234.568
        "1234567,89"     -> 1.234.568
        "1.234.567,89"   -> 1.234.568
        "1,234,567.89"   -> 1.234.568
        1234567           -> 1.234.567
    """

    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    if not texto:
        return ""

    try:
        # Caso con punto y coma
        if "." in texto and "," in texto:

            # Formato argentino: 1.234.567,89
            if texto.rfind(",") > texto.rfind("."):
                texto = texto.replace(".", "")
                texto = texto.replace(",", ".")

            # Formato inglés: 1,234,567.89
            else:
                texto = texto.replace(",", "")

        # Solo coma: 1234567,89
        elif "," in texto:
            texto = texto.replace(",", ".")

        numero = float(texto)

        # Redondear al entero más cercano
        numero = round(numero)

        # Formato argentino con punto de miles
        return f"{numero:,}".replace(",", ".")

    except ValueError:
        raise ValueError(
            f"No se pudo convertir el valor monetario: {valor}"
        )


def obtener_propietarios(grupo):
    """
    Obtiene propietarios válidos y únicos del inmueble,
    manteniendo el orden original del Excel.
    """

    propietarios = (
        grupo["propietario"]
        .dropna()
        .astype(str)
        .str.strip()
    )

    propietarios = propietarios[propietarios != ""]

    # Eliminar duplicados manteniendo el orden
    propietarios = propietarios.drop_duplicates()

    return propietarios.tolist()


def construir_nombre_propietario(propietarios):
    """
    Construye el nombre del propietario para la resolución.

    1 propietario:
        JUAN PEREZ

    2 propietarios:
        JUAN PEREZ Y OTRO

    3 o más propietarios:
        JUAN PEREZ Y OTROS
    """

    if not propietarios:
        return ""

    primer_propietario = propietarios[0]

    if len(propietarios) == 2:
        return f"{primer_propietario} Y OTRO"

    if len(propietarios) > 2:
        return f"{primer_propietario} Y OTROS"

    return primer_propietario


def obtener_nros_resoluciones(grupo):
    """
    Obtiene los números de resolución de multa del inmueble.

    Ejemplos:

    1 resolución:
        123/26

    2 resoluciones:
        123/26 y 456/26

    3 resoluciones:
        123/26, 456/26 y 789/26
    """

    resoluciones = (
        grupo["nro_res"]
        .dropna()
        .apply(normalizar_numero)
        .drop_duplicates()
        .tolist()
    )

    resoluciones = [
        f"{numero}/26"
        for numero in resoluciones
    ]

    if not resoluciones:
        return ""

    if len(resoluciones) == 1:
        return resoluciones[0]

    if len(resoluciones) == 2:
        return f"{resoluciones[0]} y {resoluciones[1]}"

    return (
        ", ".join(resoluciones[:-1])
        + f" y {resoluciones[-1]}"
    )


def obtener_tipo_inmueble(fila):
    """
    Determina la descripción del tipo de inmueble
    según las columnas 'tipoi' y 'tipo'.
    """

    tipoi = str(fila["tipoi"]).strip().upper()

    tipo = (
        str(fila["tipo"]).strip().upper()
        if pd.notna(fila["tipo"])
        else ""
    )

    if tipoi == "U":
        if tipo == "E":
            return "URBANO EDIFICADO"

        if tipo == "B":
            return "URBANO BALDIO"

        return ""

    if tipoi == "PH":
        return "PROPIEDAD HORIZONTAL"

    if tipoi == "R":
        return "RURAL"

    if tipoi == "SR":
        return "SUBRURAL"

    return ""


# ============================================================
# LEER EXCEL
# ============================================================

print("Leyendo Excel...")

df = pd.read_excel(EXCEL)

print(f"Registros encontrados: {len(df)}")


# ============================================================
# NORMALIZAR CAMPOS DE AGRUPACIÓN
# ============================================================

df["cod"] = df["cod"].astype(str).str.strip()

df["padron"] = df["padron"].apply(normalizar_padron)


# ============================================================
# AGRUPAR POR INMUEBLE
# ============================================================

inmuebles = df.groupby(
    ["cod", "padron"],
    sort=False,
    dropna=False
)

print(f"Inmuebles detectados: {inmuebles.ngroups}")
print()


# ============================================================
# GENERAR RESOLUCIONES
# ============================================================

contador_res = CONTADOR_RES_INICIAL
pdf_generados = 0


for (cod, padron), grupo in inmuebles:

    try:

        # ----------------------------------------------------
        # REGISTRO BASE
        # ----------------------------------------------------

        fila = grupo.iloc[0]


        # ----------------------------------------------------
        # PROPIETARIOS
        # ----------------------------------------------------

        propietarios = obtener_propietarios(grupo)

        nombre_propietario = construir_nombre_propietario(
            propietarios
        )

        cantidad_propietarios = len(propietarios)


        # ----------------------------------------------------
        # TEXTOS DINÁMICOS SEGÚN CANTIDAD DE PROPIETARIOS
        # ----------------------------------------------------

        if cantidad_propietarios == 1:
            texto_propietarios = "al propietario/a"
            cantidad_res = "a Resolución"
        else:
            texto_propietarios = "a los propietarios"
            cantidad_res = "as Resoluciones"


        # ----------------------------------------------------
        # RESOLUCIONES DE MULTA
        # ----------------------------------------------------

        nro_res = obtener_nros_resoluciones(grupo)


        # ----------------------------------------------------
        # NÚMERO DE NUEVA RESOLUCIÓN
        # ----------------------------------------------------

        contador_res += 1

        numero_res = str(contador_res).zfill(8)


        # ----------------------------------------------------
        # TIPO DE INMUEBLE
        # ----------------------------------------------------

        tipo_inmueble = obtener_tipo_inmueble(fila)


        # ----------------------------------------------------
        # FORMATEAR DATOS
        # ----------------------------------------------------

        vt = formatear_valor(fila["vt"])
        vm = formatear_valor(fila["vm"])
        vf = formatear_valor(fila["vf"])

        dni_cuil = (
            str(fila["numero"]).strip()
            if pd.notna(fila["numero"])
            else ""
        )

        email = (
            str(fila["email"]).strip()
            if pd.notna(fila["email"])
            else ""
        )


        # ----------------------------------------------------
        # MOSTRAR PROGRESO
        # ----------------------------------------------------

        print("----------------------------------------")
        print(f"Procesando inmueble: {cod}-{padron}")
        print(f"Propietarios detectados: {cantidad_propietarios}")
        print(f"Propietario resolución: {nombre_propietario}")
        print(f"Resoluciones de multa: {nro_res}")
        print(f"Nueva resolución: {numero_res}")


        # ----------------------------------------------------
        # CARGAR PLANTILLA
        # ----------------------------------------------------

        doc = DocxTemplate(PLANTILLA)


        # ----------------------------------------------------
        # CONTEXTO
        # ----------------------------------------------------

        contexto = {
            "res": numero_res,
            "cantidad_propietarios": texto_propietarios,
            "cantidad_res": cantidad_res,
            "numero_res": numero_res,
            "cod": cod,
            "padron": padron,
            "nombre_propietario": f'PROPIETARIO: {nombre_propietario}',           
            "nombre_propietario_2": nombre_propietario,           
            "nro_res": nro_res,
            "tipo_inmueble": tipo_inmueble,
            "vt": vt,
            "vm": vm,
            "vf": vf,
            "dni_cuil": f'DNI/CUIL N° {dni_cuil}',
            "email": f'DOMICILIO ELECTRONICO: {email}',
        }


        # ----------------------------------------------------
        # GENERAR WORD
        # ----------------------------------------------------

        doc.render(contexto)


        # ----------------------------------------------------
        # NOMBRE DEL ARCHIVO
        # ----------------------------------------------------

        nombre_archivo = (
            f"RS-2026-{numero_res}-{cod}-{padron}"
        )

        docx_path = os.path.join(
            CARPETA_DOCX,
            f"{nombre_archivo}.docx"
        )

        pdf_path = os.path.join(
            CARPETA_PDF,
            f"{nombre_archivo}.pdf"
        )


        # ----------------------------------------------------
        # GUARDAR DOCX
        # ----------------------------------------------------

        doc.save(docx_path)


        # ----------------------------------------------------
        # CONVERTIR A PDF
        # ----------------------------------------------------

        convert(
            docx_path,
            pdf_path
        )

        pdf_generados += 1

        print(f"PDF generado: {nombre_archivo}.pdf")


    except Exception as error:

        print("----------------------------------------")
        print(f"ERROR EN INMUEBLE: {cod}-{padron}")
        print(f"Detalle: {error}")


# ============================================================
# RESUMEN
# ============================================================

print()
print("========================================")
print("PROCESO FINALIZADO")
print("========================================")
print(f"Registros Excel: {len(df)}")
print(f"Inmuebles detectados: {inmuebles.ngroups}")
print(f"PDF generados: {pdf_generados}")
print("========================================")