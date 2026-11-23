from __future__ import annotations

from datetime import date, datetime
from io import BytesIO
import re
from zoneinfo import ZoneInfo

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

try:
    from fpdf import FPDF
except ImportError:
    FPDF = None


TZ_ESPAÑA = ZoneInfo("Europe/Madrid")
TIPOS_SUELO = ("Porcelánico / Cerámico", "Tarima")
TIPOS_PARED = ("Pintura", "Pladur", "Papel", "Porcelánico / Cerámico")
TIPOS_TECHO = ("Pladur", "Pintura")
MATERIAS_PRIMAS = (
    "Porcelánico / Cerámico",
    "Tarima",
    "Pintura",
    "Pladur",
    "Papel",
)
FORMATO_NUMERO_EXCEL = "#,##0.00"
FORMATO_MONEDA_EXCEL = '#,##0.00 "€"'

LIMITE_AVISO_PLANTA = 15.0
LIMITE_MAX_PLANTA = 50.0
LIMITE_MIN_ALTURA_AVISO = 2.0
LIMITE_MAX_ALTURA_AVISO = 3.5
LIMITE_MAX_ALTURA = 6.0

CATALOGO_INSTALACIONES: dict[str, list[tuple[str, str]]] = {
    "Electricidad y Telecomunicaciones": [
        ("elec_luz", "Puntos de luz (techo / aplique)"),
        ("elec_int", "Interruptores simples"),
        ("elec_conm", "Conmutados / cruzamientos"),
        ("elec_ench", "Enchufes (bases 16A)"),
        ("elec_tv", "Tomas de TV"),
        ("elec_red", "Tomas de red (Ethernet / RJ45)"),
    ],
    "Fontanería y Saneamiento": [
        ("font_lavabo", "Toma + desagüe lavabo"),
        ("font_inodoro", "Toma + desagüe inodoro"),
        ("font_bide", "Toma + desagüe bidé"),
        ("font_ducha", "Toma + desagüe plato ducha / bañera"),
        ("font_freg", "Toma + desagüe fregadero"),
        ("font_lavad", "Toma + desagüe lavadora / lavavajillas"),
        ("font_termo", "Toma termo eléctrico / caldera"),
    ],
    "Calefacción y Climatización": [
        ("clim_rad", "Radiadores (toma + elemento)"),
        ("clim_aa", "Preinstalación / split Aire Acondicionado"),
        ("clim_term", "Termostato ambiente"),
    ],
}

CATALOGO_OBRA_GENERAL: dict[str, list[tuple[str, str, str]]] = {
    "Carpintería y Cerramientos": [
        ("carp_puerta_ab", "Puerta de paso abatible (block completo)", "ud."),
        ("carp_puerta_co", "Puerta corredera empotrada (casoneto + hoja)", "ud."),
        ("carp_entrada", "Puerta de entrada blindada / acorazada", "ud."),
        ("carp_ventana", "Ventana / balconera (PVC o aluminio + vidrio)", "ud."),
        ("carp_persiana", "Persiana motorizada", "ud."),
    ],
    "Albañilería y Gestión de Residuos": [
        ("alb_contenedor", "Contenedor de escombros (alquiler y gestión)", "ud."),
        ("alb_demolicion", "Demolición y picado de revestimientos / tabiques", "jor."),
        ("alb_rozas", "Apertura y tapado de rozas / ayudas albañilería", "ml"),
        ("alb_solera", "Nivelación / recrecido de solera", "m²"),
    ],
    "Mobiliario de Cocina": [
        ("coc_bajos", "Muebles bajos de cocina", "ml"),
        ("coc_altos", "Muebles altos de cocina", "ml"),
        ("coc_columna", "Mueble columna / despensero", "ud."),
        ("coc_encimera", "Encimera de cocina", "ml"),
    ],
}


def ahora_españa() -> datetime:
    return datetime.now(TZ_ESPAÑA)


def formato_numero(valor: float, decimales: int = 2) -> str:
    texto = f"{float(valor):,.{decimales}f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def formato_euro(valor: float) -> str:
    return f"{formato_numero(valor, 2)} €"


def formato_fecha(valor: date | datetime) -> str:
    return valor.strftime("%d/%m/%Y")


def formato_fecha_hora(valor: datetime) -> str:
    return valor.strftime("%d/%m/%Y %H:%M")


def nombre_archivo_seguro(texto: str) -> str:
    texto = str(texto or "").strip() or "Sin cliente"
    texto = re.sub(r'[\\/:*?"<>|]+', "", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto or "Sin cliente"


def nombre_informe(cliente: str, instante: datetime, extension: str, prefijo: str = "Informe") -> str:
    cliente_limpio = nombre_archivo_seguro(cliente)
    fecha = instante.strftime("%d %m %y %H %M")
    return f"{prefijo} {cliente_limpio} {fecha}.{extension}"
def nuevo_id_estancia(estancias: list[dict]) -> int:
    if not estancias:
        return 1
    return max(int(e.get("id", 0)) for e in estancias) + 1


def calcular_estancia(largo: float, ancho: float, altura: float) -> dict:
    suelo = largo * ancho
    paredes = 2 * (largo + ancho) * altura
    return {
        "Largo (m)": float(largo),
        "Ancho (m)": float(ancho),
        "Altura (m)": float(altura),
        "Paredes (m²)": paredes,
        "Suelo (m²)": suelo,
        "Techo (m²)": suelo,
    }


def calcular_totales(estancias: list[dict]) -> dict:
    paredes = sum(float(e["Paredes (m²)"]) for e in estancias)
    suelos = sum(float(e["Suelo (m²)"]) for e in estancias)
    techos = sum(float(e["Techo (m²)"]) for e in estancias)

    suelos_por_tipo = {tipo: 0.0 for tipo in TIPOS_SUELO}
    paredes_por_tipo = {tipo: 0.0 for tipo in TIPOS_PARED}
    techos_por_tipo = {tipo: 0.0 for tipo in TIPOS_TECHO}
    materias_primas = {mat: 0.0 for mat in MATERIAS_PRIMAS}

    suelo_defecto, _ = TIPOS_SUELO
    pared_defecto, _, _, _ = TIPOS_PARED
    techo_defecto, _ = TIPOS_TECHO

    for estancia in estancias:
        es_descuento = bool(estancia.get("es_descuento", False))
        if not es_descuento:
            tipo_s = estancia.get("Tipo de suelo", suelo_defecto)
            sup_suelo = float(estancia["Suelo (m²)"])
            suelos_por_tipo[tipo_s] = suelos_por_tipo.get(tipo_s, 0.0) + sup_suelo
            if tipo_s in materias_primas:
                materias_primas[tipo_s] += sup_suelo

            tipo_t = estancia.get("Actuación techo", techo_defecto)
            sup_techo = float(estancia["Techo (m²)"])
            techos_por_tipo[tipo_t] = techos_por_tipo.get(tipo_t, 0.0) + sup_techo
            if tipo_t in materias_primas:
                materias_primas[tipo_t] += sup_techo

        tipo_p = estancia.get("Actuación pared", pared_defecto)
        sup_pared = float(estancia["Paredes (m²)"])
        paredes_por_tipo[tipo_p] = paredes_por_tipo.get(tipo_p, 0.0) + sup_pared
        if tipo_p in materias_primas:
            materias_primas[tipo_p] += sup_pared

    return {
        "paredes": paredes,
        "suelos": suelos,
        "techos": techos,
        "suelos_por_tipo": suelos_por_tipo,
        "paredes_por_tipo": paredes_por_tipo,
        "techos_por_tipo": techos_por_tipo,
        "materias_primas": materias_primas,
    }


def resumen_instalaciones(
    estancias_principales: list[dict],
    instalaciones_por_id: dict[int, dict[str, int]],
) -> tuple[list[dict], dict[str, int]]:
    acumulado: dict[str, int] = {}
    for items in CATALOGO_INSTALACIONES.values():
        for clave, _ in items:
            acumulado[clave] = 0

    filas_estancia: list[dict] = []
    for est in estancias_principales:
        eid = int(est["id"])
        datos_est = instalaciones_por_id.get(eid, {})
        total_est = 0
        fila: dict[str, object] = {"id": eid, "Estancia": est["Estancia"]}
        for items in CATALOGO_INSTALACIONES.values():
            for clave, _ in items:
                cant = int(datos_est.get(clave, 0))
                fila[clave] = cant
                acumulado[clave] += cant
                total_est += cant
        fila["Total puntos"] = total_est
        if total_est > 0:
            filas_estancia.append(fila)

    return filas_estancia, acumulado


def construir_partidas_valorables(
    totales_sup: dict,
    acumulado_inst: dict[str, int],
    obra_general: dict[str, float],
) -> list[dict]:
    partidas: list[dict] = []
    for mat in MATERIAS_PRIMAS:
        cant = float(totales_sup["materias_primas"].get(mat, 0.0))
        partidas.append(
            {
                "clave": f"mat_{mat}",
                "capitulo": "1. Revestimientos y Acabados",
                "concepto": f"Mano de obra y material base: {mat}",
                "unidad": "m²",
                "cantidad": round(cant, 2),
            }
        )
    for familia, items in CATALOGO_INSTALACIONES.items():
        for clave, etiqueta in items:
            cant = int(acumulado_inst.get(clave, 0))
            if cant > 0:
                partidas.append(
                    {
                        "clave": clave,
                        "capitulo": f"2. Instalaciones ({familia})",
                        "concepto": etiqueta,
                        "unidad": "ud.",
                        "cantidad": float(cant),
                    }
                )
    for familia, items in CATALOGO_OBRA_GENERAL.items():
        for clave, etiqueta, unidad in items:
            cant = float(obra_general.get(clave, 0.0))
            if cant > 0:
                partidas.append(
                    {
                        "clave": clave,
                        "capitulo": f"3. {familia}",
                        "concepto": etiqueta,
                        "unidad": unidad,
                        "cantidad": round(cant, 2),
                    }
                )
    return partidas


def comprobar_advertencias_medidas(largo: float, ancho: float, altura: float) -> list[str]:
    avisos = []
    if largo > LIMITE_AVISO_PLANTA or ancho > LIMITE_AVISO_PLANTA:
        avisos.append(
            f"Has introducido una dimensión superior a {formato_numero(LIMITE_AVISO_PLANTA, 0)} m "
            f"(Largo: {formato_numero(largo)} m, Ancho: {formato_numero(ancho)} m). "
            "Verifica que no falte una coma decimal (por ejemplo, 11,00 m en lugar de 1,10 m)."
        )
    if altura < LIMITE_MIN_ALTURA_AVISO or altura > LIMITE_MAX_ALTURA_AVISO:
        avisos.append(
            f"La altura indicada ({formato_numero(altura)} m) está fuera del rango habitual "
            f"({formato_numero(LIMITE_MIN_ALTURA_AVISO)} m – {formato_numero(LIMITE_MAX_ALTURA_AVISO)} m)."
        )
    return avisos
def limpiar_pdf(texto: object) -> str:
    reemplazos = {
        "«": '"',
        "»": '"',
        "“": '"',
        "”": '"',
        "‘": "'",
        "’": "'",
        "—": "-",
        "–": "-",
        "…": "...",
        "€": "EUR",
    }
    cadena = str(texto)
    for origen, destino in reemplazos.items():
        cadena = cadena.replace(origen, destino)
    return cadena.encode("latin-1", errors="replace").decode("latin-1")


def pdf_texto(pdf: FPDF, texto: str, ancho: float, alto: float = 6) -> None:
    pdf.cell(ancho, alto, limpiar_pdf(texto), border=0, ln=1)


def pdf_linea_acumulado(pdf: FPDF, x_inicio: float, etiqueta: str, cifra: str) -> None:
    pdf.set_x(x_inicio)
    pdf.set_font("Helvetica", "", 8.5)
    texto_etiqueta = limpiar_pdf(f"{etiqueta}: ")
    ancho_etiqueta = pdf.get_string_width(texto_etiqueta)
    pdf.cell(ancho_etiqueta, 6, texto_etiqueta, border=0, ln=0)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.cell(0, 6, limpiar_pdf(cifra), border=0, ln=1)


def generar_pdf_sintesis_cliente(
    estancias: list[dict],
    cliente: str,
    direccion: str,
    fecha_proyecto: date,
    generado: datetime,
    filas_presupuesto: list[dict],
    iva_porcentaje: int,
) -> bytes:
    if FPDF is None:
        raise RuntimeError("Instala la dependencia fpdf2 con: pip install fpdf2")

    pdf = FPDF("P", "mm", "A4")
    pdf.set_margins(15, 15, 15)
    pdf.set_auto_page_break(True, margin=15)
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, limpiar_pdf("VALORACIÓN ESTIMATIVA DE REFORMA (SÍNTESIS)"), ln=1)
    pdf.set_font("Helvetica", "I", 9)
    pdf.cell(0, 6, limpiar_pdf("Documento orientativo sin validez contractual"), ln=1)
    pdf.ln(3)

    pdf.set_font("Helvetica", "", 10)
    pdf_texto(pdf, f"Cliente: {cliente or 'Sin cliente'}", 0)
    pdf_texto(pdf, f"Vivienda / Obra: {direccion or 'No indicada'}", 0)
    pdf_texto(
        pdf,
        f"Fecha del proyecto: {formato_fecha(fecha_proyecto)}   |   Emitido: {formato_fecha_hora(generado)}",
        0,
    )
    pdf.ln(4)

    totales = calcular_totales(estancias)
    estancias_reales = [e for e in estancias if not e.get("es_descuento", False)]

    pdf.set_font("Helvetica", "B", 11)
    pdf.set_fill_color(230, 236, 245)
    pdf.cell(0, 8, limpiar_pdf("1. Resumen de superficies medidas"), border=1, ln=1, fill=True)
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(
        0,
        7,
        limpiar_pdf(
            f"Estancias medidas: {len(estancias_reales)}   |   "
            f"Suelos: {formato_numero(totales['suelos'])} m²   |   "
            f"Paredes: {formato_numero(totales['paredes'])} m²   |   "
            f"Techos: {formato_numero(totales['techos'])} m²"
        ),
        border=1,
        ln=1,
    )
    pdf.ln(4)

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, limpiar_pdf("2. Síntesis económica por capítulos"), border=1, ln=1, fill=True)

    totales_capitulo: dict[str, float] = {}
    for item in filas_presupuesto:
        cap = item["capitulo"]
        totales_capitulo[cap] = totales_capitulo.get(cap, 0.0) + float(item["importe"])

    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(242, 245, 250)
    pdf.cell(135, 8, "Capítulo de obra", border=1, align="L", fill=True)
    pdf.cell(45, 8, "Importe estimado", border=1, align="R", fill=True)
    pdf.ln()

    pdf.set_font("Helvetica", "", 9)
    base_imponible = 0.0
    for cap, importe_cap in totales_capitulo.items():
        base_imponible += importe_cap
        pdf.cell(135, 7, limpiar_pdf(cap), border=1, align="L")
        pdf.cell(45, 7, limpiar_pdf(formato_euro(importe_cap)), border=1, align="R")
        pdf.ln()

    cuota_iva = round(base_imponible * (iva_porcentaje / 100.0), 2)
    total_estimado = base_imponible + cuota_iva

    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(135, 8, "BASE ESTIMADA (Sin IVA)", border=1, align="R", fill=True)
    pdf.cell(45, 8, limpiar_pdf(formato_euro(base_imponible)), border=1, align="R", fill=True)
    pdf.ln()
    pdf.cell(135, 8, f"IVA estimado ({iva_porcentaje}%)", border=1, align="R", fill=True)
    pdf.cell(45, 8, limpiar_pdf(formato_euro(cuota_iva)), border=1, align="R", fill=True)
    pdf.ln()
    pdf.set_fill_color(220, 230, 242)
    pdf.cell(135, 9, "VALOR ESTIMATIVO TOTAL", border=1, align="R", fill=True)
    pdf.cell(45, 9, limpiar_pdf(formato_euro(total_estimado)), border=1, align="R", fill=True)
    pdf.ln(8)

    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(0, 6, limpiar_pdf("NOTA IMPORTANTE SOBRE EL ALCANCE DE ESTA VALORACIÓN:"), ln=1)
    pdf.set_font("Helvetica", "", 8.5)
    nota = (
        "1. Este documento constituye exclusivamente una valoración estimativa previa basada en las mediciones "
        "tomadas en visita, y en ningún caso tiene carácter de presupuesto cerrado ni compromiso contractual.\n"
        "2. Los importes reflejados NO incluyen el suministro de materiales de acabado y equipamiento cuya elección "
        "depende del cliente (pavimentos y revestimientos cerámicos, tarimas, sanitarios, griferías, mecanismos "
        "eléctricos, luminarias, carpinterías especiales o electrodomésticos), cuyo coste final variará en función "
        "de las calidades y modelos seleccionados."
    )
    pdf.multi_cell(0, 5, limpiar_pdf(nota), border=1)

    return bytes(pdf.output())
def generar_pdf(
    estancias: list[dict],
    cliente: str,
    direccion: str,
    fecha_proyecto: date,
    generado: datetime,
    filas_inst: list[dict],
    acumulado_inst: dict[str, int],
    filas_presupuesto: list[dict],
    iva_porcentaje: int,
) -> bytes:
    if FPDF is None:
        raise RuntimeError("Instala la dependencia fpdf2 con: pip install fpdf2")

    pdf = FPDF("L", "mm", "A4")
    pdf.set_margins(10, 10, 10)
    pdf.set_auto_page_break(True, margin=12)
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 15)
    pdf.cell(0, 9, limpiar_pdf("Informe técnico de mediciones y superficies de vivienda"), ln=1)

    pdf.set_font("Helvetica", "", 9)
    pdf_texto(pdf, f"Cliente: {cliente or 'Sin cliente'}", 0)
    pdf_texto(pdf, f"Dirección: {direccion or 'No indicada'}", 0)
    pdf_texto(
        pdf,
        f"Fecha del proyecto: {formato_fecha(fecha_proyecto)}   |   Generado: {formato_fecha_hora(generado)}",
        0,
    )
    pdf.ln(3)

    w_est, w_sue_t, w_par_t, w_tec_t = 36, 36, 36, 23
    w_lar, w_anc, w_alt = 21, 21, 20
    w_par, w_sue, w_tec = 28, 28, 28
    anchos = (w_est, w_sue_t, w_par_t, w_tec_t, w_lar, w_anc, w_alt, w_par, w_sue, w_tec)
    ancho_etiqueta_total = w_est + w_sue_t + w_par_t + w_tec_t + w_lar + w_anc + w_alt

    cabeceras = (
        "Estancia",
        "Tipo de suelo",
        "Act. pared",
        "Act. techo",
        "Largo (m)",
        "Ancho (m)",
        "Alto (m)",
        "Paredes (m²)",
        "Suelo (m²)",
        "Techo (m²)",
    )

    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(220, 230, 240)
    for titulo, ancho in zip(cabeceras, anchos):
        pdf.cell(ancho, 8, limpiar_pdf(titulo), border=1, align="C", fill=True)
    pdf.ln()

    pared_defecto, _, _, _ = TIPOS_PARED
    techo_defecto, _ = TIPOS_TECHO

    pdf.set_font("Helvetica", "", 8)
    for estancia in estancias:
        es_desc = bool(estancia.get("es_descuento", False))
        valores = (
            estancia["Estancia"],
            "—" if es_desc else estancia["Tipo de suelo"],
            estancia.get("Actuación pared", pared_defecto),
            "—" if es_desc else estancia.get("Actuación techo", techo_defecto),
            f"{formato_numero(estancia['Largo (m)'])} m",
            "—" if es_desc else f"{formato_numero(estancia['Ancho (m)'])} m",
            f"{formato_numero(estancia['Altura (m)'])} m",
            f"{formato_numero(estancia['Paredes (m²)'])} m²",
            f"{formato_numero(estancia['Suelo (m²)'])} m²",
            f"{formato_numero(estancia['Techo (m²)'])} m²",
        )
        for indice, (valor, ancho) in enumerate(zip(valores, anchos)):
            pdf.cell(ancho, 7, limpiar_pdf(valor), border=1, align="L" if indice < 4 else "R")
        pdf.ln()

    totales = calcular_totales(estancias)
    pdf.ln(3)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(220, 230, 240)
    pdf.cell(ancho_etiqueta_total, 8, "TOTALES GENERALES", border=1, align="R", fill=True)
    pdf.cell(w_par, 8, limpiar_pdf(f"{formato_numero(totales['paredes'])} m²"), border=1, align="R", fill=True)
    pdf.cell(w_sue, 8, limpiar_pdf(f"{formato_numero(totales['suelos'])} m²"), border=1, align="R", fill=True)
    pdf.cell(w_tec, 8, limpiar_pdf(f"{formato_numero(totales['techos'])} m²"), border=1, align="R", fill=True)
    pdf.ln(8)

    y_desgloses = pdf.get_y()
    if y_desgloses > 155:
        pdf.add_page()
        y_desgloses = pdf.get_y()

    pdf.set_xy(10, y_desgloses)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.cell(64, 7, "Desglose de suelos", ln=1)
    for tipo, superficie in totales["suelos_por_tipo"].items():
        pdf_linea_acumulado(pdf, 10, tipo, f"{formato_numero(superficie)} m²")

    pdf.set_xy(76, y_desgloses)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.cell(68, 7, "Desglose de paredes", ln=1)
    for tipo, superficie in totales["paredes_por_tipo"].items():
        pdf_linea_acumulado(pdf, 76, tipo, f"{formato_numero(superficie)} m²")

    pdf.set_xy(146, y_desgloses)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.cell(62, 7, "Desglose de techos", ln=1)
    for tipo, superficie in totales["techos_por_tipo"].items():
        pdf_linea_acumulado(pdf, 146, tipo, f"{formato_numero(superficie)} m²")

    pdf.set_xy(210, y_desgloses)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.cell(75, 7, "RESUMEN MATERIAS PRIMAS", ln=1)
    for mat, superficie in totales["materias_primas"].items():
        pdf_linea_acumulado(pdf, 210, mat, f"{formato_numero(superficie)} m²")

    if filas_inst:
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 14)
        pdf.cell(0, 9, limpiar_pdf("Anexo 1: Resumen de instalaciones por estancia"), ln=1)
        pdf.ln(2)
        for fila in filas_inst:
            pdf.set_font("Helvetica", "B", 10)
            pdf.set_fill_color(235, 240, 248)
            pdf.cell(0, 7, limpiar_pdf(f"Estancia: {fila['Estancia']} (Total: {fila['Total puntos']} ud.)"), border=1, ln=1, fill=True)
            pdf.set_font("Helvetica", "", 8)
            detalles = []
            for items in CATALOGO_INSTALACIONES.values():
                for clave, etiqueta in items:
                    cant = int(fila.get(clave, 0))
                    if cant > 0:
                        detalles.append(f"{etiqueta}: {cant} ud.")
            pdf.multi_cell(0, 5, limpiar_pdf("   |   ".join(detalles)), border=1)
            pdf.ln(2)

        pdf.ln(3)
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 7, "Acumulado total de instalaciones en vivienda", ln=1)
        for familia, items in CATALOGO_INSTALACIONES.items():
            activos = [(et, acumulado_inst[cl]) for cl, et in items if acumulado_inst.get(cl, 0) > 0]
            if activos:
                pdf.set_font("Helvetica", "B", 9)
                pdf.cell(0, 6, limpiar_pdf(familia), ln=1)
                for et, cant in activos:
                    pdf_linea_acumulado(pdf, 14, et, f"{cant} ud.")

    if filas_presupuesto:
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 14)
        pdf.cell(0, 9, limpiar_pdf("Anexo 2: Valoración detallada por partidas"), ln=1)
        pdf.ln(2)

        wp_cap, wp_con, wp_ud, wp_can, wp_pre, wp_imp = 65, 112, 16, 26, 28, 30
        anchos_p = (wp_cap, wp_con, wp_ud, wp_can, wp_pre, wp_imp)
        ancho_etiqueta_pres = wp_cap + wp_con + wp_ud + wp_can + wp_pre

        cab_p = ("Capítulo", "Concepto", "Ud.", "Cantidad", "Precio Ud.", "Importe")
        pdf.set_font("Helvetica", "B", 8)
        pdf.set_fill_color(220, 230, 240)
        for t, a in zip(cab_p, anchos_p):
            pdf.cell(a, 8, limpiar_pdf(t), border=1, align="C", fill=True)
        pdf.ln()

        pdf.set_font("Helvetica", "", 8)
        base_imponible = 0.0
        for item in filas_presupuesto:
            base_imponible += item["importe"]
            vals = (
                item["capitulo"],
                item["concepto"],
                item["unidad"],
                formato_numero(item["cantidad"]),
                formato_euro(item["precio"]),
                formato_euro(item["importe"]),
            )
            for idx, (v, a) in enumerate(zip(vals, anchos_p)):
                alineacion = "L" if idx < 2 else ("C" if idx == 2 else "R")
                pdf.cell(a, 7, limpiar_pdf(v), border=1, align=alineacion)
            pdf.ln()

        cuota_iva = round(base_imponible * (iva_porcentaje / 100.0), 2)
        total_final = base_imponible + cuota_iva

        pdf.ln(3)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(ancho_etiqueta_pres, 7, "BASE IMPONIBLE", border=1, align="R", fill=True)
        pdf.cell(wp_imp, 7, limpiar_pdf(formato_euro(base_imponible)), border=1, align="R", fill=True)
        pdf.ln()
        pdf.cell(ancho_etiqueta_pres, 7, f"IVA ({iva_porcentaje}%)", border=1, align="R", fill=True)
        pdf.cell(wp_imp, 7, limpiar_pdf(formato_euro(cuota_iva)), border=1, align="R", fill=True)
        pdf.ln()
        pdf.cell(ancho_etiqueta_pres, 8, "TOTAL VALORACIÓN", border=1, align="R", fill=True)
        pdf.cell(wp_imp, 8, limpiar_pdf(formato_euro(total_final)), border=1, align="R", fill=True)
        pdf.ln()

    return bytes(pdf.output())
def generar_xlsx(
    estancias: list[dict],
    cliente: str,
    direccion: str,
    fecha_proyecto: date,
    generado: datetime,
    filas_inst: list[dict],
    filas_presupuesto: list[dict],
    iva_porcentaje: int,
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "1. Superficies"

    ws["A1"] = "Informe de mediciones y superficies de vivienda"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:D1")
    ws.row_dimensions.height = 28

    datos = (
        ("Cliente", cliente or "Sin cliente"),
        ("Dirección", direccion or "No indicada"),
        ("Fecha del proyecto", fecha_proyecto),
        ("Generado", generado.replace(tzinfo=None)),
    )

    for fila, (etiqueta, valor) in enumerate(datos, start=3):
        ws.cell(fila, 1, etiqueta).font = Font(bold=True)
        celda_valor = ws.cell(fila, 2, valor)
        if etiqueta in ("Cliente", "Dirección"):
            ws.merge_cells(start_row=fila, start_column=2, end_row=fila, end_column=4)
        elif etiqueta == "Fecha del proyecto":
            celda_valor.number_format = "dd/mm/yy"
        elif etiqueta == "Generado":
            celda_valor.number_format = "dd/mm/yy hh:mm"

    inicio = 9
    columnas = (
        "Estancia",
        "Tipo de suelo",
        "Actuación pared",
        "Actuación techo",
        "Largo (m)",
        "Ancho (m)",
        "Altura (m)",
        "Paredes (m²)",
        "Suelo (m²)",
        "Techo (m²)",
    )

    for columna, texto in enumerate(columnas, 1):
        celda = ws.cell(inicio, columna, texto)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="1F4E78")
        celda.alignment = Alignment(horizontal="center")

    pared_defecto, _, _, _ = TIPOS_PARED
    techo_defecto, _ = TIPOS_TECHO

    for fila_datos, estancia in enumerate(estancias, inicio + 1):
        es_desc = bool(estancia.get("es_descuento", False))
        valores = (
            estancia["Estancia"],
            "—" if es_desc else estancia["Tipo de suelo"],
            estancia.get("Actuación pared", pared_defecto),
            "—" if es_desc else estancia.get("Actuación techo", techo_defecto),
            estancia["Largo (m)"],
            "—" if es_desc else estancia["Ancho (m)"],
            estancia["Altura (m)"],
            estancia["Paredes (m²)"],
            estancia["Suelo (m²)"],
            estancia["Techo (m²)"],
        )
        for columna, valor in enumerate(valores, 1):
            celda = ws.cell(fila_datos, columna, valor)
            if isinstance(valor, (int, float)):
                celda.alignment = Alignment(horizontal="right")
                celda.number_format = FORMATO_NUMERO_EXCEL
            else:
                celda.alignment = Alignment(horizontal="left" if columna <= 4 else "center")

    fila_total = inicio + len(estancias) + 1
    ws.cell(fila_total, 1, "TOTALES GENERALES").font = Font(bold=True)

    totales = calcular_totales(estancias)
    valores_totales = {8: totales["paredes"], 9: totales["suelos"], 10: totales["techos"]}

    for columna, valor in valores_totales.items():
        celda = ws.cell(fila_total, columna, valor)
        celda.font = Font(bold=True)
        celda.alignment = Alignment(horizontal="right")
        celda.number_format = FORMATO_NUMERO_EXCEL

    for columna in range(1, 11):
        ws.cell(fila_total, columna).fill = PatternFill("solid", fgColor="D9E1F2")

    bloques_desglose = (
        ("Desglose de suelos por material", totales["suelos_por_tipo"]),
        ("Desglose de paredes por actuación", totales["paredes_por_tipo"]),
        ("Desglose de techos por actuación", totales["techos_por_tipo"]),
        ("RESUMEN TOTAL POR MATERIA PRIMA", totales["materias_primas"]),
    )

    fila_cursor = fila_total + 3
    for titulo_bloque, diccionario in bloques_desglose:
        ws.cell(fila_cursor, 1, titulo_bloque).font = Font(bold=True)
        celda_cab = ws.cell(fila_cursor, 2, "Superficie (m²)")
        celda_cab.font = Font(bold=True)
        celda_cab.alignment = Alignment(horizontal="right")
        fila_cursor += 1
        for tipo, superficie in diccionario.items():
            ws.cell(fila_cursor, 1, tipo)
            celda_sup = ws.cell(fila_cursor, 2, superficie)
            celda_sup.alignment = Alignment(horizontal="right")
            celda_sup.number_format = FORMATO_NUMERO_EXCEL
            fila_cursor += 1
        fila_cursor += 1

    if estancias:
        ultima_fila_tabla = inicio + len(estancias)
        ws.auto_filter.ref = f"A{inicio}:J{ultima_fila_tabla}"

    ws.freeze_panes = f"A{inicio + 1}"

    anchos_hoja1 = {"A": 25, "B": 28, "C": 22, "D": 18, "E": 13, "F": 13, "G": 13, "H": 16, "I": 16, "J": 16}
    for col, ancho in anchos_hoja1.items():
        ws.column_dimensions[col].width = ancho

    for fila_excel in ws.iter_rows():
        for celda in fila_excel:
            celda.alignment = Alignment(
                horizontal=celda.alignment.horizontal or "left",
                vertical="center",
                wrap_text=True,
            )

    if filas_inst:
        ws2 = wb.create_sheet("2. Instalaciones")
        ws2["A1"] = "Resumen de instalaciones por estancia"
        ws2["A1"].font = Font(bold=True, size=14)
        ws2.merge_cells("A1:E1")

        claves_inst = []
        cabeceras_inst = ["Estancia"]
        for items in CATALOGO_INSTALACIONES.values():
            for clave, etiqueta in items:
                claves_inst.append(clave)
                cabeceras_inst.append(etiqueta)
        cabeceras_inst.append("Total puntos")

        for col_idx, texto in enumerate(cabeceras_inst, 1):
            c = ws2.cell(3, col_idx, texto)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F4E78")
            c.alignment = Alignment(horizontal="center", wrap_text=True)
            ws2.column_dimensions[c.column_letter].width = 22 if col_idx == 1 else 15

        for r_idx, fila in enumerate(filas_inst, start=4):
            ws2.cell(r_idx, 1, fila["Estancia"])
            for c_idx, clave in enumerate(claves_inst, start=2):
                celda = ws2.cell(r_idx, c_idx, int(fila.get(clave, 0)))
                celda.alignment = Alignment(horizontal="right")
            c_tot = ws2.cell(r_idx, len(cabeceras_inst), int(fila["Total puntos"]))
            c_tot.font = Font(bold=True)
            c_tot.alignment = Alignment(horizontal="right")

        fila_tot_inst = 4 + len(filas_inst)
        ws2.cell(fila_tot_inst, 1, "TOTAL VIVIENDA").font = Font(bold=True)
        for c_idx in range(2, len(cabeceras_inst) + 1):
            col_letra = ws2.cell(3, c_idx).column_letter
            c = ws2.cell(fila_tot_inst, c_idx, f"=SUM({col_letra}4:{col_letra}{fila_tot_inst - 1})")
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="right")

    if filas_presupuesto:
        ws3 = wb.create_sheet("3. Valoración")
        ws3["A1"] = "Valoración estimada de obra"
        ws3["A1"].font = Font(bold=True, size=14)
        ws3.merge_cells("A1:F1")

        cab_p = ("Capítulo", "Concepto", "Unidad", "Cantidad", "Precio unitario", "Importe")
        anchos_p = {"A": 30, "B": 42, "C": 10, "D": 14, "E": 18, "F": 18}
        for col_idx, texto in enumerate(cab_p, 1):
            c = ws3.cell(3, col_idx, texto)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F4E78")
            c.alignment = Alignment(horizontal="center")
        for col_l, ancho in anchos_p.items():
            ws3.column_dimensions[col_l].width = ancho

        for r_idx, item in enumerate(filas_presupuesto, start=4):
            ws3.cell(r_idx, 1, item["capitulo"])
            ws3.cell(r_idx, 2, item["concepto"])
            c_ud = ws3.cell(r_idx, 3, item["unidad"])
            c_ud.alignment = Alignment(horizontal="center")
            c_cant = ws3.cell(r_idx, 4, float(item["cantidad"]))
            c_cant.number_format = FORMATO_NUMERO_EXCEL
            c_cant.alignment = Alignment(horizontal="right")
            c_prec = ws3.cell(r_idx, 5, float(item["precio"]))
            c_prec.number_format = FORMATO_MONEDA_EXCEL
            c_prec.alignment = Alignment(horizontal="right")
            c_imp = ws3.cell(r_idx, 6, f"=D{r_idx}*E{r_idx}")
            c_imp.number_format = FORMATO_MONEDA_EXCEL
            c_imp.alignment = Alignment(horizontal="right")

        f_base = 4 + len(filas_presupuesto)
        ws3.cell(f_base, 5, "BASE ESTIMADA").font = Font(bold=True)
        c_base = ws3.cell(f_base, 6, f"=SUM(F4:F{f_base - 1})")
        c_base.font = Font(bold=True)
        c_base.number_format = FORMATO_MONEDA_EXCEL

        f_iva = f_base + 1
        ws3.cell(f_iva, 5, f"IVA ({iva_porcentaje}%)").font = Font(bold=True)
        c_iva = ws3.cell(f_iva, 6, f"=F{f_base}*{iva_porcentaje}/100")
        c_iva.font = Font(bold=True)
        c_iva.number_format = FORMATO_MONEDA_EXCEL

        f_tot = f_base + 2
        ws3.cell(f_tot, 5, "TOTAL ESTIMADO").font = Font(bold=True)
        c_tot = ws3.cell(f_tot, 6, f"=F{f_base}+F{f_iva}")
        c_tot.font = Font(bold=True)
        c_tot.number_format = FORMATO_MONEDA_EXCEL
        c_tot.fill = PatternFill("solid", fgColor="D9E1F2")

    salida = BytesIO()
    wb.save(salida)
    return salida.getvalue()
def aplicar_estilos_caratula() -> None:
    st.markdown(
        """
        <style>
        .stApp {
            background-color: #111827;
            color: #F3F4F6;
        }
        div[data-baseweb="input"] > div,
        div[data-baseweb="base-input"] > input {
            background-color: #1F2937 !important;
            border-color: #374151 !important;
            border-radius: 10px !important;
            color: #F9FAFB !important;
        }
        div[data-baseweb="input"] > div:focus-within {
            border-color: #2563EB !important;
            box-shadow: 0 0 0 1px #2563EB !important;
        }
        .stButton > button {
            background-color: #2563EB;
            color: #FFFFFF;
            border: none;
            border-radius: 10px;
            padding: 0.70rem 1.2rem;
            font-weight: 600;
            font-size: 1rem;
        }
        .stButton > button:hover {
            background-color: #1D4ED8;
            color: #FFFFFF;
        }
        .caratula-box {
            background-color: #161E2E;
            border: 1px solid #2A3548;
            border-radius: 14px;
            padding: 1.5rem 1.5rem 1.2rem 1.5rem;
            margin-bottom: 1rem;
        }
        .caratula-titulo {
            font-size: 1.75rem;
            font-weight: 700;
            color: #FFFFFF;
            margin-bottom: 0.25rem;
        }
        .caratula-subtitulo {
            font-size: 0.95rem;
            color: #9CA3AF;
            margin-bottom: 0.25rem;
        }
        .estado-barra {
            color: #FBBF24;
            font-size: 0.92rem;
            font-weight: 500;
            margin-top: 0.75rem;
            padding-top: 0.5rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_caratula_inicial() -> None:
    st.markdown(
        """
        <div class="caratula-box">
            <div class="caratula-titulo">Preparar medición de obra</div>
            <div class="caratula-subtitulo">Configura los datos del expediente antes de iniciar la toma de medidas</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.text_input(
        "Cliente o referencia",
        key="cliente",
        placeholder="Ej.: Reforma Integral Calle Mayor",
    )
    st.text_input(
        "Dirección / inmueble (opcional)",
        key="direccion",
        placeholder="Ej.: Madrid · Planta 3.ª B",
    )
    st.date_input(
        "Fecha del proyecto",
        key="fecha_proyecto",
        format="DD/MM/YYYY",
    )

    st.markdown("<div style='height: 0.6rem;'></div>", unsafe_allow_html=True)

    if st.button("Iniciar medición y valorar obra", use_container_width=True):
        st.session_state.estado_app = "Preparado para medir estancias"
        st.session_state.vista_activa = "trabajo"
        st.rerun()

    if st.button("Cargar vivienda tipo de ejemplo", use_container_width=True):
        if not st.session_state.cliente:
            st.session_state.cliente = "Vivienda Demo (85 m²)"
        if not st.session_state.direccion:
            st.session_state.direccion = "Calle Ejemplo 14, 2.º A"
        estancias_demo = [
            ("Salón", "Tarima", "Pintura", "Pladur", 5.50, 4.00, 2.60),
            ("Cocina", "Porcelánico / Cerámico", "Porcelánico / Cerámico", "Pladur", 3.60, 2.40, 2.60),
            ("Baño", "Porcelánico / Cerámico", "Porcelánico / Cerámico", "Pladur", 2.40, 1.80, 2.50),
            ("Dormitorio 1", "Tarima", "Pintura", "Pintura", 3.80, 3.10, 2.60),
        ]
        st.session_state.estancias = []
        for idx, (nom, sue, par, tec, lar, anc, alt) in enumerate(estancias_demo, start=1):
            d = calcular_estancia(lar, anc, alt)
            d.update({
                "id": idx,
                "id_padre": None,
                "es_descuento": False,
                "Estancia": nom,
                "Tipo de suelo": sue,
                "Actuación pared": par,
                "Actuación techo": tec,
            })
            st.session_state.estancias.append(d)
        st.session_state.estado_app = "Ejemplo cargado · Listo para editar o exportar"
        st.session_state.vista_activa = "trabajo"
        st.rerun()

    if st.button("Consultar catálogo de partidas y unidades", use_container_width=True):
        st.session_state.mostrar_catalogo = not st.session_state.get("mostrar_catalogo", False)

    if st.session_state.get("mostrar_catalogo", False):
        with st.expander("Catálogo maestro de acabados e instalaciones", expanded=True):
            st.markdown(f"**Materias primas:** {', '.join(MATERIAS_PRIMAS)}")
            for fam, items in CATALOGO_INSTALACIONES.items():
                st.markdown(f"**{fam}:** " + ", ".join(et for _, et in items))
            for fam, items in CATALOGO_OBRA_GENERAL.items():
                st.markdown(f"**{fam}:** " + ", ".join(f"{et} ({ud})" for _, et, ud in items))

    st.markdown(
        f'<div class="estado-barra">● {st.session_state.estado_app}</div>',
        unsafe_allow_html=True,
    )
def render_tab_instalaciones() -> None:
    st.subheader("⚡ Instalaciones por estancia")
    estancias_principales = [e for e in st.session_state.estancias if not e.get("es_descuento", False)]
    if not estancias_principales:
        st.info("Primero añade al menos una estancia en **1. Superficies**.")
        return

    mapa_estancias = {f"#{e['id']} - {e['Estancia']}": int(e["id"]) for e in estancias_principales}
    seleccion = st.selectbox("Selecciona la estancia a equipar", list(mapa_estancias.keys()))
    eid = mapa_estancias[seleccion]
    datos_actuales = st.session_state.instalaciones.setdefault(eid, {})

    with st.form(f"form_inst_{eid}"):
        nuevos_valores: dict[str, int] = {}
        for familia, items in CATALOGO_INSTALACIONES.items():
            st.markdown(f"**{familia}**")
            cols = st.columns(2)
            for idx, (clave, etiqueta) in enumerate(items):
                with cols[idx % 2]:
                    nuevos_valores[clave] = int(
                        st.number_input(
                            etiqueta,
                            min_value=0,
                            max_value=99,
                            value=int(datos_actuales.get(clave, 0)),
                            step=1,
                            key=f"in_{eid}_{clave}",
                        )
                    )
        if st.form_submit_button("💾 Guardar instalaciones", use_container_width=True):
            st.session_state.instalaciones[eid] = nuevos_valores
            st.success("Instalaciones guardadas.")
            st.rerun()


def render_tab_obra_general() -> None:
    st.subheader("🧱 Albañilería, Carpintería y Cocina")
    with st.form("form_obra_general"):
        nuevos_datos: dict[str, float] = {}
        for familia, items in CATALOGO_OBRA_GENERAL.items():
            st.markdown(f"**{familia}**")
            cols = st.columns(2)
            for idx, (clave, etiqueta, unidad) in enumerate(items):
                with cols[idx % 2]:
                    val_actual = st.session_state.obra_general.get(clave, 0)
                    if unidad in ("ud.", "jor."):
                        nuevos_datos[clave] = float(
                            st.number_input(
                                f"{etiqueta} ({unidad})",
                                min_value=0,
                                max_value=200,
                                value=int(val_actual),
                                step=1,
                                key=f"og_{clave}",
                            )
                        )
                    else:
                        nuevos_datos[clave] = float(
                            st.number_input(
                                f"{etiqueta} ({unidad})",
                                min_value=0.0,
                                max_value=500.0,
                                value=float(val_actual),
                                step=0.10,
                                format="%.2f",
                                key=f"og_{clave}",
                            )
                        )
        if st.form_submit_button("💾 Guardar partidas generales", use_container_width=True):
            st.session_state.obra_general = nuevos_datos
            st.success("Partidas actualizadas.")
            st.rerun()
def render_tab_precios_y_descarga(cliente: str, direccion: str, fecha_proyecto: date) -> None:
    st.subheader("💶 Valoración Estimativa y Descarga de Informes")
    estancias = st.session_state.estancias
    if not estancias:
        st.info("Añade primero las estancias en **1. Superficies**.")
        return

    estancias_principales = [e for e in estancias if not e.get("es_descuento", False)]
    totales_sup = calcular_totales(estancias)
    filas_inst, acumulado_inst = resumen_instalaciones(estancias_principales, st.session_state.instalaciones)
    partidas = construir_partidas_valorables(totales_sup, acumulado_inst, st.session_state.obra_general)

    with st.expander("💰 Precios unitarios (€) — Incluye las 5 materias primas", expanded=True):
        iva_porcentaje = st.selectbox(
            "IVA aplicable",
            options=(10, 21, 0),
            format_func=lambda x: f"{x}% (Reforma)" if x == 10 else (f"{x}% (General)" if x == 21 else "0% (Sin IVA)"),
            key="iva_porcentaje",
        )
        with st.form("form_precios"):
            nuevos_precios: dict[str, float] = {}
            capitulo_actual = None
            for p in partidas:
                if p["capitulo"] != capitulo_actual:
                    capitulo_actual = p["capitulo"]
                    st.markdown(f"**{capitulo_actual}**")
                c_desc, c_cant, c_prec = st.columns((3, 1.2, 1.5))
                c_desc.write(p["concepto"])
                c_cant.write(f"{formato_numero(p['cantidad'])} {p['unidad']}")
                nuevos_precios[p["clave"]] = float(
                    c_prec.number_input(
                        f"Precio {p['clave']}",
                        min_value=0.0,
                        max_value=100000.0,
                        value=float(st.session_state.precios.get(p["clave"], 0.0)),
                        step=1.0,
                        format="%.2f",
                        key=f"pr_{p['clave']}",
                        label_visibility="collapsed",
                    )
                )
            if st.form_submit_button("🧮 Calcular valoración", use_container_width=True):
                st.session_state.precios.update(nuevos_precios)
                st.success("Precios actualizados.")

    filas_presupuesto = []
    base_imponible = 0.0
    for p in partidas:
        precio = float(st.session_state.precios.get(p["clave"], 0.0))
        if precio > 0 and p["cantidad"] > 0:
            importe = round(p["cantidad"] * precio, 2)
            base_imponible += importe
            filas_presupuesto.append(
                {
                    "capitulo": p["capitulo"],
                    "concepto": p["concepto"],
                    "unidad": p["unidad"],
                    "cantidad": p["cantidad"],
                    "precio": precio,
                    "importe": importe,
                }
            )

    if filas_presupuesto:
        cuota_iva = round(base_imponible * (iva_porcentaje / 100.0), 2)
        m1, m2, m3 = st.columns(3)
        m1.metric("Base Estimada", formato_euro(base_imponible))
        m2.metric(f"IVA ({iva_porcentaje}%)", formato_euro(cuota_iva))
        m3.metric("TOTAL ESTIMADO", formato_euro(base_imponible + cuota_iva))

    st.markdown("---")
    generado = ahora_españa()
    pdf_sintesis = generar_pdf_sintesis_cliente(
        estancias, cliente, direccion, fecha_proyecto, generado, filas_presupuesto, iva_porcentaje
    )
    pdf_tecnico = generar_pdf(
        estancias, cliente, direccion, fecha_proyecto, generado, filas_inst, acumulado_inst, filas_presupuesto, iva_porcentaje
    )
    xlsx_bytes = generar_xlsx(
        estancias, cliente, direccion, fecha_proyecto, generado, filas_inst, filas_presupuesto, iva_porcentaje
    )

    st.download_button(
        "📋 Descargar Síntesis Estimativa Cliente (PDF)",
        data=pdf_sintesis,
        file_name=nombre_informe(cliente, generado, "pdf", "Estimacion_Cliente"),
        mime="application/pdf",
        use_container_width=True,
    )
    c_tec, c_xls = st.columns(2)
    with c_tec:
        st.download_button(
            "📄 Descargar Dossier Técnico (PDF)",
            data=pdf_tecnico,
            file_name=nombre_informe(cliente, generado, "pdf", "Dossier_Tecnico"),
            mime="application/pdf",
            use_container_width=True,
        )
    with c_xls:
        st.download_button(
            "📊 Descargar Libro Excel (.xlsx)",
            data=xlsx_bytes,
            file_name=nombre_informe(cliente, generado, "xlsx", "Mediciones"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


def interfaz() -> None:
    st.set_page_config(page_title="Medición y Valoración de Vivienda", page_icon="📐", layout="centered")
    aplicar_estilos_caratula()

    for clave, defecto in (
        ("estancias", []),
        ("instalaciones", {}),
        ("obra_general", {}),
        ("precios", {}),
        ("pendiente_confirmacion", None),
        ("cliente", ""),
        ("direccion", ""),
        ("vista_activa", "caratula"),
        ("estado_app", "Preparado para iniciar medición"),
        ("iva_porcentaje", 10),
    ):
        if clave not in st.session_state:
            st.session_state[clave] = defecto
    if "fecha_proyecto" not in st.session_state:
        st.session_state.fecha_proyecto = ahora_españa().date()

    if st.session_state.vista_activa == "caratula":
        render_caratula_inicial()
        return

    col_t, col_b = st.columns([4, 1.3])
    with col_t:
        st.title("📐 Medición y Valoración")
        st.caption(
            f"**Cliente:** {st.session_state.cliente or 'Sin cliente'}  ·  "
            f"**Obra:** {st.session_state.direccion or 'No indicada'}  ·  "
            f"**Fecha:** {formato_fecha(st.session_state.fecha_proyecto)}"
        )
    with col_b:
        st.markdown("<div style='height: 0.7rem;'></div>", unsafe_allow_html=True)
        if st.button("⚙️ Carátula", use_container_width=True):
            st.session_state.vista_activa = "caratula"
            st.rerun()

    tab1, tab2, tab3, tab4 = st.tabs(
        ("1. 📐 Superficies", "2. ⚡ Instalaciones", "3. 🧱 Albañilería y Cocina", "4. 💶 Valoración y Descarga")
    )
    with tab1:
        render_tab_superficies()
    with tab2:
        render_tab_instalaciones()
    with tab3:
        render_tab_obra_general()
    with tab4:
        render_tab_precios_y_descarga(
            st.session_state.cliente,
            st.session_state.direccion,
            st.session_state.fecha_proyecto,
        )


if __name__ == "__main__":
    interfaz()
