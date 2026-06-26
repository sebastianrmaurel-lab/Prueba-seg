"""
Proceso 1: Pago TXT al Banco
El Excel de referencia tiene TODO en una sola hoja:
  A-G  → datos del PAGOT1.txt (los llenamos desde el txt)
  P    → RUT (cuerpo)
  Q    → DV Nombre
  R    → (Nombre referencia)
  S    → Diferencia (monto)
  T    → MOTIVO
  U    → RUT COMBINADO (lo calculamos nosotros)
"""

import io
import openpyxl

CAMPOS_T1 = [
    ("MODALIDAD DE PAGO",      1,  1),
    ("COD BANCO DESTINO",      2,  3),
    ("CTA ABONO BENEFICIARIO", 5, 18),
    ("MONTO TOTAL ABONO",     23, 13),
    ("RUT BENEFICIARIO",      36, 11),
    ("NOMBRE BENEFICIARIO",   47, 40),
    ("COD SUCURSAL ENTREGA",  87,  3),
]

COL_RUT   = 16   # P
COL_DV    = 17   # Q
COL_MONTO = 19   # S


def paso1_parsear_txt(txt_bytes: bytes) -> list[dict]:
    texto = txt_bytes.decode("latin-1", errors="replace")
    filas = []
    for linea in texto.splitlines():
        if len(linea) >= 46:
            fila = {}
            for nombre, inicio, largo in CAMPOS_T1:
                fila[nombre] = linea[inicio - 1: inicio - 1 + largo].strip()
            filas.append(fila)
    return filas


def paso2_leer_referencia(xls_bytes: bytes) -> list[dict]:
    """Lee cols P(16), Q(17), S(19) del Excel. Ignora filas sin RUT."""
    wb = openpyxl.load_workbook(io.BytesIO(xls_bytes), data_only=True)
    ws = wb.active
    filas = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        rut   = row[COL_RUT - 1]   if len(row) >= COL_RUT   else None
        dv    = row[COL_DV - 1]    if len(row) >= COL_DV    else None
        monto = row[COL_MONTO - 1] if len(row) >= COL_MONTO else None
        if rut is not None and str(rut).strip() not in ("", "None"):
            filas.append({"RUT": rut, "DV": dv, "MONTO": monto})
    return filas


def _limpiar_rut(valor) -> str:
    if valor is None:
        return ""
    try:
        return str(int(float(str(valor))))
    except (ValueError, TypeError):
        return str(valor).strip()


def paso3_combinar_rut(ref_filas: list[dict]) -> list[dict]:
    for fila in ref_filas:
        cuerpo = _limpiar_rut(fila.get("RUT", ""))
        dv     = str(fila.get("DV", "") or "").strip()
        fila["RUT_COMBINADO"] = (cuerpo + dv) if cuerpo and cuerpo != "0" else ""
    return ref_filas


def _limpiar_para_match(rut_str: str) -> str:
    s = str(rut_str).strip().replace("-", "")
    while len(s) > 1 and s.startswith("0"):
        s = s[1:]
    return s


def paso4_match(pagot1_filas: list[dict], ref_filas: list[dict]) -> list[dict]:
    indice = {}
    for i, fila in enumerate(pagot1_filas):
        rut = _limpiar_para_match(fila.get("RUT BENEFICIARIO", ""))
        if rut:
            indice[rut] = i

    for fila in pagot1_filas:
        fila["_matched"] = False

    for ref in ref_filas:
        rut_comb = _limpiar_para_match(ref.get("RUT_COMBINADO", ""))
        if not rut_comb:
            continue
        if rut_comb in indice:
            idx = indice[rut_comb]
            monto_raw = str(ref.get("MONTO", "") or "").strip()
            if monto_raw:
                try:
                    pagot1_filas[idx]["MONTO TOTAL ABONO"] = str(int(float(monto_raw))).zfill(13)
                except ValueError:
                    pagot1_filas[idx]["MONTO TOTAL ABONO"] = monto_raw.zfill(13)
            pagot1_filas[idx]["_matched"] = True

    return pagot1_filas


def paso5_eliminar(filas: list[dict]) -> list[dict]:
    return [f for f in filas if f.get("_matched")]


def paso6_juntar(filas: list[dict]) -> list[dict]:
    return [f for f in filas
            if f.get("RUT BENEFICIARIO", "").strip() or f.get("NOMBRE BENEFICIARIO", "").strip()]


def _campo_fijo(valor: str, ancho: int, relleno_izq: bool = False) -> str:
    v = str(valor or "")
    if len(v) > ancho:
        return v[:ancho]
    return v.zfill(ancho) if relleno_izq else v.ljust(ancho)


def paso7_generar_txt(filas: list[dict]) -> str:
    COLS = [
        ("MODALIDAD DE PAGO",      1,  False),
        ("COD BANCO DESTINO",      3,  False),
        ("CTA ABONO BENEFICIARIO", 18, False),
        ("MONTO TOTAL ABONO",      13, True),
        ("RUT BENEFICIARIO",       11, False),
        ("NOMBRE BENEFICIARIO",    40, False),
        ("COD SUCURSAL ENTREGA",   3,  False),
    ]
    lineas = []
    for fila in filas:
        if not fila.get("RUT BENEFICIARIO", "").strip():
            continue
        linea = "".join(_campo_fijo(fila.get(col, ""), ancho, izq) for col, ancho, izq in COLS)
        lineas.append(linea)
    return "\n".join(lineas)


def procesar_pago1(txt_bytes: bytes, xls_bytes: bytes) -> str:
    pagot1 = paso1_parsear_txt(txt_bytes)
    if not pagot1:
        raise ValueError("PAGOT1.txt no tiene líneas válidas (mínimo 46 caracteres por línea).")

    ref = paso2_leer_referencia(xls_bytes)
    if not ref:
        raise ValueError("El Excel no tiene datos en la columna P desde la fila 2.")

    ref    = paso3_combinar_rut(ref)
    pagot1 = paso4_match(pagot1, ref)
    pagot1 = paso5_eliminar(pagot1)
    pagot1 = paso6_juntar(pagot1)

    if not pagot1:
        raise ValueError("No hubo coincidencias entre los RUTs del TXT y el Excel.")

    return paso7_generar_txt(pagot1)
