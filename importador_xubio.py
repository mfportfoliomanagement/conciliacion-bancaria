# -*- coding: utf-8 -*-
"""
CAPA 3 — Generador del IMPORTADOR DE ASIENTOS para Xubio.

Toma los movimientos ya clasificados (Capa 2) y arma el asiento contable AGRUPADO
por CIRCUITO y por MES, con cada cuenta contra el banco, cerrando debe = haber.

Diseño confirmado leyendo +20 importadores reales del contador (VENEZUELA).
V1: agrupa por cuenta (sin discriminar por nombre en el auxiliar). Solo filas con importe > 0.

Salida: DOS archivos separados (Xubio no acepta un .xlsx con varias solapas):
  - el IMPORTADOR (una sola hoja) -- el que se sube a Xubio.
  - el PAPEL DE TRABAJO (RESUMEN ASIENTOS, A REVISAR, TRANSF. INTERNAS, CONTROL)
    -- para revisión humana, no se sube a ningún lado.

Regla de oro: las cuentas cuyo circuito no está definido con certeza NO se inventan;
van a un bloque "A REVISAR" para imputación humana (ej. AFIP-RENTAS, Caja, Moneda Extranjera).
Tampoco se inventa el banco de una cuenta si no se puede resolver por CBU contra el
Plan de Cuentas del cliente (ver resolver_banco_por_cbu).
"""
import calendar
import datetime
import re
import unicodedata
from collections import defaultdict, OrderedDict

GB  = 'GASTOS BANCARIOS'
IMP = 'PAGO DE IMPUESTOS'
SUE = 'PAGO DE SUELDOS'
PRO = 'PAGO A PROVEEDORES'
COB = 'COBRANZAS Y DEPOSITOS'
RET = 'RETIROS Y CAJA CHICA'
INV = 'INVERSIONES EN BANCO'
OTR = 'OTROS MOVIMIENTOS'
TRP = 'TRANSFERENCIA ENTRE CUENTAS PROPIAS'
REVISAR = '__REVISAR__'   # sin circuito definido: queda para imputación humana
INTERNO = '__INTERNO__'   # transferencia entre cuentas propias: NO va al importador (se excluye a propósito)

# ---------------------------------------------------------------------------
# Normalización tolerante (mayúsculas/acentos/espacios de más — Regla de oro
# #4). Se duplica acá, chica, para que esta capa no dependa de
# conversor_extractos.py (cada Capa es independiente).
# ---------------------------------------------------------------------------
def _sin_acentos(s) -> str:
    return ''.join(c for c in unicodedata.normalize('NFD', str(s)) if unicodedata.category(c) != 'Mn')

def _norm(s: str) -> str:
    s = _sin_acentos(str(s or '')).upper()
    return re.sub(r'\s+', ' ', s).strip().rstrip(' .')

def _valor(fila: dict, *nombres_col):
    idx = {_norm(k): v for k, v in fila.items()}
    for n in nombres_col:
        v = idx.get(_norm(n))
        if v not in (None, ''):
            return v
    return None

def _digitos(s) -> str:
    return re.sub(r'\D', '', str(s or ''))

# ---------------------------------------------------------------------------
# CIRCUITO de cada cuenta contable. El nombre EXACTO de la cuenta ya NO está
# acá (sale del Plan de Cuentas del cliente -- ver cargar_plan_cuentas). El
# plan permite dejar el circuito como config estable revisable en el código:
# son 9 categorías fijas de cómo se arma el asiento, no un dato del cliente.
# ---------------------------------------------------------------------------
CIRCUITO_POR_CUENTA = {
    'Gastos bancarios':                         GB,
    'IVA Crédito Fiscal':                       GB,
    'Impuesto al Crédito Ley 25.413':           GB,
    'Impuesto al Débito Ley 25.413':             GB,
    'Percepción Ingresos Brutos Sufrida':       GB,
    'Percepción de IVA Sufrida':                GB,
    'Sircreb':                                  GB,
    'Retenciones Ingresos Brutos CABA':         GB,

    'Impuestos y Tasas':                        IMP,
    # Cuentas que arma la solapa CONCEPTOS VEP del cliente (ver
    # conversor_extractos.cargar_vep) -- mismos nombres, literales.
    'CARGAS SOCIALES A PAGAR':                  IMP,
    'IVA a pagar':                              IMP,
    'Ingresos brutos a pagar':                  IMP,
    'Intereses fiscales':                       IMP,
    'Impuesto a las Ganancias':                 IMP,

    'SUELDOS A PAGAR':                          SUE,
    'Sueldos a pagar':                          SUE,

    'Proveedores':                              PRO,
    'Acreedores Varios':                        PRO,

    'Deudores por Venta':                       COB,
    'Deudores por venta':                       COB,

    'EATON DIEGO MARTIN - Cuenta Particular':   RET,
    'SANCHEZ EDUARDO OMAR - Cuenta Particular': RET,

    'FCI':                                      INV,
    'Resultado por Inversión':                  INV,

    'Agua':                                     OTR,
    'Energía Eléctrica':                        OTR,
    'Telefonia':                                OTR,
    'ABL':                                      OTR,
    'Expensas':                                 OTR,
    'Seguros':                                  OTR,
    'GASTOS VARIOS':                            OTR,
    'SEGURIDAD':                                OTR,

    'Transferencias entre cuentas':             TRP,

    # --- Cuentas de Banco propio: transferencia entre cuentas propias, NO
    # va al importador (se registra en la cuenta que RECIBE; se excluye en
    # la que envía). ---
    'Banco Galicia en $':                       INTERNO,
    'Banco':                                    INTERNO,
    'Banco HSBC':                               INTERNO,

    # --- Sin circuito definido con certeza: van a revisión humana ---
    'AFIP-RENTAS':                              REVISAR,   # se desglosa con el VEP
    'Caja':                                     REVISAR,   # ajuste manual del contador
    'Moneda Extranjera':                        REVISAR,
}
_CIRCUITO_NORM = {_norm(k): v for k, v in CIRCUITO_POR_CUENTA.items()}

def _circuito_de(cuenta_interna: str):
    return _CIRCUITO_NORM.get(_norm(cuenta_interna))

# Orden en que se listan los circuitos en la salida (como en los importadores reales)
ORDEN_CIRCUITOS = [GB, RET, TRP, IMP, SUE, PRO, COB, INV, OTR]


def cargar_plan_cuentas(solapas: dict) -> dict:
    """`solapas` = {nombre_de_solapa: [filas...]} del archivo PLAN DE
    CUENTAS del cliente (ej. leídas con drive_io.leer_sheet). Devuelve:
      nombres : {nombre normalizado -> nombre EXACTO}  (para usar en Xubio)
      filas   : las filas de la solapa de cuentas (para buscar el CBU)."""
    for filas in solapas.values():
        if not filas:
            continue
        cols = {_norm(k) for k in filas[0].keys()}
        if _norm('Nombre') in cols:
            nombres = {}
            for fila in filas:
                nom = _valor(fila, 'Nombre')
                if nom:
                    nombres[_norm(nom)] = str(nom).strip()
            return dict(nombres=nombres, filas=filas)
    return dict(nombres={}, filas=[])


def _nucleo_cuenta(cuenta_extracto: str) -> str:
    """Primer bloque de dígitos del número de cuenta del extracto, sin ceros
       a la izquierda -- ej. '0008995-1 137-7' -> '8995'. Es el fragmento
       que más confiablemente aparece, literal, dentro del CBU (que a veces
       pierde un cero a la izquierda al guardarse como número en Sheets)."""
    primero = re.split(r'[-\s]', str(cuenta_extracto or '').strip())[0]
    return _digitos(primero).lstrip('0')


def resolver_banco_por_cbu(cuenta_extracto: str, plan_cuentas: dict):
    """Nombre contable del banco (columna 'Nombre' del Plan de Cuentas) cuyo
       CBU contiene el número de cuenta del extracto. None si ningún CBU
       cargado matchea, o si matchean varios -- nunca se adivina; ese caso
       queda para revisión humana (o para cargar el CBU que falta)."""
    nucleo = _nucleo_cuenta(cuenta_extracto)
    if len(nucleo) < 4:
        return None
    candidatos = []
    for fila in plan_cuentas.get('filas', []):
        cbu = _digitos(_valor(fila, 'CBU'))
        if cbu and nucleo in cbu:
            nombre = str(_valor(fila, 'Nombre') or '').strip()
            if nombre and nombre not in candidatos:
                candidatos.append(nombre)
    return candidatos[0] if len(candidatos) == 1 else None


def _nombre_banco(res, plan_cuentas: dict):
    return resolver_banco_por_cbu(getattr(res, 'cuenta', '') or '', plan_cuentas)


def _ultimo_dia_mes(anio: int, mes: int) -> datetime.date:
    return datetime.date(anio, mes, calendar.monthrange(anio, mes)[1])


def _r2(x) -> float:
    return round(float(x or 0.0), 2)


def construir_asientos(resultados, plan_cuentas: dict):
    """Devuelve una lista de asientos (uno por circuito y mes y cuenta bancaria).
    Cada asiento: dict(fecha, concepto, banco, circuito, lineas=[(cuenta, debe, haber)]).
    `plan_cuentas` = cargar_plan_cuentas(...) del cliente.
    """
    asientos = []
    revisar_global = []  # movimientos sin cuenta, sin circuito, o sin banco identificado
    internas_global = []  # transferencias entre cuentas propias (excluidas del importador)

    # Bancos que efectivamente aparecen en este lote (para poder asociar,
    # SOLO si son exactamente 2, la contrapartida de una transferencia
    # interna sin adivinar cuál es la otra cuenta).
    bancos_del_cliente = []
    for res in resultados:
        b = _nombre_banco(res, plan_cuentas)
        if b and b not in bancos_del_cliente:
            bancos_del_cliente.append(b)
    contra_interna = {}
    if len(bancos_del_cliente) == 2:
        contra_interna = {bancos_del_cliente[0]: bancos_del_cliente[1],
                           bancos_del_cliente[1]: bancos_del_cliente[0]}

    for res in resultados:
        banco = _nombre_banco(res, plan_cuentas)
        etiqueta_cta = str(getattr(res, 'cuenta', '') or '').strip()

        if banco is None:
            # No se pudo identificar el banco de esta cuenta (falta cargar
            # su CBU en el Plan de Cuentas, o hay varios que matchean).
            # No se arma ningún asiento para esta cuenta -- todo a revisar.
            for m in res.movimientos:
                revisar_global.append((res, m, f'banco no identificado para la cuenta '
                                                f'{etiqueta_cta}: falta cargar (o hay '
                                                f'ambigüedad en) el CBU en el Plan de Cuentas'))
            continue

        # agrupador[(anio, mes, circuito)][nombre_xubio] = [debe, haber]
        agrupador = defaultdict(lambda: defaultdict(lambda: [0.0, 0.0]))

        for m in res.movimientos:
            cuenta_int = getattr(m, 'cuenta_sugerida', '') or ''
            fecha = m.fecha
            if fecha is None:
                revisar_global.append((res, m, 'sin fecha'))
                continue
            clave_mes = (fecha.year, fecha.month)

            if not cuenta_int:
                revisar_global.append((res, m, 'sin cuenta'))
                continue
            circuito = _circuito_de(cuenta_int)
            if circuito is None:
                revisar_global.append((res, m, f'circuito no definido para la cuenta: {cuenta_int}'))
                continue
            # El nombre se usa TAL CUAL lo cargó el cliente en la planilla de
            # origen (Conceptos/Empleados/Proveedores/VEP) -- probado contra
            # datos reales: esa ortografía coincide con Xubio; la del Plan de
            # Cuentas a veces no (ej. 'CARGAS SOCIALES A PAGAR' vs 'Cargas
            # Sociales a pagar'). El Plan de Cuentas solo se usa para el
            # control de existencia (Fase 2) y para el CBU del banco.
            nombre_xubio = cuenta_int

            if circuito == INTERNO:
                # Transferencia entre cuentas propias:
                #  - RECIBIDA (crédito): se registra en ESTA cuenta como circuito TRANSFERENCIA
                #    ENTRE CUENTAS PROPIAS (banco de esta cuenta DEBE, la otra cuenta HABER).
                #  - ENVIADA (débito): se excluye (la registra la cuenta que recibe).
                if _r2(m.credito) > 0:
                    contra = contra_interna.get(banco)
                    if not contra:
                        revisar_global.append((res, m, 'transferencia entre cuentas propias '
                                                        'recibida, pero no se pudo identificar '
                                                        'sin ambigüedad la otra cuenta propia'))
                        continue
                    slot = agrupador[(clave_mes[0], clave_mes[1], TRP)][contra]
                    slot[1] = _r2(slot[1] + _r2(m.credito))   # HABER de la otra cuenta
                else:
                    internas_global.append((res, m))          # enviada: se excluye
                continue
            if circuito == REVISAR:
                revisar_global.append((res, m, f'circuito a revisar: {cuenta_int}'))
                continue

            slot = agrupador[(clave_mes[0], clave_mes[1], circuito)][nombre_xubio]
            slot[0] = _r2(slot[0] + _r2(m.debito))
            slot[1] = _r2(slot[1] + _r2(m.credito))

        # armar los asientos ordenados por mes y por ORDEN_CIRCUITOS
        claves = sorted(agrupador.keys(), key=lambda k: (k[0], k[1], ORDEN_CIRCUITOS.index(k[2]) if k[2] in ORDEN_CIRCUITOS else 99))
        for (anio, mes, circuito) in claves:
            cuentas = agrupador[(anio, mes, circuito)]
            fecha = _ultimo_dia_mes(anio, mes)
            suf = f"{mes:02d}.{anio}"
            concepto = f"{circuito} - BCO. {banco} {etiqueta_cta} del mes {suf}".replace('  ', ' ')

            # NETEADO: una sola línea por cuenta (debe - haber). Si el neto queda en 0,
            # la cuenta se cancela sola y no se lista.
            lineas = []
            neto_cuentas = 0.0   # suma de (debe - haber) de todas las cuentas del circuito
            for nombre_xubio, (debe, haber) in cuentas.items():
                neto = _r2(debe - haber)
                if neto > 0:
                    lineas.append((nombre_xubio, neto, 0.0))
                elif neto < 0:
                    lineas.append((nombre_xubio, 0.0, _r2(-neto)))
                neto_cuentas = _r2(neto_cuentas + neto)

            # Contrapartida banco (también neteada), para que el asiento cierre
            if neto_cuentas > 0:
                lineas.append((banco, 0.0, neto_cuentas))
            elif neto_cuentas < 0:
                lineas.append((banco, _r2(-neto_cuentas), 0.0))

            if lineas:
                asientos.append(dict(fecha=fecha, concepto=concepto, banco=banco,
                                     circuito=circuito, lineas=lineas))
    return asientos, revisar_global, internas_global


def _control_asiento(asiento):
    d = _r2(sum(l[1] for l in asiento['lineas']))
    h = _r2(sum(l[2] for l in asiento['lineas']))
    return d, h, _r2(d - h)


def control_cobertura(resultados, plan_cuentas: dict):
    """CONTROL GLOBAL de la Capa 3.
    Compara, por cada cuenta bancaria, cuánto movió el banco en el EXTRACTO
    (créditos - débitos de TODOS los movimientos) contra cuánto quedó reflejado
    en el IMPORTADOR (solo los movimientos que SÍ se pudieron clasificar).
    Si la diferencia no es 0, el importador está INCOMPLETO por ese monto:
    hay movimientos sin cuenta que faltan imputar (ver hoja A REVISAR).
    Devuelve una lista de dicts, uno por cuenta.
    """
    reporte = []
    for res in resultados:
        total_extracto = 0.0      # movimiento del banco EXCLUYENDO transferencias internas
        total_importado = 0.0     # lo que entró al importador (clasificado y con circuito)
        falta_monto = 0.0
        falta_cant = 0
        internas_cant = 0
        internas_monto = 0.0
        banco_id = _nombre_banco(res, plan_cuentas)
        for m in res.movimientos:
            neto = _r2(_r2(m.credito) - _r2(m.debito))
            cuenta_int = getattr(m, 'cuenta_sugerida', '') or ''
            circuito = _circuito_de(cuenta_int) if cuenta_int else None

            if circuito == INTERNO:
                # transferencia entre cuentas propias: se excluye a propósito (no es faltante)
                internas_cant += 1
                internas_monto = _r2(internas_monto + abs(neto))
                continue

            total_extracto = _r2(total_extracto + neto)
            clasificado_ok = (bool(cuenta_int) and circuito is not None and circuito != REVISAR
                               and m.fecha is not None and banco_id is not None)
            if clasificado_ok:
                total_importado = _r2(total_importado + neto)
            else:
                falta_monto = _r2(falta_monto + abs(neto))
                falta_cant += 1
        reporte.append(dict(
            cuenta=str(getattr(res, 'cuenta', '') or ''),
            banco=banco_id or '(no identificado — falta el CBU en el Plan de Cuentas)',
            total_extracto=total_extracto,
            total_importado=total_importado,
            diferencia=_r2(total_extracto - total_importado),
            faltan_movimientos=falta_cant,
            faltan_monto=falta_monto,
            internas_movimientos=internas_cant,
            internas_monto=internas_monto,
            completo=(falta_cant == 0 and banco_id is not None),
        ))
    return reporte


def generar_excel_importador(asientos, path):
    """El archivo que se SUBE a Xubio: UNA sola hoja, sin nada más (Xubio no
       acepta un .xlsx con varias solapas)."""
    import openpyxl
    from openpyxl.styles import Font

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'IMPORTADOR ASIENTOS'
    cols = ['FECHA', 'CONCEPTO', 'CIRCUITOCONTABLE', 'CUENTA', 'DEBE', 'HABER',
            'ORGANIZACION', 'CENTRODECOSTO', 'DESCRIPCION']
    ws.append(cols)
    for c in range(1, len(cols) + 1):
        ws.cell(1, c).font = Font(bold=True)
    for a in asientos:
        ws.append([a['fecha'].strftime('%d/%m/%Y'), a['concepto'], 'default', '', '', '', '', '', ''])
        for (cuenta, debe, haber) in a['lineas']:
            ws.append(['', '', '', cuenta,
                       debe if debe > 0 else '', haber if haber > 0 else '', '', '', ''])
    for col, w in zip('ABCDEFGHI', [12, 55, 16, 40, 15, 15, 14, 14, 30]):
        ws.column_dimensions[col].width = w
    wb.save(path)


def construir_wb_papel_trabajo(resultados, plan_cuentas, asientos, revisar, internas):
    """Arma (sin guardar) el Workbook del papel de trabajo -- para revisión
       humana, NO se sube a Xubio: RESUMEN ASIENTOS (con control
       debe=haber), A REVISAR, TRANSF. INTERNAS y CONTROL (cobertura
       global). Devuelve el openpyxl.Workbook para que el llamador pueda
       agregarle más hojas (ej. orquestador.py le agrega el detalle
       movimiento por movimiento de cada cuenta) antes de guardarlo."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill

    wb = openpyxl.Workbook()

    # ---------- Hoja RESUMEN ASIENTOS ----------
    wr = wb.active
    wr.title = 'RESUMEN ASIENTOS'
    wr.append(['FECHA', 'CONCEPTO', 'NOMBRE CUENTA', 'DEBE', 'HABER', 'CONTROL'])
    for c in range(1, 7):
        wr.cell(1, c).font = Font(bold=True)
    amarillo = PatternFill('solid', fgColor='FFF2CC')
    for a in asientos:
        d, h, ctrl = _control_asiento(a)
        wr.append([a['fecha'].strftime('%d/%m/%Y'), a['concepto'], '', '', '', ''])
        for (cuenta, debe, haber) in a['lineas']:
            wr.append(['', '', cuenta, debe if debe > 0 else '', haber if haber > 0 else '', ''])
        fila_total = ['', '', 'TOTAL', d, h, ctrl]
        wr.append(fila_total)
        r = wr.max_row
        for c in range(3, 7):
            wr.cell(r, c).font = Font(bold=True)
            wr.cell(r, c).fill = amarillo
    for col, w in zip('ABCDEF', [12, 55, 40, 15, 15, 12]):
        wr.column_dimensions[col].width = w

    # ---------- Hoja A REVISAR (movimientos sin circuito/cuenta/banco) ----------
    if revisar:
        wv = wb.create_sheet('A REVISAR')
        wv.append(['CUENTA (extracto)', 'FECHA', 'CONCEPTO', 'NOMBRE', 'DÉBITO', 'CRÉDITO', 'MOTIVO'])
        for c in range(1, 8):
            wv.cell(1, c).font = Font(bold=True)
        for (res, m, motivo) in revisar:
            wv.append([str(getattr(res, 'cuenta', '')),
                       m.fecha.strftime('%d/%m/%Y') if m.fecha else '',
                       m.concepto, getattr(m, 'nombre', ''),
                       _r2(m.debito) or '', _r2(m.credito) or '', motivo])
        for col, w in zip('ABCDEFG', [18, 12, 40, 28, 14, 14, 30]):
            wv.column_dimensions[col].width = w

    # ---------- Hoja TRANSFERENCIAS INTERNAS (excluidas a propósito) ----------
    if internas:
        wt = wb.create_sheet('TRANSF. INTERNAS')
        wt.append(['CUENTA (extracto)', 'FECHA', 'CONCEPTO', 'NOMBRE', 'DÉBITO', 'CRÉDITO'])
        for c in range(1, 7):
            wt.cell(1, c).font = Font(bold=True)
        for (res, m) in internas:
            wt.append([str(getattr(res, 'cuenta', '')),
                       m.fecha.strftime('%d/%m/%Y') if m.fecha else '',
                       m.concepto, getattr(m, 'nombre', ''),
                       _r2(m.debito) or '', _r2(m.credito) or ''])
        wt.append([])
        wt.append(['', '', 'Estas transferencias entre cuentas propias NO van al importador '
                   '(se registran una sola vez / aparte, para no duplicar plata interna).'])
        for col, w in zip('ABCDEF', [18, 12, 40, 28, 14, 14]):
            wt.column_dimensions[col].width = w

    # ---------- Hoja CONTROL (cobertura global) ----------
    cob = control_cobertura(resultados, plan_cuentas)
    wc = wb.create_sheet('CONTROL')
    wc.append(['CUENTA', 'BANCO IDENTIFICADO', 'MOVIÓ EL BANCO (sin internas)',
               'REFLEJADO EN IMPORTADOR', 'DIFERENCIA (falta imputar)',
               'MOV. SIN CLASIFICAR', 'TRANSF. INTERNAS', 'ESTADO'])
    for c in range(1, 9):
        wc.cell(1, c).font = Font(bold=True)
    verde = PatternFill('solid', fgColor='C6EFCE')
    rojo = PatternFill('solid', fgColor='FFC7CE')
    for r in cob:
        estado = 'COMPLETO ✓' if r['completo'] else 'INCOMPLETO — faltan cuentas'
        internas_txt = f"{r.get('internas_movimientos', 0)} (${r.get('internas_monto', 0):,.2f})"
        wc.append([r['cuenta'], r['banco'], r['total_extracto'], r['total_importado'],
                   r['diferencia'], r['faltan_movimientos'], internas_txt, estado])
        fila = wc.max_row
        wc.cell(fila, 8).fill = verde if r['completo'] else rojo
        wc.cell(fila, 8).font = Font(bold=True)
    for col, w in zip('ABCDEFGH', [22, 26, 26, 26, 26, 20, 22, 30]):
        wc.column_dimensions[col].width = w

    return wb


def generar_excel_papel_trabajo(resultados, plan_cuentas, asientos, revisar, internas, path):
    """Como construir_wb_papel_trabajo(), pero ya guardado en `path`."""
    wb = construir_wb_papel_trabajo(resultados, plan_cuentas, asientos, revisar, internas)
    wb.save(path)


def generar_excel(resultados, plan_cuentas, path_importador, path_papel_trabajo):
    """Genera los DOS archivos de la Capa 3:
       - path_importador     : UNA sola hoja, lista para subir a Xubio.
       - path_papel_trabajo  : el detalle para revisión humana (control,
                                pendientes, transferencias internas).
       `plan_cuentas` = cargar_plan_cuentas(...) del cliente."""
    asientos, revisar, internas = construir_asientos(resultados, plan_cuentas)
    generar_excel_importador(asientos, path_importador)
    generar_excel_papel_trabajo(resultados, plan_cuentas, asientos, revisar, internas, path_papel_trabajo)
    return asientos, revisar, internas
