# -*- coding: utf-8 -*-
"""
============================================================================
 CONVERSOR DE EXTRACTOS BANCARIOS  ->  TABLA ESTÁNDAR PARA CONCILIACIÓN
============================================================================

Qué hace
--------
Toma el PDF de un extracto bancario (descargado del homebanking, con texto
real adentro) y lo convierte en una tabla estándar de movimientos, una fila
por movimiento, con estas columnas:

    fecha | concepto | referencia | cuit | nombre | debito | credito | saldo

Esa tabla es la que después llena la hoja "Carga de extracto" del archivo de
conciliación. Cada banco arma su PDF distinto, así que hay UN LECTOR POR BANCO,
pero todos devuelven el mismo objeto Resultado con la misma tabla estándar.

Diseño (para que sea mantenible y sobreviva a cambios de formato)
-----------------------------------------------------------------
  * Cada lector es una función leer_<banco>(pdf) -> Resultado.
  * detectar_banco() mira el texto del PDF y elige el lector.
  * Toda la lógica frágil (posiciones X de columnas, rarezas de cada banco)
    está aislada y COMENTADA dentro de su lector. Si un banco cambia el PDF,
    se toca solo ese lector.
  * RED DE SEGURIDAD: verificar_control() reconstruye el saldo por acumulación
    (saldo_ini + créditos - débitos) y lo compara contra el saldo final que
    informa el banco. Si da 0, no se perdió ni se duplicó ningún movimiento.

Uso
---
    python conversor_extractos.py  archivo.pdf  [salida.xlsx]

    # o desde código:
    from conversor_extractos import leer_extracto, exportar_excel
    res = leer_extracto("extracto.pdf")
    exportar_excel(res, "convertido.xlsx")

Bancos soportados en este módulo: Banco Ciudad, Banco Comafi, Supervielle.
Pendientes de sumar (dejar su función leer_<banco> siguiendo el mismo patrón):
Galicia, BBVA, Macro, ICBC, Credicoop.
============================================================================
"""
from dataclasses import dataclass, field
from typing import Optional, List
import re, datetime, sys
import pdfplumber

# ---------------------------------------------------------------------------
# Modelo de datos estándar (la "interfaz" común que consume la conciliación)
# ---------------------------------------------------------------------------
@dataclass
class Movimiento:
    fecha: datetime.date
    concepto: str
    referencia: str = ''
    cuit: str = ''
    nombre: str = ''
    debito: float = 0.0
    credito: float = 0.0
    saldo: Optional[float] = None          # None si el banco no lo imprime en esa fila
    categoria: str = ''                    # p.ej. 'SUELDOS' (clasificación); '' = sin clasificar
    cuenta_sugerida: str = ''              # cuenta contable sugerida para el importador Xubio

@dataclass
class Resultado:
    banco: str
    titular: str
    cuit_titular: str
    cuenta: str
    periodo: str
    saldo_ini: float
    saldo_fin: float
    movimientos: List[Movimiento] = field(default_factory=list)

# ---------------------------------------------------------------------------
# Helpers compartidos
# ---------------------------------------------------------------------------
def num_ar(txt: str) -> float:
    """Convierte número en formato argentino a float.
       '1.234.567,89' -> 1234567.89 ; signo negativo puede ir al final ('...,72-').
       Tolera un '$' adelante y espacios de más (formato de Google Sheets,
       ej. '$ 1.234,56')."""
    txt = str(txt).strip().replace('$', '').strip()
    neg = txt.endswith('-')
    txt = txt.rstrip('-').strip()
    return (-1 if neg else 1) * float(txt.replace('.', '').replace(',', '.'))

RE_NUM = re.compile(r'^-?[\d.]+,\d{2}-?$')

def fecha_dmy2(txt: str):
    """dd/mm/yy -> date."""
    m = re.match(r'(\d{2})/(\d{2})/(\d{2})$', txt)
    return datetime.date(2000+int(m[3]), int(m[2]), int(m[1])) if m else None

def fecha_dmy4(txt: str):
    """dd/mm/yyyy -> date (año completo; lo usa el extracto Office Banking de Galicia)."""
    m = re.match(r'(\d{2})/(\d{2})/(\d{4})$', txt)
    return datetime.date(int(m[3]), int(m[2]), int(m[1])) if m else None

_MESES = {'ENE':1,'FEB':2,'MAR':3,'ABR':4,'MAY':5,'JUN':6,
          'JUL':7,'AGO':8,'SEP':9,'OCT':10,'NOV':11,'DIC':12}
def fecha_dMMMy(txt: str):
    """dd-MMM-yyyy (mes en letras, español) -> date."""
    m = re.match(r'(\d{2})-([A-Z]{3})-(\d{4})', txt.upper())
    return datetime.date(int(m[3]), _MESES[m[2]], int(m[1])) if m and m[2] in _MESES else None

def _filas_por_top(words, y_min, y_max, tol=2):
    """Agrupa palabras en filas por su coordenada 'top', con tolerancia vertical."""
    filas = {}
    for w in words:
        if not (y_min <= w['top'] <= y_max):
            continue
        key = round(w['top'])
        for k in list(filas):
            if abs(k-key) <= tol:
                key = k; break
        filas.setdefault(key, []).append(w)
    return [sorted(filas[k], key=lambda z: z['x0']) for k in sorted(filas)]

# ===========================================================================
# LECTOR: BANCO CIUDAD
# ---------------------------------------------------------------------------
# Formato prolijo, una línea por movimiento, saldo acumulado en CADA fila.
# CUIT/nombre vienen EN la misma línea (columna "Descripción de movimiento").
# Trampa propia: marca de agua vertical en el margen izquierdo (x0<28) que se
# cuela en la columna de fecha -> se descarta por posición.
# Columnas (x): fecha<82 | concepto<170 | desc>=370 | números por borde der.:
#               débito x1<236 | crédito x1<300 | saldo resto.
# ===========================================================================
def leer_ciudad(pdf) -> Resultado:
    p = pdf.pages[0]
    W = p.extract_words()
    def num_box(y0, y1, x0min):
        for w in W:
            if y0 < w['top'] < y1 and w['x0'] > x0min and RE_NUM.match(w['text']):
                return num_ar(w['text'])
        return None
    saldo_ini = num_box(195, 202, 460)
    saldo_fin = num_box(780, 786, 460)

    def col(x0, x1):
        if x0 < 82:  return 'fecha'
        if x0 < 170: return 'concepto'
        if x0 >= 370:return 'desc'
        if x1 < 236: return 'debito'
        if x1 < 300: return 'credito'
        return 'saldo'

    movs = []
    for fila in _filas_por_top([w for w in W if w['x0'] >= 28], 225, 400):
        celdas = {k: [] for k in ('fecha','concepto','debito','credito','saldo','desc')}
        for w in fila:
            celdas[col(w['x0'], w['x1'])].append(w['text'])
        fecha = fecha_dMMMy(' '.join(celdas['fecha']))
        if not fecha:
            continue
        desc = ' '.join(celdas['desc']).strip()
        mc = re.match(r'(\d{11})[-\s]*(.*)', desc)
        cuit, nombre = (mc[1], mc[2].strip()) if mc else ('', desc)
        movs.append(Movimiento(
            fecha=fecha, concepto=' '.join(celdas['concepto']).strip(),
            cuit=cuit, nombre=nombre,
            debito=num_ar(celdas['debito'][0]) if celdas['debito'] else 0.0,
            credito=num_ar(celdas['credito'][0]) if celdas['credito'] else 0.0,
            saldo=num_ar(celdas['saldo'][0]) if celdas['saldo'] else None))

    txt = p.extract_text()
    tit = re.search(r'\n(ASC .+?)\n', txt)
    cuit = re.search(r'(\d{2}-\d{8}-\d)', txt)
    cta  = re.search(r'(\d{7}/\d)', txt)
    return Resultado('Banco Ciudad', tit[1].strip() if tit else '',
                     cuit[1] if cuit else '', cta[1] if cta else '',
                     '', saldo_ini, saldo_fin, movs)

# ===========================================================================
# LECTOR: BANCO COMAFI
# ---------------------------------------------------------------------------
# Multi-página con líneas "Transporte" que arrastran el saldo entre hojas
# (se usan como límite de banda, NO son movimientos). El saldo NO aparece en
# cada fila. Movimientos que se parten en 2 líneas (ej "Pago de cheque": ref y
# nombre en un renglón, importe en el siguiente). Saldo negativo con '-' final.
# CUIT/nombre en sección aparte "Transferencias Electrónicas" -> cruce por
# fecha+importe (estilo BBVA), excluyendo los internos del banco (30604731018).
# Trampa: la palabra "Fecha" reaparece abajo (sección CHEQUES) -> el encabezado
# solo se busca en la zona de la tabla (top<270).
# Columnas (x): fecha<64 | concepto<196 | ref<290 | números por borde der.:
#               débito x1<430 | crédito x1<510 | saldo resto.
# ===========================================================================
def _comafi_transferencias(pdf):
    txt = pdf.pages[5].extract_text().split('\n')
    envs, modo = [], None
    for ln in txt:
        if 'ENVIADAS' in ln:  modo = 'env'; continue
        if 'RECIBIDAS' in ln: modo = 'rec'; continue
        if 'FONDOS COMUNES' in ln: break
        m = re.match(r'(\d{2}/\d{2}/\d{4})\s+(.*)', ln.strip())
        if not m or modo is None: continue
        imp  = re.search(r'\$\s*([\d.]+,\d{2})', m[2])
        cuit = re.search(r'\b(\d{11})\b', m[2])
        if not imp: continue
        cu = cuit[1] if cuit else ''
        if cu == '30604731018':       # movimientos internos (comisiones del banco)
            continue
        nombre = m[2][cuit.end():imp.start()].strip() if cuit else ''
        envs.append((datetime.datetime.strptime(m[1], '%d/%m/%Y').date(),
                     num_ar(imp[1]), cu, nombre))
    return envs

def leer_comafi(pdf) -> Resultado:
    envs = _comafi_transferencias(pdf)
    def col_num(x1): return 'debito' if x1 < 430 else 'credito' if x1 < 510 else 'saldo'
    saldo_ini = saldo_fin = None
    movs: List[Movimiento] = []

    for pg in range(1, 5):
        p = pdf.pages[pg]
        W = [w for w in p.extract_words() if not set(w['text']) <= set('-')]
        tops_hdr = [w['top'] for w in W if w['text'] == 'Fecha' and w['top'] < 270]
        tops_tr  = [w['top'] for w in W if w['text'] == 'Transporte']
        tops_al  = [w['top'] for w in W if w['text'] == 'al:']
        start = max(tops_hdr + [t for t in tops_tr if t < 270] or [150])
        fin_c = [t for t in tops_tr if t > 300] + tops_al
        end   = min(fin_c) if fin_c else 720

        for fila in _filas_por_top(W, start+2, end-1):
            textos = [w['text'] for w in fila]
            if 'Conceptos' in textos or 'Referencias' in textos:
                continue
            if any(w['text'] == 'Anterior' for w in fila):        # Saldo Anterior
                nums = [w for w in fila if RE_NUM.match(w['text'])]
                if nums: saldo_ini = num_ar(nums[-1]['text'])
                continue
            fecha_tok = [w for w in fila if w['x0'] < 64 and fecha_dmy2(w['text'])]
            concepto = ' '.join(w['text'] for w in fila if 64 <= w['x0'] < 196)
            ref = ' '.join(w['text'] for w in fila if 196 <= w['x0'] < 290 and w['x1'] < 290)
            nums = {'debito': None, 'credito': None, 'saldo': None}
            for w in fila:
                if RE_NUM.match(w['text']) and w['x0'] > 355:
                    nums[col_num(w['x1'])] = num_ar(w['text'])
            if fecha_tok:
                movs.append(Movimiento(
                    fecha=fecha_dmy2(fecha_tok[0]['text']), concepto=concepto.strip(),
                    referencia=ref.strip(), debito=nums['debito'] or 0.0,
                    credito=nums['credito'] or 0.0, saldo=nums['saldo']))
            elif movs:                                            # continuación
                extra = ' '.join(w['text'] for w in fila if w['x0'] < 290)
                if extra: movs[-1].concepto = (movs[-1].concepto + ' ' + extra).strip()
                if nums['debito']:  movs[-1].debito  += nums['debito']
                if nums['credito']: movs[-1].credito += nums['credito']
                if nums['saldo'] is not None: movs[-1].saldo = nums['saldo']

    for w in pdf.pages[4].extract_words():
        if w['text'] == 'al:':
            fila = [x for x in pdf.pages[4].extract_words() if abs(x['top']-w['top']) < 3]
            nums = [x for x in fila if RE_NUM.match(x['text'])]
            if nums: saldo_fin = num_ar(nums[-1]['text'])

    for m in movs:                                               # enriquecer CUIT
        if 'Transferencia Terceros' in m.concepto or ('Transf' in m.concepto and 'sueldos' in m.concepto):
            cand = [e for e in envs if e[0] == m.fecha and abs(e[1]-m.debito) < 0.01]
            if len(cand) == 1:
                m.cuit, m.nombre = cand[0][2], cand[0][3]

    t = pdf.pages[1].extract_text()
    tit  = re.search(r'TITULAR.*?\n.*?\n\s*(.+?)\s+\d{2}-\d{8}-\d', t, re.S)
    cuit = re.search(r'(\d{2}-\d{8}-\d)\s+Responsable', t)
    cta  = re.search(r'(\d{4}-\d{5}-\d)', t)
    return Resultado('Banco Comafi', (tit[1].strip() if tit else 'NEXO IT SRL'),
                     cuit[1] if cuit else '', cta[1] if cta else '',
                     '', saldo_ini, saldo_fin, movs)

# ===========================================================================
# LECTOR: SUPERVIELLE
# ---------------------------------------------------------------------------
# Una página de movimientos, texto limpio. Movimientos con líneas de
# continuación (Operación.../Cuentas Propias/CUIT-nombre) que NO traen importe.
# Trampa: la línea "Imp Ley 25413 s/Debitos" trae un importe en débito pero
# NO tiene saldo y NO mueve el balance (es informativa) -> se descarta.
# Regla: fila sin fecha + con importe pero sin saldo = informativa (descartar);
#        fila sin fecha + sin importe = continuación de texto del mov. previo.
# Columnas (x): fecha<100 | concepto<250 | ref<306 | números (x0>305) por borde
#               der.: débito x1<420 | crédito x1<470 | saldo resto.
# ===========================================================================
def leer_supervielle(pdf) -> Resultado:
    p = pdf.pages[0]
    W = [w for w in p.extract_words() if not set(w['text']) <= set('*')]
    def col_num(x1): return 'debito' if x1 < 420 else 'credito' if x1 < 470 else 'saldo'
    top_ini = min([w['top'] for w in W if w['text'] == 'anterior'] or [270])
    top_fin = max([w['top'] for w in W if w['text'] == 'ACTUAL'] or [620])

    saldo_ini = saldo_fin = None
    movs: List[Movimiento] = []
    for fila in _filas_por_top(W, top_ini-2, top_fin+2):
        txt = ' '.join(w['text'] for w in fila)
        nums = {'debito': None, 'credito': None, 'saldo': None}
        for w in fila:
            if RE_NUM.match(w['text']) and w['x0'] > 305:
                nums[col_num(w['x1'])] = num_ar(w['text'])
        if 'período anterior' in txt: saldo_ini = nums['saldo']; continue
        if 'PERIODO ACTUAL' in txt:   saldo_fin = nums['saldo']; continue
        fecha_tok = [w for w in fila if w['x0'] < 100 and fecha_dmy2(w['text'])]
        concepto = ' '.join(w['text'] for w in fila if 100 <= w['x0'] < 250)
        ref = ' '.join(w['text'] for w in fila if 250 <= w['x0'] < 306 and w['x1'] < 306)
        if fecha_tok:
            movs.append(Movimiento(
                fecha=fecha_dmy2(fecha_tok[0]['text']), concepto=concepto.strip(),
                referencia=ref.strip(), debito=nums['debito'] or 0.0,
                credito=nums['credito'] or 0.0, saldo=nums['saldo'], nombre=''))
            movs[-1]._cont = ''
        else:
            if nums['debito'] or nums['credito']:       # informativa -> descartar
                continue
            if movs:                                    # continuación de texto
                extra = ' '.join(w['text'] for w in fila if w['x0'] >= 100)
                movs[-1]._cont = (getattr(movs[-1], '_cont', '') + ' ' + extra).strip()

    for m in movs:
        mc = re.search(r'\b(\d{11})\b\s*(.*)', getattr(m, '_cont', ''))
        if mc: m.cuit, m.nombre = mc[1], mc[2].strip()

    t = p.extract_text()
    tit  = re.search(r'([A-ZÑ ]*OXXON[A-ZÑ ]*)', t)
    cuit = re.search(r'C\.U\.I\.T\.\s*0?(\d{2}-\d{8}-\d)', t)
    cta  = re.search(r'Nro\.:\s*([\d-]+)', t)
    return Resultado('Supervielle', (tit[1].strip() if tit else ''),
                     cuit[1] if cuit else '', cta[1] if cta else '',
                     '', saldo_ini, saldo_fin, movs)

# ===========================================================================
# LECTORES REINTEGRADOS: GALICIA, BBVA, MACRO
# ---------------------------------------------------------------------------
# Portados desde el desarrollo original. Galicia = 1 cuenta por PDF.
# BBVA y Macro = VARIAS cuentas por PDF -> devuelven una LISTA de Resultado.
# Agrupación de líneas por cercanía vertical (tol=4) para no perder importes
# por micro-desalineación (fue el bug que en BBVA hacía perder movimientos).
# ===========================================================================
_CUIT11 = re.compile(r'^\d{11}$')

def _cluster(words, tol=4):
    ws = sorted(words, key=lambda w: w['top']); out=[]; cur=[]; ref=None
    for w in ws:
        if ref is None or abs(w['top']-ref) <= tol:
            cur.append(w)
            if ref is None: ref=w['top']
        else:
            out.append(cur); cur=[w]; ref=w['top']
    if cur: out.append(cur)
    return [sorted(l, key=lambda z: z['x0']) for l in out]

def _fecha_ddmm(txt, anio):
    m = re.match(r'(\d{2})/(\d{2})$', txt)
    return datetime.date(int(anio), int(m[2]), int(m[1])) if m else None

# ---------------------------------------------------------------------------
# GALICIA — bloques multilínea (nombre / CUIT / CBU apilados bajo el movimiento)
# Columnas por borde derecho: crédito x1<400 | débito x1<500 | saldo resto.
# Totales y saldo final en la fila "Total". Saldo negativo con "-" al final.
# ---------------------------------------------------------------------------
def leer_galicia(pdf) -> Resultado:
    DATE = re.compile(r'^\d{2}/\d{2}/\d{2}$')
    def zona(x1): return 'credito' if x1 < 400 else ('debito' if x1 < 500 else 'saldo')
    t0 = pdf.pages[0].extract_text() or ''
    mcu = re.search(r'Responsable Impositivo\s*:\s*([\d\-]+)', t0)
    cuit_tit = mcu.group(1) if mcu else ''
    mct = re.search(r'N[°º]\s*([\d\-]+\s*[\d\-]+)', t0)
    cuenta = mct.group(1).strip() if mct else ''
    fechas = re.findall(r'\d{2}/\d{2}/\d{4}', t0)
    periodo = f"{fechas[0]} a {fechas[1]}" if len(fechas) >= 2 else ''
    mtt = re.search(r'\n([A-ZÑ0-9&\.\- ]{3,40})\s*\nResumen de Cuenta', t0)
    titular = mtt.group(1).strip() if mtt else ''

    movs=[]; cur=None; ended=False; tot=None
    for page in pdf.pages:
        for toks in _cluster(page.extract_words()):
            txt = ' '.join(w['text'] for w in toks).strip()
            if txt.startswith('Total'):
                amts=[num_ar(w['text']) for w in toks if RE_NUM.match(w['text'])]
                if len(amts) >= 3: tot=(abs(amts[0]), abs(amts[1]), amts[2])
                if cur: movs.append(cur); cur=None
                ended=True; continue
            if ended: continue
            first=toks[0]
            if DATE.match(first['text']) and first['x0'] < 80:
                if cur: movs.append(cur)
                cur=dict(fecha=first['text'], desc=[], credito=0.0, debito=0.0, saldo=None, det=[])
                for w in toks[1:]:
                    x0,x1,tx = w['x0'], w['x1'], w['text']
                    if RE_NUM.match(tx):
                        v=num_ar(tx); z=zona(x1)
                        if z=='credito': cur['credito']=abs(v)
                        elif z=='debito': cur['debito']=abs(v)
                        else: cur['saldo']=v
                    elif 80 <= x0 < 220:
                        cur['desc'].append(tx)
            else:
                if cur is None: continue
                det=[w['text'] for w in toks if 80 <= w['x0'] < 300]
                if det: cur['det'].append(' '.join(det))
    if cur and not ended: movs.append(cur)

    out=[]
    for m in movs:
        det=m['det']; nombre=det[0] if det else ''
        cu=next((d.strip() for d in det if _CUIT11.match(d.strip())), '')
        out.append(Movimiento(fecha_dmy2(m['fecha']), ' '.join(m['desc']).strip(),
                              '', cu, nombre, m['debito'], m['credito'], m['saldo']))
    saldo_fin = tot[2] if tot else (out[-1].saldo if out else 0.0)
    saldo_ini = round(out[0].saldo - (out[0].credito - out[0].debito), 2) if out and out[0].saldo is not None else 0.0
    return Resultado('Banco Galicia', titular, cuit_tit, cuenta, periodo, saldo_ini, saldo_fin, out)

# ---------------------------------------------------------------------------
# GALICIA "OFFICE BANKING" (export "Movimientos de CC") — FORMATO DISTINTO al
# "Resumen de Cuenta Corriente" clásico. Diferencias que obligan a un lector propio:
#   * Columnas en orden Fecha | Descripción | DÉBITOS | CRÉDITOS | Saldos
#     (débito ANTES que crédito; en el clásico es al revés).
#   * Fecha con AÑO COMPLETO (dd/mm/aaaa), no dd/mm/aa.
#   * Importe con SIGNO: "+ $ ..." = crédito ; "- $ ..." = débito. El saldo va sin signo.
#   * NO hay fila "Total" final. El saldo aparece en CADA fila -> el control de saldo
#     fila por fila es la garantía de integridad.
#   * No trae nombre del titular ni su CUIT en la cabecera (quedan vacíos).
# Posiciones reales medidas sobre el PDF (borde derecho x1 de cada número):
#   débito  ~329 | crédito ~428 | saldo ~534 . Se clasifica por el SIGNO (robusto),
#   con respaldo por posición si faltara el signo.
# ---------------------------------------------------------------------------
def leer_galicia_ob(pdf) -> Resultado:
    from collections import defaultdict
    HDR = {'Fecha', 'Descripción', 'Débitos', 'Créditos', 'Saldos'}
    DATE4 = re.compile(r'^\d{2}/\d{2}/\d{4}$')

    # Cabecera de la cuenta (solo en página 1)
    t0 = pdf.pages[0].extract_text() or ''
    mcta = re.search(r'Movimientos de CC.*?\$?\s*([\d\-]+\s+[\d\-]+)', t0)
    cuenta = mcta.group(1).strip() if mcta else ''

    movs = []; cur = None
    for page in pdf.pages:
        lineas = defaultdict(list)
        for w in page.extract_words():
            lineas[round(w['top'])].append(w)
        for top in sorted(lineas):
            ws = sorted(lineas[top], key=lambda w: w['x0'])
            textos = [w['text'] for w in ws]
            # Saltar encabezado de columnas y el pie de página que se repiten
            if HDR & set(textos):
                continue
            if any('descarga' in t.lower() or 'Banking' in t for t in textos):
                continue
            first = ws[0]
            if DATE4.match(first['text']) and first['x0'] < 90:
                # Nueva fila de movimiento
                if cur: movs.append(cur)
                desc = ' '.join(w['text'] for w in ws
                                if 100 <= w['x0'] < 300 and not RE_NUM.match(w['text'])
                                and w['text'] not in ('+', '-', '$')).rstrip(' -+').strip()
                # Débito/crédito/saldo se determinan por la POSICIÓN de la columna (borde
                # derecho x1), NO por el signo "+/-". Esto es robusto ante conceptos que
                # tienen un guion en el nombre (ej. "NAVE - VENTA CON TARJETA"), que antes
                # hacían leer un crédito como débito.
                #   débito  x1~329 | crédito x1~428 | saldo x1~534
                nums = [(num_ar(w['text']), w['x1']) for w in ws if RE_NUM.match(w['text'])]
                saldo = None; deb = cre = 0.0
                for val, x1 in nums:
                    if x1 >= 481:        # columna Saldos (la más a la derecha)
                        saldo = abs(val)
                    elif x1 >= 378:      # columna Créditos
                        cre = abs(val)
                    else:                # columna Débitos
                        deb = abs(val)
                cur = dict(fecha=first['text'], desc=desc, deb=deb, cre=cre, saldo=saldo, det=[])
            else:
                # Línea de detalle (nombre, CUIT, CBU, VARIOS, etc.) del movimiento en curso
                if cur is None: continue
                det = [w['text'] for w in ws if 100 <= w['x0'] < 300]
                if det: cur['det'].append(' '.join(det))
    if cur: movs.append(cur)

    out = []
    for m in movs:
        det = m['det']; nombre = det[0] if det else ''
        cu = next((d.strip() for d in det if _CUIT11.match(d.strip())), '')
        # El detalle (líneas debajo del movimiento: nombre, CUIT, "HONORARIOS", "ALQUILERES",
        # CBU, etc.) se guarda en 'referencia' para que las reglas de clasificación puedan verlo.
        referencia = ' '.join(det[1:]) if len(det) > 1 else ''
        out.append(Movimiento(fecha_dmy4(m['fecha']), m['desc'], referencia, cu, nombre,
                              m['deb'], m['cre'], m['saldo']))
    periodo = ''
    if out:
        fs = [m.fecha for m in out if m.fecha]
        if fs: periodo = f"{min(fs).strftime('%d/%m/%Y')} a {max(fs).strftime('%d/%m/%Y')}"
    saldo_fin = out[-1].saldo if out and out[-1].saldo is not None else 0.0
    saldo_ini = round(out[0].saldo - (out[0].credito - out[0].debito), 2) if out and out[0].saldo is not None else 0.0
    return Resultado('Banco Galicia', '', '', cuenta, periodo, saldo_ini, saldo_fin, out)

# ---------------------------------------------------------------------------
# BBVA — 1 línea por movimiento, VARIAS CUENTAS por PDF. Débitos con signo.
# Columnas por borde derecho: débito x1<460 | crédito x1<540 | saldo resto.
# Cada cuenta: SALDO ANTERIOR ... movimientos ... SALDO AL / TOTAL MOVIMIENTOS.
# Se descartan los encabezados-resumen (no tienen SALDO ANTERIOR).
# La fecha viene dd/mm (sin año): se completa con el año del período.
# ---------------------------------------------------------------------------
def leer_bbva(pdf) -> List[Resultado]:
    DATE = re.compile(r'^\d{2}/\d{2}$')
    def zona(x1): return 'debito' if x1 < 460 else ('credito' if x1 < 540 else 'saldo')
    txt_all = '\n'.join((p.extract_text() or '') for p in pdf.pages)
    t0 = pdf.pages[0].extract_text() or ''
    mt = re.search(r'([A-ZÑ][A-ZÑ0-9&\.\- ]+?)\s*\((\d{2}-\d{8}-\d)\)', t0)
    titular = mt.group(1).strip() if mt else ''
    cuit_tit = mt.group(2) if mt else ''
    ma = re.search(r'(20\d\d)', t0); anio = ma.group(1) if ma else '2026'
    mp = re.search(r'MES:\s*([A-ZÑ]+ 20\d\d)', txt_all.upper()); periodo = mp.group(1).title() if mp else ''

    accts=[]; cur=None; inside=False
    for page in pdf.pages:
        for toks in _cluster(page.extract_words()):
            txt=' '.join(w['text'] for w in toks)
            m=re.search(r'(CC (?:\$|U\$S) [\d\-/]+)\s*\(Cta\.Cte', txt)
            if m:
                cur=dict(cuenta=m.group(1), saldo_ini=None, saldo_fin=None, movs=[])
                accts.append(cur); inside=False; continue
            if cur is None: continue
            up=txt.upper()
            if up.startswith('SALDO ANTERIOR'):
                a=[num_ar(w['text']) for w in toks if RE_NUM.match(w['text'])]
                if a: cur['saldo_ini']=a[-1]
                inside=True; continue
            if up.startswith('SALDO AL'):
                a=[num_ar(w['text']) for w in toks if RE_NUM.match(w['text'])]
                if a: cur['saldo_fin']=a[-1]
                inside=False; continue
            if up.startswith('TOTAL MOVIMIENTOS'):
                inside=False; continue
            first=toks[0]
            if inside and DATE.match(first['text']) and first['x0'] < 90 and first['text'] != '00/00':
                mv=dict(fecha=first['text'], origen='', concepto=[], debito=0.0, credito=0.0, saldo=None)
                for w in toks[1:]:
                    x0,x1,tx = w['x0'], w['x1'], w['text']
                    if RE_NUM.match(tx) and x1 > 380:
                        v=num_ar(tx); z=zona(x1)
                        if z=='debito': mv['debito']=abs(v)
                        elif z=='credito': mv['credito']=abs(v)
                        else: mv['saldo']=v
                    elif 90 <= x0 < 133: mv['origen'] += tx+' '
                    elif 133 <= x0 < 385: mv['concepto'].append(tx)
                mv['concepto']=' '.join(mv['concepto']).strip(); mv['origen']=mv['origen'].strip()
                cur['movs'].append(mv)

    res=[]
    for a in accts:
        if a['saldo_ini'] is None: continue     # encabezado-resumen duplicado
        movimientos=[Movimiento(_fecha_ddmm(mv['fecha'], anio), mv['concepto'], mv['origen'],
                                '', '', mv['debito'], mv['credito'], mv['saldo']) for mv in a['movs']]
        sfin = a['saldo_fin'] if a['saldo_fin'] is not None else (movimientos[-1].saldo if movimientos else a['saldo_ini'])
        res.append(Resultado('Banco BBVA', titular, cuit_tit, a['cuenta'], periodo,
                             a['saldo_ini'], sfin, movimientos))
    return res

# ---------------------------------------------------------------------------
# MACRO — monospace ancho fijo, VARIAS CUENTAS por PDF, cuenta cortada entre
# páginas (el encabezado de cuenta se repite: se mantiene el estado si es la
# misma). NO publica totales de movimientos -> control por continuidad de saldo.
# Columnas por borde derecho: débito x1<410 | crédito x1<490 | saldo resto.
# ---------------------------------------------------------------------------
def leer_macro(pdf) -> List[Resultado]:
    DATE = re.compile(r'^\d{2}/\d{2}/\d{2}$')
    ACC  = re.compile(r'CUENTA CORRIENTE.*NRO\.?:\s*([\d\-]+)')
    def zona(x1): return 'debito' if x1 < 410 else ('credito' if x1 < 490 else 'saldo')
    t0 = pdf.pages[0].extract_text() or ''
    titular=''
    lineas=[l.strip() for l in t0.split('\n')]
    for i,l in enumerate(lineas):
        if 'Sr(es)' in l:
            for j in range(i+1, min(i+4, len(lineas))):
                if lineas[j]: titular=lineas[j]; break
            break
    mc=re.search(r'C\.?U\.?I\.?T\.?\s*:?\s*(\d{11})', t0); cuit_tit=mc.group(1) if mc else ''
    mp=re.search(r'Periodo del Extracto:\s*([\d/]+ al [\d/]+)', t0); periodo=mp.group(1) if mp else ''

    accts=[]; cur=None; inside=False; cur_num=None
    for page in pdf.pages:
        for toks in _cluster(page.extract_words()):
            txt=' '.join(w['text'] for w in toks)
            m=ACC.search(txt)
            if m:
                num=m.group(1)
                if num != cur_num:
                    cur=dict(cuenta=num, saldo_ini=None, saldo_fin=None, movs=[])
                    accts.append(cur); cur_num=num; inside=False
                continue
            if cur is None: continue
            up=txt.upper()
            if 'SALDO ULTIMO EXTRACTO' in up:
                a=[num_ar(w['text']) for w in toks if RE_NUM.match(w['text'])]
                if a: cur['saldo_ini']=a[-1]
                inside=True; continue
            if 'SALDO FINAL AL DIA' in up:
                a=[num_ar(w['text']) for w in toks if RE_NUM.match(w['text'])]
                if a: cur['saldo_fin']=a[-1]
                inside=False; continue
            first=toks[0]
            if inside and DATE.match(first['text']) and first['x0'] < 80:
                mv=dict(fecha=first['text'], desc=[], ref='', debito=0.0, credito=0.0, saldo=None)
                for w in toks[1:]:
                    x0,x1,tx = w['x0'], w['x1'], w['text']
                    if RE_NUM.match(tx) and x1 > 320:
                        v=num_ar(tx); z=zona(x1)
                        if z=='debito': mv['debito']=abs(v)
                        elif z=='credito': mv['credito']=abs(v)
                        else: mv['saldo']=v
                    elif 67 <= x0 < 260: mv['desc'].append(tx)
                    elif 260 <= x0 < 339: mv['ref'] += tx+' '
                mv['desc']=' '.join(mv['desc']).strip(); mv['ref']=mv['ref'].strip()
                cur['movs'].append(mv)

    res=[]
    for a in accts:
        if not a['movs']: continue              # cuentas sin movimientos: se omiten
        movimientos=[Movimiento(fecha_dmy2(mv['fecha']), mv['desc'], mv['ref'],
                                '', '', mv['debito'], mv['credito'], mv['saldo']) for mv in a['movs']]
        sfin = a['saldo_fin'] if a['saldo_fin'] is not None else (movimientos[-1].saldo if movimientos else 0.0)
        res.append(Resultado('Banco Macro', titular, cuit_tit, a['cuenta'], periodo,
                             a['saldo_ini'] or 0.0, sfin, movimientos))
    return res

# ===========================================================================
# IMPUTACIÓN DE PROVEEDORES  (CUIT -> importe -> nombre)
# ---------------------------------------------------------------------------
# Usa el archivo de proveedores del cliente (listado + cuenta corriente).
#   1) Por CUIT: el extracto trae el CUIT de la transferencia -> proveedor.
#   2) Por importe exacto: el débito coincide con una factura de compra.
#   3) Si varios proveedores comparten importe: se desempata por nombre en el texto.
# Si no se identifica, queda "a imputar" para revisión humana (NUNCA se adivina).
# ===========================================================================
def cargar_proveedores(solapas: dict) -> dict:
    """`solapas` = {nombre_de_solapa: [filas...]} del archivo 'Listado y Ca
    Cte de Proveedores' del cliente (las filas ya vienen leídas, ej. con
    drive_io.leer_sheet — esta función no sabe de dónde salieron).
    No importa cómo se llamen las solapas: se identifican por sus columnas,
    igual que antes cuando se leía el Excel.
    Devuelve índices para imputar: por_cuit {cuit11 -> nombre},
    por_importe {importe -> {nombres}}."""
    por_cuit = {}; nombres = set(); por_importe = {}
    for filas in solapas.values():
        if not filas:
            continue
        cols = {_norm_concepto(k) for k in filas[0].keys()}
        # Hoja directorio: Nombre + Número de Identificación (CUIT)
        if _norm_concepto('Nombre') in cols and _norm_concepto('Número de Identificación') in cols:
            for fila in filas:
                nom = _valor(fila, 'Nombre'); cid = _valor(fila, 'Número de Identificación')
                if nom and cid:
                    cu = re.sub(r'\D', '', str(cid))
                    if len(cu) == 11: por_cuit[cu] = str(nom).strip()
                    nombres.add(str(nom).strip().upper())
        # Hoja cuenta corriente: importes de Facturas de Compra por proveedor.
        # Filtro EXACTO que pide el plan: solo filas con debehaber = -1
        # ("Factura de Compra"), nunca los pagos.
        elif {_norm_concepto('Proveedor'), _norm_concepto('Haber'), _norm_concepto('debehaber')} <= cols:
            for fila in filas:
                prov = _valor(fila, 'Proveedor'); hab = _valor(fila, 'Haber')
                dh = _valor(fila, 'debehaber')
                try:
                    es_factura_compra = int(float(dh)) == -1
                except (TypeError, ValueError):
                    es_factura_compra = False
                if prov and hab and es_factura_compra:
                    try:
                        importe = round(float(hab), 2)
                    except (TypeError, ValueError):
                        continue
                    por_importe.setdefault(importe, set()).add(str(prov).strip().upper())
    return dict(por_cuit=por_cuit, por_importe=por_importe, nombres=nombres)

def cargar_empleados_socios(solapas: dict) -> dict:
    """`solapas` = {nombre_de_solapa: [filas...]} del mismo archivo de
    Proveedores (la solapa EMPLEADOS Y SOCIOS vive ahí). Devuelve
    {cuit11 -> cuenta contable}. Reemplaza a los diccionarios hardcodeados
    SOCIOS_CUIT / SUELDOS_CUIT: la cuenta (Cuenta Particular del socio,
    Sueldos a Pagar, etc.) sale directo de la columna CUENTA CONTABLE que
    carga el cliente — no hay ninguna distinción hardcodeada por persona."""
    out = {}
    for filas in solapas.values():
        if not filas:
            continue
        cols = {_norm_concepto(k) for k in filas[0].keys()}
        if _norm_concepto('CUIT') in cols and _norm_concepto('CUENTA CONTABLE') in cols:
            for fila in filas:
                cid = _valor(fila, 'CUIT'); cta = _valor(fila, 'CUENTA CONTABLE')
                if cid and cta:
                    cu = re.sub(r'\D', '', str(cid))
                    if len(cu) == 11: out[cu] = str(cta).strip()
    return out

# ===========================================================================
# IMPUTACIÓN DE VEP / IMPUESTOS AFIP  (por importe -> concepto -> cuenta)
# Los pagos de AFIP salen del extracto como un importe (sin decir qué impuesto).
# Se cruza contra el listado de VEP del cliente (hoja 'VEP' del liquidador:
# Importe + Descripcion, ej. 'IVA DJ01/26', 'SIJPDJ02/26', 'IIBBBA02/26').
# Si el mismo importe está en varios VEP, se desempata por fecha más cercana.
# ===========================================================================
def _parse_fecha_vep(txt):
    """'Fecha de Pago' de la solapa VEP viene en DOS formatos mezclados
       ('13/11/2024 17:16' y '2026-06-11 19:49:39' — Google Sheets los
       formatea distinto según cómo se cargó la fila). Sin match -> None
       (esa fila deja de poder desempatar por fecha, pero si es la única
       con ese importe igual se usa)."""
    if not txt:
        return None
    txt = str(txt).strip()
    m = re.match(r'(\d{2})/(\d{2})/(\d{4})', txt)
    if m:
        d, mo, y = m.groups()
    else:
        m = re.match(r'(\d{4})-(\d{2})-(\d{2})', txt)
        if not m:
            return None
        y, mo, d = m.groups()
    try:
        return datetime.date(int(y), int(mo), int(d))
    except ValueError:
        return None

def cargar_vep(solapas: dict) -> dict:
    """`solapas` = {nombre_de_solapa: [filas...]} del archivo 'Vep Pagados'
    del cliente. Devuelve:
      por_importe : {importe -> [(descripcion, fecha_o_None), ...]}
                    (solo las filas con Estado = 'Pagado')
      conceptos   : [(raiz_normalizada, cuenta), ...] de la solapa
                    CONCEPTOS VEP, ordenada por raíz más larga primero
                    (para que una raíz más específica gane antes que una
                    más corta que también matchea)."""
    por_importe = {}; conceptos = []
    for filas in solapas.values():
        if not filas:
            continue
        cols = {_norm_concepto(k) for k in filas[0].keys()}
        if _norm_concepto('Importe') in cols and (_norm_concepto('Descripcion') in cols or _norm_concepto('Descripción') in cols):
            for fila in filas:
                estado = _valor(fila, 'Estado')
                if estado and _norm_concepto(estado) != _norm_concepto('Pagado'):
                    continue  # solo VEP efectivamente pagados
                imp = _valor(fila, 'Importe')
                if imp is None:
                    continue
                try:
                    imp = round(num_ar(str(imp)), 2)
                except (TypeError, ValueError):
                    continue
                desc = str(_valor(fila, 'Descripcion', 'Descripción') or '').strip()
                fec = _parse_fecha_vep(_valor(fila, 'Fecha de Pago'))
                por_importe.setdefault(imp, []).append((desc, fec))
        elif _norm_concepto('CONCEPTO VEP') in cols and _norm_concepto('CUENTA CONTABLE') in cols:
            for fila in filas:
                raiz = _valor(fila, 'CONCEPTO VEP'); cta = _valor(fila, 'CUENTA CONTABLE')
                if raiz and cta:  # celda vacía (ej. ARCA, MULTA, ART sin definir) -> no se carga, queda pendiente
                    conceptos.append((_norm_concepto(raiz), str(cta).strip()))
    conceptos.sort(key=lambda t: -len(t[0]))
    return dict(por_importe=por_importe, conceptos=conceptos)

def _cuenta_por_concepto_vep(desc: str, conceptos: list):
    """Busca en `conceptos` (de cargar_vep) la primera raíz con la que el
       concepto puntual del VEP (ej. 'SIJPDJ05/26') EMPIEZA. Sin match ->
       None (pendiente; nunca se adivina)."""
    d = _norm_concepto(desc)
    for raiz, cta in conceptos:
        if raiz and d.startswith(raiz):
            return cta
    return None

# Firmas por texto EN MAYÚSCULAS. Ojo: el extracto de Galicia puede contener la
# palabra "BBVA" (en un detalle de transferencia), y en el PDF de BBVA el nombre
# y CUIT del banco vienen como imagen. Por eso se detecta por texto estructural
# propio de cada banco, no por el nombre del banco.
_FIRMAS = [
    ('Banco Comafi', leer_comafi,     lambda t: 'COMAFI' in t or '30-60473101-8' in t),
    ('Supervielle',  leer_supervielle,lambda t: 'SUPERVIELLE' in t or '33-50000517-9' in t or 'IAUREG010000' in t or 'AJBXP' in t),
    ('Banco Ciudad', leer_ciudad,     lambda t: '30-99903208-3' in t or 'BANCOCIUDAD' in t or 'BANCO CIUDAD' in t),
    ('Banco Macro',  leer_macro,      lambda t: 'MACRO' in t or '30-50001008-4' in t),
    ('Banco BBVA',   leer_bbva,       lambda t: 'CTA.CTE.BANCARIA' in t or 'CUENTA PYME' in t or 'SALDO ANTERIOR' in t),
    ('Banco Galicia',leer_galicia,    lambda t: 'RESUMEN DE CUENTA CORRIENTE' in t),
    ('Banco Galicia',leer_galicia_ob, lambda t: 'MOVIMIENTOS DE CC' in t and 'OFFICE BANKING' in t),
]

def detectar_banco(pdf):
    texto = "\n".join((pdf.pages[i].extract_text() or '') for i in range(min(2, len(pdf.pages)))).upper()
    for nombre, lector, test in _FIRMAS:
        if test(texto):
            return nombre, lector
    raise ValueError("No pude identificar el banco de este PDF. "
                     "Agregá su firma y su lector en el módulo.")

def leer_extracto(path: str) -> List[Resultado]:
    """Devuelve SIEMPRE una lista de Resultado (una por cuenta).
       Bancos de una sola cuenta devuelven una lista de un elemento."""
    pdf = pdfplumber.open(path)
    _, lector = detectar_banco(pdf)
    out = lector(pdf)
    return out if isinstance(out, list) else [out]

# ===========================================================================
# Control de integridad + exportación a Excel
# ===========================================================================
def verificar_control(res: Resultado):
    tot_d = sum(m.debito for m in res.movimientos)
    tot_c = sum(m.credito for m in res.movimientos)
    calc  = res.saldo_ini + tot_c - tot_d
    dif   = round(calc - res.saldo_fin, 2)
    acum, parciales = res.saldo_ini, 0
    for m in res.movimientos:
        acum += m.credito - m.debito
        if m.saldo is not None and abs(acum - m.saldo) > 0.01:
            parciales += 1
    return dict(total_debitos=tot_d, total_creditos=tot_c, saldo_calculado=calc,
                diferencia=dif, ok=(dif == 0 and parciales == 0),
                saldos_parciales_mal=parciales)

# ===========================================================================
# CLASIFICACIÓN CONTABLE — TODO sale de los 4 archivos del cliente (Fase 2).
# Cero reglas de negocio en el código: si algo no se puede resolver con los
# datos del cliente, el movimiento queda "a imputar" para revisión humana.
# ===========================================================================
import unicodedata as _ud
def _sin_acentos(s: str) -> str:
    return ''.join(c for c in _ud.normalize('NFD', str(s)) if _ud.category(c) != 'Mn')

def _texto_mov(m: Movimiento) -> str:
    # En Galicia el nombre/detalle (p.ej. 'HABERES') va en 'nombre'; se incluye.
    return f"{m.concepto} {m.nombre} {m.referencia}"

def _norm_concepto(s: str) -> str:
    """Normaliza texto para comparar TOLERANDO mayúsculas, acentos y espacios
       de más (Regla de oro #4). Se usa para conceptos, cuentas, nombres de
       prestador/proveedor y nombres de columna de los Sheets — cualquier
       texto que carguen personas y pueda tener variaciones de tipeo."""
    s = _sin_acentos(str(s or '')).upper()
    s = re.sub(r'\s+', ' ', s).strip()
    return s.rstrip(' .')

def _valor(fila: dict, *nombres_col):
    """Valor de una fila (dict, como las que devuelve drive_io.leer_sheet)
       por nombre de columna, tolerante a mayúsculas/acentos/espacios en el
       encabezado. Prueba cada nombre en `nombres_col` en orden."""
    idx = {_norm_concepto(k): v for k, v in fila.items()}
    for n in nombres_col:
        v = idx.get(_norm_concepto(n))
        if v not in (None, ''):
            return v
    return None

def cargar_conceptos(solapas: dict) -> dict:
    """`solapas` = {nombre_de_solapa: [filas...]} del archivo 'CONCEPTO
    GASTO Y ASIGNACION CUENTA' del cliente. No importa cómo se llamen las
    solapas: se identifican por sus columnas.
    Devuelve:
      mapeo     : {concepto normalizado -> cuenta}   (solapa principal, cuenta ÚNICA)
      ambiguos  : {concepto normalizado, ...}         (solapa principal, varias cuentas -> "por jerarquía")
      servicios : {prestador normalizado -> cuenta}   (solapa SERVICIOS)
    Concepto con celda de cuenta vacía, o ambiguo, NO se auto-asigna acá
    (requiere el paso siguiente en clasificar(): contraparte / VEP / humano)."""
    tmp, servicios = {}, {}
    for filas in solapas.values():
        if not filas:
            continue
        cols = {_norm_concepto(k) for k in filas[0].keys()}
        if _norm_concepto('CONCEPTO EXTRACTO') in cols and _norm_concepto('CUENTA CONTABLE') in cols:
            for fila in filas:
                con = _valor(fila, 'CONCEPTO EXTRACTO'); cta = _valor(fila, 'CUENTA CONTABLE')
                if con and cta:
                    tmp.setdefault(_norm_concepto(con), set()).add(str(cta).strip())
        elif _norm_concepto('PRESTADOR') in cols and _norm_concepto('CUENTA CONTABLE') in cols:
            for fila in filas:
                prest = _valor(fila, 'PRESTADOR'); cta = _valor(fila, 'CUENTA CONTABLE')
                if prest and cta:
                    servicios[_norm_concepto(prest)] = str(cta).strip()
    mapeo = {k: next(iter(v)) for k, v in tmp.items() if len(v) == 1}
    ambiguos = {k for k, v in tmp.items() if len(v) > 1}
    return dict(mapeo=mapeo, ambiguos=ambiguos, servicios=servicios)

def cargar_plan_cuentas(solapas: dict) -> set:
    """`solapas` = {nombre_de_solapa: [filas...]} del archivo PLAN DE
    CUENTAS del cliente. Devuelve el conjunto de nombres de cuenta válidos
    (normalizados) para el control cruzado de clasificar()."""
    cuentas = set()
    for filas in solapas.values():
        if not filas:
            continue
        cols = {_norm_concepto(k) for k in filas[0].keys()}
        if _norm_concepto('Nombre') in cols:
            for fila in filas:
                nom = _valor(fila, 'Nombre')
                if nom:
                    cuentas.add(_norm_concepto(nom))
    return cuentas

def clasificar(res: Resultado, conceptos: dict, empleados_socios: dict,
               proveedores: dict, vep: dict, plan_cuentas: set = None,
               tol_dias_vep: int = 45) -> Resultado:
    """Asigna cuenta contable a cada movimiento leyendo TODO de los 4
    archivos del cliente (nada hardcodeado). Por cada movimiento, en orden:
      1) Concepto EXACTO en la tabla de conceptos del cliente, si tiene una
         sola cuenta posible (comparación tolerante).
      2) Si no (no está, o está "por jerarquía" con varias cuentas) ->
         se busca la contraparte, en este orden:
           a) Empleados y Socios, por CUIT (solo débitos).
           b) Proveedores: CUIT -> razón social -> importe exacto de
              factura de compra (2+ importes iguales sin desempatar por
              nombre -> sigue sin resolver, NO se adivina).
           c) Servicios: nombre del prestador (EDESUR, AYSA, ...) dentro
              del detalle del movimiento.
      3) Si sigue sin cuenta -> se prueba como pago de VEP/AFIP: importe
         exacto contra el listado de VEP pagados del cliente (desempate por
         fecha de pago más cercana si hay varios), y de ahí la cuenta por
         la solapa CONCEPTOS VEP (la Descripcion del VEP "empieza con" la
         raíz cargada, ej. 'SIJPDJ05/26' empieza con 'SIJP').
      4) Si nada de esto resolvió -> se etiqueta 'PENDIENTE DE IMPUTAR'
         (sin cuenta) para revisión humana. Nunca se adivina.
    Control cruzado: si se pasa `plan_cuentas` (de cargar_plan_cuentas) y la
    cuenta resuelta no figura ahí, el movimiento se manda igual a pendiente
    y se avisa por print (atrapa cuentas mal cargadas, ej. 'AFIP-RENTAS')."""
    mapeo = conceptos.get('mapeo', {})
    servicios = conceptos.get('servicios', {})
    por_cuit_emp = empleados_socios or {}
    por_cuit_prov = (proveedores or {}).get('por_cuit', {})
    por_importe_prov = (proveedores or {}).get('por_importe', {})
    por_importe_vep = (vep or {}).get('por_importe', {})
    conceptos_vep = (vep or {}).get('conceptos', [])

    def _dif_fecha(item, fmov):
        f = item[1]
        return abs((f - fmov).days) if (f and fmov) else 10 ** 6

    for m in res.movimientos:
        if m.categoria:
            continue
        cuit_norm = re.sub(r'\D', '', str(m.cuit or ''))

        # 1) Concepto exacto (sin ambigüedad) en la tabla del cliente.
        for k in (_norm_concepto(m.concepto), _norm_concepto(f"{m.concepto} {m.nombre}")):
            if k in mapeo:
                m.categoria = 'CONCEPTO'; m.cuenta_sugerida = mapeo[k]
                break
        if m.categoria:
            continue

        # 2a) Empleados y Socios, por CUIT.
        if m.debito > 0 and cuit_norm in por_cuit_emp:
            m.categoria = 'EMPLEADO/SOCIO'; m.cuenta_sugerida = por_cuit_emp[cuit_norm]
            continue

        # 2b) Proveedores: CUIT -> importe exacto (desempate por nombre en el detalle).
        matched = None
        if len(cuit_norm) == 11 and cuit_norm in por_cuit_prov:
            matched = por_cuit_prov[cuit_norm]
        elif m.debito > 0:
            cand = por_importe_prov.get(round(m.debito, 2))
            if cand and len(cand) == 1:
                matched = next(iter(cand))
            elif cand and len(cand) > 1:
                txt = _texto_mov(m).upper()
                hits = [n for n in cand if n in txt]
                if len(hits) == 1:
                    matched = hits[0]
        if matched:
            m.categoria = 'PROVEEDOR'; m.cuenta_sugerida = 'Proveedores'
            if not m.nombre:
                m.nombre = matched
            continue

        # 2c) Servicios: nombre del prestador en el detalle del movimiento.
        if servicios:
            txt_detalle = _norm_concepto(_texto_mov(m))
            for prestador_norm, cta in servicios.items():
                if prestador_norm and prestador_norm in txt_detalle:
                    m.categoria = 'SERVICIO'; m.cuenta_sugerida = cta
                    break
            if m.categoria:
                continue

        # 3) VEP: importe exacto -> concepto (solapa CONCEPTOS VEP, "empieza con").
        if m.debito > 0:
            cand = por_importe_vep.get(round(m.debito, 2))
            if cand:
                if len(cand) == 1:
                    elegido = cand[0]
                else:
                    mejor = min(cand, key=lambda it: _dif_fecha(it, m.fecha))
                    elegido = mejor if _dif_fecha(mejor, m.fecha) <= tol_dias_vep else None
                if elegido:
                    cta = _cuenta_por_concepto_vep(elegido[0], conceptos_vep)
                    if cta:
                        m.categoria = 'IMPUESTO AFIP'; m.cuenta_sugerida = cta; m.nombre = elegido[0]
                        continue

        # 4) Nada resolvió -> pendiente de imputar (etiqueta visual, sin cuenta).
        #    Mensaje mas especifico segun lo que se sabe del movimiento, para
        #    orientar a la persona sobre que planilla del cliente completar
        #    (no es una cuenta -- solo un diagnostico, nunca se adivina).
        cuit_final = re.sub(r'\D', '', str(m.cuit or ''))
        if len(cuit_final) == 11:
            m.categoria = (f"PENDIENTE DE IMPUTAR: CUIT {m.cuit} no está en "
                            f"Proveedores ni en Empleados y Socios")
        elif (m.nombre or '').strip():
            m.categoria = (f"PENDIENTE DE IMPUTAR: '{m.nombre}' no está identificado "
                            f"como proveedor/prestador en ninguna planilla")
        else:
            m.categoria = f"PENDIENTE DE IMPUTAR: concepto '{m.concepto}' no está en la planilla de Conceptos"

    # Control cruzado con el Plan de Cuentas del cliente.
    if plan_cuentas is not None:
        for m in res.movimientos:
            if m.cuenta_sugerida and _norm_concepto(m.cuenta_sugerida) not in plan_cuentas:
                print(f"[clasificar] Aviso: la cuenta '{m.cuenta_sugerida}' (movimiento "
                      f"'{m.concepto}' del {m.fecha}) no existe en el Plan de Cuentas del "
                      f"cliente -> se manda a pendientes.")
                m.categoria = 'PENDIENTE DE IMPUTAR (cuenta no existe en Plan de Cuentas)'
                m.cuenta_sugerida = ''
    return res

def _sanitizar_hoja(nombre, usados):
    for ch in '\\/?*[]:':
        nombre = nombre.replace(ch, '-')
    nombre = (nombre or 'Cuenta')[:28] or 'Cuenta'
    base=nombre; i=2
    while nombre in usados:
        nombre=f"{base[:25]}_{i}"; i+=1
    usados.add(nombre)
    return nombre

def _escribir_hoja(ws, res):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    AR=Font(name='Arial',size=10); ARB=Font(name='Arial',size=10,bold=True)
    TIT=Font(name='Arial',size=12,bold=True,color='FFFFFF'); HF=Font(name='Arial',size=10,bold=True,color='FFFFFF')
    FILL=PatternFill('solid',fgColor='1F4E78'); GREEN=PatternFill('solid',fgColor='C6EFCE')
    YEL=PatternFill('solid',fgColor='FFF2CC')
    MONEY='#,##0.00;[Red]-#,##0.00'; BD=Border(bottom=Side(style='thin',color='D9D9D9')); CEN=Alignment(horizontal='center')
    BLUE=PatternFill('solid',fgColor='DDEBF7'); ORANGE=PatternFill('solid',fgColor='FCE4D6')
    ws['A1']=f"{res.banco.upper()} — {res.titular} (CUIT {res.cuit_titular}) · Cta {res.cuenta} · {res.periodo}"
    ws['A1'].font=TIT; ws['A1'].fill=FILL; ws.merge_cells('A1:K1')
    ws['A1'].alignment=Alignment(horizontal='left',vertical='center'); ws.row_dimensions[1].height=20
    cols=['Fecha','Concepto','Referencia','CUIT','Nombre','Cuenta sugerida',
          'Débitos','Créditos','Saldo (banco)','Saldo calculado','Control']
    for j,c in enumerate(cols,1):
        cc=ws.cell(3,j,c); cc.font=HF; cc.fill=FILL; cc.alignment=CEN
    r=4; first=r
    for m in res.movimientos:
        if m.fecha is not None: ws.cell(r,1,m.fecha).number_format='dd/mm/yyyy'
        ws.cell(r,2,m.concepto); ws.cell(r,3,m.referencia); ws.cell(r,4,m.cuit); ws.cell(r,5,m.nombre)
        etiqueta = m.cuenta_sugerida or m.categoria
        cs=ws.cell(r,6,etiqueta)
        if m.cuenta_sugerida: cs.fill=BLUE
        elif m.categoria: cs.fill=ORANGE
        ws.cell(r,7,m.debito or None).number_format=MONEY
        ws.cell(r,8,m.credito or None).number_format=MONEY
        if m.saldo is not None: ws.cell(r,9,m.saldo).number_format=MONEY
        ws.cell(r,10, f'=$N$2+H{r}-G{r}' if r==first else f'=J{r-1}+H{r}-G{r}').number_format=MONEY
        ws.cell(r,11, f'=IF(I{r}="","",I{r}-J{r})').number_format=MONEY
        for j in range(1,12): ws.cell(r,j).font=AR; ws.cell(r,j).border=BD
        r+=1
    last=r-1
    n_cuit=sum(1 for m in res.movimientos if m.cuit)
    n_cta =sum(1 for m in res.movimientos if m.cuenta_sugerida)
    n_imp =sum(1 for m in res.movimientos if m.categoria and not m.cuenta_sugerida)
    n_sin =sum(1 for m in res.movimientos if not m.categoria)
    filas=[('Saldo inicial',res.saldo_ini),('Total Débitos',f'=SUM(G{first}:G{last})'),
           ('Total Créditos',f'=SUM(H{first}:H{last})'),('Saldo final calculado','=N2+N4-N3'),
           ('Saldo final real (banco)',res.saldo_fin),('CONTROL (debe dar 0)','=N5-N6'),
           ('Cant. movimientos',f'=COUNT(A{first}:A{last})'),
           ('Con cuenta sugerida',n_cta),('A imputar (revisar)',n_imp),('Sin clasificar',n_sin)]
    for i,(lbl,val) in enumerate(filas):
        rr=2+i; ws.cell(rr,13,lbl).font=ARB
        c=ws.cell(rr,14,val); c.font=AR; c.number_format=MONEY if i<6 else '0'
    ws['N2'].fill=YEL; ws['N6'].fill=YEL; ws['N7'].fill=GREEN
    for j,w in enumerate([11,32,13,13,20,22,15,15,15,15,12],1):
        ws.column_dimensions[get_column_letter(j)].width=w
    ws.column_dimensions['M'].width=24; ws.column_dimensions['N'].width=16
    ws.freeze_panes='A4'

def exportar_excel(resultados, path):
    """Acepta un Resultado o una lista. Escribe UNA HOJA por cuenta.
       OJO: ya NO clasifica acá adentro (antes lo hacía con criterios
       hardcodeados). Clasificar requiere los datos del cliente (Fase 2:
       cargar_conceptos/cargar_empleados_socios/cargar_proveedores/cargar_vep
       + clasificar) -- eso lo hace el llamador ANTES de exportar."""
    import openpyxl
    if not isinstance(resultados, list): resultados=[resultados]
    wb=openpyxl.Workbook(); primera=True; usados=set()
    for res in resultados:
        ws = wb.active if primera else wb.create_sheet(); primera=False
        ws.title=_sanitizar_hoja(res.cuenta, usados)
        _escribir_hoja(ws, res)
    wb.save(path)

# ===========================================================================
# CLI
# ===========================================================================
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Uso: python conversor_extractos.py archivo.pdf [salida.xlsx]"); sys.exit(1)
    entrada = sys.argv[1]
    resultados = leer_extracto(entrada)
    print(f"Banco   : {resultados[0].banco}")
    print(f"Titular : {resultados[0].titular}  (CUIT {resultados[0].cuit_titular})")
    print(f"Cuentas : {len(resultados)}")
    todo_ok=True
    for res in resultados:
        ctrl=verificar_control(res); ok=ctrl['ok']; todo_ok = todo_ok and ok
        print(f"  · Cta {res.cuenta:22} {len(res.movimientos):3} movs | "
              f"saldo {res.saldo_ini:,.2f} -> {res.saldo_fin:,.2f} | "
              f"control {ctrl['diferencia']:,.2f}  {'OK ✓' if ok else 'REVISAR ✗'}")
    print("Nota    : este modo de línea de comandos solo lee y controla el PDF -- ya "
          "no clasifica sin los archivos del cliente (ver Fase 2 / drive_io).")
    import os
    salida = sys.argv[2] if len(sys.argv) > 2 else os.path.basename(entrada).rsplit('.',1)[0] + '_convertido.xlsx'
    exportar_excel(resultados, salida)
    print(f"Excel   : {salida}   ({'TODO OK' if todo_ok else 'HAY CUENTAS A REVISAR'})")
