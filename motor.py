# -*- coding: utf-8 -*-
"""
Motor Control ORAMI
- Lee el buzón transacciones@nacionalissimo.com (IMAP)
- Correo del banco (Banorte, transferencia SPEI) -> crea recarga PENDIENTE + manda notificacion push
- Correo de ORAMI (xlsx estado de cuenta) -> agrega/actualiza movimientos y marca recargas VERIFICADAS por clave de rastreo
Se ejecuta en GitHub Actions (cron).
"""
import os, ssl, imaplib, email, re, json, io, zipfile
import xml.etree.ElementTree as ET
from email.header import decode_header
from datetime import datetime, timedelta

import firebase_admin
from firebase_admin import credentials, firestore, messaging

# ---------- Config desde variables de entorno (GitHub Secrets) ----------
IMAP_HOST = os.environ.get("IMAP_HOST", "mail.nacionalissimo.com")
IMAP_PORT = int(os.environ.get("IMAP_PORT", "993"))
IMAP_USER = os.environ["IMAP_USER"]
IMAP_PASS = os.environ["IMAP_PASS"]
SA_JSON   = os.environ["FIREBASE_SA"]   # contenido JSON de la cuenta de servicio

SALDO_INICIAL = 23820.0  # se ajustara con el ciclo mensual mas adelante

# ---------- Firebase ----------
cred = credentials.Certificate(json.loads(SA_JSON))
firebase_admin.initialize_app(cred)
db = firestore.client()

def log(*a): print("[motor]", *a, flush=True)

def comision(t): return round((t/1.16)*0.055, 2)
def abono_neto(t): return round(t - comision(t), 2)

# ---------- Utilidades de correo ----------
def decode(s):
    if not s: return ""
    parts = decode_header(s); out = ""
    for txt, enc in parts:
        out += txt.decode(enc or "utf-8", "ignore") if isinstance(txt, bytes) else txt
    return out

def get_body(msg):
    """Devuelve el texto plano del correo (quitando HTML si hace falta)."""
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            if ct in ("text/plain", "text/html") and "attachment" not in str(part.get("Content-Disposition")):
                try: body += part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", "ignore")
                except Exception: pass
    else:
        try: body = msg.get_payload(decode=True).decode(msg.get_content_charset() or "utf-8", "ignore")
        except Exception: body = str(msg.get_payload())
    body = re.sub(r"<[^>]+>", " ", body)   # quitar tags HTML
    body = re.sub(r"&nbsp;", " ", body)
    return re.sub(r"\s+", " ", body)

def get_xlsx_attachment(msg):
    for part in msg.walk():
        fn = decode(part.get_filename())
        if fn and fn.lower().endswith(".xlsx"):
            return part.get_payload(decode=True)
    return None

# ---------- Notificaciones push ----------
def enviar_push(titulo, texto):
    tokens = [d.to_dict().get("token") for d in db.collection("tokens").stream()]
    tokens = [t for t in tokens if t]
    enviados = 0
    for t in tokens:
        try:
            messaging.send(messaging.Message(
                notification=messaging.Notification(title=titulo, body=texto),
                token=t,
                webpush=messaging.WebpushConfig(notification=messaging.WebpushNotification(icon="icon.svg"))
            ))
            enviados += 1
        except Exception as e:
            log("push fallo:", e)
    log(f"push enviado a {enviados}/{len(tokens)} dispositivos")

# ---------- Procesar correo del BANCO (transferencia) ----------
def procesar_banco(body):
    importe = re.search(r"Importe a Transferir:\s*\$?\s*([\d,]+\.\d{2})", body)
    clave   = re.search(r"Clave de Rastreo:\s*([A-Z0-9]+)", body)
    fechah  = re.search(r"Fecha y Hora de Operaci[oó]n:\s*([0-9]{1,2}/\w+/[0-9]{4}[^A-Za-z]*\d{2}:\d{2}:\d{2})", body)
    ref     = re.search(r"N[uú]mero de Referencia:\s*(\d+)", body)
    benef   = re.search(r"Nombre del Beneficiario:\s*([^N]+?)\s+(?:CLABE|RFC)", body)
    if not (importe and clave):
        log("correo banco sin importe/clave, ignorado"); return
    monto = float(importe.group(1).replace(",", ""))
    cve = clave.group(1)
    doc_id = "rec-" + cve
    ref_doc = db.collection("movimientos").document(doc_id)
    if ref_doc.get().exists:
        log("transferencia ya registrada:", cve); return
    fh = fechah.group(1).strip() if fechah else ""
    try:
        dt = datetime.strptime(re.sub(r"\s+", " ", fh)[:20], "%d/%b/%Y %H:%M:%S")
    except Exception:
        dt = datetime.now()
    orden = (dt - datetime(1899,12,30)).total_seconds()/86400.0
    MESES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
    ref_doc.set({
        "tipo":"abono","servicio":"Recarga (transferencia)",
        "fecha":"%d-%s"%(dt.day, MESES[dt.month-1]),
        "fechaFull":"%d-%s-%d"%(dt.day, MESES[dt.month-1], dt.year),
        "hora":dt.strftime("%H:%M:%S"),"recibo":ref.group(1) if ref else "",
        "banco":"SPEI a ORAMI - Clave "+cve,
        "transferencia":monto,"estado":"pendiente","orden":orden,
        "comprobante":{
            "banco":"Banorte en su Empresa","operacion":"Transferencia SPEI",
            "fechaHora":fh,"ordenante":"MIGUEL ANGEL VALLEJO FLORES",
            "beneficiario":(benef.group(1).strip() if benef else "CONSULTORIA EMPRESARIAL ORAMI S.C."),
            "clabe":"","bancoDestino":"BAJIO","importe":"$%s MN"%importe.group(1),
            "referencia":ref.group(1) if ref else "","concepto":"ORAMI",
            "aplicacion":fh.split(" ")[0] if fh else "","clave":cve
        }
    })
    log("recarga creada:", cve, monto)
    enviar_push("Transferencia registrada", "Transferiste $%s a ORAMI. Recarga pendiente de confirmar." % importe.group(1))

# ---------- Procesar correo de ORAMI (xlsx) ----------
def parse_xlsx(data):
    ns = {'a':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    z = zipfile.ZipFile(io.BytesIO(data))
    shared = []
    if 'xl/sharedStrings.xml' in z.namelist():
        r = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in r.findall('a:si', ns):
            shared.append(''.join(t.text or '' for t in si.findall('.//a:t', ns)))
    sheet = sorted([n for n in z.namelist() if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')])[0]
    root = ET.fromstring(z.read(sheet))
    def ci(ref):
        L = ''.join(c for c in ref if c.isalpha()); idx=0
        for c in L: idx=idx*26+(ord(c)-64)
        return idx-1
    rows=[]
    for row in root.findall('.//a:row', ns):
        cells={}; mx=-1
        for c in row.findall('a:c', ns):
            v=c.find('a:v', ns); val=''
            if v is not None and v.text is not None:
                val = shared[int(v.text)] if c.get('t')=='s' else v.text
            k=ci(c.get('r')); cells[k]=val; mx=max(mx,k)
        rows.append([cells.get(i,'') for i in range(mx+1)])
    return rows

def servicio(d):
    d=d.upper()
    if 'SPEI RECIBIDO' in d: return 'Recarga (transferencia)'
    if 'FACEBK' in d or 'FACEBOOK' in d or 'METAPAY' in d: return 'Facebook'
    if 'SHOPIFY' in d: return 'Shopify'
    if 'ANTHROPIC' in d or 'CLAUDE' in d: return 'Claude (Anthropic)'
    if 'MICROSOFT' in d: return 'Microsoft'
    if 'MERCADOLIBRE' in d or 'MERPAGO' in d: return 'Mercado Libre'
    if 'RAILWAY' in d: return 'Railway'
    if 'ADOBE' in d: return 'Adobe'
    if 'MAILCHIMP' in d: return 'Mailchimp'
    return 'Otro'
def merchant(d):
    m=re.search(r'5161020004119535 (.+?) Tarjeta', d)
    return m.group(1).strip() if m else d[:50]

def procesar_orami(data):
    rows = parse_xlsx(data)
    MESES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
    nuevos=0; verif=0
    for r in rows:
        if not r or not str(r[0]).strip().isdigit(): continue
        serial=float(r[1]); dt=datetime(1899,12,30)+timedelta(days=serial)
        desc=r[4] if len(r)>4 else ''; cargo=r[5] if len(r)>5 else ''; abono=r[6] if len(r)>6 else ''
        recibo=str(r[3]) if len(r)>3 else ''
        base={"orden":serial,"servicio":servicio(desc),
              "fecha":"%d-%s"%(dt.day,MESES[dt.month-1]),
              "fechaFull":"%d-%s-%d"%(dt.day,MESES[dt.month-1],dt.year),
              "hora":(datetime(1899,12,30)+timedelta(days=float(r[2]))).strftime("%H:%M:%S") if len(r)>2 and r[2] else dt.strftime("%H:%M:%S"),
              "recibo":recibo}
        if cargo not in ('',None):
            doc_id="mov-"+recibo
            base.update({"tipo":"cargo","monto":round(float(cargo),2),"banco":merchant(desc),"estado":"verificada"})
            db.collection("movimientos").document(doc_id).set(base, merge=True); nuevos+=1
        elif abono not in ('',None):
            # buscar clave de rastreo en la descripcion para casar con la recarga del banco
            cve = re.search(r'Clave de Rastreo:\s*([A-Z0-9]+)', desc)
            if cve:
                rec = db.collection("movimientos").document("rec-"+cve.group(1))
                if rec.get().exists:
                    rec.update({"estado":"verificada"}); verif+=1; continue
            doc_id="mov-"+recibo
            base.update({"tipo":"abono","transferencia":round(float(abono),2),"banco":"SPEI recibido","estado":"verificada"})
            db.collection("movimientos").document(doc_id).set(base, merge=True); nuevos+=1
    log(f"ORAMI procesado: {nuevos} movimientos, {verif} recargas verificadas")
    enviar_push("Estado de cuenta de ORAMI",
                f"Llegó el reporte de ORAMI: {nuevos} movimientos y {verif} recargas verificadas.")

# ---------- Main ----------
def main():
    ctx = ssl.create_default_context()
    M = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=ctx)
    M.login(IMAP_USER, IMAP_PASS)
    M.select("INBOX")
    typ, data = M.search(None, "UNSEEN")
    ids = data[0].split()
    log(f"correos nuevos: {len(ids)}")
    for num in ids:
        typ, d = M.fetch(num, "(RFC822)")
        msg = email.message_from_bytes(d[0][1])
        frm = decode(msg.get("From","")).lower()
        subj = decode(msg.get("Subject",""))
        log("correo de:", frm, "| asunto:", subj)
        try:
            xlsx = get_xlsx_attachment(msg)
            if "banorte" in frm or ("transferencia" in subj.lower() and "spei" in subj.lower()):
                procesar_banco(get_body(msg))
            elif xlsx is not None:
                procesar_orami(xlsx)
            else:
                log("correo no reconocido, se ignora")
        except Exception as e:
            log("ERROR procesando correo:", e)
        M.store(num, "+FLAGS", "\\Seen")
    M.logout()
    log("listo")

if __name__ == "__main__":
    main()
