import os
import re
import sys
import json
import shutil
import sqlite3
import datetime
import unicodedata
from difflib import SequenceMatcher
import pandas as pd
from fpdf import FPDF
from langchain_core.tools import tool

# --- PERCORSI DINAMICI (COMPATIBILE SIA CON PYTHON SIA CON PYINSTALLER .EXE) ---
if getattr(sys, "frozen", False):
    CARTELLA_BASE = os.path.dirname(sys.executable)
else:
    CARTELLA_BASE = os.path.dirname(os.path.abspath(__file__))

FILE_DB = os.path.join(CARTELLA_BASE, "gestionale.db")
CARTELLA_PDF = os.path.join(CARTELLA_BASE, "Documenti_PDF")
CARTELLA_BACKUP = os.path.join(CARTELLA_BASE, "Backup_DB")
FILE_LOGO = os.path.join(CARTELLA_BASE, "logo_azienda.png")

os.makedirs(CARTELLA_PDF, exist_ok=True)
os.makedirs(CARTELLA_BACKUP, exist_ok=True)


def get_conn():
    conn = sqlite3.connect(FILE_DB)
    conn.row_factory = sqlite3.Row
    return conn


def inizializza_database():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS magazzino (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            Codice TEXT UNIQUE,
            Descrizione TEXT,
            Giacenza INTEGER DEFAULT 0,
            Prezzo_Acq_Medio REAL DEFAULT 0.0,
            Prezzo_Listino REAL DEFAULT 0.0,
            Extra_JSON TEXT DEFAULT '{}'
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS documenti (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            Data TEXT,
            Tipo TEXT,
            Ruolo TEXT DEFAULT 'Cliente',
            Numero INTEGER,
            Cliente TEXT,
            Voci TEXT,
            Importo_Totale REAL DEFAULT 0.0,
            Fatturato TEXT DEFAULT 'NO',
            Dettaglio_JSON TEXT DEFAULT '[]'
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS soggetti (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            Nome TEXT UNIQUE,
            Ruolo TEXT DEFAULT 'Cliente',
            Partita_IVA TEXT DEFAULT '',
            Codice_Fiscale TEXT DEFAULT '',
            Indirizzo TEXT DEFAULT '',
            Citta TEXT DEFAULT '',
            Email_PEC TEXT DEFAULT '',
            Telefono TEXT DEFAULT ''
        )
    """)

    # Migrazione automatica colonne per database esistenti
    col_mag = [r[1] for r in cur.execute("PRAGMA table_info(magazzino)").fetchall()]
    if "Extra_JSON" not in col_mag:
        cur.execute("ALTER TABLE magazzino ADD COLUMN Extra_JSON TEXT DEFAULT '{}'")

    col_doc = [r[1] for r in cur.execute("PRAGMA table_info(documenti)").fetchall()]
    if "Dettaglio_JSON" not in col_doc:
        cur.execute("ALTER TABLE documenti ADD COLUMN Dettaglio_JSON TEXT DEFAULT '[]'")
    if "Ruolo" not in col_doc:
        cur.execute("ALTER TABLE documenti ADD COLUMN Ruolo TEXT DEFAULT 'Cliente'")

    col_sog = [r[1] for r in cur.execute("PRAGMA table_info(soggetti)").fetchall()]
    for c_req in ["Ruolo", "Partita_IVA", "Codice_Fiscale", "Indirizzo", "Citta", "Email_PEC", "Telefono"]:
        if c_req not in col_sog:
            cur.execute(f"ALTER TABLE soggetti ADD COLUMN {c_req} TEXT DEFAULT ''")

    conn.commit()
    conn.close()


inizializza_database()


# ==============================================================================
# BACKUP & SICUREZZA DATABASE
# ==============================================================================
def esegui_backup_automatico(forza: bool = False):
    if not os.path.exists(FILE_DB):
        return
    oggi_str = datetime.date.today().strftime("%Y%m%d")
    destinazione = os.path.join(CARTELLA_BACKUP, f"gestionale_backup_{oggi_str}.db")
    if forza or not os.path.exists(destinazione):
        try:
            shutil.copy2(FILE_DB, destinazione)
            # Mantiene al massimo gli ultimi 30 backup giornalieri
            files = sorted(
                [os.path.join(CARTELLA_BACKUP, f) for f in os.listdir(CARTELLA_BACKUP) if f.endswith(".db")],
                key=os.path.getmtime
            )
            while len(files) > 30:
                os.remove(files.pop(0))
        except Exception as e:
            print(f"Errore backup automatico: {e}")


def esporta_backup_manuale(percorso_destinazione: str):
    if not os.path.exists(FILE_DB):
        inizializza_database()
    shutil.copy2(FILE_DB, percorso_destinazione)


def ripristina_backup_db(percorso_sorgente: str):
    if os.path.exists(FILE_DB):
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(FILE_DB, os.path.join(CARTELLA_BACKUP, f"pre_ripristino_{ts}.db"))
    shutil.copy2(percorso_sorgente, FILE_DB)
    inizializza_database()


esegui_backup_automatico(forza=False)


# ==============================================================================
# RICERCA INTELLIGENTE (FUZZY) DI ARTICOLI E SOGGETTI
# ==============================================================================
def _norm(s) -> str:
    """Normalizza un testo: minuscolo, senza accenti, spazi e punteggiatura.
    'Acro bat Bianco 14 Litri' e 'ACROBAT bianco 14litri' diventano identici."""
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]", "", s.lower())


def _numeri(s: str) -> list:
    return re.findall(r"\d+", s)


def _cerca_fuzzy(testo: str, righe: list, campi: list, soglia: float = 0.72, max_risultati: int = 5):
    """
    Cerca `testo` tra le `righe` (lista di dict) confrontando i `campi`.
    Restituisce (stato, elemento, candidati):
      - ("trovato", dict, [])      -> corrispondenza certa (esatta o molto netta)
      - ("incerto", None, [(score, dict), ...]) -> piu' candidati o solo simili: serve chiedere all'utente
      - ("nessuno", None, [])      -> nessuna corrispondenza ragionevole
    """
    t = _norm(testo)
    if not t:
        return "nessuno", None, []

    esatti = []
    punteggi = []
    for d in righe:
        best = 0.0
        esatto = False
        for c in campi:
            v = _norm(d.get(c, ""))
            if not v:
                continue
            if v == t:
                esatto = True
                best = 1.0
                break
            if (len(t) >= 5 and t in v) or (len(v) >= 5 and v in t):
                s = 0.9
            else:
                sm = SequenceMatcher(None, t, v)
                if sm.real_quick_ratio() < soglia or sm.quick_ratio() < soglia:
                    s = 0.0
                else:
                    s = sm.ratio()
            # Se i numeri sono diversi (es. 14 litri vs 5 litri) NON e' mai una corrispondenza certa
            if s < 1.0 and _numeri(t) != _numeri(v):
                s = min(s, 0.85)
            best = max(best, s)
        if esatto:
            esatti.append(d)
        if best >= soglia:
            punteggi.append((best, d))

    if len(esatti) == 1:
        return "trovato", esatti[0], []
    if len(esatti) > 1:
        return "incerto", None, [(1.0, d) for d in esatti[:max_risultati]]
    if not punteggi:
        return "nessuno", None, []

    punteggi.sort(key=lambda x: -x[0])
    top = punteggi[0][0]
    secondo = punteggi[1][0] if len(punteggi) > 1 else 0.0
    if top >= 0.9 and (top - secondo) >= 0.08:
        return "trovato", punteggi[0][1], []
    return "incerto", None, punteggi[:max_risultati]


def _carica_articoli(cur) -> list:
    cur.execute("SELECT * FROM magazzino")
    return [dict(r) for r in cur.fetchall()]


def _carica_soggetti(cur) -> list:
    cur.execute("SELECT * FROM soggetti")
    return [dict(r) for r in cur.fetchall()]


def _risolvi_articolo(cur, chiave: str, cache: list = None):
    righe = cache if cache is not None else _carica_articoli(cur)
    return _cerca_fuzzy(chiave, righe, ["Codice", "Descrizione"])


def _risolvi_soggetto(cur, nome: str, cache: list = None):
    righe = cache if cache is not None else _carica_soggetti(cur)
    return _cerca_fuzzy(nome, righe, ["Nome"])


def _msg_candidati_articoli(chiave: str, cands: list) -> str:
    elenco = "\n".join(
        f"- [{d['Codice']}] {d['Descrizione']} (giacenza {d['Giacenza']}, somiglianza {s:.0%})"
        for s, d in cands
    )
    return (
        f"Nessuna corrispondenza CERTA per l'articolo '{chiave}'. Articoli simili trovati:\n{elenco}\n"
        f"NON procedere: chiedi all'utente quale di questi intende, oppure se vuole creare un NUOVO articolo."
    )


def _msg_candidati_soggetti(nome: str, cands: list) -> str:
    elenco = "\n".join(f"- {d['Nome']} ({d['Ruolo']}, somiglianza {s:.0%})" for s, d in cands)
    return (
        f"Nessuna corrispondenza CERTA per il soggetto '{nome}'. Soggetti simili in anagrafica:\n{elenco}\n"
        f"NON procedere: chiedi all'utente quale intende, oppure se si tratta di un NUOVO soggetto."
    )


def _genera_codice_libero(cur, prefisso: str = "ART-") -> str:
    cur.execute("SELECT COUNT(*) FROM magazzino")
    n = cur.fetchone()[0] + 1
    while True:
        cod = f"{prefisso}{n:04d}"
        cur.execute("SELECT 1 FROM magazzino WHERE LOWER(Codice) = LOWER(?)", (cod,))
        if not cur.fetchone():
            return cod
        n += 1


# ==============================================================================
# IMPORTAZIONE EXCEL UNIVERSALE (MAGAZZINO & ANAGRAFICA)
# ==============================================================================
def _pulisci_numero(val) -> float:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    s = str(val).strip().replace("€", "").replace(" ", "")
    if not s or s.lower() == "nan":
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def importa_magazzino_da_excel(
    percorso_excel: str,
    mappatura_core: dict,
    colonne_extra: list,
    modalita: str = "sostituisci"
) -> int:
    df = pd.read_excel(percorso_excel).fillna("")
    conn = get_conn()
    cur = conn.cursor()

    if modalita == "sostituisci":
        cur.execute("DELETE FROM magazzino")

    col_cod = mappatura_core.get("Codice", "")
    col_desc = mappatura_core.get("Descrizione", "")
    col_giac = mappatura_core.get("Giacenza", "")
    col_acq = mappatura_core.get("Prezzo_Acq_Medio", "")
    col_list = mappatura_core.get("Prezzo_Listino", "")

    importati = 0
    for idx, row in df.iterrows():
        codice = str(row[col_cod]).strip() if (col_cod and col_cod in df.columns) else ""
        descrizione = str(row[col_desc]).strip() if (col_desc and col_desc in df.columns) else ""

        if not codice and not descrizione:
            continue
        if not codice:
            codice = f"ART-{idx + 1:04d}"
        if not descrizione:
            descrizione = f"Articolo {codice}"

        giacenza = int(round(_pulisci_numero(row[col_giac]))) if (col_giac and col_giac in df.columns) else 0
        p_acq = round(_pulisci_numero(row[col_acq]), 2) if (col_acq and col_acq in df.columns) else 0.0
        p_list = round(_pulisci_numero(row[col_list]), 2) if (col_list and col_list in df.columns) else 0.0

        if p_acq > 0 and p_list == 0:
            p_list = round(p_acq * 1.22 * 1.30, 2)
        elif p_list > 0 and p_acq == 0:
            p_acq = round(p_list / 1.22 / 1.30, 2)

        extra_dict = {}
        for ec in colonne_extra:
            if ec in df.columns:
                val_ec = str(row[ec]).strip()
                if val_ec and val_ec.lower() != "nan":
                    extra_dict[ec] = val_ec

        extra_json = json.dumps(extra_dict, ensure_ascii=False)

        cur.execute("SELECT id, Extra_JSON FROM magazzino WHERE LOWER(Codice) = LOWER(?)", (codice,))
        esistente = cur.fetchone()
        if esistente:
            try:
                old_extra = json.loads(esistente["Extra_JSON"] or "{}")
            except Exception:
                old_extra = {}
            old_extra.update(extra_dict)
            cur.execute("""
                UPDATE magazzino
                SET Descrizione = ?, Giacenza = ?, Prezzo_Acq_Medio = ?, Prezzo_Listino = ?, Extra_JSON = ?
                WHERE id = ?
            """, (descrizione, giacenza, p_acq, p_list, json.dumps(old_extra, ensure_ascii=False), esistente["id"]))
        else:
            cur.execute("""
                INSERT INTO magazzino (Codice, Descrizione, Giacenza, Prezzo_Acq_Medio, Prezzo_Listino, Extra_JSON)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (codice, descrizione, giacenza, p_acq, p_list, extra_json))

        importati += 1

    conn.commit()
    conn.close()
    return importati


def importa_anagrafica_da_excel(
    percorso_excel: str,
    mappatura: dict,
    ruolo_default: str = "Cliente",
    modalita: str = "unisci"
) -> int:
    df = pd.read_excel(percorso_excel).fillna("")
    conn = get_conn()
    cur = conn.cursor()

    if modalita == "sostituisci":
        cur.execute("DELETE FROM soggetti")

    col_nome = mappatura.get("Nome", "")
    if not col_nome or col_nome not in df.columns:
        conn.close()
        raise ValueError("È obbligatorio associare la colonna 'Ragione Sociale / Nome'.")

    importati = 0
    for _, row in df.iterrows():
        nome = str(row[col_nome]).strip()
        if not nome or nome.lower() == "nan":
            continue

        def leggi_campo(chiave):
            c = mappatura.get(chiave, "")
            val = str(row[c]).strip() if (c and c in df.columns) else ""
            return "" if val.lower() == "nan" else val

        ruolo_raw = leggi_campo("Ruolo").lower()
        if "forn" in ruolo_raw:
            ruolo = "Fornitore"
        elif "clien" in ruolo_raw:
            ruolo = "Cliente"
        else:
            ruolo = ruolo_default

        piva = leggi_campo("Partita_IVA")
        cf = leggi_campo("Codice_Fiscale")
        indirizzo = leggi_campo("Indirizzo")
        citta = leggi_campo("Citta")
        pec = leggi_campo("Email_PEC")
        tel = leggi_campo("Telefono")

        cur.execute("SELECT id FROM soggetti WHERE LOWER(Nome) = LOWER(?)", (nome,))
        esistente = cur.fetchone()
        if esistente:
            cur.execute("""
                UPDATE soggetti
                SET Ruolo = ?,
                    Partita_IVA = CASE WHEN ? != '' THEN ? ELSE Partita_IVA END,
                    Codice_Fiscale = CASE WHEN ? != '' THEN ? ELSE Codice_Fiscale END,
                    Indirizzo = CASE WHEN ? != '' THEN ? ELSE Indirizzo END,
                    Citta = CASE WHEN ? != '' THEN ? ELSE Citta END,
                    Email_PEC = CASE WHEN ? != '' THEN ? ELSE Email_PEC END,
                    Telefono = CASE WHEN ? != '' THEN ? ELSE Telefono END
                WHERE id = ?
            """, (ruolo, piva, piva, cf, cf, indirizzo, indirizzo, citta, citta, pec, pec, tel, tel, esistente["id"]))
        else:
            cur.execute("""
                INSERT INTO soggetti (Nome, Ruolo, Partita_IVA, Codice_Fiscale, Indirizzo, Citta, Email_PEC, Telefono)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (nome, ruolo, piva, cf, indirizzo, citta, pec, tel))
        importati += 1

    conn.commit()
    conn.close()
    return importati


# ==============================================================================
# GENERATORE PDF CONTABILI FORMALI (FPDF2)
# ==============================================================================
def _pulisci_testo_pdf(testo: str) -> str:
    s = str(testo or "").replace("€", "EUR").replace("’", "'").replace("–", "-")
    return s.encode("latin-1", "replace").decode("latin-1")


class DocumentoPDF(FPDF):
    def header(self):
        if os.path.exists(FILE_LOGO):
            try:
                self.image(FILE_LOGO, x=12, y=10, h=16)
            except Exception:
                pass
        self.set_font("Helvetica", "B", 15)
        self.set_text_color(30, 35, 42)
        self.cell(0, 7, "EDIL BARRESE", ln=True, align="R")
        self.set_font("Helvetica", "", 9)
        self.set_text_color(100, 105, 115)
        self.cell(0, 5, "Commercio Materiali Edili, Ferramenta e Forniture", ln=True, align="R")
        self.ln(6)
        self.set_draw_color(200, 35, 51)
        self.set_line_width(0.6)
        self.line(12, self.get_y(), 198, self.get_y())
        self.ln(6)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(130, 130, 130)
        self.cell(0, 10, f"Documento generato dal Sistema Gestionale - Pagina {self.page_no()}", align="C")


def _genera_pdf_documento(
    tipo: str,
    numero: int,
    data_str: str,
    soggetto_dict: dict,
    righe: list,
    totale_imponibile: float,
    totale_iva: float,
    totale_documento: float,
    note: str = ""
) -> str:
    pdf = DocumentoPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=20)

    # Intestazione Documento & Spett.le Cliente
    y_start = pdf.get_y()
    pdf.set_fill_color(245, 247, 250)
    pdf.rect(12, y_start, 88, 30, style="F")
    pdf.set_xy(15, y_start + 3)
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(200, 35, 51)
    pdf.cell(82, 6, _pulisci_testo_pdf(f"{tipo.upper()} N. {numero}"), ln=True)
    pdf.set_x(15)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(40, 40, 40)
    pdf.cell(82, 6, _pulisci_testo_pdf(f"Data Documento: {data_str}"), ln=True)

    # Box Destinatario
    pdf.rect(106, y_start, 92, 30, style="F")
    pdf.set_xy(109, y_start + 3)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(110, 110, 110)
    pdf.cell(86, 5, "SPETT.LE DESTINATARIO:", ln=True)
    pdf.set_x(109)
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(20, 20, 20)
    pdf.cell(86, 5, _pulisci_testo_pdf(soggetto_dict.get("Nome", "")), ln=True)

    dettagli_dest = []
    if soggetto_dict.get("Indirizzo"):
        dettagli_dest.append(soggetto_dict["Indirizzo"])
    if soggetto_dict.get("Citta"):
        dettagli_dest.append(soggetto_dict["Citta"])
    if soggetto_dict.get("Partita_IVA"):
        dettagli_dest.append(f"P.IVA: {soggetto_dict['Partita_IVA']}")
    elif soggetto_dict.get("Codice_Fiscale"):
        dettagli_dest.append(f"C.F.: {soggetto_dict['Codice_Fiscale']}")

    pdf.set_font("Helvetica", "", 8.5)
    for riga_d in dettagli_dest[:3]:
        pdf.set_x(109)
        pdf.cell(86, 4.5, _pulisci_testo_pdf(riga_d), ln=True)

    pdf.set_y(y_start + 36)

    # Tabella Articoli
    pdf.set_fill_color(30, 35, 42)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(25, 8, "CODICE", border=0, align="C", fill=True)
    pdf.cell(83, 8, "DESCRIZIONE ARTICOLO", border=0, align="L", fill=True)
    pdf.cell(16, 8, "Q.TA", border=0, align="C", fill=True)
    pdf.cell(24, 8, "PREZZO UN.", border=0, align="R", fill=True)
    pdf.cell(14, 8, "IVA%", border=0, align="C", fill=True)
    pdf.cell(24, 8, "IMPORTO", border=0, align="R", fill=True, ln=True)

    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(30, 30, 30)
    for idx, r in enumerate(righe):
        if idx % 2 == 1:
            pdf.set_fill_color(248, 249, 251)
            fill = True
        else:
            fill = False
        cod = _pulisci_testo_pdf(r.get("codice", "-"))[:14]
        desc = _pulisci_testo_pdf(r.get("descrizione", ""))[:46]
        qta = r.get("quantita", 1)
        p_netto = float(r.get("prezzo_unitario_netto", 0.0))
        iva_p = int(r.get("iva", 22))
        imp_riga = round(p_netto * qta, 2)

        pdf.cell(25, 7, cod, align="C", fill=fill)
        pdf.cell(83, 7, desc, align="L", fill=fill)
        pdf.cell(16, 7, str(qta), align="C", fill=fill)
        pdf.cell(24, 7, f"{p_netto:.2f}", align="R", fill=fill)
        pdf.cell(14, 7, f"{iva_p}%", align="C", fill=fill)
        pdf.cell(24, 7, f"{imp_riga:.2f}", align="R", fill=fill, ln=True)

    pdf.ln(6)
    pdf.set_draw_color(200, 200, 200)
    pdf.line(120, pdf.get_y(), 198, pdf.get_y())
    pdf.ln(3)

    # Riepilogo Totali
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(155, 6, "Totale Imponibile:", align="R")
    pdf.cell(31, 6, f"EUR {totale_imponibile:.2f}", align="R", ln=True)

    pdf.cell(155, 6, "Totale IVA:", align="R")
    pdf.cell(31, 6, f"EUR {totale_iva:.2f}", align="R", ln=True)

    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(200, 35, 51)
    pdf.cell(155, 8, "TOTALE DOCUMENTO:", align="R")
    pdf.cell(31, 8, f"EUR {totale_documento:.2f}", align="R", ln=True)

    if note:
        pdf.ln(4)
        pdf.set_font("Helvetica", "I", 8.5)
        pdf.set_text_color(90, 90, 90)
        pdf.multi_cell(0, 5, _pulisci_testo_pdf(f"Note: {note}"))

    nome_pulito = soggetto_dict.get("Nome", "Cliente").replace(" ", "_").replace(".", "").replace("'", "").replace("/", "_")
    nome_file = f"{tipo.upper()}_{numero}_{nome_pulito}.pdf"
    percorso_out = os.path.join(CARTELLA_PDF, nome_file)
    pdf.output(percorso_out)
    return percorso_out


# ==============================================================================
# TOOLS LANGCHAIN PER L'AGENTE AI
# ==============================================================================
@tool
def verifica_o_registra_soggetto(
    nome: str,
    ruolo: str = "Cliente",
    partita_iva: str = "",
    codice_fiscale: str = "",
    indirizzo: str = "",
    citta: str = "",
    email_pec: str = "",
    telefono: str = "",
    conferma_nuovo: bool = False
) -> str:
    """
    Controlla se un Cliente o Fornitore esiste già in anagrafica (la ricerca tollera refusi, spazi e maiuscole).
    Se non esiste, lo registra automaticamente con i dati disponibili.
    Se esistono soggetti SIMILI ma non identici, NON registra nulla e restituisce l'elenco: chiedi all'utente
    quale intende; solo se conferma che è un soggetto nuovo, richiama il tool con conferma_nuovo=True.
    """
    conn = get_conn()
    cur = conn.cursor()
    stato, d, cands = _risolvi_soggetto(cur, nome)

    if stato == "trovato":
        conn.close()
        return (
            f"Soggetto trovato in anagrafica: '{d['Nome']}' | Ruolo: {d['Ruolo']} | "
            f"P.IVA: {d.get('Partita_IVA') or 'N/D'} | C.F.: {d.get('Codice_Fiscale') or 'N/D'} | "
            f"Indirizzo: {d.get('Indirizzo') or ''} {d.get('Citta') or ''} | "
            f"PEC/Email: {d.get('Email_PEC') or 'N/D'} | Tel: {d.get('Telefono') or 'N/D'}"
        )

    if stato == "incerto" and not conferma_nuovo:
        conn.close()
        return _msg_candidati_soggetti(nome, cands)

    ruolo_norm = "Fornitore" if "forn" in ruolo.lower() else "Cliente"
    try:
        cur.execute("""
            INSERT INTO soggetti (Nome, Ruolo, Partita_IVA, Codice_Fiscale, Indirizzo, Citta, Email_PEC, Telefono)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (nome.strip(), ruolo_norm, partita_iva, codice_fiscale, indirizzo, citta, email_pec, telefono))
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        return f"Errore: esiste già un soggetto con nome '{nome.strip()}'."
    conn.close()
    return f"Nuovo {ruolo_norm} '{nome.strip()}' registrato in anagrafica."


@tool
def modifica_o_elimina_soggetto(
    nome_attuale: str,
    nuovo_nome: str = "",
    nuovo_ruolo: str = "",
    nuova_partita_iva: str = "",
    nuovo_codice_fiscale: str = "",
    nuovo_indirizzo: str = "",
    nuova_citta: str = "",
    nuova_email_pec: str = "",
    nuovo_telefono: str = "",
    aggiorna_anche_documenti: bool = True,
    elimina_soggetto: bool = False
) -> str:
    """
    Modifica i dati anagrafici di un Cliente o Fornitore esistente (Ragione Sociale, Ruolo, Partita IVA, Codice Fiscale, Indirizzo, Città, PEC, Telefono) oppure lo elimina.
    Se `nuovo_nome` viene specificato e `aggiorna_anche_documenti=True`, aggiorna anche il nome nei documenti già registrati.
    """
    conn = get_conn()
    cur = conn.cursor()

    stato, soggetto, cands = _risolvi_soggetto(cur, nome_attuale)
    if stato == "nessuno":
        conn.close()
        return f"Nessun cliente o fornitore trovato con nome '{nome_attuale}'."
    if stato == "incerto":
        conn.close()
        return _msg_candidati_soggetti(nome_attuale, cands)

    nome_vecchio = soggetto["Nome"]

    if elimina_soggetto:
        cur.execute("DELETE FROM soggetti WHERE id = ?", (soggetto["id"],))
        conn.commit()
        conn.close()
        return f"Soggetto '{nome_vecchio}' eliminato definitivamente dall'anagrafica."

    nome_fin = nuovo_nome.strip() if nuovo_nome.strip() else soggetto["Nome"]
    ruolo_fin = nuovo_ruolo.strip().capitalize() if nuovo_ruolo.strip() else soggetto["Ruolo"]
    if ruolo_fin not in ("Cliente", "Fornitore"):
        ruolo_fin = soggetto["Ruolo"]

    piva_fin = nuova_partita_iva.strip() if nuova_partita_iva.strip() else (soggetto.get("Partita_IVA") or "")
    cf_fin = nuovo_codice_fiscale.strip() if nuovo_codice_fiscale.strip() else (soggetto.get("Codice_Fiscale") or "")
    ind_fin = nuovo_indirizzo.strip() if nuovo_indirizzo.strip() else (soggetto.get("Indirizzo") or "")
    citta_fin = nuova_citta.strip() if nuova_citta.strip() else (soggetto.get("Citta") or "")
    pec_fin = nuova_email_pec.strip() if nuova_email_pec.strip() else (soggetto.get("Email_PEC") or "")
    tel_fin = nuovo_telefono.strip() if nuovo_telefono.strip() else (soggetto.get("Telefono") or "")

    try:
        cur.execute("""
            UPDATE soggetti
            SET Nome = ?, Ruolo = ?, Partita_IVA = ?, Codice_Fiscale = ?,
                Indirizzo = ?, Citta = ?, Email_PEC = ?, Telefono = ?
            WHERE id = ?
        """, (nome_fin, ruolo_fin, piva_fin, cf_fin, ind_fin, citta_fin, pec_fin, tel_fin, soggetto["id"]))

        doc_aggiornati = 0
        if nuovo_nome.strip() and nuovo_nome.strip() != nome_vecchio and aggiorna_anche_documenti:
            cur.execute("UPDATE documenti SET Cliente = ? WHERE LOWER(Cliente) = LOWER(?)", (nome_fin, nome_vecchio))
            doc_aggiornati = cur.rowcount

        conn.commit()
        conn.close()

        msg = f"Anagrafica aggiornata per '{nome_fin}' ({ruolo_fin})"
        if piva_fin:
            msg += f" | P.IVA: {piva_fin}"
        if ind_fin or citta_fin:
            msg += f" | Sede: {ind_fin} {citta_fin}".strip()
        if doc_aggiornati > 0:
            msg += f" | Aggiornato anche su {doc_aggiornati} documenti storici."
        return msg
    except Exception as e:
        conn.close()
        return f"Errore durante l'aggiornamento anagrafico: {e}"


@tool
def cerca_articolo_magazzino(termine_ricerca: str) -> str:
    """
    Cerca articoli nel magazzino per Codice, Descrizione oppure valori contenuti negli attributi extra (es. Reparto, Deposito, Ubicazione, Fornitore).
    La ricerca tollera refusi, spazi diversi e parole in ordine sparso.
    """
    conn = get_conn()
    cur = conn.cursor()

    parole = [p.lower() for p in termine_ricerca.strip().split() if p.strip()]
    if not parole:
        conn.close()
        return "Specifica un termine di ricerca."

    cond = "(LOWER(Codice) LIKE ? OR LOWER(Descrizione) LIKE ? OR LOWER(Extra_JSON) LIKE ?)"
    where = " AND ".join([cond] * len(parole))
    params = []
    for p in parole:
        params.extend([f"%{p}%"] * 3)

    cur.execute(f"""
        SELECT Codice, Descrizione, Giacenza, Prezzo_Acq_Medio, Prezzo_Listino, Extra_JSON
        FROM magazzino
        WHERE {where}
        LIMIT 25
    """, params)
    righe = [dict(r) for r in cur.fetchall()]
    intestazione = ""

    if not righe:
        # Fallback fuzzy: tollera refusi e spazi (es. "acro bat bianco 14 litri")
        tutti = _carica_articoli(cur)
        stato, art, cands = _cerca_fuzzy(termine_ricerca, tutti, ["Codice", "Descrizione"])
        if stato == "trovato":
            righe = [art]
        elif stato == "incerto":
            righe = [d for _, d in cands]
            intestazione = f"Nessuna corrispondenza esatta per '{termine_ricerca}'. Articoli simili:\n"
    conn.close()

    if not righe:
        return f"Nessun articolo trovato in magazzino per '{termine_ricerca}'."

    risultati = []
    for d in righe:
        extra_str = ""
        try:
            ex = json.loads(d.get("Extra_JSON") or "{}")
            if ex:
                extra_str = " | Extra: " + ", ".join([f"{k}={v}" for k, v in ex.items()])
        except Exception:
            pass
        risultati.append(
            f"- [{d['Codice']}] {d['Descrizione']} | Giacenza: {d['Giacenza']} | "
            f"P.Acq.Medio: {float(d['Prezzo_Acq_Medio'] or 0):.2f} EUR | "
            f"Listino(IVATO): {float(d['Prezzo_Listino'] or 0):.2f} EUR{extra_str}"
        )
    return intestazione + "\n".join(risultati)


@tool
def modifica_articolo_magazzino(
    codice_o_descrizione: str,
    nuovo_codice: str = "",
    nuova_descrizione: str = "",
    nuovo_prezzo_listino: float = -1.0,
    nuovo_prezzo_acquisto: float = -1.0
) -> str:
    """
    Modifica Codice, Descrizione, Prezzo_Listino (IVA inclusa) o Prezzo_Acq_Medio di un articolo esistente.
    """
    conn = get_conn()
    cur = conn.cursor()
    stato, d, cands = _risolvi_articolo(cur, codice_o_descrizione)
    if stato == "nessuno":
        conn.close()
        return f"Articolo '{codice_o_descrizione}' non trovato in magazzino."
    if stato == "incerto":
        conn.close()
        return _msg_candidati_articoli(codice_o_descrizione, cands)

    cod_fin = nuovo_codice.strip() if nuovo_codice.strip() else d["Codice"]
    desc_fin = nuova_descrizione.strip() if nuova_descrizione.strip() else d["Descrizione"]
    acq_fin = round(nuovo_prezzo_acquisto, 2) if nuovo_prezzo_acquisto >= 0 else d["Prezzo_Acq_Medio"]
    list_fin = round(nuovo_prezzo_listino, 2) if nuovo_prezzo_listino >= 0 else d["Prezzo_Listino"]

    try:
        cur.execute("""
            UPDATE magazzino
            SET Codice = ?, Descrizione = ?, Prezzo_Acq_Medio = ?, Prezzo_Listino = ?
            WHERE id = ?
        """, (cod_fin, desc_fin, acq_fin, list_fin, d["id"]))
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        return f"Errore: esiste già un articolo con codice '{cod_fin}'."
    conn.close()
    return f"Articolo aggiornato: [{cod_fin}] {desc_fin} | P.Acq: {float(acq_fin or 0):.2f} EUR | Listino: {float(list_fin or 0):.2f} EUR"


@tool
def aggiorna_attributi_extra_articolo(codice_o_descrizione: str, attributi_json: str) -> str:
    """
    Modifica o aggiunge attributi extra dinamici (es. '{"Magazzino": "Deposito Nord", "Corsia": "C3", "Reparto": "Idraulica"}') a un articolo.
    """
    try:
        nuovi_attr = json.loads(attributi_json)
        if not isinstance(nuovi_attr, dict):
            return "Errore: attributi_json deve essere un oggetto JSON chiave-valore."
    except Exception as e:
        return f"Errore parsing JSON attributi: {e}"

    conn = get_conn()
    cur = conn.cursor()
    stato, d, cands = _risolvi_articolo(cur, codice_o_descrizione)
    if stato == "nessuno":
        conn.close()
        return f"Articolo '{codice_o_descrizione}' non trovato in magazzino."
    if stato == "incerto":
        conn.close()
        return _msg_candidati_articoli(codice_o_descrizione, cands)

    try:
        extra_attuale = json.loads(d.get("Extra_JSON") or "{}")
    except Exception:
        extra_attuale = {}

    for k, v in nuovi_attr.items():
        if v is None or str(v).strip() == "":
            extra_attuale.pop(k, None)
        else:
            extra_attuale[k] = str(v).strip()

    cur.execute("UPDATE magazzino SET Extra_JSON = ? WHERE id = ?", (json.dumps(extra_attuale, ensure_ascii=False), d["id"]))
    conn.commit()
    conn.close()
    return f"Attributi aggiornati per [{d['Codice']}] {d['Descrizione']}: {extra_attuale}"


@tool
def aggiorna_giacenza(
    codice_o_descrizione: str,
    variazione_quantita: int,
    descrizione_nuovo_articolo: str = "",
    prezzo_acquisto: float = 0.0,
    prezzo_listino: float = 0.0,
    attributi_extra_json: str = "{}",
    conferma_nuovo_articolo: bool = False
) -> str:
    """
    Aggiunge (+N) o sottrae (-N) quantità in magazzino. La ricerca dell'articolo tollera refusi, spazi e maiuscole
    (es. 'acro bat bianco 14 litri' trova 'acrobat bianco 14litri').
    Se l'articolo NON esiste e non ci sono articoli simili, con variazione_quantita > 0 lo crea automaticamente.
    Se esistono articoli SIMILI ma non identici, NON modifica e NON crea nulla: restituisce l'elenco dei candidati.
    In quel caso chiedi all'utente quale intende; solo se conferma esplicitamente che è un articolo NUOVO,
    richiama il tool con conferma_nuovo_articolo=True.
    """
    conn = get_conn()
    cur = conn.cursor()

    stato, d, cands = _risolvi_articolo(cur, codice_o_descrizione)

    if stato == "incerto" and not conferma_nuovo_articolo:
        conn.close()
        return _msg_candidati_articoli(codice_o_descrizione, cands)

    try:
        extra_in = json.loads(attributi_extra_json) if attributi_extra_json else {}
    except Exception:
        extra_in = {}

    if stato == "trovato":
        giac_vecchia = int(d["Giacenza"] or 0)
        giac_nuova = giac_vecchia + int(variazione_quantita)

        acq_vecchio = float(d["Prezzo_Acq_Medio"] or 0.0)
        list_vecchio = float(d["Prezzo_Listino"] or 0.0)

        if variazione_quantita > 0 and prezzo_acquisto > 0:
            if giac_vecchia > 0 and acq_vecchio > 0:
                acq_nuovo = round(((giac_vecchia * acq_vecchio) + (variazione_quantita * prezzo_acquisto)) / (giac_vecchia + variazione_quantita), 2)
            else:
                acq_nuovo = round(prezzo_acquisto, 2)
            list_nuovo = round(prezzo_listino, 2) if prezzo_listino > 0 else round(prezzo_acquisto * 1.22 * 1.30, 2)
        else:
            acq_nuovo = acq_vecchio
            list_nuovo = round(prezzo_listino, 2) if prezzo_listino > 0 else list_vecchio

        try:
            ex_old = json.loads(d.get("Extra_JSON") or "{}")
        except Exception:
            ex_old = {}
        ex_old.update(extra_in)

        cur.execute("""
            UPDATE magazzino
            SET Giacenza = ?, Prezzo_Acq_Medio = ?, Prezzo_Listino = ?, Extra_JSON = ?
            WHERE id = ?
        """, (giac_nuova, acq_nuovo, list_nuovo, json.dumps(ex_old, ensure_ascii=False), d["id"]))
        conn.commit()
        conn.close()
        return f"Giacenza aggiornata per [{d['Codice']}] {d['Descrizione']}: da {giac_vecchia} a {giac_nuova} pz (P.Acq.Medio: {acq_nuovo:.2f} EUR, Listino: {list_nuovo:.2f} EUR)."

    # --- Articolo non esistente (o creazione esplicitamente confermata) ---
    if variazione_quantita < 0 and not conferma_nuovo_articolo:
        conn.close()
        return (
            f"Articolo '{codice_o_descrizione}' non trovato in magazzino: non posso scaricare {abs(variazione_quantita)} pz "
            f"da un articolo inesistente. Chiedi all'utente di verificare il nome."
        )

    chiave = codice_o_descrizione.strip()
    if len(chiave) <= 15 and " " not in chiave:
        cod_nuovo = chiave.upper()
        cur.execute("SELECT 1 FROM magazzino WHERE LOWER(Codice) = LOWER(?)", (cod_nuovo,))
        if cur.fetchone():
            cod_nuovo = _genera_codice_libero(cur)
    else:
        cod_nuovo = _genera_codice_libero(cur)

    desc_nuova = descrizione_nuovo_articolo.strip() if descrizione_nuovo_articolo.strip() else chiave
    acq_nuovo = round(prezzo_acquisto, 2) if prezzo_acquisto > 0 else 0.0
    list_nuovo = round(prezzo_listino, 2) if prezzo_listino > 0 else (round(acq_nuovo * 1.22 * 1.30, 2) if acq_nuovo > 0 else 0.0)

    cur.execute("""
        INSERT INTO magazzino (Codice, Descrizione, Giacenza, Prezzo_Acq_Medio, Prezzo_Listino, Extra_JSON)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (cod_nuovo, desc_nuova, int(variazione_quantita), acq_nuovo, list_nuovo, json.dumps(extra_in, ensure_ascii=False)))
    conn.commit()
    conn.close()
    return f"Nuovo articolo inserito in magazzino: [{cod_nuovo}] {desc_nuova} | Giacenza: {variazione_quantita} pz | Listino: {list_nuovo:.2f} EUR."


@tool
def emetti_documento(
    tipo_documento: str,
    cliente: str,
    righe_json: str,
    prezzi_ivato_inclusa: bool = True,
    scarica_magazzino: bool = True,
    note: str = "",
    data_documento: str = ""
) -> str:
    """
    Emette un documento di vendita verso un Cliente ('DDT', 'FATTURA' o 'NOTA_CREDITO'), genera il PDF formale in Documenti_PDF/ e aggiorna le giacenze di magazzino.
    `righe_json` deve essere una lista JSON di oggetti con chiavi: 'descrizione' (o 'codice'), 'quantita', 'prezzo_unitario' (opzionale: se 0 usa il listino di magazzino), 'iva' (default 22).
    Se un articolo o il cliente sono ambigui/simili a più voci esistenti, il tool NON emette nulla e restituisce i candidati: chiedi all'utente quale intende.
    """
    tipo = tipo_documento.strip().upper().replace(" ", "_")
    if tipo not in ("DDT", "FATTURA", "NOTA_CREDITO"):
        return "Errore: tipo_documento deve essere 'DDT', 'FATTURA' o 'NOTA_CREDITO'."

    try:
        righe_input = json.loads(righe_json)
    except Exception as e:
        return f"Errore formato righe_json: {e}"

    conn = get_conn()
    cur = conn.cursor()

    # --- FASE 1: verifica di soggetto e articoli PRIMA di scrivere qualsiasi cosa ---
    stato_s, sog_db, cands_s = _risolvi_soggetto(cur, cliente)
    if stato_s == "incerto":
        conn.close()
        return _msg_candidati_soggetti(cliente, cands_s)

    articoli_cache = _carica_articoli(cur)
    risolti = []
    problemi = []
    for r in righe_input:
        chiave_art = str(r.get("codice") or r.get("descrizione") or "").strip()
        stato_a, art_db, cands_a = _risolvi_articolo(cur, chiave_art, cache=articoli_cache)
        if stato_a == "incerto":
            problemi.append(_msg_candidati_articoli(chiave_art, cands_a))
        risolti.append((r, chiave_art, art_db))
    if problemi:
        conn.close()
        return "\n\n".join(problemi)

    # --- FASE 2: registrazione ---
    if stato_s == "trovato":
        soggetto_dict = sog_db
    else:
        cur.execute("INSERT INTO soggetti (Nome, Ruolo) VALUES (?, 'Cliente')", (cliente.strip(),))
        soggetto_dict = {"Nome": cliente.strip(), "Ruolo": "Cliente"}

    # Progressivo documento per tipo e ruolo Cliente
    cur.execute("SELECT COALESCE(MAX(Numero), 0) FROM documenti WHERE UPPER(Tipo) = ? AND Ruolo = 'Cliente'", (tipo,))
    numero_doc = cur.fetchone()[0] + 1
    data_str = data_documento.strip() if data_documento.strip() else datetime.date.today().strftime("%d/%m/%Y")

    righe_elaborate = []
    tot_imponibile = 0.0
    tot_iva = 0.0
    voci_riassunto = []

    segno_mag = 1 if tipo == "NOTA_CREDITO" else -1

    for r, chiave_art, art_d in risolti:
        qta = int(r.get("quantita", 1))
        iva_perc = int(r.get("iva", 22))
        prezzo_in = float(r.get("prezzo_unitario", 0.0))

        if art_d:
            codice_art = art_d["Codice"]
            desc_art = r.get("descrizione") or art_d["Descrizione"]
            if prezzo_in <= 0:
                prezzo_in = float(art_d["Prezzo_Listino"] or 0.0)
                ivato = True
            else:
                ivato = prezzi_ivato_inclusa

            if scarica_magazzino:
                cur.execute("UPDATE magazzino SET Giacenza = COALESCE(Giacenza, 0) + ? WHERE id = ?", (segno_mag * qta, art_d["id"]))
        else:
            codice_art = "-"
            desc_art = chiave_art
            ivato = prezzi_ivato_inclusa

        if ivato:
            prezzo_netto = round(prezzo_in / (1 + iva_perc / 100.0), 4)
        else:
            prezzo_netto = round(prezzo_in, 4)

        imp_riga = round(prezzo_netto * qta, 2)
        iva_riga = round(imp_riga * (iva_perc / 100.0), 2)

        tot_imponibile = round(tot_imponibile + imp_riga, 2)
        tot_iva = round(tot_iva + iva_riga, 2)

        righe_elaborate.append({
            "codice": codice_art,
            "descrizione": desc_art,
            "quantita": qta,
            "prezzo_unitario_netto": round(prezzo_netto, 2),
            "iva": iva_perc
        })
        voci_riassunto.append(f"{qta}x {desc_art}")

    tot_documento = round(tot_imponibile + tot_iva, 2)
    fatturato_flag = "SI" if tipo in ("FATTURA", "NOTA_CREDITO") else "NO"

    percorso_pdf = _genera_pdf_documento(
        tipo=tipo,
        numero=numero_doc,
        data_str=data_str,
        soggetto_dict=soggetto_dict,
        righe=righe_elaborate,
        totale_imponibile=tot_imponibile,
        totale_iva=tot_iva,
        totale_documento=tot_documento,
        note=note
    )

    cur.execute("""
        INSERT INTO documenti (Data, Tipo, Ruolo, Numero, Cliente, Voci, Importo_Totale, Fatturato, Dettaglio_JSON)
        VALUES (?, ?, 'Cliente', ?, ?, ?, ?, ?, ?)
    """, (
        data_str,
        tipo,
        numero_doc,
        soggetto_dict["Nome"],
        ", ".join(voci_riassunto),
        tot_documento,
        fatturato_flag,
        json.dumps(righe_elaborate, ensure_ascii=False)
    ))
    conn.commit()
    conn.close()

    return (
        f"Emesso {tipo} N. {numero_doc} per '{soggetto_dict['Nome']}' in data {data_str}. "
        f"Totale documento: {tot_documento:.2f} EUR (Imponibile {tot_imponibile:.2f} + IVA {tot_iva:.2f}). "
        f"PDF salvato in: {os.path.basename(percorso_pdf)}."
    )


@tool
def registra_documento_fornitore(
    fornitore: str,
    tipo_documento: str,
    numero_documento: str,
    data_documento: str,
    righe_json: str,
    totale_documento: float,
    percorso_file_originale: str = "",
    aggiorna_anche_magazzino: bool = True,
    partita_iva_fornitore: str = "",
    indirizzo_fornitore: str = ""
) -> str:
    """
    Registra un documento ricevuto da un Fornitore ('FATTURA', 'DDT' o 'NOTA_CREDITO'), archivia una copia del file in Documenti_PDF/ e carica le quantità acquistate in magazzino.
    Se il fornitore o un articolo risultano ambigui/simili a più voci esistenti, il tool NON registra nulla e restituisce i candidati: chiedi all'utente quale intende.
    """
    tipo = tipo_documento.strip().upper().replace(" ", "_")
    if "FATT" in tipo:
        tipo = "FATTURA"
    elif "NOTA" in tipo:
        tipo = "NOTA_CREDITO"
    else:
        tipo = "DDT"

    try:
        righe_input = json.loads(righe_json)
    except Exception as e:
        return f"Errore formato righe_json: {e}"

    conn = get_conn()
    cur = conn.cursor()

    # --- FASE 1: verifiche preliminari (nessuna scrittura) ---
    stato_f, row_forn, cands_f = _risolvi_soggetto(cur, fornitore)
    if stato_f == "incerto":
        conn.close()
        return _msg_candidati_soggetti(fornitore, cands_f)

    articoli_cache = _carica_articoli(cur)
    risolti = []
    problemi = []
    for r in righe_input:
        cod = str(r.get("codice") or "").strip()
        desc = str(r.get("descrizione") or "").strip()
        art = None
        stato_a = "nessuno"
        cands_a = []
        if aggiorna_anche_magazzino and (cod or desc):
            if cod:
                stato_a, art, cands_a = _cerca_fuzzy(cod, articoli_cache, ["Codice"])
                if stato_a == "incerto":
                    # il codice fornitore non corrisponde con certezza: prova con la descrizione
                    stato_a, art, cands_a = ("nessuno", None, [])
            if not art and desc:
                stato_a, art, cands_a = _risolvi_articolo(cur, desc, cache=articoli_cache)
            if stato_a == "incerto":
                problemi.append(_msg_candidati_articoli(desc or cod, cands_a))
        risolti.append((r, cod, desc, art))
    if problemi:
        conn.close()
        return "\n\n".join(problemi)

    # --- FASE 2: registrazione ---
    if stato_f == "trovato":
        row_forn = row_forn
        nome_forn_ufficiale = row_forn["Nome"]
        if partita_iva_fornitore and not row_forn.get("Partita_IVA"):
            cur.execute("UPDATE soggetti SET Partita_IVA = ? WHERE id = ?", (partita_iva_fornitore, row_forn["id"]))
    else:
        nome_forn_ufficiale = fornitore.strip()
        cur.execute("""
            INSERT INTO soggetti (Nome, Ruolo, Partita_IVA, Indirizzo)
            VALUES (?, 'Fornitore', ?, ?)
        """, (nome_forn_ufficiale, partita_iva_fornitore, indirizzo_fornitore))

    segno_mag = -1 if tipo == "NOTA_CREDITO" else 1
    voci_riassunto = []
    righe_salvate = []

    for r, cod, desc, ad in risolti:
        qta = int(r.get("quantita", 1))
        p_netto = float(r.get("prezzo_unitario_netto", r.get("prezzo_unitario", 0.0)))
        sconto = float(r.get("sconto", 0.0))
        if sconto > 0:
            p_netto = round(p_netto * (1 - sconto / 100.0), 4)
        iva_p = int(r.get("iva", 22))

        voci_riassunto.append(f"{qta}x {desc or cod}")
        righe_salvate.append({
            "codice": ad["Codice"] if ad else cod,
            "descrizione": desc or (ad["Descrizione"] if ad else cod),
            "quantita": qta,
            "prezzo_unitario_netto": round(p_netto, 2),
            "iva": iva_p
        })

        if aggiorna_anche_magazzino and (cod or desc):
            if ad:
                cur.execute("SELECT Giacenza, Prezzo_Acq_Medio, Prezzo_Listino FROM magazzino WHERE id = ?", (ad["id"],))
                cur_row = cur.fetchone()
                g_old = int(cur_row["Giacenza"] or 0)
                g_new = g_old + (segno_mag * qta)
                acq_old = float(cur_row["Prezzo_Acq_Medio"] or 0.0)
                list_old = float(cur_row["Prezzo_Listino"] or 0.0)
                if segno_mag > 0 and p_netto > 0:
                    if g_old > 0 and acq_old > 0:
                        acq_new = round(((g_old * acq_old) + (qta * p_netto)) / (g_old + qta), 2)
                    else:
                        acq_new = round(p_netto, 2)
                    list_new = round(p_netto * (1 + iva_p / 100.0) * 1.30, 2) if list_old == 0 else list_old
                else:
                    acq_new = acq_old
                    list_new = list_old

                cur.execute("""
                    UPDATE magazzino SET Giacenza = ?, Prezzo_Acq_Medio = ?, Prezzo_Listino = ? WHERE id = ?
                """, (g_new, acq_new, list_new, ad["id"]))
            elif segno_mag > 0:
                cod_gen = cod if cod else _genera_codice_libero(cur)
                cur.execute("SELECT 1 FROM magazzino WHERE LOWER(Codice) = LOWER(?)", (cod_gen,))
                if cur.fetchone():
                    cod_gen = _genera_codice_libero(cur)
                list_calc = round(p_netto * (1 + iva_p / 100.0) * 1.30, 2) if p_netto > 0 else 0.0
                cur.execute("""
                    INSERT INTO magazzino (Codice, Descrizione, Giacenza, Prezzo_Acq_Medio, Prezzo_Listino, Extra_JSON)
                    VALUES (?, ?, ?, ?, ?, '{}')
                """, (cod_gen, desc or cod_gen, qta, round(p_netto, 2), list_calc))

    # Archiviazione file originale in Documenti_PDF/
    num_pulito = re.sub(r"[^0-9]", "", str(numero_documento)) or "1"
    num_int = int(num_pulito)
    if percorso_file_originale and os.path.exists(percorso_file_originale):
        est = os.path.splitext(percorso_file_originale)[1]
        nome_forn_clean = nome_forn_ufficiale.replace(" ", "_").replace(".", "").replace("'", "").replace("/", "_")
        dest_allegato = os.path.join(CARTELLA_PDF, f"{tipo}_{num_int}_{nome_forn_clean}{est}")
        try:
            shutil.copy2(percorso_file_originale, dest_allegato)
        except Exception as e:
            print(f"Errore copia allegato fornitore: {e}")

    cur.execute("""
        INSERT INTO documenti (Data, Tipo, Ruolo, Numero, Cliente, Voci, Importo_Totale, Fatturato, Dettaglio_JSON)
        VALUES (?, ?, 'Fornitore', ?, ?, ?, ?, 'SI', ?)
    """, (
        data_documento.strip(),
        tipo,
        num_int,
        nome_forn_ufficiale,
        ", ".join(voci_riassunto),
        round(float(totale_documento), 2),
        json.dumps(righe_salvate, ensure_ascii=False)
    ))
    conn.commit()
    conn.close()
    return f"Registrato {tipo} N. {num_int} del fornitore '{nome_forn_ufficiale}' (Totale: {totale_documento:.2f} EUR) e aggiornato il magazzino."


@tool
def emetti_fattura_differita(cliente: str) -> str:
    """
    Raggruppa tutti i DDT non ancora fatturati ('Fatturato = NO') di un Cliente in un'unica FATTURA riepilogativa.
    """
    conn = get_conn()
    cur = conn.cursor()

    stato, sog, cands = _risolvi_soggetto(cur, cliente)
    if stato == "incerto":
        conn.close()
        return _msg_candidati_soggetti(cliente, cands)
    if stato == "nessuno":
        conn.close()
        return f"Nessun cliente trovato in anagrafica con nome '{cliente}'."

    nome_cliente = sog["Nome"]
    cur.execute("""
        SELECT * FROM documenti
        WHERE UPPER(Tipo) = 'DDT' AND Ruolo = 'Cliente' AND UPPER(Fatturato) = 'NO'
          AND LOWER(Cliente) = LOWER(?)
    """, (nome_cliente,))
    ddt_list = [dict(r) for r in cur.fetchall()]

    if not ddt_list:
        conn.close()
        return f"Nessun DDT in sospeso da fatturare trovato per '{nome_cliente}'."

    soggetto_dict = sog

    righe_unite = []
    numeri_ddt = []
    tot_imp = 0.0
    tot_iva = 0.0

    for ddt in ddt_list:
        numeri_ddt.append(str(ddt["Numero"]))
        try:
            det = json.loads(ddt.get("Dettaglio_JSON") or "[]")
        except Exception:
            det = []
        for r in det:
            righe_unite.append(r)
            imp_r = round(float(r.get("prezzo_unitario_netto", 0)) * int(r.get("quantita", 1)), 2)
            iva_r = round(imp_r * (int(r.get("iva", 22)) / 100.0), 2)
            tot_imp = round(tot_imp + imp_r, 2)
            tot_iva = round(tot_iva + iva_r, 2)

    tot_doc = round(tot_imp + tot_iva, 2)
    cur.execute("SELECT COALESCE(MAX(Numero), 0) FROM documenti WHERE UPPER(Tipo) = 'FATTURA' AND Ruolo = 'Cliente'")
    num_fatt = cur.fetchone()[0] + 1
    oggi_str = datetime.date.today().strftime("%d/%m/%Y")

    note_rif = f"Fattura differita riepilogativa dei DDT N. {', '.join(numeri_ddt)}"
    percorso_pdf = _genera_pdf_documento(
        tipo="FATTURA",
        numero=num_fatt,
        data_str=oggi_str,
        soggetto_dict=soggetto_dict,
        righe=righe_unite,
        totale_imponibile=tot_imp,
        totale_iva=tot_iva,
        totale_documento=tot_doc,
        note=note_rif
    )

    ids_ddt = [d["id"] for d in ddt_list]
    cur.executemany("UPDATE documenti SET Fatturato = 'SI' WHERE id = ?", [(i,) for i in ids_ddt])

    voci_str = ", ".join([f"{r['quantita']}x {r['descrizione']}" for r in righe_unite])
    cur.execute("""
        INSERT INTO documenti (Data, Tipo, Ruolo, Numero, Cliente, Voci, Importo_Totale, Fatturato, Dettaglio_JSON)
        VALUES (?, 'FATTURA', 'Cliente', ?, ?, ?, ?, 'SI', ?)
    """, (oggi_str, num_fatt, nome_cliente, voci_str, tot_doc, json.dumps(righe_unite, ensure_ascii=False)))

    conn.commit()
    conn.close()
    return f"Emessa Fattura Differita N. {num_fatt} per '{nome_cliente}' (DDT inclusi: {', '.join(numeri_ddt)}). Totale: {tot_doc:.2f} EUR. File: {os.path.basename(percorso_pdf)}."


@tool
def modifica_documento(
    tipo_documento: str,
    numero: int,
    ruolo: str = "",
    nuova_data: str = "",
    nuovo_soggetto: str = "",
    nuovo_importo: float = -1.0
) -> str:
    """
    Modifica Data, Cliente/Fornitore o Importo Totale di un documento già registrato.
    """
    conn = get_conn()
    cur = conn.cursor()
    if ruolo:
        cur.execute("SELECT * FROM documenti WHERE UPPER(Tipo) = UPPER(?) AND Numero = ? AND LOWER(Ruolo) = LOWER(?)", (tipo_documento.strip(), numero, ruolo.strip()))
    else:
        cur.execute("SELECT * FROM documenti WHERE UPPER(Tipo) = UPPER(?) AND Numero = ?", (tipo_documento.strip(), numero))
    trovati = [dict(r) for r in cur.fetchall()]

    if len(trovati) == 0:
        conn.close()
        return f"Nessun documento '{tipo_documento}' N. {numero} trovato."
    if len(trovati) > 1:
        conn.close()
        return f"Esistono PIÙ documenti '{tipo_documento}' N. {numero} (sia Cliente che Fornitore). Specifica quale intende l'utente."

    doc = trovati[0]
    data_f = nuova_data.strip() if nuova_data.strip() else doc["Data"]
    sog_f = nuovo_soggetto.strip() if nuovo_soggetto.strip() else doc["Cliente"]
    imp_f = round(nuovo_importo, 2) if nuovo_importo >= 0 else doc["Importo_Totale"]

    cur.execute("UPDATE documenti SET Data = ?, Cliente = ?, Importo_Totale = ? WHERE id = ?", (data_f, sog_f, imp_f, doc["id"]))
    conn.commit()
    conn.close()
    return f"Documento {doc['Tipo']} N. {numero} ({doc['Ruolo']}) aggiornato: Data={data_f}, Soggetto='{sog_f}', Totale={float(imp_f or 0):.2f} EUR."


@tool
def elimina_documenti(
    tipo_documento: str,
    numero: int,
    ruolo: str = "",
    ripristina_magazzino: bool = True
) -> str:
    """
    Elimina un documento dal registro e ripristina automaticamente le giacenze di magazzino coinvolte.
    """
    conn = get_conn()
    cur = conn.cursor()
    if ruolo:
        cur.execute("SELECT * FROM documenti WHERE UPPER(Tipo) = UPPER(?) AND Numero = ? AND LOWER(Ruolo) = LOWER(?)", (tipo_documento.strip(), numero, ruolo.strip()))
    else:
        cur.execute("SELECT * FROM documenti WHERE UPPER(Tipo) = UPPER(?) AND Numero = ?", (tipo_documento.strip(), numero))
    trovati = [dict(r) for r in cur.fetchall()]

    if len(trovati) == 0:
        conn.close()
        return f"Nessun documento '{tipo_documento}' N. {numero} trovato."
    if len(trovati) > 1:
        conn.close()
        return f"Esistono PIÙ documenti '{tipo_documento}' N. {numero} (Cliente e Fornitore). Chiedi all'utente quale vuole eliminare."

    doc = trovati[0]
    ripristinati = []
    non_ripristinati = []

    if ripristina_magazzino:
        try:
            righe_det = json.loads(doc.get("Dettaglio_JSON") or "[]")
        except Exception:
            righe_det = []

        # Se era una vendita Cliente (DDT/FATTURA) il magazzino era stato scaricato (-qta), quindi ripristiniamo +qta.
        # Se era un acquisto Fornitore (DDT/FATTURA) il magazzino era stato caricato (+qta), quindi ripristiniamo -qta.
        if doc["Ruolo"] == "Cliente":
            moltiplicatore = -1 if doc["Tipo"].upper() == "NOTA_CREDITO" else 1
        else:
            moltiplicatore = 1 if doc["Tipo"].upper() == "NOTA_CREDITO" else -1

        articoli_cache = _carica_articoli(cur)
        for r in righe_det:
            cod = str(r.get("codice") or "").strip()
            desc = str(r.get("descrizione") or "").strip()
            qta = int(r.get("quantita", 0))
            if qta == 0:
                continue

            art = None
            if cod and cod != "-":
                stato, art, _ = _cerca_fuzzy(cod, articoli_cache, ["Codice"])
                if stato != "trovato":
                    art = None
            if not art and desc:
                stato, art, _ = _cerca_fuzzy(desc, articoli_cache, ["Codice", "Descrizione"])
                if stato != "trovato":
                    art = None

            if art:
                cur.execute("UPDATE magazzino SET Giacenza = COALESCE(Giacenza, 0) + ? WHERE id = ?", (moltiplicatore * qta, art["id"]))
                cur.execute("SELECT Giacenza FROM magazzino WHERE id = ?", (art["id"],))
                nuova_g = cur.fetchone()["Giacenza"]
                ripristinati.append(f"{art['Descrizione']} (nuova giacenza: {nuova_g})")
            else:
                non_ripristinati.append(desc or cod)

    # Rimuove anche eventuale PDF/allegato associato
    nome_clean = doc["Cliente"].replace(" ", "_").replace(".", "").replace("'", "").replace("/", "_")
    prefisso = f"{doc['Tipo'].upper()}_{doc['Numero']}_{nome_clean}"
    if os.path.exists(CARTELLA_PDF):
        for fn in os.listdir(CARTELLA_PDF):
            if fn.startswith(prefisso + "."):
                try:
                    os.remove(os.path.join(CARTELLA_PDF, fn))
                except Exception:
                    pass

    cur.execute("DELETE FROM documenti WHERE id = ?", (doc["id"],))
    conn.commit()
    conn.close()

    extra_msg = f" Giacenze ripristinate: {', '.join(ripristinati)}." if ripristinati else ""
    if non_ripristinati:
        extra_msg += f" ATTENZIONE: articoli non trovati in magazzino, giacenza NON ripristinata: {', '.join(non_ripristinati)}."
    return f"Eliminato {doc['Tipo']} N. {doc['Numero']} ({doc['Ruolo']} - {doc['Cliente']}).{extra_msg}"


@tool
def apri_documento_pdf(tipo_documento: str, numero: int) -> str:
    """
    Apre con il visualizzatore predefinito di sistema il file PDF/allegato di un documento salvato in Documenti_PDF/.
    """
    prefisso = f"{tipo_documento.strip().upper()}_{numero}_"
    if os.path.exists(CARTELLA_PDF):
        for fname in os.listdir(CARTELLA_PDF):
            if fname.startswith(prefisso):
                percorso = os.path.join(CARTELLA_PDF, fname)
                try:
                    os.startfile(percorso)
                    return f"Ho aperto il documento '{fname}'."
                except Exception as e:
                    return f"Impossibile aprire il file '{fname}': {e}"
    return f"Nessun file PDF trovato per {tipo_documento} N. {numero}."


TUTTI_I_TOOLS = [
    verifica_o_registra_soggetto,
    modifica_o_elimina_soggetto,
    cerca_articolo_magazzino,
    modifica_articolo_magazzino,
    aggiorna_attributi_extra_articolo,
    aggiorna_giacenza,
    emetti_documento,
    registra_documento_fornitore,
    emetti_fattura_differita,
    modifica_documento,
    elimina_documenti,
    apri_documento_pdf
]
