import os
import re
import sys
import json
import base64
import threading
import datetime
import pandas as pd
from PIL import Image
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk

from langchain_anthropic import ChatAnthropic
from langchain_core.messages import SystemMessage, HumanMessage, AIMessage
from langgraph.prebuilt import create_react_agent

from database_tools import (
    CARTELLA_BASE,
    FILE_DB,
    CARTELLA_PDF,
    CARTELLA_BACKUP,
    FILE_LOGO,
    TUTTI_I_TOOLS,
    get_conn,
    importa_magazzino_da_excel,
    importa_anagrafica_da_excel,
    esegui_backup_automatico,
    esporta_backup_manuale,
    ripristina_backup_db
)

# --- IMPOSTAZIONI GRAFICHE E PALETTE COLORI ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

COL_BG_MAIN = "#121417"
COL_CARD = "#1B1F24"
COL_CARD_LIGHT = "#242930"
COL_ACCENT = "#C82333"
COL_ACCENT_HOVER = "#A71D2A"
COL_BLUE = "#2563EB"
COL_BLUE_HOVER = "#1D4ED8"
COL_GREEN = "#15803D"
COL_GREEN_HOVER = "#166534"
COL_PURPLE = "#7C3AED"
COL_PURPLE_HOVER = "#6D28D9"
COL_TEXT_MUTED = "#9CA3AF"
COL_BORDER = "#2D333B"

FILE_CHIAVE = os.path.join(CARTELLA_BASE, "chiave_api.txt")

try:
    with open(FILE_CHIAVE, "r") as f:
        API_KEY = f.read().strip()
except FileNotFoundError:
    API_KEY = "MANCANTE"

SYSTEM_PROMPT = """Sei un assistente gestionale aziendale. Hai accesso diretto agli strumenti del database per:
1. verifica_o_registra_soggetto: controlla se un'azienda esiste già in anagrafica/storico e se è 'Cliente' o 'Fornitore'.
2. modifica_o_elimina_soggetto: modifica Ragione Sociale, Ruolo (Cliente/Fornitore), Partita IVA, Codice Fiscale, Indirizzo, Città, Email/PEC o Telefono di un soggetto in anagrafica (aggiornando anche i documenti storici se la ragione sociale cambia), oppure lo elimina.
3. cerca_articolo_magazzino: legge codice, descrizione, giacenza, prezzi e TUTTI gli attributi extra dinamici (es. Reparto, Magazzino/Deposito, Ubicazione/Corsia, Peso, Fornitore, ecc.).
4. modifica_articolo_magazzino: modifica codice, descrizione e prezzo di listino (IVA INCLUSA) in magazzino.
5. aggiorna_attributi_extra_articolo: modifica o aggiunge campi extra dinamici di un articolo (es. spostamento magazzino/deposito, cambio corsia/scaffale, reparto, peso, note).
6. aggiorna_giacenza: modifica la quantità in magazzino e crea/aggiorna l'articolo con descrizione, prezzo_acquisto, prezzo_listino ed eventuali attributi_extra.
7. emetti_documento: crea PDF di DDT, FATTURA o NOTA_CREDITO EMESSI DALLA NOSTRA AZIENDA verso Clienti.
8. registra_documento_fornitore: registra un documento RICEVUTO DA UN FORNITORE (Fattura d'acquisto, DDT fornitore, Nota di credito fornitore), archiviando il file allegato e caricando il magazzino.
9. emetti_fattura_differita: raggruppa i DDT in sospeso di un cliente in un'unica fattura.
10. modifica_documento / elimina_documenti: corregge o cancella documenti già esistenti nel registro.
11. apri_documento_pdf: apre un documento salvato quando l'utente lo chiede.

REGOLE OPERATIVE RIGIDE:
- GESTIONE ANAGRAFICA CLIENTI E FORNITORI:
  * Quando l'utente chiede di correggere il nome/ragione sociale di un cliente o fornitore (es. "correggi GiuseppeTambe spa in Giuseppe Tambè S.p.A.") o di aggiornare Partita IVA, Codice Fiscale, Indirizzo, Città, PEC o Telefono, usa subito `modifica_o_elimina_soggetto`.
- ATTRIBUTI EXTRA DINAMICI E MULTI-MAGAZZINO:
  * Ogni azienda può avere colonne personalizzate nel magazzino (es. Magazzino, Deposito, Reparto, Ubicazione, Corsia, Peso_Kg, Scorta_Minima, Fornitore_Abituale).
  * Quando l'utente chiede informazioni su un reparto, un deposito/magazzino o un'ubicazione (es. "cosa c'è nel Deposito Nord?", "dov'è il martello?", "sposta l'articolo X nella Corsia C"), usa `cerca_articolo_magazzino` e `aggiorna_attributi_extra_articolo`.
- DOCUMENTI CARICATI / RICEVUTI DAI FORNITORI (CICLO PASSIVO):
  * Quando l'utente allega un PDF o un'immagine di un documento ricevuto da un fornitore, analizza attentamente OGNI dettaglio: Fornitore, Tipo documento, Numero documento, Data, tutte le righe articolo (Codice, Descrizione, Quantità, Prezzo unitario netto, Sconto, % IVA), Spese accessorie e Totale documento.
  * Controlla con `verifica_o_registra_soggetto` se il fornitore esiste già e con `cerca_articolo_magazzino` se i codici articolo esistono già in magazzino.
  * IMPORTANTE: NON registrare subito nulla! Presenta prima all'utente un riepilogo strutturato e chiaro di tutti i dati estratti e CHIEDI CONFERMA.
  * Solo dopo che l'utente risponde "Sì" / conferma, chiama `registra_documento_fornitore` passando tutti i dati estratti, il `percorso_file_originale` indicato nel messaggio e `aggiorna_anche_magazzino=True`.
  * NON usare MAI `emetti_documento` per documenti emessi da fornitori!
- ACQUISTI SENZA DOCUMENTO ALLEGATO:
  * Quando cerchi un articolo in magazzino con `cerca_articolo_magazzino`, usa sempre il nome al SINGOLARE (es. "martello" invece di "martelli") o la parola chiave principale.
  * Se l'utente dice semplicemente "ho comprato 3 martelli da carpentiere" (senza chiedere di registrare un documento), limita l'azione a `aggiorna_giacenza` (+3) e non emettere alcun PDF.
  * Se dichiara anche il costo d'acquisto (es. "ho comprato 10 martelli a 13 euro"), passa `prezzo_acquisto=13` ad `aggiorna_giacenza` (calcolerà listino = 13 * 1.22 * 1.3).
- VENDITE EMESSE DA NOI (CICLO ATTIVO):
  * Il `Prezzo_Listino` in magazzino è il prezzo consigliato al pubblico GIA' IVATO (IVA 22% inclusa). Se l'utente non specifica un prezzo diverso, usa `Prezzo_Listino` con `prezzi_ivato_inclusa=True`. Se invece l'utente indica un prezzo di vendita personalizzato (es. "vendiglielo a 18 euro"), usa esattamente il prezzo indicato dall'utente.
  * Se l'utente dice "ho venduto 3 pezzi di X" senza menzionare documenti, usa solo `aggiorna_giacenza` (-3).
- RICONOSCIMENTO SOGGETTI: Prima di registrare o emettere qualsiasi documento, verifica sempre il soggetto con `verifica_o_registra_soggetto`.
- MODIFICHE E CANCELLAZIONI:
  * Se l'utente chiede di correggere un documento o un articolo esistente, usa subito `modifica_documento` o `modifica_articolo_magazzino`.
  * Se il tool `elimina_documenti` o `modifica_documento` segnala che esistono PIÙ documenti con lo stesso numero (es. uno del Cliente e uno del Fornitore), FERMATI e chiedi all'utente quale dei due intende prima di procedere.
  * Se l'utente chiede di eliminare o annullare un documento ("cancella questo documento", "fai finta che non sia mai esistito"), chiama `elimina_documenti` con `ripristina_magazzino=True` (NON modificare manualmente il magazzino a parte, perché `elimina_documenti` ripristina già in automatico giacenze, prezzi di acquisto medi, listini e anagrafica). Riporta poi chiaramente all'utente tutto ciò che è stato ripristinato.
- OPACITÀ E ZERO LOG: Quando aggiorni il magazzino, sovrascrivi solo la giacenza senza log storici.

Rispondi sempre in modo sintetico, diretto e professionale."""


def interpreta_data(testo_data: str):
    s = str(testo_data).strip()
    if not s:
        return None

    if s.isdigit():
        if len(s) == 8:
            giorno, mese, anno = int(s[0:2]), int(s[2:4]), int(s[4:8])
        elif len(s) == 6:
            giorno, mese, anno = int(s[0:2]), int(s[2:4]), 2000 + int(s[4:6])
        else:
            return "INVALID"
        try:
            return pd.Timestamp(year=anno, month=mese, day=giorno)
        except ValueError:
            return "INVALID"

    parti = re.split(r"[/\-.]", s)
    if len(parti) == 3 and all(p.isdigit() for p in parti):
        giorno, mese, anno = int(parti[0]), int(parti[1]), int(parti[2])
        if anno < 100:
            anno += 2000
        try:
            return pd.Timestamp(year=anno, month=mese, day=giorno)
        except ValueError:
            return "INVALID"

    dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
    return dt if pd.notna(dt) else "INVALID"


def prepara_contenuto_multimodale(testo_utente: str, percorso_allegato: str):
    if not percorso_allegato or not os.path.exists(percorso_allegato):
        return testo_utente

    estensione = os.path.splitext(percorso_allegato)[1].lower()
    with open(percorso_allegato, "rb") as f:
        dati_b64 = base64.standard_b64encode(f.read()).decode("utf-8")

    testo_completo = (
        f"{testo_utente}\n\n"
        f"[SISTEMA - FILE ALLEGATO DALL'UTENTE: '{os.path.basename(percorso_allegato)}' | "
        f"PERCORSO COMPLETO DA PASSARE A `percorso_file_originale`: '{percorso_allegato}']"
    )

    if estensione == ".pdf":
        return [
            {"type": "text", "text": testo_completo},
            {
                "type": "document",
                "source": {
                    "type": "base64",
                    "media_type": "application/pdf",
                    "data": dati_b64,
                },
            },
        ]
    elif estensione in (".png", ".jpg", ".jpeg", ".webp"):
        mime = "image/jpeg" if estensione in (".jpg", ".jpeg") else f"image/{estensione[1:]}"
        return [
            {"type": "text", "text": testo_completo},
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": mime,
                    "data": dati_b64,
                },
            },
        ]
    return testo_completo


# ==============================================================================
# FINESTRA GUIDATA IMPORTAZIONE MAGAZZINO EXCEL UNIVERSALE
# ==============================================================================
class ImportaMagazzinoDialog(ctk.CTkToplevel):
    def __init__(self, parent, percorso_excel: str, callback_completato):
        super().__init__(parent)
        self.title("Importazione Magazzino Excel Universale")
        self.geometry("760x700")
        self.configure(fg_color=COL_BG_MAIN)
        self.grab_set()

        self.percorso_excel = percorso_excel
        self.callback_completato = callback_completato

        try:
            self.df_preview = pd.read_excel(percorso_excel)
        except Exception as e:
            messagebox.showerror("Errore file Excel", f"Impossibile leggere il file selezionato:\n{e}")
            self.destroy()
            return

        self.colonne_file = [str(c) for c in self.df_preview.columns]
        opzioni_menu = ["(Nessuna / Genera automaticamente)"] + self.colonne_file

        ctk.CTkLabel(
            self,
            text=f"📥 Importazione Guidata: {os.path.basename(percorso_excel)}",
            font=("Segoe UI", 16, "bold"),
            text_color="white"
        ).pack(anchor="w", padx=20, pady=(16, 4))

        ctk.CTkLabel(
            self,
            text=f"Trovati {len(self.df_preview)} articoli e {len(self.colonne_file)} colonne. Associa i campi base e scegli le colonne extra da conservare:",
            font=("Segoe UI", 11),
            text_color=COL_TEXT_MUTED
        ).pack(anchor="w", padx=20, pady=(0, 12))

        frame_core = ctk.CTkFrame(self, fg_color=COL_CARD, corner_radius=10, border_width=1, border_color=COL_BORDER)
        frame_core.pack(fill="x", padx=20, pady=6)

        ctk.CTkLabel(
            frame_core, text="1. ASSOCIAZIONE CAMPI PRINCIPALI (CORE)",
            font=("Segoe UI", 11, "bold"), text_color="#60A5FA"
        ).pack(anchor="w", padx=14, pady=(10, 6))

        campi_core = [
            ("Codice", "Codice Articolo / SKU:", ["cod", "sku", "rif", "id", "articolo"]),
            ("Descrizione", "Descrizione Prodotto:", ["desc", "nome", "prodotto", "voce", "titolo"]),
            ("Giacenza", "Quantità / Giacenza:", ["giac", "qta", "quant", "disp", "esist", "stock"]),
            ("Prezzo_Acq_Medio", "Prezzo Acquisto Netto:", ["acq", "costo", "netto", "forn"]),
            ("Prezzo_Listino", "Prezzo Vendita / Listino:", ["list", "vend", "pubbl", "prezzo", "ivato"])
        ]

        self.combo_core = {}
        for chiave, etichetta, parole_chiave in campi_core:
            riga = ctk.CTkFrame(frame_core, fg_color="transparent")
            riga.pack(fill="x", padx=14, pady=4)

            ctk.CTkLabel(riga, text=etichetta, width=210, anchor="w", font=("Segoe UI", 12)).pack(side="left")
            combo = ctk.CTkOptionMenu(
                riga, values=opzioni_menu, width=320, height=30,
                fg_color=COL_BG_MAIN, button_color=COL_BLUE,
                command=lambda v: self.aggiorna_lista_extra()
            )
            scelta_auto = "(Nessuna / Genera automaticamente)"
            for col in self.colonne_file:
                if col.lower() == chiave.lower():
                    scelta_auto = col
                    break
            if scelta_auto == "(Nessuna / Genera automaticamente)":
                for col in self.colonne_file:
                    if any(pk in col.lower() for pk in parole_chiave):
                        scelta_auto = col
                        break
            combo.set(scelta_auto)
            combo.pack(side="left", padx=8)
            self.combo_core[chiave] = combo

        ctk.CTkLabel(frame_core, text="", height=4).pack()

        frame_extra = ctk.CTkFrame(self, fg_color=COL_CARD, corner_radius=10, border_width=1, border_color=COL_BORDER)
        frame_extra.pack(fill="both", expand=True, padx=20, pady=8)

        ctk.CTkLabel(
            frame_extra, text="2. COLONNE EXTRA AZIENDALI DA CONSERVARE (Multi-magazzino, Reparto, Ubicazione, Peso...)",
            font=("Segoe UI", 11, "bold"), text_color="#34D399"
        ).pack(anchor="w", padx=14, pady=(10, 4))

        self.scroll_extra = ctk.CTkScrollableFrame(frame_extra, fg_color=COL_BG_MAIN, height=110)
        self.scroll_extra.pack(fill="both", expand=True, padx=14, pady=(4, 10))
        self.vars_extra = {}
        self.aggiorna_lista_extra()

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(fill="x", padx=20, pady=(4, 16))

        self.var_modalita = ctk.StringVar(value="sostituisci")
        ctk.CTkRadioButton(
            footer, text="Sostituisci magazzino attuale", variable=self.var_modalita, value="sostituisci",
            font=("Segoe UI", 11)
        ).pack(side="left", padx=(0, 14))
        ctk.CTkRadioButton(
            footer, text="Unisci / Aggiorna esistenti", variable=self.var_modalita, value="unisci",
            font=("Segoe UI", 11)
        ).pack(side="left")

        ctk.CTkButton(
            footer, text="✅ Avvia Importazione", width=180, height=38,
            fg_color=COL_GREEN, hover_color=COL_GREEN_HOVER,
            font=("Segoe UI", 12, "bold"), command=self.esegui_importazione
        ).pack(side="right")

    def aggiorna_lista_extra(self):
        for w in self.scroll_extra.winfo_children():
            w.destroy()
        self.vars_extra.clear()

        usate_core = {c.get() for c in self.combo_core.values() if c.get() != "(Nessuna / Genera automaticamente)"}
        rimanenti = [c for c in self.colonne_file if c not in usate_core]

        if not rimanenti:
            ctk.CTkLabel(
                self.scroll_extra,
                text="Tutte le colonne del file sono già state associate ai campi principali.",
                font=("Segoe UI", 11, "italic"), text_color=COL_TEXT_MUTED
            ).pack(anchor="w", padx=8, pady=8)
            return

        for col in rimanenti:
            var = ctk.BooleanVar(value=True)
            esempio = ""
            non_vuoti = self.df_preview[col].dropna()
            if not non_vuoti.empty:
                esempio = f"  (es. '{str(non_vuoti.iloc[0])[:30]}')"
            chk = ctk.CTkCheckBox(
                self.scroll_extra,
                text=f"{col}{esempio}",
                variable=var,
                font=("Segoe UI", 11)
            )
            chk.pack(anchor="w", padx=8, pady=3)
            self.vars_extra[col] = var

    def esegui_importazione(self):
        mappatura = {}
        for k, combo in self.combo_core.items():
            val = combo.get()
            mappatura[k] = "" if val == "(Nessuna / Genera automaticamente)" else val

        extra_selezionate = [col for col, var in self.vars_extra.items() if var.get()]

        try:
            n = importa_magazzino_da_excel(
                percorso_excel=self.percorso_excel,
                mappatura_core=mappatura,
                colonne_extra=extra_selezionate,
                modalita=self.var_modalita.get()
            )
            messagebox.showinfo(
                "Importazione Completata",
                f"Importati con successo {n} articoli nel database definitivo!\n"
                f"Da questo momento tutti i dati sono salvati al sicuro nel gestionale e non avrai più bisogno del vecchio file Excel."
            )
            self.callback_completato()
            self.destroy()
        except Exception as e:
            messagebox.showerror("Errore Importazione", f"Si è verificato un errore:\n{e}")


# ==============================================================================
# FINESTRA GUIDATA IMPORTAZIONE ANAGRAFICA CLIENTI / FORNITORI
# ==============================================================================
class ImportaAnagraficaDialog(ctk.CTkToplevel):
    def __init__(self, parent, percorso_excel: str, callback_completato=None):
        super().__init__(parent)
        self.title("Importazione Anagrafica Clienti / Fornitori")
        self.geometry("680x640")
        self.configure(fg_color=COL_BG_MAIN)
        self.grab_set()
        self.percorso_excel = percorso_excel
        self.callback_completato = callback_completato

        try:
            self.df_preview = pd.read_excel(percorso_excel)
        except Exception as e:
            messagebox.showerror("Errore file Excel", f"Impossibile leggere il file:\n{e}")
            self.destroy()
            return

        colonne = [str(c) for c in self.df_preview.columns]
        opzioni = ["(Nessuna)"] + colonne

        ctk.CTkLabel(
            self, text=f"👥 Importazione Anagrafica: {os.path.basename(percorso_excel)}",
            font=("Segoe UI", 16, "bold"), text_color="white"
        ).pack(anchor="w", padx=20, pady=(16, 4))

        ctk.CTkLabel(
            self, text=f"Trovati {len(self.df_preview)} nominativi. Associa le colonne del tuo file Excel:",
            font=("Segoe UI", 11), text_color=COL_TEXT_MUTED
        ).pack(anchor="w", padx=20, pady=(0, 12))

        box = ctk.CTkFrame(self, fg_color=COL_CARD, corner_radius=10, border_width=1, border_color=COL_BORDER)
        box.pack(fill="x", padx=20, pady=6)

        campi = [
            ("Nome", "Ragione Sociale / Nome *:", ["ragione", "nome", "cliente", "fornitore", "azienda", "denominazione"]),
            ("Ruolo", "Tipo (Cliente/Fornitore):", ["ruolo", "tipo", "categoria"]),
            ("Partita_IVA", "Partita IVA:", ["p.iva", "piva", "partita", "vat"]),
            ("Codice_Fiscale", "Codice Fiscale:", ["codice fiscale", "c.f.", "cf", "fiscale"]),
            ("Indirizzo", "Indirizzo (Via/Piazza):", ["indirizzo", "via", "sede"]),
            ("Citta", "CAP / Città / Provincia:", ["citta", "città", "comune", "localita", "cap"]),
            ("Email_PEC", "Email / PEC / SDI:", ["pec", "mail", "sdi", "univoco"]),
            ("Telefono", "Telefono / Cellulare:", ["tel", "cell", "phone"])
        ]

        self.combo_campi = {}
        for chiave, label, keywords in campi:
            riga = ctk.CTkFrame(box, fg_color="transparent")
            riga.pack(fill="x", padx=14, pady=4)
            ctk.CTkLabel(riga, text=label, width=210, anchor="w", font=("Segoe UI", 12)).pack(side="left")
            combo = ctk.CTkOptionMenu(riga, values=opzioni, width=320, height=30, fg_color=COL_BG_MAIN, button_color=COL_BLUE)
            scelta = "(Nessuna)"
            for col in colonne:
                if any(kw in col.lower() for kw in keywords):
                    scelta = col
                    break
            combo.set(scelta)
            combo.pack(side="left", padx=8)
            self.combo_campi[chiave] = combo

        ctk.CTkLabel(box, text="", height=4).pack()

        opt_frame = ctk.CTkFrame(self, fg_color=COL_CARD, corner_radius=10, border_width=1, border_color=COL_BORDER)
        opt_frame.pack(fill="x", padx=20, pady=8)

        ctk.CTkLabel(
            opt_frame, text="Se il file non ha una colonna 'Tipo', importa tutti come:",
            font=("Segoe UI", 11, "bold"), text_color="#60A5FA"
        ).pack(anchor="w", padx=14, pady=(10, 4))

        self.var_ruolo = ctk.StringVar(value="Cliente")
        r_row = ctk.CTkFrame(opt_frame, fg_color="transparent")
        r_row.pack(fill="x", padx=14, pady=(0, 10))
        ctk.CTkRadioButton(r_row, text="Clienti", variable=self.var_ruolo, value="Cliente").pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(r_row, text="Fornitori", variable=self.var_ruolo, value="Fornitore").pack(side="left")

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(fill="x", padx=20, pady=(8, 16))
        ctk.CTkButton(
            footer, text="✅ Importa Anagrafica", width=190, height=38,
            fg_color=COL_GREEN, hover_color=COL_GREEN_HOVER,
            font=("Segoe UI", 12, "bold"), command=self.esegui
        ).pack(side="right")

    def esegui(self):
        mappatura = {k: ("" if c.get() == "(Nessuna)" else c.get()) for k, c in self.combo_campi.items()}
        try:
            n = importa_anagrafica_da_excel(self.percorso_excel, mappatura, ruolo_default=self.var_ruolo.get())
            messagebox.showinfo("Anagrafica Importata", f"Importati con successo {n} soggetti nel database!\nL'AI ora riconosce automaticamente i loro dati.")
            if self.callback_completato:
                self.callback_completato()
            self.destroy()
        except Exception as e:
            messagebox.showerror("Errore", f"Errore durante l'importazione:\n{e}")


# ==============================================================================
# FINESTRA GESTIONE BACKUP & RIPRISTINO DATABASE
# ==============================================================================
class GestioneBackupDialog(ctk.CTkToplevel):
    def __init__(self, parent, callback_dopo_ripristino):
        super().__init__(parent)
        self.title("Sicurezza Dati • Backup & Ripristino Database")
        self.geometry("540x360")
        self.configure(fg_color=COL_BG_MAIN)
        self.grab_set()
        self.callback_dopo_ripristino = callback_dopo_ripristino

        ctk.CTkLabel(
            self, text="💾 Protezione e Sicurezza Database",
            font=("Segoe UI", 16, "bold"), text_color="white"
        ).pack(anchor="w", padx=22, pady=(18, 6))

        ctk.CTkLabel(
            self,
            text="Il gestionale salva automaticamente una copia di sicurezza ogni giorno nella cartella 'Backup_DB'.\nPuoi anche salvare una copia su chiavetta USB/Cloud o ripristinare un salvataggio precedente.",
            font=("Segoe UI", 11), text_color=COL_TEXT_MUTED, justify="left", wraplength=490
        ).pack(anchor="w", padx=22, pady=(0, 16))

        card1 = ctk.CTkFrame(self, fg_color=COL_CARD, corner_radius=10, border_width=1, border_color=COL_BORDER)
        card1.pack(fill="x", padx=22, pady=6)
        ctk.CTkLabel(card1, text="Esporta Copia di Sicurezza (Consigliato prima di grandi modifiche)", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=14, pady=(10, 4))
        ctk.CTkButton(
            card1, text="💾 Salva Copia Database su PC / Chiavetta USB", height=36,
            fg_color=COL_GREEN, hover_color=COL_GREEN_HOVER, font=("Segoe UI", 12, "bold"),
            command=self.salva_copia
        ).pack(fill="x", padx=14, pady=(2, 12))

        card2 = ctk.CTkFrame(self, fg_color=COL_CARD, corner_radius=10, border_width=1, border_color=COL_BORDER)
        card2.pack(fill="x", padx=22, pady=8)
        ctk.CTkLabel(card2, text="Ripristina da un Backup Esistente (es. cambio PC)", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=14, pady=(10, 4))
        ctk.CTkButton(
            card2, text="♻️ Seleziona File Backup (.db) da Ripristinare", height=36,
            fg_color=COL_BLUE, hover_color=COL_BLUE_HOVER, font=("Segoe UI", 12, "bold"),
            command=self.ripristina_copia
        ).pack(fill="x", padx=14, pady=(2, 12))

    def salva_copia(self):
        percorso = filedialog.asksaveasfilename(
            title="Salva Copia di Sicurezza Database",
            defaultextension=".db",
            initialfile=f"Backup_Gestionale_{datetime.date.today().strftime('%Y%m%d')}.db",
            filetypes=[("Database Gestionale", "*.db")]
        )
        if percorso:
            try:
                esporta_backup_manuale(percorso)
                messagebox.showinfo("Backup Completato", f"Copia di sicurezza salvata in:\n{percorso}")
            except Exception as e:
                messagebox.showerror("Errore Backup", f"Impossibile salvare il backup:\n{e}")

    def ripristina_copia(self):
        percorso = filedialog.askopenfilename(
            title="Seleziona file di backup (.db)",
            initialdir=CARTELLA_BACKUP,
            filetypes=[("Database Gestionale", "*.db")]
        )
        if percorso:
            if messagebox.askyesno("Conferma Ripristino", "Vuoi sostituire i dati attuali con quelli del backup selezionato?\n(Verrà comunque salvata una copia di sicurezza dello stato attuale)."):
                try:
                    ripristina_backup_db(percorso)
                    messagebox.showinfo("Ripristino Completato", "Database ripristinato con successo!")
                    self.callback_dopo_ripristino()
                    self.destroy()
                except Exception as e:
                    messagebox.showerror("Errore Ripristino", f"Impossibile ripristinare:\n{e}")


# ==============================================================================
# INTERFACCIA GRAFICA PRINCIPALE
# ==============================================================================
class AppGestionale(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("EDIL BARRESE • Gestionale AI & Magazzino")
        self.geometry("1280x850")
        self.minsize(1020, 700)
        self.configure(fg_color=COL_BG_MAIN)

        self.after(10, self.massimizza_finestra)
        self.protocol("WM_DELETE_WINDOW", self.alla_chiusura_app)

        # Motore AI economico ed efficiente: Claude Haiku 5.5 con Prompt Caching attivo
        self.agent_ai = None
        if API_KEY != "MANCANTE" and API_KEY:
            llm = ChatAnthropic(
                model="claude-haiku-5-5",
                api_key=API_KEY,
                max_tokens=3072,
                model_kwargs={
                    "extra_headers": {"anthropic-beta": "prompt-caching-2024-07-31"}
                }
            )
            self.agent_ai = create_react_agent(llm, TUTTI_I_TOOLS)

        self.storico_ram = [
            SystemMessage(
                content=[
                    {
                        "type": "text",
                        "text": SYSTEM_PROMPT,
                        "cache_control": {"type": "ephemeral"}
                    }
                ]
            )
        ]
        self.stato_ordinamento = {}
        self.file_allegato_corrente = ""

        self.cache_magazzino = []
        self.colonne_extra_attive = []
        self.cache_documenti = []
        self.cache_anagrafica = []

        self.configura_stile_tabelle()
        self.costruisci_header()

        # Banner Prima Configurazione (mostrato solo se il magazzino è completamente vuoto)
        self.banner_primo_avvio = ctk.CTkFrame(
            self, fg_color="#1E3A8A", corner_radius=10, border_width=1, border_color="#3B82F6"
        )
        ctk.CTkLabel(
            self.banner_primo_avvio,
            text="👋 PRIMA CONFIGURAZIONE: Il magazzino è vuoto. Importa una sola volta il tuo file Excel aziendale per iniziare:",
            font=("Segoe UI", 12, "bold"), text_color="white"
        ).pack(side="left", padx=16, pady=10)
        ctk.CTkButton(
            self.banner_primo_avvio,
            text="📥 Importa Magazzino Excel Ora",
            height=32, fg_color=COL_GREEN, hover_color=COL_GREEN_HOVER,
            font=("Segoe UI", 12, "bold"), command=self.apri_importazione_excel
        ).pack(side="right", padx=14, pady=8)

        self.tabview = ctk.CTkTabview(
            self,
            fg_color=COL_CARD,
            segmented_button_fg_color=COL_BG_MAIN,
            segmented_button_selected_color=COL_ACCENT,
            segmented_button_selected_hover_color=COL_ACCENT_HOVER,
            segmented_button_unselected_color=COL_CARD_LIGHT,
            segmented_button_unselected_hover_color="#2F3640",
            corner_radius=12,
            border_width=1,
            border_color=COL_BORDER,
            command=self.al_cambio_scheda
        )
        self.tabview.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.tabview._segmented_button.configure(font=("Segoe UI", 13, "bold"), height=38)

        self.tab_chat = self.tabview.add("  💬  Assistente AI  ")
        self.tab_magazzino = self.tabview.add("  📦  Magazzino  ")
        self.tab_documenti = self.tabview.add("  📄  Documenti & Contabilità  ")
        self.tab_anagrafica = self.tabview.add("  👥  Anagrafica  ")

        self.costruisci_scheda_chat()
        self.costruisci_scheda_magazzino()
        self.costruisci_scheda_documenti()
        self.costruisci_scheda_anagrafica()

        self.ricarica_magazzino_da_db()
        self.ricarica_documenti_da_db()
        self.ricarica_anagrafica_da_db()

    def alla_chiusura_app(self):
        esegui_backup_automatico(forza=True)
        self.destroy()

    def massimizza_finestra(self):
        try:
            self.state("zoomed")
        except Exception:
            try:
                self.attributes("-zoomed", True)
            except Exception:
                w, h = self.winfo_screenwidth(), self.winfo_screenheight()
                self.geometry(f"{w}x{h}+0+0")

    def costruisci_header(self):
        header = ctk.CTkFrame(self, fg_color=COL_CARD, corner_radius=12, border_width=1, border_color=COL_BORDER, height=74)
        header.pack(fill="x", padx=20, pady=(18, 12))
        header.pack_propagate(False)

        accent_bar = ctk.CTkFrame(header, fg_color=COL_ACCENT, width=6, corner_radius=3)
        accent_bar.pack(side="left", fill="y", padx=(10, 12), pady=12)

        if os.path.exists(FILE_LOGO):
            try:
                img_pil = Image.open(FILE_LOGO)
                altezza_logo = 48
                larghezza_logo = int(img_pil.width * (altezza_logo / img_pil.height))
                self.logo_ctk = ctk.CTkImage(light_image=img_pil, dark_image=img_pil, size=(larghezza_logo, altezza_logo))
                lbl_logo = ctk.CTkLabel(header, image=self.logo_ctk, text="")
                lbl_logo.pack(side="left", padx=(0, 12), pady=10)
            except Exception as e:
                print(f"Errore caricamento logo header: {e}")

        title_box = ctk.CTkFrame(header, fg_color="transparent")
        title_box.pack(side="left", pady=10)

        ctk.CTkLabel(
            title_box, text="EDIL BARRESE", font=("Segoe UI", 19, "bold"), text_color="white"
        ).pack(anchor="w")
        ctk.CTkLabel(
            title_box, text="Sistema Gestionale Intelligente • Schema Universale & AI",
            font=("Segoe UI", 11), text_color=COL_TEXT_MUTED
        ).pack(anchor="w")

        stato_ok = (self.agent_ai is not None)
        badge_color = "#14532D" if stato_ok else "#7F1D1D"
        dot_color = "#22C55E" if stato_ok else "#EF4444"
        testo_stato = "●  Motore AI Attivo" if stato_ok else "●  Chiave API Mancante"

        badge = ctk.CTkFrame(header, fg_color=badge_color, corner_radius=20, height=30)
        badge.pack(side="right", padx=(8, 18))
        ctk.CTkLabel(
            badge, text=testo_stato, font=("Segoe UI", 11, "bold"), text_color=dot_color
        ).pack(padx=14, pady=4)

        ctk.CTkButton(
            header,
            text="💾 Backup / Ripristina DB",
            height=32,
            width=175,
            fg_color=COL_CARD_LIGHT,
            hover_color="#323942",
            border_width=1,
            border_color=COL_BORDER,
            font=("Segoe UI", 11, "bold"),
            command=self.apri_gestione_backup
        ).pack(side="right", padx=6)

        ctk.CTkButton(
            header,
            text="👥 Importa Anagrafica",
            height=32,
            width=165,
            fg_color=COL_PURPLE,
            hover_color=COL_PURPLE_HOVER,
            font=("Segoe UI", 11, "bold"),
            command=self.apri_importazione_anagrafica
        ).pack(side="right", padx=6)

    def apri_gestione_backup(self):
        GestioneBackupDialog(self, callback_dopo_ripristino=self.ricarica_tutto_dopo_ripristino)

    def ricarica_tutto_dopo_ripristino(self):
        self.ricarica_magazzino_da_db()
        self.ricarica_documenti_da_db()
        self.ricarica_anagrafica_da_db()

    def configura_stile_tabelle(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Treeview",
            background=COL_CARD,
            foreground="#F3F4F6",
            rowheight=34,
            fieldbackground=COL_CARD,
            bordercolor=COL_BORDER,
            borderwidth=0,
            font=("Segoe UI", 10)
        )
        style.map('Treeview', background=[('selected', '#1E3A8A')], foreground=[('selected', 'white')])
        style.configure(
            "Treeview.Heading",
            background="#111317",
            foreground="#E5E7EB",
            relief="flat",
            padding=(8, 8),
            font=("Segoe UI", 10, "bold")
        )
        style.map("Treeview.Heading", background=[('active', '#1F242D')])

    def ordina_colonna(self, tree: ttk.Treeview, col: str, tipo_dato: str, intestazioni_base: dict, is_magazzino: bool = False):
        chiave_stato = (id(tree), col)
        decrescente = self.stato_ordinamento.get(chiave_stato, False)

        righe = [(tree.set(item, col), item) for item in tree.get_children("")]

        def chiave_sort(coppia):
            val = str(coppia[0]).strip()
            if tipo_dato == "num":
                try:
                    pulito = val.replace("€", "").replace(" ", "")
                    if "," in pulito and "." in pulito:
                        pulito = pulito.replace(".", "").replace(",", ".")
                    else:
                        pulito = pulito.replace(",", ".")
                    return float(pulito)
                except ValueError:
                    return 0.0
            elif tipo_dato == "date":
                dt = interpreta_data(val)
                return dt if (dt is not None and dt != "INVALID") else pd.Timestamp.min
            else:
                return val.lower()

        righe.sort(key=chiave_sort, reverse=decrescente)

        for idx, (_, item) in enumerate(righe):
            tree.move(item, "", idx)
            if is_magazzino:
                try:
                    giac = int(float(tree.set(item, "Giacenza")))
                except ValueError:
                    giac = 999
                tag = "lowstock" if giac < 0 else ("evenrow" if idx % 2 == 0 else "oddrow")
            else:
                tag = "evenrow" if idx % 2 == 0 else "oddrow"
            tree.item(item, tags=(tag,))

        for c_id, titolo_pulito in intestazioni_base.items():
            if c_id == col:
                freccia = " ▼" if decrescente else " ▲"
                tree.heading(c_id, text=titolo_pulito + freccia)
            else:
                tree.heading(c_id, text=titolo_pulito)

        self.stato_ordinamento[chiave_stato] = not decrescente

    def crea_kpi_card(self, parent, titolo, valore_iniziale, colore_valore="white"):
        card = ctk.CTkFrame(parent, fg_color=COL_CARD_LIGHT, corner_radius=10, border_width=1, border_color=COL_BORDER)
        card.pack(side="left", fill="x", expand=True, padx=5)
        ctk.CTkLabel(card, text=titolo.upper(), font=("Segoe UI", 10, "bold"), text_color=COL_TEXT_MUTED).pack(
            anchor="w", padx=14, pady=(10, 2)
        )
        lbl_val = ctk.CTkLabel(card, text=valore_iniziale, font=("Segoe UI", 19, "bold"), text_color=colore_valore)
        lbl_val.pack(anchor="w", padx=14, pady=(0, 10))
        return lbl_val

    def al_cambio_scheda(self):
        scheda_attiva = self.tabview.get()
        if "Magazzino" in scheda_attiva:
            self.ricarica_magazzino_da_db()
        elif "Documenti" in scheda_attiva:
            self.ricarica_documenti_da_db()
        elif "Anagrafica" in scheda_attiva:
            self.ricarica_anagrafica_da_db()

    # =========================================================================
    # SCHEDA 1: CHAT AGENTE AI
    # =========================================================================
    def costruisci_scheda_chat(self):
        suggerimenti_frame = ctk.CTkFrame(self.tab_chat, fg_color="transparent")
        suggerimenti_frame.pack(fill="x", padx=12, pady=(8, 4))

        ctk.CTkLabel(
            suggerimenti_frame, text="Azioni rapide:", font=("Segoe UI", 11, "bold"), text_color=COL_TEXT_MUTED
        ).pack(side="left", padx=(0, 8))

        esempi = [
            ("📄 Nuovo DDT", "Emetti un DDT per "),
            ("🧾 Fattura Differita", "Emetti fattura differita per "),
            ("📦 Vendita al Banco", "Ho venduto 1 pezzo di "),
            ("🚚 Carico Acquisto", "Ho comprato ")
        ]
        for label_btn, testo_prefill in esempi:
            ctk.CTkButton(
                suggerimenti_frame,
                text=label_btn,
                font=("Segoe UI", 11),
                height=28,
                fg_color=COL_CARD_LIGHT,
                hover_color="#323942",
                border_width=1,
                border_color=COL_BORDER,
                corner_radius=14,
                command=lambda t=testo_prefill: self.inserisci_suggerimento(t)
            ).pack(side="left", padx=4)

        ctk.CTkButton(
            suggerimenti_frame,
            text="🧹 Nuova Chat",
            font=("Segoe UI", 11, "bold"),
            height=28,
            fg_color="#374151",
            hover_color="#4B5563",
            corner_radius=14,
            command=self.azzera_memoria_chat
        ).pack(side="right", padx=4)

        self.chat_box = ctk.CTkTextbox(
            self.tab_chat,
            state="disabled",
            font=("Segoe UI", 13),
            fg_color=COL_BG_MAIN,
            text_color="#F3F4F6",
            corner_radius=10,
            border_width=1,
            border_color=COL_BORDER,
            wrap="word",
            padx=12,
            pady=12
        )
        self.chat_box.pack(pady=8, padx=12, fill="both", expand=True)

        self.chat_box._textbox.tag_config("user_tag", foreground="#60A5FA", font=("Segoe UI", 12, "bold"))
        self.chat_box._textbox.tag_config("ai_tag", foreground="#34D399", font=("Segoe UI", 12, "bold"))
        self.chat_box._textbox.tag_config("sys_tag", foreground="#F87171", font=("Segoe UI", 12, "bold"))
        self.chat_box._textbox.tag_config("sep_tag", foreground="#2D333B")

        self.barra_allegato = ctk.CTkFrame(self.tab_chat, fg_color="#1E3A8A", corner_radius=8, height=32)
        self.lbl_allegato = ctk.CTkLabel(
            self.barra_allegato, text="", font=("Segoe UI", 11, "bold"), text_color="white"
        )
        self.lbl_allegato.pack(side="left", padx=12, pady=4)
        ctk.CTkButton(
            self.barra_allegato, text="✕ Rimuovi", width=75, height=22, fg_color="#1D4ED8", hover_color="#1E40AF",
            font=("Segoe UI", 10, "bold"), command=self.rimuovi_allegato
        ).pack(side="right", padx=8, pady=4)

        self.frame_inferiore = ctk.CTkFrame(self.tab_chat, fg_color="transparent")
        self.frame_inferiore.pack(pady=(4, 12), padx=12, fill="x")

        self.bottone_allega = ctk.CTkButton(
            self.frame_inferiore,
            text="📎 Carica Doc Fornitore",
            width=175,
            height=46,
            fg_color=COL_BLUE,
            hover_color=COL_BLUE_HOVER,
            corner_radius=10,
            font=("Segoe UI", 12, "bold"),
            command=self.seleziona_file_fornitore
        )
        self.bottone_allega.pack(side="left", padx=(0, 10))

        self.input_testo = ctk.CTkEntry(
            self.frame_inferiore,
            placeholder_text="Scrivi un comando oppure allega un PDF/foto di un documento fornitore...",
            font=("Segoe UI", 13),
            fg_color=COL_BG_MAIN,
            border_color=COL_BORDER,
            corner_radius=10,
            height=46
        )
        self.input_testo.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.input_testo.bind("<Return>", lambda event: self.avvia_elaborazione())

        self.bottone_invia = ctk.CTkButton(
            self.frame_inferiore,
            text="Invia Comando  ➤",
            width=155,
            height=46,
            fg_color=COL_ACCENT,
            hover_color=COL_ACCENT_HOVER,
            corner_radius=10,
            font=("Segoe UI", 13, "bold"),
            command=self.avvia_elaborazione
        )
        self.bottone_invia.pack(side="right")

        if self.agent_ai is None:
            self.aggiungi_messaggio("Sistema", "ERRORE: File 'chiave_api.txt' non trovato o vuoto nella cartella dell'applicazione.")
            self.bottone_invia.configure(state="disabled")
        else:
            self.aggiungi_messaggio(
                "Sistema",
                "Benvenuto nel Gestionale AI. Puoi emettere documenti di vendita, caricare documenti fornitore oppure gestire magazzino e anagrafica in linguaggio naturale."
            )

    def seleziona_file_fornitore(self):
        percorso = filedialog.askopenfilename(
            title="Seleziona Documento Fornitore (PDF o Immagine)",
            filetypes=[
                ("Documenti e Immagini", "*.pdf *.png *.jpg *.jpeg *.webp"),
                ("File PDF", "*.pdf"),
                ("Immagini", "*.png *.jpg *.jpeg *.webp")
            ]
        )
        if percorso:
            self.file_allegato_corrente = percorso
            nome_file = os.path.basename(percorso)
            self.lbl_allegato.configure(text=f"📎 Documento pronto per l'analisi: {nome_file}")
            self.barra_allegato.pack(before=self.frame_inferiore, fill="x", padx=12, pady=(2, 4))
            if not self.input_testo.get().strip():
                self.input_testo.insert(0, "Analizza questo documento del fornitore e dimmi cosa contiene prima di registrarlo.")
            self.input_testo.focus()

    def rimuovi_allegato(self):
        self.file_allegato_corrente = ""
        self.barra_allegato.pack_forget()

    def inserisci_suggerimento(self, testo):
        self.input_testo.delete(0, "end")
        self.input_testo.insert(0, testo)
        self.input_testo.focus()

    def azzera_memoria_chat(self):
        self.storico_ram = [self.storico_ram[0]]
        self.rimuovi_allegato()
        self.chat_box.configure(state="normal")
        self.chat_box._textbox.delete("1.0", "end")
        self.chat_box.configure(state="disabled")
        self.aggiungi_messaggio(
            "Sistema",
            "Memoria conversazione azzerata (i dati e i PDF salvati nella cartella restano intatti)."
        )

    def aggiungi_messaggio(self, mittente, testo, e_attesa: bool = False):
        self.chat_box.configure(state="normal")
        if e_attesa:
            self.idx_inizio_attesa = self.chat_box._textbox.index("end-1c")

        tag = "user_tag" if mittente == "Tu" else ("ai_tag" if mittente == "AI" else "sys_tag")
        etichetta = "TU" if mittente == "Tu" else ("ASSISTENTE AI" if mittente == "AI" else "SISTEMA")

        self.chat_box._textbox.insert("end", f"{etichetta}\n", tag)
        self.chat_box._textbox.insert("end", f"{testo}\n\n")
        self.chat_box._textbox.insert("end", "─" * 65 + "\n\n", "sep_tag")

        if e_attesa:
            self.idx_fine_attesa = self.chat_box._textbox.index("end-1c")

        self.chat_box.configure(state="disabled")
        self.chat_box.see("end")

    def avvia_elaborazione(self):
        comando = self.input_testo.get().strip()
        allegato = self.file_allegato_corrente

        if not comando and not allegato:
            return
        if not comando and allegato:
            comando = "Analizza questo documento del fornitore."

        self.input_testo.delete(0, "end")
        self.rimuovi_allegato()

        testo_visualizzato = f"[📎 Allegato: {os.path.basename(allegato)}]\n{comando}" if allegato else comando
        self.aggiungi_messaggio("Tu", testo_visualizzato)
        self.aggiungi_messaggio("AI", "⏳ Elaborazione e analisi in corso...", e_attesa=True)
        self.bottone_invia.configure(state="disabled")
        self.bottone_allega.configure(state="disabled")

        contenuto_msg = prepara_contenuto_multimodale(comando, allegato)
        self.storico_ram.append(HumanMessage(content=contenuto_msg))

        # Conserva al massimo gli ultimi 6 messaggi oltre al SystemMessage
        if len(self.storico_ram) > 7:
            self.storico_ram = [self.storico_ram[0]] + self.storico_ram[-6:]

        threading.Thread(target=self.worker_thread, args=(allegato,), daemon=True).start()

    def worker_thread(self, allegato_inviato: str = ""):
        try:
            risultato = self.agent_ai.invoke({"messages": self.storico_ram})
            risposta = risultato["messages"][-1].content
            self.storico_ram.append(AIMessage(content=risposta))

            # Rimuove i byte Base64 del PDF/immagine dalla RAM (il file fisico resta su disco!)
            if allegato_inviato:
                for idx in range(len(self.storico_ram) - 1, 0, -1):
                    msg = self.storico_ram[idx]
                    if isinstance(msg, HumanMessage) and isinstance(msg.content, list):
                        testi = [b.get("text", "") for b in msg.content if isinstance(b, dict) and b.get("type") == "text"]
                        self.storico_ram[idx] = HumanMessage(content="\n".join(testi))
                        break
        except Exception as e:
            risposta = f"Errore durante l'esecuzione: {str(e)}"

        self.after(0, self.mostra_risposta_ai, risposta)

    def mostra_risposta_ai(self, risposta):
        self.chat_box.configure(state="normal")
        if hasattr(self, "idx_inizio_attesa") and hasattr(self, "idx_fine_attesa"):
            try:
                self.chat_box._textbox.delete(self.idx_inizio_attesa, self.idx_fine_attesa)
            except Exception:
                pass
        self.chat_box.configure(state="disabled")

        self.aggiungi_messaggio("AI", risposta)
        self.bottone_invia.configure(state="normal")
        self.bottone_allega.configure(state="normal")
        self.input_testo.focus()

        self.ricarica_magazzino_da_db()
        self.ricarica_documenti_da_db()
        self.ricarica_anagrafica_da_db()

    # =========================================================================
    # SCHEDA 2: VISUALIZZATORE MAGAZZINO DINAMICO (CORE + EXTRA_JSON)
    # =========================================================================
    def costruisci_scheda_magazzino(self):
        kpi_row = ctk.CTkFrame(self.tab_magazzino, fg_color="transparent")
        kpi_row.pack(fill="x", padx=7, pady=(10, 6))

        self.kpi_mag_articoli = self.crea_kpi_card(kpi_row, "Referenze a Catalogo", "0")
        self.kpi_mag_pezzi = self.crea_kpi_card(kpi_row, "Pezzi Totali in Giacenza", "0", "#60A5FA")
        self.kpi_mag_valore_acq = self.crea_kpi_card(kpi_row, "Valore a Costo (Acq. Medio)", "€ 0,00", "#FBBF24")

        barra_ricerca = ctk.CTkFrame(self.tab_magazzino, fg_color=COL_CARD_LIGHT, corner_radius=10, border_width=1, border_color=COL_BORDER)
        barra_ricerca.pack(fill="x", padx=12, pady=8)

        ctk.CTkLabel(barra_ricerca, text="🔍", font=("Segoe UI", 14)).pack(side="left", padx=(12, 6), pady=8)
        self.cerca_mag_entry = ctk.CTkEntry(
            barra_ricerca,
            placeholder_text="Cerca per Codice, Descrizione, Reparto, Magazzino, Ubicazione...",
            fg_color=COL_BG_MAIN,
            border_color=COL_BORDER,
            height=34,
            width=390
        )
        self.cerca_mag_entry.pack(side="left", padx=(0, 10), pady=8)
        self.cerca_mag_entry.bind("<KeyRelease>", lambda e: self.filtra_e_mostra_magazzino())

        ctk.CTkButton(
            barra_ricerca,
            text="📥 Importa Excel",
            width=140,
            height=34,
            fg_color=COL_PURPLE,
            hover_color=COL_PURPLE_HOVER,
            font=("Segoe UI", 12, "bold"),
            command=self.apri_importazione_excel
        ).pack(side="right", padx=(6, 12), pady=8)

        ctk.CTkButton(
            barra_ricerca,
            text="📊 Esporta Excel",
            width=135,
            height=34,
            fg_color=COL_GREEN,
            hover_color=COL_GREEN_HOVER,
            font=("Segoe UI", 12, "bold"),
            command=self.esporta_magazzino_excel
        ).pack(side="right", padx=4, pady=8)

        ctk.CTkButton(
            barra_ricerca,
            text="↻ Aggiorna",
            width=105,
            height=34,
            fg_color=COL_BLUE,
            hover_color=COL_BLUE_HOVER,
            font=("Segoe UI", 12, "bold"),
            command=self.ricarica_magazzino_da_db
        ).pack(side="right", padx=4, pady=8)

        self.frame_tab_mag = ctk.CTkFrame(self.tab_magazzino, fg_color=COL_BG_MAIN, corner_radius=10, border_width=1, border_color=COL_BORDER)
        self.frame_tab_mag.pack(fill="both", expand=True, padx=12, pady=(4, 12))

        self.tree_mag = None
        self.headers_mag = {}

    def _costruisci_colonne_treeview_magazzino(self, extra_cols: list):
        for w in self.frame_tab_mag.winfo_children():
            w.destroy()

        colonne_base = ["Codice", "Descrizione", "Giacenza", "Prezzo_Acq_Medio", "Prezzo_Listino"]
        tutte_colonne = colonne_base + extra_cols

        self.tree_mag = ttk.Treeview(self.frame_tab_mag, columns=tutte_colonne, show="headings")

        self.headers_mag = {
            "Codice": "CODICE",
            "Descrizione": "DESCRIZIONE ARTICOLO",
            "Giacenza": "GIACENZA",
            "Prezzo_Acq_Medio": "P. ACQ. MEDIO (€)",
            "Prezzo_Listino": "LISTINO (€)"
        }
        tipi_mag = {
            "Codice": "str",
            "Descrizione": "str",
            "Giacenza": "num",
            "Prezzo_Acq_Medio": "num",
            "Prezzo_Listino": "num"
        }

        for ec in extra_cols:
            self.headers_mag[ec] = ec.upper().replace("_", " ")
            tipi_mag[ec] = "str"

        for col_id, titolo in self.headers_mag.items():
            self.tree_mag.heading(
                col_id,
                text=titolo,
                command=lambda c=col_id, t=tipi_mag[col_id]: self.ordina_colonna(
                    self.tree_mag, c, t, self.headers_mag, is_magazzino=True
                )
            )

        self.tree_mag.column("Codice", width=110, anchor="center")
        self.tree_mag.column("Descrizione", width=330, anchor="w")
        self.tree_mag.column("Giacenza", width=95, anchor="center")
        self.tree_mag.column("Prezzo_Acq_Medio", width=135, anchor="e")
        self.tree_mag.column("Prezzo_Listino", width=120, anchor="e")

        for ec in extra_cols:
            self.tree_mag.column(ec, width=145, anchor="center")

        self.tree_mag.tag_configure("oddrow", background="#181C20")
        self.tree_mag.tag_configure("evenrow", background="#1E2329")
        self.tree_mag.tag_configure("lowstock", background="#3B181C", foreground="#FCA5A5")

        scroll_y = ttk.Scrollbar(self.frame_tab_mag, orient="vertical", command=self.tree_mag.yview)
        scroll_x = ttk.Scrollbar(self.frame_tab_mag, orient="horizontal", command=self.tree_mag.xview)
        self.tree_mag.configure(yscroll=scroll_y.set, xscroll=scroll_x.set)

        scroll_y.pack(side="right", fill="y", pady=2, padx=2)
        scroll_x.pack(side="bottom", fill="x", padx=2, pady=2)
        self.tree_mag.pack(side="left", fill="both", expand=True, padx=(2, 0), pady=(2, 0))

    def apri_importazione_excel(self):
        percorso = filedialog.askopenfilename(
            title="Seleziona il file Excel del Magazzino Aziendale",
            filetypes=[("File Excel", "*.xlsx *.xls")]
        )
        if percorso:
            ImportaMagazzinoDialog(self, percorso, callback_completato=self.ricarica_magazzino_da_db)

    def ricarica_magazzino_da_db(self):
        nuove_extra = []
        viste = set()
        try:
            conn = get_conn()
            cur = conn.cursor()
            cur.execute("SELECT Codice, Descrizione, Giacenza, Prezzo_Acq_Medio, Prezzo_Listino, Extra_JSON FROM magazzino")
            righe = cur.fetchall()
            conn.close()

            self.cache_magazzino = []
            for r in righe:
                d = dict(r)
                try:
                    extra_dict = json.loads(d.get("Extra_JSON") or "{}")
                except Exception:
                    extra_dict = {}
                d["_extra"] = extra_dict
                for k in extra_dict.keys():
                    if k not in viste:
                        viste.add(k)
                        nuove_extra.append(k)
                self.cache_magazzino.append(d)
        except Exception as e:
            print(f"Errore lettura DB magazzino: {e}")

        # Mostra o nasconde automaticamente il banner di Prima Configurazione
        if len(self.cache_magazzino) == 0:
            if not self.banner_primo_avvio.winfo_ismapped():
                self.banner_primo_avvio.pack(before=self.tabview, fill="x", padx=20, pady=(0, 10))
        else:
            if self.banner_primo_avvio.winfo_ismapped():
                self.banner_primo_avvio.pack_forget()

        if self.tree_mag is None or nuove_extra != self.colonne_extra_attive:
            self.colonne_extra_attive = nuove_extra
            self._costruisci_colonne_treeview_magazzino(self.colonne_extra_attive)

        self.filtra_e_mostra_magazzino()

    def filtra_e_mostra_magazzino(self):
        if self.tree_mag is None:
            return

        for riga in self.tree_mag.get_children():
            self.tree_mag.delete(riga)

        for c_id, titolo in self.headers_mag.items():
            self.tree_mag.heading(c_id, text=titolo)

        filtro = self.cerca_mag_entry.get().strip().lower()
        righe_filtrate = []
        for r in self.cache_magazzino:
            if not filtro:
                righe_filtrate.append(r)
                continue
            testo_ricerca = (
                f"{r['Codice']} {r['Descrizione']} " +
                " ".join([f"{k} {v}" for k, v in r.get("_extra", {}).items()])
            ).lower()
            if filtro in testo_ricerca:
                righe_filtrate.append(r)

        tot_articoli = len(righe_filtrate)
        tot_pezzi = 0
        valore_totale_acq = 0.0

        for idx, row in enumerate(righe_filtrate):
            giac = int(row['Giacenza'] or 0)
            prezzo_list = float(row['Prezzo_Listino'] or 0.0)
            prezzo_acq = float(row['Prezzo_Acq_Medio'] or round(prezzo_list / 1.3 / 1.22, 2))

            val_acq_riga = giac * prezzo_acq

            tot_pezzi += giac
            valore_totale_acq += val_acq_riga

            extra_dict = row.get("_extra", {})
            valori_riga = [
                row['Codice'],
                row['Descrizione'],
                giac,
                f"{prezzo_acq:.2f}",
                f"{prezzo_list:.2f}"
            ] + [str(extra_dict.get(ec, "")) for ec in self.colonne_extra_attive]

            tag = "lowstock" if giac < 0 else ("evenrow" if idx % 2 == 0 else "oddrow")
            self.tree_mag.insert("", "end", values=valori_riga, tags=(tag,))

        self.kpi_mag_articoli.configure(text=str(tot_articoli))
        self.kpi_mag_pezzi.configure(text=f"{tot_pezzi:,}".replace(",", "."))
        self.kpi_mag_valore_acq.configure(text=f"€ {valore_totale_acq:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    def esporta_magazzino_excel(self):
        percorso = filedialog.asksaveasfilename(
            title="Esporta Magazzino Completo in Excel",
            defaultextension=".xlsx",
            initialfile=f"Magazzino_Completo_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            filetypes=[("File Excel", "*.xlsx")]
        )
        if percorso:
            righe_export = []
            for r in self.cache_magazzino:
                riga_piatta = {
                    "Codice": r["Codice"],
                    "Descrizione": r["Descrizione"],
                    "Giacenza": r["Giacenza"],
                    "Prezzo_Acq_Medio": r["Prezzo_Acq_Medio"],
                    "Prezzo_Listino": r["Prezzo_Listino"]
                }
                for ec in self.colonne_extra_attive:
                    riga_piatta[ec] = r.get("_extra", {}).get(ec, "")
                righe_export.append(riga_piatta)
            pd.DataFrame(righe_export).to_excel(percorso, index=False)
            messagebox.showinfo("Esportazione completata", f"Magazzino (con tutte le colonne extra) esportato in:\n{percorso}")

    # =========================================================================
    # SCHEDA 3: VISUALIZZATORE DOCUMENTI
    # =========================================================================
    def costruisci_scheda_documenti(self):
        kpi_doc_row = ctk.CTkFrame(self.tab_documenti, fg_color="transparent")
        kpi_doc_row.pack(fill="x", padx=7, pady=(10, 6))

        self.kpi_doc_count = self.crea_kpi_card(kpi_doc_row, "Documenti Visualizzati", "0")
        self.kpi_doc_ddt_sospesi = self.crea_kpi_card(kpi_doc_row, "DDT da Fatturare (Filtro)", "0", "#FBBF24")

        pannello_filtri = ctk.CTkFrame(self.tab_documenti, fg_color=COL_CARD_LIGHT, corner_radius=10, border_width=1, border_color=COL_BORDER)
        pannello_filtri.pack(fill="x", padx=12, pady=6)

        riga1 = ctk.CTkFrame(pannello_filtri, fg_color="transparent")
        riga1.pack(fill="x", padx=12, pady=(10, 6))

        self.filtro_tipo_doc = ctk.CTkSegmentedButton(
            riga1,
            values=["Tutti", "DDT", "FATTURA", "NOTA_CREDITO"],
            selected_color=COL_ACCENT,
            selected_hover_color=COL_ACCENT_HOVER,
            font=("Segoe UI", 11, "bold"),
            height=32,
            command=lambda v: self.filtra_e_mostra_documenti()
        )
        self.filtro_tipo_doc.set("Tutti")
        self.filtro_tipo_doc.pack(side="left", padx=(0, 12))

        self.filtro_ruolo_doc = ctk.CTkSegmentedButton(
            riga1,
            values=["Tutti", "Cliente", "Fornitore"],
            selected_color=COL_BLUE,
            selected_hover_color=COL_BLUE_HOVER,
            font=("Segoe UI", 11, "bold"),
            height=32,
            command=lambda v: self.filtra_e_mostra_documenti()
        )
        self.filtro_ruolo_doc.set("Tutti")
        self.filtro_ruolo_doc.pack(side="left", padx=(0, 12))

        self.cerca_doc_entry = ctk.CTkEntry(
            riga1,
            placeholder_text="🔍 Filtra per Ragione Sociale...",
            fg_color=COL_BG_MAIN,
            border_color=COL_BORDER,
            height=32
        )
        self.cerca_doc_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.cerca_doc_entry.bind("<KeyRelease>", lambda e: self.filtra_e_mostra_documenti())

        ctk.CTkButton(
            riga1,
            text="📊 Esporta Excel",
            width=130,
            height=32,
            fg_color=COL_GREEN,
            hover_color=COL_GREEN_HOVER,
            font=("Segoe UI", 11, "bold"),
            command=self.esporta_documenti_excel
        ).pack(side="right")

        riga2 = ctk.CTkFrame(pannello_filtri, fg_color="transparent")
        riga2.pack(fill="x", padx=12, pady=(4, 10))

        ctk.CTkLabel(riga2, text="📅 Periodo:", font=("Segoe UI", 12, "bold"), text_color="#E5E7EB").pack(side="left", padx=(0, 6))

        self.combo_periodo = ctk.CTkOptionMenu(
            riga2,
            values=["Tutte le date", "Oggi", "Questo Mese", "Mese Scorso", "Quest'Anno"],
            fg_color=COL_BG_MAIN,
            button_color="#323942",
            button_hover_color="#3E4652",
            font=("Segoe UI", 11),
            height=30,
            width=145,
            command=self.applica_periodo_rapido
        )
        self.combo_periodo.set("Tutte le date")
        self.combo_periodo.pack(side="left", padx=(0, 14))

        ctk.CTkLabel(riga2, text="Dal:", font=("Segoe UI", 11), text_color=COL_TEXT_MUTED).pack(side="left", padx=(0, 4))
        self.data_dal_entry = ctk.CTkEntry(
            riga2, placeholder_text="es. 01/11/26", width=110, height=30, fg_color=COL_BG_MAIN, border_color=COL_BORDER, font=("Segoe UI", 11)
        )
        self.data_dal_entry.pack(side="left", padx=(0, 10))
        self.data_dal_entry.bind("<Return>", lambda e: self.filtra_e_mostra_documenti())

        ctk.CTkLabel(riga2, text="Al:", font=("Segoe UI", 11), text_color=COL_TEXT_MUTED).pack(side="left", padx=(0, 4))
        self.data_al_entry = ctk.CTkEntry(
            riga2, placeholder_text="es. 30/11/26", width=110, height=30, fg_color=COL_BG_MAIN, border_color=COL_BORDER, font=("Segoe UI", 11)
        )
        self.data_al_entry.pack(side="left", padx=(0, 10))
        self.data_al_entry.bind("<Return>", lambda e: self.filtra_e_mostra_documenti())

        ctk.CTkButton(
            riga2, text="Applica Date", width=95, height=30, fg_color=COL_BLUE, hover_color=COL_BLUE_HOVER,
            font=("Segoe UI", 11, "bold"), command=self.filtra_e_mostra_documenti
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            riga2, text="Reset Filtri", width=85, height=30, fg_color="#374151", hover_color="#4B5563",
            font=("Segoe UI", 11), command=self.reset_filtri_documenti
        ).pack(side="left")

        ctk.CTkLabel(
            riga2, text="💡 Clic su colonna per ordinare • Doppio clic per aprire PDF",
            font=("Segoe UI", 11, "italic"), text_color=COL_TEXT_MUTED
        ).pack(side="right")

        frame_tab_doc = ctk.CTkFrame(self.tab_documenti, fg_color=COL_BG_MAIN, corner_radius=10, border_width=1, border_color=COL_BORDER)
        frame_tab_doc.pack(fill="both", expand=True, padx=12, pady=(4, 12))

        colonne_doc = ("Data", "Tipo", "Ruolo", "Numero", "Cliente", "Voci", "Importo_Totale", "Fatturato")
        self.tree_doc = ttk.Treeview(frame_tab_doc, columns=colonne_doc, show="headings")

        self.headers_doc = {
            "Data": "DATA",
            "Tipo": "DOCUMENTO",
            "Ruolo": "SOGGETTO",
            "Numero": "N.",
            "Cliente": "RAGIONE SOCIALE",
            "Voci": "DETTAGLIO ARTICOLI",
            "Importo_Totale": "TOTALE (€)",
            "Fatturato": "FATT."
        }
        tipi_doc = {
            "Data": "date",
            "Tipo": "str",
            "Ruolo": "str",
            "Numero": "num",
            "Cliente": "str",
            "Voci": "str",
            "Importo_Totale": "num",
            "Fatturato": "str"
        }

        for col_id, titolo in self.headers_doc.items():
            self.tree_doc.heading(
                col_id,
                text=titolo,
                command=lambda c=col_id, t=tipi_doc[col_id]: self.ordina_colonna(
                    self.tree_doc, c, t, self.headers_doc, is_magazzino=False
                )
            )

        self.tree_doc.column("Data", width=105, anchor="center")
        self.tree_doc.column("Tipo", width=120, anchor="center")
        self.tree_doc.column("Ruolo", width=95, anchor="center")
        self.tree_doc.column("Numero", width=65, anchor="center")
        self.tree_doc.column("Cliente", width=210, anchor="w")
        self.tree_doc.column("Voci", width=320, anchor="w")
        self.tree_doc.column("Importo_Totale", width=110, anchor="e")
        self.tree_doc.column("Fatturato", width=75, anchor="center")

        self.tree_doc.tag_configure("oddrow", background="#181C20")
        self.tree_doc.tag_configure("evenrow", background="#1E2329")
        self.tree_doc.bind("<Double-1>", self.apri_pdf_da_riga)

        scrollbar_doc = ttk.Scrollbar(frame_tab_doc, orient="vertical", command=self.tree_doc.yview)
        self.tree_doc.configure(yscroll=scrollbar_doc.set)

        self.tree_doc.pack(side="left", fill="both", expand=True, padx=(2, 0), pady=2)
        scrollbar_doc.pack(side="right", fill="y", pady=2, padx=2)

    def ricarica_documenti_da_db(self):
        try:
            conn = get_conn()
            cur = conn.cursor()
            cur.execute("SELECT Data, Tipo, Ruolo, Numero, Cliente, Voci, Importo_Totale, Fatturato FROM documenti ORDER BY id DESC")
            self.cache_documenti = [dict(r) for r in cur.fetchall()]
            conn.close()
        except Exception as e:
            print(f"Errore lettura DB documenti: {e}")
        self.filtra_e_mostra_documenti()

    def applica_periodo_rapido(self, scelta):
        oggi = datetime.date.today()
        self.data_dal_entry.delete(0, "end")
        self.data_al_entry.delete(0, "end")

        if scelta == "Oggi":
            s = oggi.strftime("%d/%m/%Y")
            self.data_dal_entry.insert(0, s)
            self.data_al_entry.insert(0, s)
        elif scelta == "Questo Mese":
            inizio = oggi.replace(day=1)
            self.data_dal_entry.insert(0, inizio.strftime("%d/%m/%Y"))
            self.data_al_entry.insert(0, oggi.strftime("%d/%m/%Y"))
        elif scelta == "Mese Scorso":
            primo_corrente = oggi.replace(day=1)
            fine_scorso = primo_corrente - datetime.timedelta(days=1)
            inizio_scorso = fine_scorso.replace(day=1)
            self.data_dal_entry.insert(0, inizio_scorso.strftime("%d/%m/%Y"))
            self.data_al_entry.insert(0, fine_scorso.strftime("%d/%m/%Y"))
        elif scelta == "Quest'Anno":
            inizio_anno = oggi.replace(month=1, day=1)
            self.data_dal_entry.insert(0, inizio_anno.strftime("%d/%m/%Y"))
            self.data_al_entry.insert(0, oggi.strftime("%d/%m/%Y"))

        self.filtra_e_mostra_documenti()

    def reset_filtri_documenti(self):
        self.filtro_tipo_doc.set("Tutti")
        self.filtro_ruolo_doc.set("Tutti")
        self.cerca_doc_entry.delete(0, "end")
        self.combo_periodo.set("Tutte le date")
        self.data_dal_entry.delete(0, "end")
        self.data_al_entry.delete(0, "end")
        self.data_dal_entry.configure(border_color=COL_BORDER)
        self.data_al_entry.configure(border_color=COL_BORDER)
        self.filtra_e_mostra_documenti()

    def apri_pdf_da_riga(self, event):
        selezione = self.tree_doc.selection()
        if not selezione:
            return
        valori = self.tree_doc.item(selezione[0], "values")
        tipo = str(valori[1]).upper()
        numero = str(valori[3])
        cliente = str(valori[4]).replace(" ", "_").replace(".", "").replace("'", "").replace("/", "_")
        prefisso = f"{tipo}_{numero}_{cliente}"

        if os.path.exists(CARTELLA_PDF):
            for fname in os.listdir(CARTELLA_PDF):
                if fname.startswith(prefisso + "."):
                    os.startfile(os.path.join(CARTELLA_PDF, fname))
                    return

        messagebox.showinfo("Allegato non trovato", f"Nessun file '{prefisso}.*' trovato nella cartella Documenti_PDF.")

    def filtra_e_mostra_documenti(self):
        for riga in self.tree_doc.get_children():
            self.tree_doc.delete(riga)

        for c_id, titolo in self.headers_doc.items():
            self.tree_doc.heading(c_id, text=titolo)

        tipo_sel = self.filtro_tipo_doc.get()
        ruolo_sel = self.filtro_ruolo_doc.get()
        testo_sel = self.cerca_doc_entry.get().strip().lower()

        dal_raw = self.data_dal_entry.get().strip()
        al_raw = self.data_al_entry.get().strip()

        dt_dal = interpreta_data(dal_raw)
        dt_al = interpreta_data(al_raw)

        if dt_dal == "INVALID":
            self.data_dal_entry.configure(border_color="#EF4444")
        else:
            self.data_dal_entry.configure(border_color=COL_BORDER)
            if dt_dal is not None:
                self.data_dal_entry.delete(0, "end")
                self.data_dal_entry.insert(0, dt_dal.strftime("%d/%m/%Y"))

        if dt_al == "INVALID":
            self.data_al_entry.configure(border_color="#EF4444")
        else:
            self.data_al_entry.configure(border_color=COL_BORDER)
            if dt_al is not None:
                self.data_al_entry.delete(0, "end")
                self.data_al_entry.insert(0, dt_al.strftime("%d/%m/%Y"))

        ddt_da_fatturare = 0
        mostrati = 0

        for row in self.cache_documenti:
            if tipo_sel != "Tutti" and str(row['Tipo']).upper() != tipo_sel:
                continue
            if ruolo_sel != "Tutti" and str(row['Ruolo']).capitalize() != ruolo_sel:
                continue
            if testo_sel and testo_sel not in str(row['Cliente']).lower():
                continue

            if (dt_dal is not None and dt_dal != "INVALID") or (dt_al is not None and dt_al != "INVALID"):
                dt_riga = interpreta_data(row['Data'])
                if dt_riga is None or dt_riga == "INVALID":
                    continue
                if dt_dal is not None and dt_dal != "INVALID" and dt_riga < dt_dal:
                    continue
                if dt_al is not None and dt_al != "INVALID" and dt_riga > dt_al:
                    continue

            imp = float(row['Importo_Totale'] or 0.0)
            if str(row['Tipo']).upper() == 'DDT' and str(row['Fatturato']).upper() == 'NO':
                ddt_da_fatturare += 1

            tag = "evenrow" if mostrati % 2 == 0 else "oddrow"
            self.tree_doc.insert(
                "", "end",
                values=(
                    row['Data'], row['Tipo'], row['Ruolo'], row['Numero'],
                    row['Cliente'], row['Voci'], f"{imp:.2f}", row['Fatturato']
                ),
                tags=(tag,)
            )
            mostrati += 1

        self.kpi_doc_count.configure(text=str(mostrati))
        self.kpi_doc_ddt_sospesi.configure(text=str(ddt_da_fatturare))

    def esporta_documenti_excel(self):
        percorso = filedialog.asksaveasfilename(
            title="Esporta Registro Documenti in Excel",
            defaultextension=".xlsx",
            initialfile=f"Registro_Documenti_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            filetypes=[("File Excel", "*.xlsx")]
        )
        if percorso:
            pd.DataFrame(self.cache_documenti).to_excel(percorso, index=False)
            messagebox.showinfo("Esportazione completata", f"Registro esportato in:\n{percorso}")

    # =========================================================================
    # SCHEDA 4: ANAGRAFICA CLIENTI & FORNITORI
    # =========================================================================
    def costruisci_scheda_anagrafica(self):
        kpi_ana_row = ctk.CTkFrame(self.tab_anagrafica, fg_color="transparent")
        kpi_ana_row.pack(fill="x", padx=7, pady=(10, 6))

        self.kpi_ana_tot = self.crea_kpi_card(kpi_ana_row, "Totale Soggetti Registrati", "0")
        self.kpi_ana_clienti = self.crea_kpi_card(kpi_ana_row, "Clienti Attivi", "0", "#60A5FA")
        self.kpi_ana_fornitori = self.crea_kpi_card(kpi_ana_row, "Fornitori Registrati", "0", "#34D399")

        barra_ana = ctk.CTkFrame(self.tab_anagrafica, fg_color=COL_CARD_LIGHT, corner_radius=10, border_width=1, border_color=COL_BORDER)
        barra_ana.pack(fill="x", padx=12, pady=8)

        self.filtro_ruolo_ana = ctk.CTkSegmentedButton(
            barra_ana,
            values=["Tutti", "Cliente", "Fornitore"],
            selected_color=COL_BLUE,
            selected_hover_color=COL_BLUE_HOVER,
            font=("Segoe UI", 11, "bold"),
            height=34,
            command=lambda v: self.filtra_e_mostra_anagrafica()
        )
        self.filtro_ruolo_ana.set("Tutti")
        self.filtro_ruolo_ana.pack(side="left", padx=(12, 12), pady=8)

        self.cerca_ana_entry = ctk.CTkEntry(
            barra_ana,
            placeholder_text="🔍 Cerca per Ragione Sociale, P.IVA, Città, PEC o Telefono...",
            fg_color=COL_BG_MAIN,
            border_color=COL_BORDER,
            height=34,
            width=360
        )
        self.cerca_ana_entry.pack(side="left", fill="x", expand=True, padx=(0, 10), pady=8)
        self.cerca_ana_entry.bind("<KeyRelease>", lambda e: self.filtra_e_mostra_anagrafica())

        ctk.CTkButton(
            barra_ana,
            text="👥 Importa Excel",
            width=140,
            height=34,
            fg_color=COL_PURPLE,
            hover_color=COL_PURPLE_HOVER,
            font=("Segoe UI", 12, "bold"),
            command=self.apri_importazione_anagrafica
        ).pack(side="right", padx=(6, 12), pady=8)

        ctk.CTkButton(
            barra_ana,
            text="📊 Esporta Excel",
            width=135,
            height=34,
            fg_color=COL_GREEN,
            hover_color=COL_GREEN_HOVER,
            font=("Segoe UI", 12, "bold"),
            command=self.esporta_anagrafica_excel
        ).pack(side="right", padx=4, pady=8)

        ctk.CTkButton(
            barra_ana,
            text="↻ Aggiorna",
            width=105,
            height=34,
            fg_color=COL_BLUE,
            hover_color=COL_BLUE_HOVER,
            font=("Segoe UI", 12, "bold"),
            command=self.ricarica_anagrafica_da_db
        ).pack(side="right", padx=4, pady=8)

        frame_tab_ana = ctk.CTkFrame(self.tab_anagrafica, fg_color=COL_BG_MAIN, corner_radius=10, border_width=1, border_color=COL_BORDER)
        frame_tab_ana.pack(fill="both", expand=True, padx=12, pady=(4, 12))

        colonne_ana = ("Nome", "Ruolo", "Partita_IVA", "Codice_Fiscale", "Indirizzo", "Citta", "Email_PEC", "Telefono")
        self.tree_ana = ttk.Treeview(frame_tab_ana, columns=colonne_ana, show="headings")

        self.headers_ana = {
            "Nome": "RAGIONE SOCIALE / NOME",
            "Ruolo": "TIPO",
            "Partita_IVA": "PARTITA IVA",
            "Codice_Fiscale": "CODICE FISCALE",
            "Indirizzo": "INDIRIZZO",
            "Citta": "CITTÀ / CAP",
            "Email_PEC": "EMAIL / PEC / SDI",
            "Telefono": "TELEFONO"
        }

        for col_id, titolo in self.headers_ana.items():
            self.tree_ana.heading(
                col_id,
                text=titolo,
                command=lambda c=col_id: self.ordina_colonna(
                    self.tree_ana, c, "str", self.headers_ana, is_magazzino=False
                )
            )

        self.tree_ana.column("Nome", width=230, anchor="w")
        self.tree_ana.column("Ruolo", width=95, anchor="center")
        self.tree_ana.column("Partita_IVA", width=125, anchor="center")
        self.tree_ana.column("Codice_Fiscale", width=140, anchor="center")
        self.tree_ana.column("Indirizzo", width=210, anchor="w")
        self.tree_ana.column("Citta", width=150, anchor="w")
        self.tree_ana.column("Email_PEC", width=190, anchor="w")
        self.tree_ana.column("Telefono", width=120, anchor="center")

        self.tree_ana.tag_configure("oddrow", background="#181C20")
        self.tree_ana.tag_configure("evenrow", background="#1E2329")

        scroll_y = ttk.Scrollbar(frame_tab_ana, orient="vertical", command=self.tree_ana.yview)
        scroll_x = ttk.Scrollbar(frame_tab_ana, orient="horizontal", command=self.tree_ana.xview)
        self.tree_ana.configure(yscroll=scroll_y.set, xscroll=scroll_x.set)

        scroll_y.pack(side="right", fill="y", pady=2, padx=2)
        scroll_x.pack(side="bottom", fill="x", padx=2, pady=2)
        self.tree_ana.pack(side="left", fill="both", expand=True, padx=(2, 0), pady=(2, 0))

    def ricarica_anagrafica_da_db(self):
        try:
            conn = get_conn()
            cur = conn.cursor()
            cur.execute("SELECT Nome, Ruolo, Partita_IVA, Codice_Fiscale, Indirizzo, Citta, Email_PEC, Telefono FROM soggetti ORDER BY Nome ASC")
            self.cache_anagrafica = [dict(r) for r in cur.fetchall()]
            conn.close()
        except Exception as e:
            print(f"Errore lettura DB anagrafica: {e}")
        self.filtra_e_mostra_anagrafica()

    def filtra_e_mostra_anagrafica(self):
        for riga in self.tree_ana.get_children():
            self.tree_ana.delete(riga)

        for c_id, titolo in self.headers_ana.items():
            self.tree_ana.heading(c_id, text=titolo)

        ruolo_sel = self.filtro_ruolo_ana.get()
        filtro = self.cerca_ana_entry.get().strip().lower()

        tot_clienti = sum(1 for r in self.cache_anagrafica if str(r.get("Ruolo", "")).capitalize() == "Cliente")
        tot_fornitori = sum(1 for r in self.cache_anagrafica if str(r.get("Ruolo", "")).capitalize() == "Fornitore")

        mostrati = 0
        for row in self.cache_anagrafica:
            if ruolo_sel != "Tutti" and str(row.get("Ruolo", "")).capitalize() != ruolo_sel:
                continue
            if filtro:
                testo = " ".join(str(v) for v in row.values()).lower()
                if filtro not in testo:
                    continue

            tag = "evenrow" if mostrati % 2 == 0 else "oddrow"
            self.tree_ana.insert(
                "", "end",
                values=(
                    row.get("Nome", ""),
                    row.get("Ruolo", ""),
                    row.get("Partita_IVA", ""),
                    row.get("Codice_Fiscale", ""),
                    row.get("Indirizzo", ""),
                    row.get("Citta", ""),
                    row.get("Email_PEC", ""),
                    row.get("Telefono", "")
                ),
                tags=(tag,)
            )
            mostrati += 1

        self.kpi_ana_tot.configure(text=str(len(self.cache_anagrafica)))
        self.kpi_ana_clienti.configure(text=str(tot_clienti))
        self.kpi_ana_fornitori.configure(text=str(tot_fornitori))

    def esporta_anagrafica_excel(self):
        percorso = filedialog.asksaveasfilename(
            title="Esporta Anagrafica Clienti e Fornitori in Excel",
            defaultextension=".xlsx",
            initialfile=f"Anagrafica_Completa_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            filetypes=[("File Excel", "*.xlsx")]
        )
        if percorso:
            pd.DataFrame(self.cache_anagrafica).to_excel(percorso, index=False)
            messagebox.showinfo("Esportazione completata", f"Anagrafica esportata in:\n{percorso}")

    def apri_importazione_anagrafica(self):
        percorso = filedialog.askopenfilename(
            title="Seleziona il file Excel Anagrafica Clienti / Fornitori",
            filetypes=[("File Excel", "*.xlsx *.xls")]
        )
        if percorso:
            ImportaAnagraficaDialog(self, percorso, callback_completato=self.ricarica_anagrafica_da_db)


if __name__ == "__main__":
    app = AppGestionale()
    app.mainloop()