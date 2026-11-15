# 🏢 Gestionale AI & Magazzino Universale

Applicazione desktop per la gestione intelligente di **Magazzino, Documenti Contabili (DDT, Fatture, Note di Credito) e Anagrafica Clienti/Fornitori**, dotata di un assistente AI integrato (basato su **Claude Haiku 5.5**) capace di operare in linguaggio naturale e leggere documenti PDF o immagini ricevuti dai fornitori.

---

## ✨ Funzionalità Principali

* **💬 Assistente AI Operativo:**
* Emette documenti di vendita (**DDT, Fatture, Note di Credito**) generando automaticamente il PDF formale e scaricando il magazzino.
* Raggruppa i DDT in sospeso di un cliente in un'unica **Fattura Differita**.
* Legge PDF e foto di **fatture o DDT ricevuti dai fornitori**, propone un riepilogo per conferma e carica automaticamente le giacenze e i costi medi d'acquisto.
* Aggiorna giacenze, prezzi, attributi personalizzati e dati anagrafici tramite comandi testuali.


* **📦 Magazzino Universale Dinamico:**
* Importazione guidata da qualsiasi file Excel aziendale con associazione colonne e conservazione automatica di tutte le colonne extra personalizzate (es. *Reparto, Deposito, Ubicazione/Corsia, Peso, Scorta Minima*).
* Ricerca istantanea su tutti i campi ed esportazione completa in `.xlsx`.


* **📄 Registro Documenti & Contabilità:**
* Filtri per tipologia (*DDT, Fattura, Nota di Credito*), soggetto (*Cliente/Fornitore*), ragione sociale e intervallo date.
* Apertura diretta del PDF allegato con doppio clic sulla riga.


* **👥 Anagrafica Clienti & Fornitori:**
* Importazione/esportazione Excel e modifica rapida via chat AI (con aggiornamento automatico anche sullo storico documenti).


* **💾 Sicurezza & Backup:**
* Backup giornaliero automatico in `Backup_DB/`, esportazione manuale e ripristino da file `.db`.



---

## 📂 Struttura della Cartella

Per utilizzare il gestionale, assicurati di avere nella cartella principale:

* `gestionale_ai.py` — Interfaccia grafica principale (`CustomTkinter`) e motore AI (`LangGraph` + `Anthropic`).
* `database_tools.py` — Gestione database SQLite, importazione Excel, generazione PDF e strumenti dell'agente AI.
* `requirements.txt` — Elenco delle dipendenze Python necessarie.
* `chiave_api.txt` *(da creare)* — File di testo contenente la propria API Key di Anthropic (vedi `chiave_api.example.txt`).
* `logo_azienda.png` *(opzionale)* — Logo aziendale mostrato nell'intestazione dell'app e nei PDF generati.

*(I file `gestionale.db`, `Documenti_PDF/` e `Backup_DB/` vengono creati automaticamente al primo avvio).*

---

## 🚀 Installazione e Avvio (da Sorgente Python)

1. **Clona o scarica il repository** sul tuo computer.
2. **Installa le librerie richieste** aprendo il terminale nella cartella e digitando:
`pip install -r requirements.txt`
3. **Configura la chiave API:**
* Rinomina `chiave_api.example.txt` in **`chiave_api.txt`**.
* Apri `chiave_api.txt` e incolla la tua API Key di Anthropic (senza spazi o a capo aggiuntivi).


4. *(Opzionale)* Sostituisci `logo_azienda.png` con il logo della tua azienda.
5. **Avvia il gestionale** eseguendo:
`python gestionale_ai.py`
6. Al primo avvio, usa i pulsanti **📥 Importa Excel** nelle schede *Magazzino* e *Anagrafica* (puoi testare l'importazione usando `esempio_magazzino.xlsx` ed `esempio_anagrafica.xlsx`).

---

## 📦 Compilazione Eseguibile Windows (`.exe`)

Per generare una cartella autonoma eseguibile senza Python installato, esegui nel terminale:

`pyinstaller --noconfirm --onedir --windowed --name "Gestionale_AI" --collect-all customtkinter --collect-all fpdf --hidden-import "database_tools" gestionale_ai.py`

Al termine, copia **`chiave_api.txt`** (e l'eventuale **`logo_azienda.png`**) dentro la cartella `dist/Gestionale_AI/` prima di avviare `Gestionale_AI.exe`.
