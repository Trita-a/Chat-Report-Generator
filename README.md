# Chat Report Generator

**Chat Report Generator** è un tool in Python progettato per convertire le esportazioni delle chat (dai report Excel di Cellebrite o formati simili) in report HTML puliti, leggibili e ricercabili.

Replica lo stile visivo delle app di messaggistica più diffuse (Signal e WhatsApp) per offrire un'esperienza di lettura familiare per analisi forensi o revisioni.

<p align="center">
  <img src="preview.png" width="45%" alt="Anteprima Report Chat - Signal" />
  <img src="preview_whatsapp.png" width="45%" alt="Anteprima Report Chat - WhatsApp" />
</p>

## ✨ Funzionalità

*   **Eseguibile Standalone**: Funziona su Windows senza bisogno di installare Python.
*   **Doppio Stile**:
    *   **Signal**: Tema autentico blu/bianco con avatar rotondi.
    *   **WhatsApp**: Classico stile con bolle e sfondo predefinito.
*   **Ricerca HTML (Premium)**: Barra di ricerca integrata nel report HTML per filtrare le chat o cercare messaggi specifici istantaneamente.
*   **Ricerca Avanzata**: Trova messaggi, parole chiave e traduzioni, evidenziandoli e scorrendo direttamente al punto esatto della conversazione.
*   **Supporto Media**: Gestisce immagini, video e allegati audio se presenti nell'esportazione.
*   **Parsing Intelligente**: Rileva automaticamente i partecipanti "Proprietario" e "Contatto" dalle intestazioni Cellebrite.

## 🚀 Utilizzo

### 1. Usando l'Eseguibile (Windows)
Avvia semplicemente `ChatReportGenerator.exe`.
1.  **Seleziona File Excel**: Scegli la tua esportazione `.xlsx`.
2.  **Seleziona Stile**: Scegli tra "Signal" o "WhatsApp".
3.  **Genera**: Il tool creerà una cartella con il report HTML.

### 2. Esecuzione dal Codice Sorgente
1.  Installa le dipendenze:
    ```bash
    pip install -r requirements.txt
    ```
2.  Avvia lo script:
    ```bash
    python ChatReportGenerator.py
    ```

## 🛠️ Compilazione

Per compilare l'eseguibile da solo:

```bash
pyinstaller --noconsole --onefile --clean --name "ChatReportGenerator" ChatReportGenerator.py
```

## 📂 Struttura del Progetto

*   `ChatReportGenerator.py`: Script principale consolidato (Interfaccia + Logica).
*   `ChatReportGenerator_Pandas.py`: Versione di backup che usa Pandas.
*   `dist/`: Cartella di output per l'eseguibile.

## 📝 Licenza
Copyright © 2026 William Tritapepe. Tutti i diritti riservati.
