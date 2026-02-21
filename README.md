# PDF → PPTX Converter

Converte ogni pagina di un PDF in una slide PowerPoint con testo nativo selezionabile e sfondo fedele all'originale.

## Come funziona

Per ogni pagina del PDF l'app applica una pipeline a tre strati:

1. **Sfondo** — la pagina viene renderizzata come immagine; i glifi di testo vengono cancellati campionando il colore di sfondo circostante (giallo, colorato, bianco, ecc.), lasciando intatti grafica vettoriale e immagini embedded.
2. **Testo** — il testo nativo del PDF (posizione, font, dimensione, colore, grassetto/corsivo) viene estratto e ricreato come textbox reale in PowerPoint → testo selezionabile, copiabile e ricercabile.
3. **Fallback** — le pagine senza testo nativo (scansioni) vengono inserite come sola immagine.

## Installazione

```bash
pip install streamlit PyMuPDF python-pptx Pillow
```

## Avvio

```bash
streamlit run app.py
```

## Impostazioni

| Campo | Descrizione |
|---|---|
| **DPI** | Risoluzione del layer di sfondo (72–300). 150 è un buon compromesso. |
| **Nome file** | Nome del file `.pptx` scaricabile. |

## Limitazioni

- Le pagine scansionate (senza testo nativo) vengono inserite come immagine senza testo selezionabile.
- Layout molto complessi (tabelle annidate, testo su path curvi) possono non essere perfettamente fedeli.
- Font non installati in PowerPoint vengono sostituiti dal sistema.
