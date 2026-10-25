# Guida semplice per avviare e usare la webapp

Questa guida spiega come aprire la webapp, fare login e usare le funzioni principali senza conoscere il codice.

In Markdown i titoli sono scritti così: `#` per il titolo principale, `##` per i titoli di sezione e `###` per i sottotitoli. Questa guida usa proprio questa struttura.

## 1. Cosa serve prima di iniziare

Per usare la webapp servono:

- Python installato sul computer.
- Connessione internet, se vuoi usare le funzioni AI con Gemini.
- Una chiave API Gemini nel file `.env`.

Il file `.env` deve stare nella cartella `SicuraFlow WebApp` e deve contenere una riga simile:

```env
GEMINI_API_KEY=la_tua_chiave_api
```

## 2. Come avviare la webapp

### Metodo più semplice

1. Apri la cartella `SicuraFlow WebApp`.
2. Fai doppio click su `avvia_app.bat`.
3. Si apre una finestra nera del terminale.
4. Lascia aperta quella finestra.
5. Fai doppio click su file `index.html` oppure apri il browser e vai su:

```text
http://localhost:8000
```

Se tutto è corretto, comparirà la schermata di login.

### Se il file `.bat` dà errore

Nella cartella è presente `requirements.txt`, mentre il file `.bat` cerca `requirements.txt`.

Se l’installazione automatica non parte, puoi fare così:

1. Apri il terminale nella cartella `SicuraFlow Web App`.
2. Installa le dipendenze con:

```powershell
pip install -r "requirements.txt"
```

3. Avvia il server con:

```powershell
python server.py
```

4. Apri il browser su:

```text
http://localhost:8000
```

## 3. Login disponibili

Prima di entrare con uno degli utenti slezionare SETUP nella schermata iniziale ed inserire la propria api key.
Successivemante entrare con questi utenti demo.

| Profilo | Username | Password | Cosa può fare |
|---|---|---|---|
| Operatore | `martina` | `password` | Crea nuove segnalazioni e vede le proprie |
| Operatore | `marco` | `password` | Crea nuove segnalazioni e vede le proprie |
| Operatore RSPP | `elena` | `password` | Crea nuove segnalazioni e vede le proprie |
| Manager | `manager1` | `password` | Vede dashboard, cartelle, analisi qualità e notifiche |

## 4. Profilo Operatore

Il profilo operatore serve per creare una nuova segnalazione partendo da un racconto audio.

### Creare una nuova segnalazione

1. Fai login come operatore.
2. Entra nella scheda `Nuova segnalazione`.
3. Scegli la categoria:
   - `Incidente`
   - `Near miss`
   - `Operazione di manutenzione`
   - `Rottura macchinario`
4. Scegli come fornire l’audio:
   - `Carica file audio`, se hai già un file.
   - `Registra audio dal microfono`, se vuoi registrare direttamente.
5. Attendi l’analisi AI.
6. Controlla la trascrizione e il report strutturato.
7. Se serve, premi `Modifica report` e correggi i campi.
8. Premi `Conferma report`.

Alla conferma, la segnalazione viene salvata in una cartella incidente. Svolgere questa operazione per tutti e tre 
gli operator usando i corrispettivi file audio associati. il sistema in automatico riconosce che si tratta del medesimo incidente
raccontato da più persone differenti e metterà in maniera autonoma i report generati all'interno della stessa cartella incidente.

### Vedere le proprie segnalazioni

1. Vai su `Le mie segnalazioni`.
2. Trovi solo le cartelle che contengono report creati dall’utente con cui hai fatto login.
3. Clicca una riga per aprire il dettaglio.
4. Nel dettaglio puoi vedere:
   - trascrizione audio;
   - report strutturato;
   - categoria;
   - reparto;
   - gravità;
   - cause, conseguenze e azioni correttive.

## 5. Profilo Manager

Il manager ha quattro schede principali:

- `Dashboard`
- `Analisi cartelle`
- `Analisi Qualità`
- `Notifiche`

## 6. Dashboard Manager

La dashboard mostra una vista generale della situazione.

Qui trovi:

- indicatori sintetici;
- stato delle azioni correttive;
- grafici sugli incidenti;
- mappa/visione impianto;
- ultimi incidenti registrati.

Puoi usare i filtri dei grafici, per esempio settimana, mese o anno, quando disponibili.

## 7. Analisi cartelle

Questa scheda serve per consultare le cartelle incidente già create dagli operatori.

### Aprire una cartella

1. Vai su `Analisi cartelle`.
2. Usa i filtri se vuoi cercare una cartella specifica:
   - codice incidente;
   - categoria;
   - reparto;
   - data;
   - gravità.
3. Clicca sulla riga della cartella.

Nel dettaglio trovi:

- tutti i report associati a quell'incidente;
- autore del report;
- trascrizione audio originale;
- report strutturato;
- confronto tra versioni dello stesso incidente.

### Generare un report unificato

Se una cartella contiene più report, compare il pulsante `Genera Report Unificato`.

1. Apri la cartella incidente.
2. Premi `Genera Report Unificato`.
3. Il sistema fonde i report associati.
4. Il report unificato viene aggiunto alla cartella.

### Usare il chatbot sulla cartella

Dentro il dettaglio di una cartella trovi la sezione `Interroga il chatbot`.

1. Apri una cartella incidente.
2. Scrivi una domanda sul caso.
3. Il chatbot risponde usando le informazioni della cartella aperta.
4. Le domande e risposte vengono salvate nell’archivio chatbot.

Esempi di domande:

- Qual è la causa principale?
- Quali azioni correttive sono suggerite?
- Ci sono DPI mancanti?
- Qual è la gravità stimata?

## 8. Analisi Qualità

La scheda `Analisi Qualità` del manager serve per valutare la qualità delle estrazioni AI, la qualità di come viene generato il
report unificato e dell’archivio chatbot.

Importante: il file audio usato per valutare la qualità delle estrazioni AI qui viene caricato manualmente e temporaneamente. Non prende in automatico i file già presenti in `Analisi cartelle`.

### Fare una estrazione multipla da audio

1. Vai nel profilo manager.
2. Apri la scheda `Analisi Qualità`.
3. Nella sezione `Estrazione multipla`, carica un file audio.
4. Imposta il numero di estrazioni `N`.
   - Un valore tipico è `3`.
   - Più estrazioni richiedono più tempo.
5. Premi `Avvia Analisi`.
6. Attendi la barra di avanzamento.
7. Alla fine compare una tabella con:
   - campi dell’ontologia;
   - risultati di ogni run;
   - Ground Truth Proxy;
   - metriche di qualità.

### Cosa sono le metriche della tabella qualità

Nella tabella trovi metriche come:

- `Cohen Kappa globale`
- `Precision`
- `Recall`
- `F1 Score`
- `Consistenza Lessicale`
- `CER`
- `WER`
- `DTA`
- `EPR`

Puoi cliccare sul nome della metrica per leggere cosa misura e come viene calcolata.

I colori aiutano a leggere velocemente il risultato:

- verde: ottimo;
- giallo: buono;
- arancione: scarso o da controllare;
- rosso: critico.

### Analisi qualità delle risposte chatbot

Sotto l’estrazione multipla c’è la sezione `Analisi qualita risposte chatbot`.

Questa analizza l’archivio delle domande e risposte fatte al chatbot nelle cartelle incidente.

Le metriche principali sono:

- `Indice qualità complessivo`
- `Completezza`
- `Profondità risposta`
- `Copertura cartelle`
- `Domande ripetute`
- `Distribuzione sorgenti`

Puoi cliccare su ogni parametro per vedere:

- cosa misura;
- come viene calcolato.

### Archivio chatbot

Nella stessa scheda trovi anche `Archivio chatbot`.

Qui puoi:

- vedere le ultime domande e risposte salvate;
- esportare l’archivio in CSV;
- svuotare l’archivio.

Per esportare:

1. Vai su `Analisi Qualità`.
2. Scorri fino a `Archivio chatbot`.
3. Premi `Esporta CSV`.

Per cancellare:

1. Premi `Svuota`.
2. Conferma l’operazione.

## 9. Notifiche Manager

La scheda `Notifiche` mostra gli avvisi disponibili per il manager.

Da qui puoi:

- filtrare o consultare notifiche;
- segnare notifiche come lette;
- aprire la scheda `Analisi Qualità` da alcuni avvisi;
- svuotare le notifiche.

## 10. Dove vengono salvati i dati

La webapp salva molte informazioni localmente nel browser, tramite `localStorage`.

Questo significa che:

- i dati restano sullo stesso browser;
- se cambi browser o computer non li ritrovi automaticamente;
- se cancelli i dati del browser, puoi perdere segnalazioni e archivio chatbot.

Gli audio caricati per le analisi vengono salvati temporaneamente nella cartella `temp_audio` durante l’elaborazione e poi rimossi dal server.

## 11. Problemi comuni

### La pagina non si apre

Controlla che il server sia acceso.

La finestra nera del terminale deve rimanere aperta. Poi apri:

```text
http://localhost:8000
```

### L’analisi AI non parte

Controlla:

- che il file `.env` contenga `GEMINI_API_KEY`;
- che il computer sia connesso a internet;
- che il file audio sia valido;
- che il server sia ancora attivo.

### Il `.bat` non installa le dipendenze

Usa il comando manuale:

```powershell
pip install -r "requirements.txt"
```

Poi avvia:

```powershell
python server.py
```

### Non vedo le segnalazioni create

Verifica con quale utente hai fatto login.

Gli operatori vedono solo le proprie segnalazioni nella scheda `Le mie segnalazioni`.

Il manager vede tutte le cartelle nella scheda `Analisi cartelle`.

## 12. Come chiudere la webapp

1. Torna alla finestra nera del terminale.
2. Premi `CTRL + C`.
3. Se chiede conferma, premi `S` o `Y` e poi `Invio`.

La webapp non sarà più raggiungibile finché non riavvii il server.
