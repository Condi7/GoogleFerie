# GoogleFerie
Un piccolo programma di gestione delle ferie, scritto utilizzando gli strumenti di Google Workspace

# Istruzioni
- I - **Creare il form per la richiesta ferie/permessi su Google Forms**
     - I-A - Per un corretto funzionamento del programma è fortemente consigliato utilizzare **ALMENO** le seguenti domande:
        - Nome (risposta breve) - **OBBLIGATORIO**
        - Cognome (risposta breve) - **OBBLIGATORIO**
        - E-mail aziendale (risposta breve) - **OBBLIGATORIO**
        - Tipologia (domanda a risposta multipla con due opzioni: Ferie o Permesso) - **OBBLIGATORIO**
        - Data inizio (data) - **OBBLIGATORIO**
        - Ora inizio (ora) - opzionale
        - Data fine (data) - **OBBLIGATORIO**
        - Ora fine (ora) - opzionale
        - Note aggiuntive/motivazioni (risposta lunga) - opzionale, ma suggerito (qui il richiedente può inserire cose del tipo "Maternità", "Visita medica", ecc...)
    - I-B - Salvare il form su Google Drive
    - I-C - Pubblicare il form, selezionando le opzioni di condivisione desiderate (ad esempio è possibile rendere visibile il form ad un particolare dominio di mail come @bluecube.it)

- II - **Creare il foglio Google, partendo dal form appena creato**
    - II-A - Nel form appena creato cliccare su "Risposte" e creare un foglio Google collegato al form (ovviamente questo passo deve essere fatto da chi ha creato il form)
    - II-B - Aggiungere le colonne "Stato", "Data approvazione" e "Data rifiuto"
    - II-C - Nella colonna "Stato" mettere dei valori fissi in un elenco a discesa (**CAMPI OBBLIGATORI** = In attesa, Approvata, Rifiutata, Annullata - è possibile aggiungere altri stati in seguito, ma dovranno               essere implementati nello script per fare in modo che funzionino correttamente)
    - II-D - Creare un foglio di nome "Archivio"

- III - **Inserire e modificare lo script**
    - III-A - Dal foglio Google delle risposte, cliccare su "Estensioni" e poi su "Google Scripts"
    - III-B - Creare un nuovo file e copiare il codice presente su questa repository
              (**ATTENZIONE!** - Il seguente script funziona solamente se vengono rispettati i parametri indicati in precedenza, altrimenti                 bisognerà cambiare gli indici nel codice per farli                     combaciare ai campi del form diversi da quanto indicato sopra.)
    - III-C - Cambiare i primi tre parametri presenti nel file con quelli corretti, per farlo:
        
        - Aprire Google Calendar, creare un nuovo calendario, e dalle opzioni del calendario appena creato cliccare su "Opzioni e condivisione", copiare l'id del calendario e incollarlo nello script nel parametro           CALENDAR_ID
        - Sostituire il parametro APPROVER_EMAIL con la mail aziendale del responsabile delle ferie
        - Sostituire il paremetro SHEET_NAME con il nome del **foglio** (non del **file**!) Google
        - Se necessario, è possibile modificare il testo e l'oggetto delle mail automaticamente generate ed inviate (leggere il codice e modificare dove viene indicato dai commenti)

    - III-D - Inserire due nuovi attivatori (trigger) dalla barra sinistra di Google Scripts, andando a modificare **SOLO** i parametri "Scegli quale funzione eseguire" e "Seleziona il tipo di evento":

        - **Primo trigger**: Scegli quale funzione eseguire: onFormSubmit(), Seleziona il tipo di evento: All'invio del modulo
        - **Secondo trigger**: Scegli quale funzione eseguire: onEdit(). Selezione il tipo di evento: Alla modifica
     
        - **Nota**: Vanno quindi lasciati invariati i campi "Viene eseguito durante il deployment" (sempre **head**), e "Seleziona l'origine dell'evento" (sempre "**Da foglio di lavoro**).
        - La frequenza con la quale vengono inviate notifiche via mail riguardo gli incident dei trigger è personale, e non va ad influire sul codice.

# Features aggiuntive
-  Visto che tutto il programma rimane gestibile dal foglio Google, possono essere implementati dei fogli aggiuntivi che possono tornare utili agli admin.
- Un primo esempio potrebbe essere un foglio con il totale delle ore suddivise per nome e cognome del richiedente, in modo tale da vedere il totale ore di ferie e permessi effettuati da ogni dipendente.
- Suggerimento di formula per ottenere questo (se i nomi delle colonne coincidono con quelli sopracitati): 
- =SE(O(A2=""; B2=""); ""; SOMMA.PIÙ.SE('Risposte del modulo 1'!$N$2:$N; 'Risposte del modulo 1'!$B$2:$B; A2; 'Risposte del modulo 1'!$C$2:$C; B2))

# Cosa c'è di nuovo?

- **Implementato l'archivio**
- **Creato un nuovo foglio per la visualizzazione veloce delle ore totali di ferie e permessi**
- **Eseguito refactoring del codice**
- **Inserito controllo sul form, per evitare che vengano inviati nomi, cognomi od email con uno spazio all'inizio o alla fine (questo creerebbe dei problemi nel foglio Google)**

# Funzionalità in fase di sviluppo (10/02/2026 - 17.53)
- Lo stato della richiesta "archiviata"
- Possibilità di avere diversi approvatori delle ferie e dei permessi.
- Visualizzazione mensile delle ore di ferie e permessi per ogni dipendente.
- Funzione onOpen() per archiviare automaticamente le richieste di ferie che sono scadute.
