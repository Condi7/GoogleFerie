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

- III - **Inserire e modificare lo script**
    - III-A - Dal foglio Google delle risposte, cliccare su "Estensioni" e poi su "Google Scripts"
    - III-B - Creare un nuovo file e copiare il codice presente su questa repository
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
