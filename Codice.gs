// CONFIGURAZIONE
const CALENDAR_ID = "1577ef16117de84a4d94758bd273332ec3d22f538f37feb25963c4d065518aaf@group.calendar.google.com";
const APPROVER_EMAIL = "paoloconde@outlook.it";
const SHEET_NAME = "Risposte del modulo 1"; // Deve essere uguale al nome del foglio Google (non del file!)

// Funzione eseguita al submit del form
// 1) Imposta lo stato inziale su "In attesa"
// 2) Manda una mail al responsabile delle ferie, indicando che c'è un nuovo evento in attesa di approvazione

function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const row = e.range.getRow();
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const timestamp = values[0]; // Valore di default in tutti i fogli Google creati da un form
  // Se necessario, cambiare i valori sottostanti in modo tale che coincidono con quelli del form
  const nome = values[1];
  const cognome = values[2];
  const email = values[3];
  const tipo = values[4];
  const dataInizio = values[5];
  const oraInizio = values[6];
  const dataFine = values[7];
  const oraFine = values[8];
  const note = values[9] || "";

  // Imposta lo stato iniziale su In attesa nella colonna "Stato" (11)
  sheet.getRange(row, 11).setValue("In attesa");

  // Notifica al responsabile con link alla riga
  const ssUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const rowLink = ssUrl + "#gid=" + sheet.getSheetId() + "&range=" + row + ":" + row;
  // Qui si può modificare il messaggio della mail
  const message = `Nuova richiesta ${tipo} da ${nome+" "+cognome}\nEmail: ${email}\nDate: ${dataInizio} ${oraInizio || ""} - ${dataFine} ${oraFine || ""}\nNote: ${note}\nApri il foglio per approvare: ${rowLink}`;
  MailApp.sendEmail(APPROVER_EMAIL, "Richiesta assenza in attesa di approvazione", message);
}

// Funzione che reagisce alle modifiche del foglio, controlla cambi di Stato
// Questo consente di mandare una seconda mail una volta che il foglio viene modificato
// (cioè quando il responsabile ha confermato o rifiutato la richiesta)

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  const editedCol = range.getColumn();
  const editedRow = range.getRow();

  // Sostituire con l'indice della colonna "Stato"
  if (editedCol !== 11) return;

  const stato = sheet.getRange(editedRow, 11).getValue().toString().toLowerCase(); // Sostituire con l'indice della colonna stato
  // Anche qui, modificare in base alle colonne presenti nel foglio Google
  const nome = sheet.getRange(editedRow, 2).getValue();
  const cognome = sheet.getRange(editedRow, 3).getValue();
  const email = sheet.getRange(editedRow, 4).getValue();
  const tipo = sheet.getRange(editedRow, 5).getValue();
  const dataInizio = sheet.getRange(editedRow, 6).getValue();
  const oraInizio = sheet.getRange(editedRow, 7).getValue();
  const dataFine = sheet.getRange(editedRow, 8).getValue();
  const oraFine = sheet.getRange(editedRow, 9).getValue();
  const note = sheet.getRange(editedRow, 11).getValue() || "";

// Utility per combinare data e ora
function combineDateTime(dateVal, timeVal, ssTimeZone) {
  if (!dateVal && !timeVal) return null;

  // Ottieni oggetto data base dalla dataVal
  let dateObj = null;
  if (dateVal instanceof Date && !isNaN(dateVal)) {
    dateObj = new Date(dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate());
  } else if (typeof dateVal === "string" && dateVal.trim() !== "") {
    // prova formato gg/mm/aaaa o ISO
    let d = new Date(dateVal);
    if (isNaN(d)) {
      const parts = dateVal.split(/[\/\-.]/);
      if (parts.length >= 3) {
        const day = parseInt(parts[0],10), month = parseInt(parts[1],10)-1, year = parseInt(parts[2],10);
        d = new Date(year, month, day);
      }
    }
    if (!isNaN(d)) dateObj = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  // Se non c'è data ma c'è ora, usa oggi nel timezone del foglio
  if (!dateObj && timeVal) {
    const now = Utilities.formatDate(new Date(), ssTimeZone, "yyyy-MM-dd");
    const parts = now.split("-");
    dateObj = new Date(parseInt(parts[0],10), parseInt(parts[1],10)-1, parseInt(parts[2],10));
  }

  // Se non c'è ora, ritorna la data a mezzanotte
  if (!timeVal || timeVal === "") return dateObj;

  // Estrai ore e minuti dall'ora (gestisce Date, numero seriale, stringa "HH:MM")
  let hours = 0, minutes = 0;
  if (timeVal instanceof Date && !isNaN(timeVal)) {
    hours = timeVal.getHours();
    minutes = timeVal.getMinutes();
  } else if (typeof timeVal === "number") {
    // valore seriale di Sheets: 1 giorno = 1, moltiplica per 24h
    const totalMinutes = Math.round(timeVal * 24 * 60);
    hours = Math.floor(totalMinutes / 60);
    minutes = totalMinutes % 60;
  } else if (typeof timeVal === "string") {
    const hm = timeVal.trim().match(/(\d{1,2}):(\d{2})/);
    if (hm) {
      hours = parseInt(hm[1],10);
      minutes = parseInt(hm[2],10);
    } else {
      // fallback: prova a parsare come Date
      const tmp = new Date(timeVal);
      if (!isNaN(tmp)) {
        hours = tmp.getHours();
        minutes = tmp.getMinutes();
      }
    }
  }

  // Costruisci il Date finale usando la dataObj e le ore estratte
  dateObj.setHours(hours, minutes, 0, 0);
  return dateObj;
}

  // Se approvata, crea evento e notifica il richiedente
  if (stato === "approvata") {
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

    // Determina se evento è tutto il giorno (nessuna ora fornita) o con orari
    const hasStartTime = !!oraInizio;
    const hasEndTime = !!oraFine;

    let startDateTime = combineDateTime(dataInizio, oraInizio);
    let endDateTime = combineDateTime(dataFine, oraFine);

    // Se nessuna ora fornita, useremo createAllDayEvent e la data di fine è esclusiva
    if (!hasStartTime && !hasEndTime) {
      const start = new Date(dataInizio);
      const end = new Date(dataFine);
      end.setDate(end.getDate() + 1); // end esclusivo

      // Controllo duplicati: cerca eventi nello stesso intervallo con stesso titolo
      const title = `${tipo} - ${cognome}`;
      const events = calendar.getEvents(start, end);
      const exists = events.some(ev => ev.getTitle() === title && ev.isAllDayEvent());

      if (!exists) {
        const ev = calendar.createAllDayEvent(title, start, end, {description: `Richiedente: ${nome+" "+cognome}\nEmail: ${email}\nNote: ${note}`});
        if (email) ev.addGuest(email);
        sheet.getRange(editedRow, 12).setValue(new Date()); // Colonna K: Data approvazione
        MailApp.sendEmail(email, `${tipo} approvata`, `La tua richiesta ${tipo} è stata approvata per le date ${dataInizio} - ${dataFine}`);
      } else {
        MailApp.sendEmail(APPROVER_EMAIL, "Evento duplicato rilevato", `Tentativo di creare evento duplicato per ${nome+" "+cognome} nelle date ${dataInizio} - ${dataFine}`);
      }
    } else {
      // Se sono presenti orari, assicurati che endDateTime sia dopo startDateTime
      if (!endDateTime) {
        // se manca ora fine, considera fine come fine della giornata
        endDateTime = new Date(dataFine);
        endDateTime.setHours(23,59,59,999);
      }
      if (!startDateTime) {
        // se manca ora inizio, considera inizio come inizio della giornata
        startDateTime = new Date(dataInizio);
        startDateTime.setHours(0,0,0,0);
      }

      if (endDateTime <= startDateTime) {
        MailApp.sendEmail(APPROVER_EMAIL, "Errore date/orari", `La richiesta di ${nome+" "+cognome} ha orario di fine precedente o uguale all'inizio. Controlla il foglio.`);
        return;
      }

      // Controllo sovrapposizioni: cerca eventi che si sovrappongono all'intervallo
      const events = calendar.getEvents(startDateTime, endDateTime);
      const title = `${tipo} - ${cognome}`;
      const exists = events.some(ev => ev.getTitle() === title);

      if (!exists) {
        const ev = calendar.createEvent(title, startDateTime, endDateTime, {description: `Richiedente: ${nome+" "+cognome}\nEmail: ${email}\nNote: ${note}`});
        if (email) ev.addGuest(email);
        sheet.getRange(editedRow, 12).setValue(new Date()); // Colonna K: Data approvazione
        MailApp.sendEmail(email, `${tipo} approvata`, `La tua richiesta ${tipo} è stata approvata per il periodo ${startDateTime} - ${endDateTime}`);
      } else {
        MailApp.sendEmail(APPROVER_EMAIL, "Evento duplicato rilevato", `Tentativo di creare evento duplicato per ${nome+" "+cognome} nell'intervallo selezionato.`);
      }
    }
  }

  // Se rifiutata, notifica il richiedente e registra data rifiuto
  if (stato === "rifiutata") {
    MailApp.sendEmail(email, `${tipo} rifiutata`, `La tua richiesta ${tipo} per il periodo ${dataInizio} - ${dataFine} è stata rifiutata dal responsabile.`);
    sheet.getRange(editedRow, 12).setValue(new Date()); // Colonna L: Data rifiuto
  }
}
