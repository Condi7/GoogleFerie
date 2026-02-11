// CONFIGURAZIONE
const CALENDAR_ID =
  "1577ef16117de84a4d94758bd273332ec3d22f538f37feb25963c4d065518aaf@group.calendar.google.com";
const APPROVER_EMAIL = "paoloconde@outlook.it";

// !IMPORTANTE
// I nomi devono corrispondere a quelli dei fogli Google presenti
const SHEET_MODULE = "Risposte del modulo 1";
const SHEET_ARCHIVE = "Archivio";

function getRowData(sheet, row) {
  const values = sheet
    .getRange(row, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  return {
    timestamp: values[0],
    nome: values[1],
    cognome: values[2],
    email: values[3],
    tipo: values[4],
    dataInizio: values[5],
    oraInizio: values[6],
    dataFine: values[7],
    oraFine: values[8],
    note: values[9] || "",
    stato: values[10],
  };
}

// Utility per combinare data e ora
function combineDateTime(dateVal, timeVal, ssTimeZone) {
  if (!dateVal && !timeVal) return null;

  // Ottieni oggetto data base dalla dataVal
  let dateObj = null;
  if (dateVal instanceof Date && !isNaN(dateVal)) {
    dateObj = new Date(
      dateVal.getFullYear(),
      dateVal.getMonth(),
      dateVal.getDate(),
    );
  } else if (typeof dateVal === "string" && dateVal.trim() !== "") {
    // prova formato gg/mm/aaaa o ISO
    let d = new Date(dateVal);
    if (isNaN(d)) {
      const parts = dateVal.split(/[\/\-.]/);
      if (parts.length >= 3) {
        const day = parseInt(parts[0], 10),
          month = parseInt(parts[1], 10) - 1,
          year = parseInt(parts[2], 10);
        d = new Date(year, month, day);
      }
    }
    if (!isNaN(d))
      dateObj = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  // Se non c'è data ma c'è ora, usa oggi nel timezone del foglio
  if (!dateObj && timeVal) {
    const now = Utilities.formatDate(new Date(), ssTimeZone, "yyyy-MM-dd");
    const parts = now.split("-");
    dateObj = new Date(
      parseInt(parts[0], 10),
      parseInt(parts[1], 10),
      parseInt(parts[2], 10),
    );
  }

  // Se non c'è ora, ritorna la data a mezzanotte
  if (!timeVal || timeVal === "") return dateObj;

  // Estrai ore e minuti dall'ora (gestisce Date, numero seriale, stringa "HH:MM")
  let hours = 0,
    minutes = 0;
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
      hours = parseInt(hm[1], 10);
      minutes = parseInt(hm[2], 10);
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

function sendNotificationEmail(type, params) {
  const {
    approverEmail = null,
    startDateTime,
    endDateTime,
    rowLink,
    ...rowData
  } = params;

  let subject = "";
  let body = "";

  switch (type) {
    case "APPROVED_ALL_DAY":
      subject = `${rowData.tipo} approvata`;
      body = `La tua richiesta ${rowData.tipo} è stata approvata per le date ${rowData.dataInizio} - ${rowData.dataFine}`;
      break;

    case "APPROVED_TIMED":
      subject = `${rowData.tipo} approvata`;
      body = `La tua richiesta ${rowData.tipo} è stata approvata per il periodo ${startDateTime} - ${endDateTime}`;
      break;

    case "REJECTED":
      subject = `${rowData.tipo} rifiutata`;
      body = `La tua richiesta ${rowData.tipo} per il periodo ${rowData.dataInizio} - ${rowData.dataFine} è stata rifiutata dal responsabile.`;
      break;

    case "DELETED":
      subject = `${rowData.tipo} annullata`;
      body = `La tua richiesta ${rowData.tipo} per il periodo ${rowData.dataInizio} - ${rowData.dataFine} è stata annullata dal responsabile.`;
      break;

    case "DUPLICATE_EVENT":
      subject = "Evento duplicato rilevato";
      body = `Tentativo di creare evento duplicato per ${rowData.nome} ${rowData.cognome}.\n\nPeriodo: ${rowData.dataInizio} - ${rowData.dataFine}`;
      break;

    case "DATE_ERROR":
      subject = "Errore date/orari";
      body = `La richiesta di ${rowData.nome} ${rowData.cognome} ha orari non validi.\nControlla il foglio.`;
      break;

    case "NEW_REQUEST":
      subject = "Richiesta assenza in attesa di approvazione";
      body = [
        `Nuova richiesta ${rowData.tipo} da ${rowData.nome} ${rowData.cognome}`,
        `Email: ${rowData.email}`,
        `Date: ${rowData.dataInizio} - ${rowData.dataFine}`,
        `Note: ${rowData.note || "-"}`,
        `Apri il foglio per approvare: ${rowLink}`,
      ].join("\n");
      break;

    default:
      return;
  }

  MailApp.sendEmail(approverEmail || rowData.email, subject, body);
}

function handleApprovedRequest(sheet, rowData, row) {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  // Determina se evento è tutto il giorno (nessuna ora fornita) o con orari
  const hasStartTime = !!rowData.oraInizio;
  const hasEndTime = !!rowData.oraFine;
  const timeZone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  let startDateTime = combineDateTime(
    rowData.dataInizio,
    rowData.oraInizio,
    timeZone,
  );
  let endDateTime = combineDateTime(
    rowData.dataFine,
    rowData.oraFine,
    timeZone,
  );
  const title = `${rowData.tipo} - ${rowData.cognome}`;

  // EVENTO TUTTO IL GIORNO
  if (!hasStartTime && !hasEndTime) {
    const start = new Date(rowData.dataInizio);
    const end = new Date(rowData.dataFine);

    end.setDate(end.getDate() + 1); // end esclusivo

    // Controllo duplicati
    const events = calendar.getEvents(start, end);
    const exists = events.some(
      (ev) => ev.getTitle() === title && ev.isAllDayEvent(),
    );

    if (!exists) {
      const ev = calendar.createAllDayEvent(title, start, end, {
        description: `Richiedente: ${rowData.nome} ${rowData.cognome}\nEmail: ${rowData.email}\nNote: ${rowData.note}`,
      });
      if (rowData.email) ev.addGuest(rowData.email);
      sheet.getRange(row, 12).setValue(new Date()); // Colonna K: Data approvazione

      sendNotificationEmail("APPROVED_ALL_DAY", { ...rowData });
    } else {
      sendNotificationEmail("DUPLICATE_EVENT", { ...rowData });
    }

    // EVENTO CON ORARI
  } else {
    if (!endDateTime) {
      endDateTime = new Date(rowData.dataFine);
      endDateTime.setHours(23, 59, 59, 999);
    }
    if (!startDateTime) {
      startDateTime = new Date(rowData.dataInizio);
      startDateTime.setHours(0, 0, 0, 0);
    }

    // Controllo validità orari
    if (endDateTime <= startDateTime) {
      sendNotificationEmail("DATE_ERROR", { ...rowData });
      return;
    }

    // Controllo duplicati
    const events = calendar.getEvents(startDateTime, endDateTime);
    const exists = events.some((ev) => ev.getTitle() === title);

    if (!exists) {
      const ev = calendar.createEvent(title, startDateTime, endDateTime, {
        description: `Richiedente: ${rowData.nome} ${rowData.cognome}\nEmail: ${rowData.email}\nNote: ${rowData.note}`,
      });

      if (rowData.email) ev.addGuest(rowData.email);
      sheet.getRange(row, 12).setValue(new Date()); // Colonna K: Data approvazione

      sendNotificationEmail("APPROVED_TIMED", {
        ...rowData,
        startDateTime,
        endDateTime,
      });
    } else {
      sendNotificationEmail("DUPLICATE_EVENT", { ...rowData });
    }
  }
}

function handleRejectedRequest(sheet, rowData, row) {
  sendNotificationEmail("REJECTED", { ...rowData });
  sheet.getRange(row, 13).setValue(new Date()); // Colonna L: Data rifiuto
}

function handleDeletedRequest(sheet, rowData, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = ss.getSheetByName(SHEET_ARCHIVE);
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  if (!archiveSheet) {
    Logger.log("Foglio 'Archivio' non trovato");
    return;
  }

  // Cancella evento dal calendario (se esiste)
  const eventId = sheet.getRange(row, 13).getValue(); // colonna M

  if (eventId) {
    try {
      const ev = calendar.getEventById(eventId);
      if (ev) ev.deleteEvent();
    } catch (e) {
      Logger.log("Evento non trovato o già eliminato");
    }
  }

  // Copia l'intera riga nell'archivio
  const lastCol = sheet.getLastColumn();
  const rowValues = sheet.getRange(row, 1, 1, lastCol).getValues(); 
  const datArchiviazione = new Date()

  archiveSheet.appendRow([...rowValues[0], datArchiviazione]);
  sheet.deleteRow(row);

  sendNotificationEmail("DELETED", { ...rowData });
}

function handleAllRequests(sheet, stato, rowData, row) {
  switch (stato.toUpperCase()) {
    case "APPROVATA":
      handleApprovedRequest(sheet, rowData, row);
      break;
    case "RIFIUTATA":
      handleRejectedRequest(sheet, rowData, row);
      break;
    case "ANNULLATA":
      handleDeletedRequest(sheet, rowData, row);
      break;
    default:
      return;
  }
}

// Funzione eseguita al submit del form
// 1) Imposta lo stato inziale su "In attesa"
// 2) Manda una mail al responsabile delle ferie, indicando che c'è un nuovo evento in attesa di approvazione

function onFormSubmit(e) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MODULE);
  const row = e.range.getRow();
  const rowData = getRowData(sheet, row);

  // Imposta lo stato iniziale su In attesa nella colonna "Stato" (11)
  sheet.getRange(row, 11).setValue("In attesa");

  // Notifica al responsabile con link alla riga
  const ssUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const rowLink = `${ssUrl}#gid=${sheet.getSheetId()}&range${row}:${row}`;

  sendNotificationEmail("NEW_REQUEST", {
    approverEmail: APPROVER_EMAIL,
    ...rowData,
    rowLink,
  });
}

// Funzione che reagisce alle modifiche del foglio, controlla cambi di Stato
// Questo consente di mandare una seconda mail una volta che il foglio viene modificato
// (cioè quando il responsabile ha confermato o rifiutato la richiesta)

function onEdit(e) {
  if (!e) return;

  const range = e.range;
  const sheet = range.getSheet();

  if (sheet.getName() !== SHEET_MODULE) return;
  if (range.getColumn() !== 11) return; // Colonna Stato
  if (range.getRow() === 1) return; // Header

  const stato = String(range.getValue()).trim().toLowerCase();
  const row = range.getRow();
  const rowData = getRowData(sheet, row);

  handleAllRequests(sheet, stato, rowData, row);
}
