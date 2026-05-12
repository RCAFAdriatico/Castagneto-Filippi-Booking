// ================================================================
// FILIPPI TROPHY BEACH SPRINT - BOOKING APPS SCRIPT
// Castagneto Carducci 2026 (4a Tappa)
// Rowing Club Adriatico Fano ASD (creator)
// ================================================================

const ADMIN_EMAIL_PRIMARY = "rent@filippiboats.com";
const ADMIN_EMAIL_CC      = "beachsprintfano@gmail.com";
const ADMIN_EMAIL         = ADMIN_EMAIL_PRIMARY + ", " + ADMIN_EMAIL_CC;

const SHEET_BOOKINGS    = "Prenotazioni";
const SHEET_CONFIG      = "Config";
const SHEET_CLUBS       = "Clubs";

// Bank data
const BENEFICIARY = "European Rowing Coastal Challenge";
const IBAN = "IT58P0846170689000010979287";
const BIC = "CCRTIT2TCAS";
const BANK = "Castagneto Banca 1910 - Credito Cooperativo S.C.";

// ================================================================
// ENTRY POINTS
// ================================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.parameter.data);
    const action = data.action;

    if (action === "book")       return jsonResponse(handleBook(data));
    if (action === "cancel")     return jsonResponse(handleCancel(data));
    if (action === "markPaid")   return jsonResponse(handleMarkPaid(data));
    if (action === "moveSlot")   return jsonResponse(handleMoveSlot(data));
    if (action === "summary")    return jsonResponse(handleSummary(data));
    if (action === "getBookings")return jsonResponse(handleGetBookings(data));
    if (action === "setConfig")  return jsonResponse(handleSetConfig(data));
    if (action === "registerClub") return jsonResponse(handleRegisterClub(data));

    return jsonResponse({ ok: false, error: "Unknown action: " + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function doGet(e) {
  return jsonResponse({ ok: true, msg: "Filippi Trophy Booking API - Castagneto 2026" });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================================================================
// INIT SHEETS (run manually once)
// ================================================================
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sBook = ss.getSheetByName(SHEET_BOOKINGS);
  if (!sBook) {
    sBook = ss.insertSheet(SHEET_BOOKINGS);
    sBook.getRange(1, 1, 1, 16).setValues([[
      "Timestamp","Club","Referente","Email","Telefono",
      "Specialita","Giorno","Sessione","Inizio","Fine",
      "DK","SlotMin","Stato","Pagato","Note","Prezzo"
    ]]);
    sBook.setFrozenRows(1);
    sBook.getRange(1, 1, 1, 16).setFontWeight("bold").setBackground("#1e3348").setFontColor("#ffffff");
  }

  let sCfg = ss.getSheetByName(SHEET_CONFIG);
  if (!sCfg) {
    sCfg = ss.insertSheet(SHEET_CONFIG);
    sCfg.getRange(1, 1, 1, 3).setValues([["chiave","valore","descrizione"]]);
    sCfg.setFrozenRows(1);
    sCfg.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#1e3348").setFontColor("#ffffff");

    const defaults = [
      ["prezzo_1x", 15, "Prezzo singolo 1x (EUR)"],
      ["prezzo_2x", 30, "Prezzo doppio 2x (EUR)"],
      ["prezzo_4x+", 60, "Prezzo quattro 4x+ (EUR)"],
      ["max_1x", 6, "Max 1x per slot"],
      ["max_2x", 5, "Max 2x per slot"],
      ["max_4x+", 1, "Max 4x+ per slot"],
      ["admin_password", "Castagneto2026", "Password admin app"],
      ["admin_user", "FilippiCastagneto26", "Username admin"],
      ["event_name", "Filippi Trophy Castagneto Carducci 2026", "Nome evento"],
      ["event_dates", "3-7 Giugno 2026", "Date evento"],
      ["organizer", "European Rowing Coastal Challenge", "Organizzatore"],
      ["organizer_vat", "01956180499", "P.IVA/CF organizzatore"],
      ["organizer_sdi", "W7YVJK9", "Codice SDI"],
      ["iban", IBAN, "IBAN"],
      ["bic", BIC, "BIC/SWIFT"],
      ["bank", BANK, "Banca"],
      ["beneficiary", BENEFICIARY, "Beneficiario"],
      ["contact_name", "Claudio Sulas", "Persona di riferimento"],
      ["contact_email", "rent@filippiboats.com", "Email contatto pubblica"],
      ["contact_phone", "+39 328 865 9759", "Telefono contatto"],
    ];
    sCfg.getRange(2, 1, defaults.length, 3).setValues(defaults);
  }

  let sClubs = ss.getSheetByName(SHEET_CLUBS);
  if (!sClubs) {
    sClubs = ss.insertSheet(SHEET_CLUBS);
    sClubs.getRange(1, 1, 1, 5).setValues([["Club","Referente","Email","Telefono","Note"]]);
    sClubs.setFrozenRows(1);
    sClubs.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#1e3348").setFontColor("#ffffff");
  }

  SpreadsheetApp.getUi().alert("Inizializzazione completata! Tre fogli creati: Prenotazioni, Config, Clubs.");
}

// ================================================================
// HANDLERS
// ================================================================
function handleBook(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_BOOKINGS);
  const ts = new Date();

  let totalEur = 0;
  const slotLines = [];

  data.slots.forEach(slot => {
    const price = Number(slot.price || 0);
    totalEur += price;
    sheet.appendRow([
      ts, data.club, data.referent || "", data.email || "", data.phone || "",
      slot.boat, slot.dayLabel, slot.sessLabel, slot.start, slot.end,
      slot.dk, slot.slotMin, "Confermato", "No", "", price
    ]);
    slotLines.push("• " + slot.boat + " – " + slot.dayLabel + " (" + slot.sessLabel + ") – "
      + slot.start + "–" + slot.end + " · €" + price.toFixed(2));
  });

  registerClubIfNew(data);

  const corpo =
    "Club: " + data.club + "\n" +
    "Contact: " + (data.referent || "-") + "\n" +
    "Email: " + (data.email || "-") + "\n" +
    "Phone: " + (data.phone || "-") + "\n\n" +
    "BOOKED SLOTS:\n" + slotLines.join("\n") + "\n\n" +
    "Total: €" + totalEur.toFixed(2) + "\n\n" +
    "PAYMENT\n" +
    "Beneficiary: " + BENEFICIARY + "\n" +
    "Bank: " + BANK + "\n" +
    "IBAN: " + IBAN + "\n" +
    "BIC: " + BIC + "\n" +
    "Reference: TRAINING RENTAL " + data.club + "\n\n" +
    "Please download, sign and send the waiver + proof of payment to:\n" +
    "rent@filippiboats.com within 24 hours.";

  // Email all'admin (TO + CC)
  GmailApp.sendEmail(
    ADMIN_EMAIL_PRIMARY,
    "Nuova prenotazione – " + data.club + " – Filippi Trophy Castagneto 2026",
    "Nuova prenotazione ricevuta.\n\n" + corpo,
    { name: "Filippi Trophy Booking", replyTo: data.email || ADMIN_EMAIL_PRIMARY, cc: ADMIN_EMAIL_CC }
  );

  // Email di conferma al club
  if (data.email) {
    try {
      GmailApp.sendEmail(
        data.email,
        "Conferma prenotazione – Filippi Trophy Castagneto 2026",
        "Gentile " + data.club + ",\n\nLa tua prenotazione e' stata registrata con successo.\n\n" + corpo +
        "\n\nGrazie!\nFilippi Trophy Booking - Castagneto Carducci 2026",
        { name: "Filippi Trophy Booking", replyTo: ADMIN_EMAIL_PRIMARY }
      );
    } catch(e) {
      Logger.log("Errore invio email cliente: " + e);
    }
  }

  return { ok: true, total: totalEur };
}

function handleCancel(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BOOKINGS);
  const rows = sheet.getDataRange().getValues();
  let cancelled = 0;
  for (let i = rows.length - 1; i > 0; i--) {
    const r = rows[i];
    if (r[1] === data.club && r[5] === data.boat &&
        String(r[10]) === String(data.dk) && Number(r[11]) === Number(data.slotMin)) {
      sheet.deleteRow(i + 1);
      cancelled++;
    }
  }
  if (cancelled > 0) {
    GmailApp.sendEmail(ADMIN_EMAIL_PRIMARY,
      "Cancellazione – " + data.club,
      data.club + " ha cancellato: " + data.boat + " – " + (data.start || "") + "–" + (data.end || ""),
      { name: "Filippi Trophy Booking", cc: ADMIN_EMAIL_CC });
  }
  return { ok: true, cancelled: cancelled };
}

function handleMarkPaid(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BOOKINGS);
  const rows = sheet.getDataRange().getValues();
  let marked = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r[1] === data.club && r[5] === data.boat &&
        String(r[10]) === String(data.dk) && Number(r[11]) === Number(data.slotMin)) {
      sheet.getRange(i + 1, 14).setValue(data.paid ? "Si" : "No");
      marked++;
    }
  }
  return { ok: true, marked: marked };
}

function handleMoveSlot(data) {
  return { ok: true, msg: "MoveSlot handled" };
}

function handleSummary(data) {
  GmailApp.sendEmail(ADMIN_EMAIL_PRIMARY, data.subject || "Riepilogo prenotazioni", data.message || "Nessuna prenotazione",
    { name: "Filippi Trophy Booking", cc: ADMIN_EMAIL_CC });
  return { ok: true };
}

function handleGetBookings(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BOOKINGS);
  const rows = sheet.getDataRange().getValues();
  const bookings = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    bookings.push({
      timestamp: r[0], club: r[1], referent: r[2], email: r[3], phone: r[4],
      boat: r[5], dayLabel: r[6], sessLabel: r[7], start: r[8], end: r[9],
      dk: r[10], slotMin: r[11], status: r[12], paid: r[13], notes: r[14],
      price: Number(r[15] || 0)
    });
  }
  return { ok: true, bookings: bookings };
}

function handleSetConfig(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.key) {
      sheet.getRange(i + 1, 2).setValue(data.value);
      return { ok: true, updated: true };
    }
  }
  sheet.appendRow([data.key, data.value, data.description || ""]);
  return { ok: true, created: true };
}

function handleRegisterClub(data) {
  registerClubIfNew(data);
  return { ok: true };
}

function registerClubIfNew(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CLUBS);
  if (!sheet) return;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.club) return;
  }
  sheet.appendRow([data.club, data.referent || "", data.email || "", data.phone || "", ""]);
}
