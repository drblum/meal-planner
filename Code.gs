const RECIPIENT_EMAIL = "daniel.jin.blum@gmail.com";
const DINNER_START_HOUR = 19; // 7 PM
const DINNER_DURATION_MINS = 60;
const TIMEZONE = "Asia/Shanghai";

const SPREADSHEET_ID = "";
const CALENDAR_NAME = "Family Meals";
const EVENT_MARKER = "FamilyMealPlanner";

const SHEETS = {
  dishes: "Dishes",
  weekPlan: "WeekPlan",
  history: "History"
};

const HEADERS = {
  Dishes: ["id", "name", "cuisine", "type", "lastServed", "notes", "tags", "rating"],
  WeekPlan: ["weekStartDate", "weekData"],
  History: ["weekStartDate", "weekLabel", "anchorCuisine", "weekData"]
};

const CUISINE_EMOJI = {
  Korean: "🇰🇷",
  Japanese: "🇯🇵",
  Italian: "🇮🇹",
  American: "🇺🇸",
  Chinese: "🇨🇳",
  Mediterranean: "🌿",
  "Southeast Asian": "🌶️",
  Other: "🍽️"
};

function setup() {
  const ss = getSpreadsheet_();
  Object.keys(SHEETS).forEach(function(key) {
    const name = SHEETS[key];
    const sheet = getOrCreateSheet_(ss, name);
    ensureHeaders_(sheet, HEADERS[name]);
  });
  getOrCreateCalendar_();
  return ok_({ message: "Setup complete" });
}

function doGet(e) {
  try {
    const action = getParam_(e, "action");
    if (action === "getDishes") return ok_(getDishes_());
    if (action === "getWeekPlan") return ok_(getWeekPlan_(getParam_(e, "weekStart")));
    if (action === "getHistory") return ok_(getHistory_());
    if (action === "ping") return ok_({ ok: true, now: nowIso_() });
    return error_("Unknown GET action: " + action, 400);
  } catch (err) {
    return error_(err.message || String(err), 500);
  }
}

function doPost(e) {
  try {
    const body = parseBody_(e);
    const action = body.action || getParam_(e, "action");
    if (action === "saveDish") return ok_(saveDish_(body.dish || body));
    if (action === "deleteDish") return ok_(deleteDish_(body.id));
    if (action === "saveWeekPlan") return ok_(saveWeekPlan_(body));
    if (action === "archiveWeek") return ok_(archiveWeek_(body));
    if (action === "sendWeeklyEmail") return ok_(sendWeeklyEmail_(body));
    return error_("Unknown POST action: " + action, 400);
  } catch (err) {
    return error_(err.message || String(err), 500);
  }
}

function getDishes_() {
  const sheet = getSheet_(SHEETS.dishes);
  return rowsToObjects_(sheet).map(function(row) {
    row.id = String(row.id || "").trim();
    row.name = String(row.name || "").trim();
    row.cuisine = row.cuisine || "Other";
    row.type = row.type || "main";
    row.lastServed = normalizeDateValue_(row.lastServed);
    row.notes = row.notes || "";
    row.tags = row.tags || "";
    row.rating = row.rating || "";
    return row;
  }).filter(function(row) {
    return row.id || row.name;
  });
}

function getWeekPlan_(weekStart) {
  if (!weekStart) throw new Error("weekStart is required");
  const sheet = getSheet_(SHEETS.weekPlan);
  const found = findRowByKey_(sheet, 1, weekStart);
  if (!found) return null;
  return {
    weekStartDate: weekStart,
    weekData: parseJsonSafe_(sheet.getRange(found.row, 2).getValue()) || null
  };
}

function getHistory_() {
  const rows = rowsToObjects_(getSheet_(SHEETS.history)).map(function(row) {
    return {
      weekStartDate: normalizeDateValue_(row.weekStartDate),
      weekLabel: row.weekLabel || "",
      anchorCuisine: row.anchorCuisine || "",
      weekData: parseJsonSafe_(row.weekData) || null
    };
  }).filter(function(row) {
    return row.weekStartDate;
  });
  rows.sort(function(a, b) {
    return b.weekStartDate.localeCompare(a.weekStartDate);
  });
  return rows;
}

function saveDish_(dish) {
  if (!dish) throw new Error("dish is required");
  dish.id = String(dish.id || makeId_()).trim();
  if (!dish.name) throw new Error("Dish name is required");
  dish.cuisine = dish.cuisine || "Other";
  dish.type = dish.type || "main";
  const sheet = getSheet_(SHEETS.dishes);
  const values = [
    dish.id,
    dish.name,
    dish.cuisine,
    dish.type,
    dish.lastServed || "",
    dish.notes || "",
    dish.tags || "",
    dish.rating || ""
  ];
  const found = findRowByKey_(sheet, 1, dish.id);
  if (found) {
    sheet.getRange(found.row, 1, 1, values.length).setValues([values]);
  } else {
    sheet.appendRow(values);
  }
  return { dish: objectFromHeaders_(HEADERS.Dishes, values) };
}

function deleteDish_(id) {
  if (!id) throw new Error("id is required");
  const sheet = getSheet_(SHEETS.dishes);
  const found = findRowByKey_(sheet, 1, id);
  if (found) sheet.deleteRow(found.row);
  return { deleted: !!found, id: id };
}

function saveWeekPlan_(body) {
  const weekStartDate = body.weekStartDate || body.weekStart;
  const weekData = body.weekData || body;
  if (!weekStartDate) throw new Error("weekStartDate is required");
  if (!weekData || !weekData.plan) throw new Error("weekData.plan is required");

  upsertWeekPlan_(weekStartDate, weekData);
  syncCalendar_(weekStartDate, weekData);
  updateDishLastServed_(weekData.plan);
  archiveWeek_({
    weekStartDate: weekStartDate,
    weekLabel: makeWeekLabel_(weekStartDate),
    anchorCuisine: weekData.anchorCuisine || "",
    weekData: weekData
  });
  return { saved: true, weekStartDate: weekStartDate };
}

function archiveWeek_(body) {
  const weekStartDate = body.weekStartDate || body.weekStart;
  const weekData = body.weekData || body;
  if (!weekStartDate) throw new Error("weekStartDate is required");
  if (!weekData) throw new Error("weekData is required");
  const sheet = getSheet_(SHEETS.history);
  const row = [
    weekStartDate,
    body.weekLabel || makeWeekLabel_(weekStartDate),
    body.anchorCuisine || weekData.anchorCuisine || "",
    JSON.stringify(weekData)
  ];
  upsertByFirstColumn_(sheet, weekStartDate, row);
  return { archived: true, weekStartDate: weekStartDate };
}

function sendWeeklyEmail_(body) {
  const weekStartDate = body.weekStartDate || body.weekStart;
  const weekData = body.weekData || body;
  const recipient = body.recipient || RECIPIENT_EMAIL;
  if (!weekStartDate) throw new Error("weekStartDate is required");
  if (!weekData || !weekData.plan) throw new Error("weekData.plan is required");
  if (!recipient || recipient === "your@email.com") {
    throw new Error("Set RECIPIENT_EMAIL at the top of Code.gs before sending email");
  }

  const days = dayNames_();
  const rows = days.map(function(day, index) {
    const date = addDays_(parseDate_(weekStartDate), index);
    const meal = weekData.plan[day] || {};
    const main = getMealMain_(meal);
    const cuisine = getMealCuisine_(meal);
    const emoji = meal.travel ? "✈️" : (main ? cuisineEmoji_(cuisine) : "");
    const dishName = meal.travel ? "Travel" : meal.blank ? "Blank" : main ? main.name : "";
    const sides = (meal.sides || []).map(function(side) { return side.name || side; }).join(", ");
    return "<tr>" +
      "<td>" + escapeHtml_(day) + "</td>" +
      "<td>" + escapeHtml_(formatDate_(date)) + "</td>" +
      "<td>" + escapeHtml_(emoji + " " + cuisine) + "</td>" +
      "<td><strong>" + escapeHtml_(dishName || "—") + "</strong></td>" +
      "<td>" + escapeHtml_(sides || "—") + "</td>" +
      "</tr>";
  }).join("");

  const subject = "Family Meals: " + makeWeekLabel_(weekStartDate);
  const htmlBody = "<div style=\"font-family:Arial,sans-serif;color:#1a1a1a\">" +
    "<h2 style=\"margin:0 0 8px\">Family Meals</h2>" +
    "<p style=\"margin:0 0 18px;color:#666\">" + escapeHtml_(makeWeekLabel_(weekStartDate)) + "</p>" +
    "<table cellpadding=\"8\" cellspacing=\"0\" style=\"border-collapse:collapse;width:100%;max-width:760px\">" +
    "<thead><tr style=\"background:#f5f5f3;text-align:left\">" +
    "<th>Day</th><th>Date</th><th>Cuisine</th><th>Dish</th><th>Sides</th>" +
    "</tr></thead><tbody>" + rows + "</tbody></table></div>";

  GmailApp.sendEmail(recipient, subject, stripHtml_(htmlBody), { htmlBody: htmlBody });
  return { sent: true, recipient: recipient };
}

function upsertWeekPlan_(weekStartDate, weekData) {
  const sheet = getSheet_(SHEETS.weekPlan);
  upsertByFirstColumn_(sheet, weekStartDate, [weekStartDate, JSON.stringify(weekData)]);
}

function syncCalendar_(weekStartDate, weekData) {
  const calendar = getOrCreateCalendar_();
  const base = parseDate_(weekStartDate);
  const days = dayNames_();
  days.forEach(function(day, index) {
    const date = addDays_(base, index);
    const meal = (weekData.plan && weekData.plan[day]) || {};
    if (meal.blank || meal.travel || !getMealMain_(meal)) {
      deleteMealEvents_(calendar, date);
      return;
    }
    upsertMealEvent_(calendar, date, day, meal);
  });
}

function upsertMealEvent_(calendar, date, day, meal) {
  const main = getMealMain_(meal);
  const cuisine = getMealCuisine_(meal);
  const title = cuisineEmoji_(cuisine) + " Dinner: " + main.name;
  const sideNames = (meal.sides || []).map(function(side) { return side.name || side; });
  const marker = eventMarkerForDate_(date);
  const description = [
    marker,
    "Cuisine: " + cuisine,
    sideNames.length ? "Sides: " + sideNames.join(", ") : "Sides: —",
    meal.note ? "Note: " + meal.note : ""
  ].filter(Boolean).join("\n");

  const start = new Date(date.getFullYear(), date.getMonth(), date.getDate(), DINNER_START_HOUR, 0, 0);
  const end = new Date(start.getTime() + DINNER_DURATION_MINS * 60000);
  const existing = findMealEvents_(calendar, date);
  if (existing.length) {
    const event = existing[0];
    event.setTitle(title);
    event.setTime(start, end);
    event.setDescription(description);
    existing.slice(1).forEach(function(extra) { extra.deleteEvent(); });
  } else {
    calendar.createEvent(title, start, end, { description: description });
  }
}

function deleteMealEvents_(calendar, date) {
  findMealEvents_(calendar, date).forEach(function(event) {
    event.deleteEvent();
  });
}

function findMealEvents_(calendar, date) {
  const start = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0);
  const end = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59);
  const marker = eventMarkerForDate_(date);
  return calendar.getEvents(start, end).filter(function(event) {
    return String(event.getDescription() || "").indexOf(marker) !== -1;
  });
}

function updateDishLastServed_(plan) {
  const sheet = getSheet_(SHEETS.dishes);
  const data = sheet.getDataRange().getValues();
  const idToRow = {};
  for (var i = 1; i < data.length; i++) {
    idToRow[String(data[i][0])] = i + 1;
  }
  Object.keys(plan || {}).forEach(function(day) {
    const meal = plan[day] || {};
    if (meal.blank || meal.travel) return;
    const main = getMealMain_(meal);
    if (!main || !main.id || !meal.date) return;
    const row = idToRow[String(main.id)];
    if (row) sheet.getRange(row, 5).setValue(meal.date);
  });
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error("No active spreadsheet. Bind this script to a Google Sheet or set SPREADSHEET_ID.");
  return active;
}

function getSheet_(name) {
  const sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) throw new Error("Missing sheet: " + name + ". Run setup() first.");
  ensureHeaders_(sheet, HEADERS[name]);
  return sheet;
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureHeaders_(sheet, headers) {
  const range = sheet.getRange(1, 1, 1, headers.length);
  const existing = range.getValues()[0];
  const needsHeaders = existing.join("") === "" || headers.some(function(header, index) {
    return existing[index] !== header;
  });
  if (needsHeaders) {
    range.setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function rowsToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(function(header) { return String(header); });
  return values.slice(1).filter(function(row) {
    return row.some(function(cell) { return cell !== ""; });
  }).map(function(row) {
    return objectFromHeaders_(headers, row);
  });
}

function objectFromHeaders_(headers, row) {
  const obj = {};
  headers.forEach(function(header, index) {
    obj[header] = row[index];
  });
  return obj;
}

function findRowByKey_(sheet, column, key) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, column, lastRow - 1, 1).getValues();
  const wanted = String(key);
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === wanted) return { row: i + 2 };
  }
  return null;
}

function upsertByFirstColumn_(sheet, key, row) {
  const found = findRowByKey_(sheet, 1, key);
  if (found) {
    sheet.getRange(found.row, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function getOrCreateCalendar_() {
  const calendars = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  return calendars.length ? calendars[0] : CalendarApp.createCalendar(CALENDAR_NAME, { timeZone: TIMEZONE });
}

function parseBody_(e) {
  if (!e || !e.postData || !e.postData.contents) return {};
  return JSON.parse(e.postData.contents);
}

function getParam_(e, name) {
  return e && e.parameter ? e.parameter[name] : "";
}

function ok_(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function error_(message, status) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: message, status: status || 500 }))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseJsonSafe_(value) {
  if (!value) return null;
  if (typeof value === "object") return value;
  try {
    return JSON.parse(String(value));
  } catch (err) {
    return null;
  }
}

function getMealMain_(meal) {
  if (!meal) return null;
  if (meal.main && typeof meal.main === "object") return meal.main;
  if (meal.main) return { name: String(meal.main), cuisine: meal.cuisine || "Other" };
  return null;
}

function getMealCuisine_(meal) {
  const main = getMealMain_(meal);
  return (main && main.cuisine) || meal.cuisine || "Other";
}

function cuisineEmoji_(cuisine) {
  return CUISINE_EMOJI[cuisine] || CUISINE_EMOJI.Other;
}

function eventMarkerForDate_(date) {
  return EVENT_MARKER + ":" + formatDate_(date);
}

function dayNames_() {
  return ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
}

function parseDate_(value) {
  const parts = String(value).split("-").map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function addDays_(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function formatDate_(date) {
  return Utilities.formatDate(date, TIMEZONE, "yyyy-MM-dd");
}

function normalizeDateValue_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") return formatDate_(value);
  return String(value);
}

function makeWeekLabel_(weekStartDate) {
  const start = parseDate_(weekStartDate);
  const end = addDays_(start, 6);
  return Utilities.formatDate(start, TIMEZONE, "MMM d") + " - " + Utilities.formatDate(end, TIMEZONE, "MMM d, yyyy");
}

function makeId_() {
  return "dish_" + Utilities.getUuid().replace(/-/g, "").slice(0, 18);
}

function nowIso_() {
  return Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
}

function escapeHtml_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function stripHtml_(html) {
  return String(html || "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
}
