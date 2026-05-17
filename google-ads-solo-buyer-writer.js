const SPREADSHEET_ID = "1l7w0VqphJhQnESwZaH9ILQ7B2Ktqx7Qk0iazWGXcAw1";

const STATS_SHEET_NAME = "Buyer_Stats";
const API_VERSION = "v22";

const PERIOD_OPTIONS = [
  { id: "TODAY", label: "Today" },
  { id: "YESTERDAY", label: "Yesterday" },
  { id: "LAST_7_DAYS", label: "Last 7 days" },
  { id: "THIS_MONTH", label: "This month" },
  { id: "LAST_MONTH", label: "Last month" },
];

const HEADERS = [
  "runDate",
  "period",
  "periodLabel",
  "accountId",
  "accountName",
  "currency",
  "impressions",
  "clicks",
  "ctr",
  "conversions",
  "conversionRate",
  "averageCpc",
  "averageCpm",
  "cost",
  "costPerConversion",
];

function main() {
  const account = AdsApp.currentAccount();
  const meta = {
    runDate: getTodayDateString_(),
    accountId: account.getCustomerId(),
    accountName: safeGetAccountName_(account),
    currency: account.getCurrencyCode(),
  };

  const rows = PERIOD_OPTIONS.map((period) => getAccountStatsRow_(meta, period));

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet_(ss, STATS_SHEET_NAME);
  writeRows_(sheet, rows, meta.accountId, meta.runDate);

  Logger.log(
    "Updated " + STATS_SHEET_NAME +
      ". Account: " + meta.accountId +
      ", rows: " + rows.length
  );
}

function getAccountStatsRow_(meta, period) {
  const query = `
    SELECT
      metrics.impressions,
      metrics.clicks,
      metrics.conversions,
      metrics.cost_micros
    FROM campaign
    WHERE segments.date DURING ${period.id}
      AND campaign.status != 'REMOVED'
  `;

  const report = AdsApp.report(query, { apiVersion: API_VERSION });
  const it = report.rows();

  let impressions = 0;
  let clicks = 0;
  let conversions = 0;
  let cost = 0;

  while (it.hasNext()) {
    const r = it.next();
    impressions += toInt_(r["metrics.impressions"]);
    clicks += toInt_(r["metrics.clicks"]);
    conversions += toFloat_(r["metrics.conversions"]);
    cost += microsToMoney_(r["metrics.cost_micros"]);
  }

  return [
    meta.runDate,
    period.id,
    period.label,
    meta.accountId,
    meta.accountName,
    meta.currency,
    impressions,
    clicks,
    ratio_(clicks, impressions),
    conversions,
    ratio_(conversions, clicks),
    clicks > 0 ? cost / clicks : 0,
    impressions > 0 ? cost / impressions * 1000 : 0,
    cost,
    conversions > 0 ? cost / conversions : 0,
  ];
}

function writeRows_(sheet, newRows, accountId, runDate) {
  const lastRow = sheet.getLastRow();
  let existingData = [];

  if (lastRow > 1) {
    existingData = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  }

  const filteredData = existingData.filter((row) => {
    const rowRunDate = normalizeSheetDate_(row[0]);
    const rowAccountId = String(row[3]);
    return !(rowRunDate === runDate && rowAccountId === String(accountId));
  });

  const allData = filteredData.concat(newRows);

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
  }
  if (allData.length > 0) {
    sheet.getRange(2, 1, allData.length, HEADERS.length).setValues(allData);
  }

  formatSheet_(sheet);
}

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
  return sheet;
}

function formatSheet_(sheet) {
  sheet.autoResizeColumns(1, HEADERS.length);
  sheet.getRange("I:I").setNumberFormat("0.00%");
  sheet.getRange("K:K").setNumberFormat("0.00%");
  sheet.getRange("L:O").setNumberFormat("0.00");
}

function getTodayDateString_() {
  const now = new Date();
  const timeZone = AdsApp.currentAccount().getTimeZone();
  return Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
}

function safeGetAccountName_(account) {
  try {
    return typeof account.getName === "function" ? account.getName() : "";
  } catch (_) {
    return "";
  }
}

function normalizeSheetDate_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]") {
    const timeZone = AdsApp.currentAccount().getTimeZone();
    return Utilities.formatDate(value, timeZone, "yyyy-MM-dd");
  }
  return String(value).trim();
}

function normalizeNumberString_(value) {
  if (value === null || value === undefined) return "";
  const s = String(value).trim().replace(/[^\d.,-]/g, "");
  if (s.indexOf(".") >= 0 && s.indexOf(",") >= 0) return s.replace(/,/g, "");
  if (s.indexOf(",") >= 0 && s.indexOf(".") < 0) return s.replace(",", ".");
  return s;
}

function toInt_(value) {
  const n = parseInt(normalizeNumberString_(value), 10);
  return Number.isFinite(n) ? n : 0;
}

function toFloat_(value) {
  const n = parseFloat(normalizeNumberString_(value));
  return Number.isFinite(n) ? n : 0;
}

function microsToMoney_(value) {
  return toFloat_(value) / 1000000;
}

function ratio_(part, total) {
  return total > 0 ? part / total : 0;
}
