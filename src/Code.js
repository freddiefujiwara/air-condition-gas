const SPREADSHEET_ID = "1RtJaJDjeTK61RYpIR2JS5Jv25RVUrEI3_PYnzAesbFY";
const SHEET_NAMES = {
  conditions: "conditions",
  status: "status",
  setting: "setting",
};

// Home server details (base URLs and device IDs) live in script
// properties, not in this public source code:
//   SWITCHBOT_AC_URL     e.g. <base>/switchbot-ac/-d/<device id>
//   SWITCHBOT_CUSTOM_URL e.g. <base>/switchbot-custom/-d/<device id>
const getProperty_ = (name) => {
  const value = PropertiesService.getScriptProperties().getProperty(name);
  if (!value) {
    throw new Error(`script property ${name} is not set`);
  }
  return value;
};

const CONDITIONS_COLUMNS = {
  date: 1,
  temperature: 2,
  humidity: 3,
};

const STATUS_CELL = {
  row: 1,
  column: 1,
};

const getSheets = () => {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    conditions: spreadsheet.getSheetByName(SHEET_NAMES.conditions),
    status: spreadsheet.getSheetByName(SHEET_NAMES.status),
    setting: spreadsheet.getSheetByName(SHEET_NAMES.setting),
  };
};

const buildAcUrl = (command) =>
  `${getProperty_("SWITCHBOT_AC_URL")}/-a/${command}`;

const buildCustomUrl = (command) =>
  `${getProperty_("SWITCHBOT_CUSTOM_URL")}/-c/${command}`;

const getConditionsLastRow = (sheets = getSheets()) =>
  sheets.conditions.getLastRow();

const appendConditionRow = ({ temperature, humidity }, sheets = getSheets()) => {
  const rowIndex = getConditionsLastRow(sheets) + 1;
  sheets.conditions
    .getRange(rowIndex, CONDITIONS_COLUMNS.date)
    .setValue(new Date());
  sheets.conditions
    .getRange(rowIndex, CONDITIONS_COLUMNS.temperature)
    .setValue(temperature ?? 0);
  sheets.conditions
    .getRange(rowIndex, CONDITIONS_COLUMNS.humidity)
    .setValue(humidity ?? 0);
};

const setStatus = (nextStatus, sheets = getSheets()) => {
  sheets.status
    .getRange(STATUS_CELL.row, STATUS_CELL.column)
    .setValue(nextStatus);
};

const getStatus = (sheets = getSheets()) =>
  sheets.status.getRange(STATUS_CELL.row, STATUS_CELL.column).getValue();

const getSettings = (sheets = getSheets()) => {
  const values = sheets.setting.getDataRange().getValues();
  const headers = values.shift() ?? [];
  const row = values.shift() ?? [];

  return headers.reduce((acc, header, index) => {
    acc[header] = row[index];
    return acc;
  }, {});
};

const formatJapanTime = (value) => {
  let dt;
  try {
    dt = new Date(
      new Date(value).toLocaleString("en-US", { timeZone: "Asia/Tokyo" })
    );
  } catch (ignored) {
    // A cell value that cannot even be converted (an object without a
    // usable toString, a symbol, ...) must not kill the whole payload;
    // the Date field simply reads empty.
    return "";
  }
  return `${dt.toLocaleDateString()} ${dt.toLocaleTimeString()}`;
};

const serializeConditions = (headers, rows) =>
  rows.map((row) =>
    row.reduce((acc, column, index) => {
      const header = headers[index];
      if (header === "Date") {
        acc[header] = formatJapanTime(column);
      } else {
        acc[header] = column;
      }
      return acc;
    }, {})
  );

const buildJsonpResponse = (callback, data) =>
  `${callback}&&${callback}(${JSON.stringify(data)});`;

// The write path (no callback) changes the spreadsheet that turnOnAC
// acts on, so it requires the shared token stored in the WRITE_TOKEN
// script property. Callers (the home server ingest and the ha
// frontend) send it as the `token` parameter. Fail closed: when the
// property is not set, every write is rejected.
const isAuthorizedWrite = (params) => {
  const expected = PropertiesService.getScriptProperties().getProperty(
    "WRITE_TOKEN"
  );
  return Boolean(expected) && params.token === expected;
};

// Same callback shape accepted by the other freddiefujiwara.com GAS
// endpoints: a plain identifier path, nothing else.
const CALLBACK_PATTERN =
  /^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/;

const isValidCallback = (callback) =>
  typeof callback === "string" && CALLBACK_PATTERN.test(callback);

const getRowValues = (sheet, row, lastColumn) =>
  sheet.getRange(row, 1, 1, lastColumn).getValues()[0];

const buildConditionsPayload = (sheets) => {
  const lastRow = getConditionsLastRow(sheets);
  const conditionsSheet = sheets.conditions;
  const lastColumn = conditionsSheet.getLastColumn();
  const headers = getRowValues(conditionsSheet, 1, lastColumn);
  const rowValues = getRowValues(conditionsSheet, lastRow, lastColumn);

  return {
    conditions: serializeConditions(headers, [rowValues]),
    status: getStatus(sheets),
    setting: getSettings(sheets),
  };
};

// Sheet cells are outside input: only a finite number or a non-blank
// numeric string counts as a value. Anything else (blank, boolean,
// Date, array, object) is no value at all — a bare Number() would
// invent one (true -> 1, a Date -> its epoch milliseconds, "" -> 0)
// and the automation would act on the invented number.
const numericCellValue_ = (value) => {
  const parsed =
    typeof value === "number"
      ? value
      : typeof value === "string" && value.trim() !== ""
        ? Number(value)
        : NaN;
  return Number.isFinite(parsed) ? parsed : NaN;
};

const getLatestTemperature = (sheets) =>
  numericCellValue_(
    sheets.conditions
      .getRange(
        getConditionsLastRow(sheets),
        CONDITIONS_COLUMNS.temperature
      )
      .getValue()
  );

const AC_ACTIONS = [
  {
    name: "hot_on",
    shouldRun: ({ temperature, statusValue, settings }) =>
      temperature <= settings.hot_buttom &&
      statusValue === "off" &&
      Number.isFinite(numericCellValue_(settings.hot)),
    acCommand: ({ settings }) => `${settings.hot},5,1,on`,
    customCommand: "Hot",
    nextStatus: "hot",
  },
  {
    name: "hot_off",
    shouldRun: ({ temperature, statusValue, settings }) =>
      temperature >= settings.hot_up && statusValue === "hot",
    acCommand: () => "25,5,1,off",
    customCommand: "Off",
    nextStatus: "off",
  },
  {
    name: "cool_off",
    shouldRun: ({ temperature, statusValue, settings }) =>
      temperature <= settings.cool_buttom && statusValue === "cool",
    acCommand: () => "25,2,1,off",
    customCommand: "Off",
    nextStatus: "off",
  },
  {
    name: "cool_on",
    shouldRun: ({ temperature, statusValue, settings }) =>
      temperature >= settings.cool_up &&
      statusValue === "off" &&
      Number.isFinite(numericCellValue_(settings.cool)),
    acCommand: ({ settings }) => `${settings.cool},2,1,on`,
    customCommand: "Cool",
    nextStatus: "cool",
  },
];

// The status cell is the automation's state: turnOnAC only acts on
// "off" / "hot" / "cool", and the ha frontend's fourth value
// "control" parks the A/C under manual control. Anything else stored
// here would silently disable the automation and be displayed by ha
// as the A/C state, so an unknown status is rejected like a bad
// token and the stored state is left untouched.
const KNOWN_STATUSES = ["off", "hot", "cool", "control"];

export function doGet(e) {
  const output = ContentService.createTextOutput();
  const params = e?.parameter ?? {};

  if (params.callback === undefined) {
    output.setMimeType(ContentService.MimeType.TEXT);
    if (!isAuthorizedWrite(params)) {
      output.setContent("NG");
      return output;
    }
    const sheets = getSheets();
    if (params.s === "status") {
      if (!KNOWN_STATUSES.includes(params.t)) {
        output.setContent("NG");
        return output;
      }
      setStatus(params.t, sheets);
    } else {
      appendConditionRow({
        temperature: params.t,
        humidity: params.h,
      }, sheets);
    }
    output.setContent("OK");
    return output;
  }

  const sheets = getSheets();
  const payload = buildConditionsPayload(sheets);
  if (!isValidCallback(params.callback)) {
    // Never reflect an untrusted callback into a JavaScript response;
    // fall back to plain JSON, which carries the same payload.
    output.setMimeType(ContentService.MimeType.JSON);
    output.setContent(JSON.stringify(payload));
    return output;
  }

  output.setMimeType(ContentService.MimeType.JAVASCRIPT);
  output.setContent(buildJsonpResponse(params.callback, payload));
  return output;
}

export function turnOnAC() {
  const sheets = getSheets();
  const temperature = getLatestTemperature(sheets);
  const settings = getSettings(sheets);
  const statusValue = getStatus(sheets);

  const action = AC_ACTIONS.find((candidate) =>
    candidate.shouldRun({ temperature, statusValue, settings })
  );

  if (action) {
    const options = {
      headers: { "X-Api-Token": getProperty_("HOME_API_TOKEN") },
    };
    UrlFetchApp.fetch(buildAcUrl(action.acCommand({ settings })), options);
    UrlFetchApp.fetch(buildCustomUrl(action.customCommand), options);
    setStatus(action.nextStatus, sheets);
  }

  Logger.log(settings);
  Logger.log(temperature);
}
