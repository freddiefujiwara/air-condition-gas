const SPREADSHEET_ID = "1RtJaJDjeTK61RYpIR2JS5Jv25RVUrEI3_PYnzAesbFY";
const SHEET_NAMES = {
  conditions: "conditions",
  status: "status",
  setting: "setting",
};

const SWITCHBOT = {
  acBaseUrl: "http://a.ze.gs/switchbot-ac",
  acDevicePath: "-d/02-202307290753-22894539",
  customBaseUrl: "http://a.ze.gs/switchbot-custom",
  customDevicePath: "-d/03-202401251013-58699638",
};

const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
const sheets = {
  conditions: spreadsheet.getSheetByName(SHEET_NAMES.conditions),
  status: spreadsheet.getSheetByName(SHEET_NAMES.status),
  setting: spreadsheet.getSheetByName(SHEET_NAMES.setting),
};

const buildAcUrl = (command) =>
  `${SWITCHBOT.acBaseUrl}/${SWITCHBOT.acDevicePath}/-a/${command}`;

const buildCustomUrl = (command, repeat = 1) => {
  const segments = Array.from({ length: repeat }, () =>
    `${SWITCHBOT.customDevicePath}/-c/${command}`
  );
  return `${SWITCHBOT.customBaseUrl}/${segments.join("/")}`;
};

const getConditionsLastRow = () => sheets.conditions.getLastRow();

const appendConditionRow = ({ temperature, humidity }) => {
  const rowIndex = getConditionsLastRow() + 1;
  sheets.conditions.getRange(rowIndex, 1).setValue(new Date());
  sheets.conditions.getRange(rowIndex, 2).setValue(temperature ?? 0);
  sheets.conditions.getRange(rowIndex, 3).setValue(humidity ?? 0);
};

const setStatus = (nextStatus) => {
  sheets.status.getRange(1, 1).setValue(nextStatus);
};

const getStatus = () => sheets.status.getRange(1, 1).getValue();

const getSettings = () => {
  const values = sheets.setting.getDataRange().getValues();
  const headers = values.shift() ?? [];
  const row = values.shift() ?? [];

  return headers.reduce((acc, header, index) => {
    acc[header] = row[index];
    return acc;
  }, {});
};

const formatJapanTime = (value) => {
  const dt = new Date(
    new Date(value).toLocaleString("en-US", { timeZone: "Asia/Tokyo" })
  );
  return `${dt.toLocaleDateString()} ${dt.toLocaleTimeString()}`;
};

const serializeConditions = (values) => {
  const headers = values.shift() ?? [];
  return values.map((row) =>
    row.reduce((acc, column, index) => {
      const header = headers[index];
      if (column instanceof Date) {
        acc[header] = formatJapanTime(column);
      } else {
        acc[header] = column;
      }
      return acc;
    }, {})
  );
};

const buildJsonpResponse = (callback, data) =>
  `${callback}&&${callback}(${JSON.stringify(data)});`;

const AC_ACTIONS = [
  {
    name: "hot_on",
    shouldRun: ({ temperature, statusValue, settings }) =>
      temperature <= settings.hot_buttom && statusValue === "off",
    acCommand: ({ settings }) => `${settings.hot},5,1,on`,
    customCommand: "Hot",
    customRepeat: 2,
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
      temperature >= settings.cool_up && statusValue === "off",
    acCommand: ({ settings }) => `${settings.cool},2,1,on`,
    customCommand: "Cool",
    nextStatus: "cool",
  },
];

function doGet(e) {
  const output = ContentService.createTextOutput();
  const params = e?.parameter ?? {};

  if (params.callback === undefined) {
    output.setMimeType(ContentService.MimeType.TEXT);
    if (params.s === "status") {
      setStatus(params.t);
      output.setContent("OK");
      return;
    }
    appendConditionRow({
      temperature: params.t,
      humidity: params.h,
    });
    output.setContent("OK");
    return output;
  }

  output.setMimeType(ContentService.MimeType.JAVASCRIPT);
  const values = sheets.conditions.getDataRange().getValues();
  const result = serializeConditions(values);
  output.setContent(buildJsonpResponse(params.callback, result));
  return output;
}

function turnOnAC() {
  const temperature = Number(
    sheets.conditions.getRange(getConditionsLastRow(), 2).getValue()
  );
  const settings = getSettings();
  const statusValue = getStatus();

  const action = AC_ACTIONS.find((candidate) =>
    candidate.shouldRun({ temperature, statusValue, settings })
  );

  if (action) {
    UrlFetchApp.fetch(buildAcUrl(action.acCommand({ settings })));
    UrlFetchApp.fetch(
      buildCustomUrl(action.customCommand, action.customRepeat ?? 1)
    );
    setStatus(action.nextStatus);
  }

  Logger.log(settings);
  Logger.log(temperature);
}
