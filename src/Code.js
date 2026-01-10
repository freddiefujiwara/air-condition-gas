const SPREADSHEET_ID = "1RtJaJDjeTK61RYpIR2JS5Jv25RVUrEI3_PYnzAesbFY";
const sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
const conditions = sheet.getSheetByName("conditions");
const status = sheet.getSheetByName("status");
const setting = sheet.getSheetByName("setting");
const conditionsLastRow = conditions.getLastRow();

function doGet(e) {
  const output = ContentService.createTextOutput();

  if (e.parameter.callback === undefined) {
    output.setMimeType(ContentService.MimeType.TEXT);
    if ("status" === e.parameter.s) {
      status.getRange(1, 1).setValue(e.parameter.t);
      output.setContent("OK");
      return;
    }
    conditions.getRange(conditionsLastRow + 1, 1).setValue(new Date()); // Datetime
    conditions.getRange(conditionsLastRow + 1, 2).setValue(e.parameter.t || 0); // Temperature
    conditions.getRange(conditionsLastRow + 1, 3).setValue(e.parameter.h || 0); // Humidity
    output.setContent("OK");
  } else {
    output.setMimeType(ContentService.MimeType.JAVASCRIPT);
    const values = conditions.getDataRange().getValues();
    const headers = values.shift();
    const result = values.map((row) => {
      let data = {};
      row.map((column, index) => {
        if ("Datetime" === headers[index]) {
          const dt = new Date(new Date(column).toLocaleString("en-US", {
            timeZone: "Asia/Tokyo"
          }));
          data[headers[index]] = `${dt.toLocaleDateString()} ${dt.toLocaleTimeString()}`;
        } else {
          data[headers[index]] = column;
        }
      });
      return data;
    });
    output.setContent(e.parameter.callback + "&&" + e.parameter.callback + "(" + JSON.stringify(result) + ");");
  }
  return output;
}

function turnOnAC() {
  const temperature = 1.0 * conditions.getRange(conditionsLastRow, 2).getValue();

  const values = setting.getDataRange().getValues();
  const headers = values.shift();
  const settings = values.map((row) => {
    let data = {};
    row.map((column, index) => {
      data[headers[index]] = column;
    });
    return data;
  }).shift();
  const stat = status.getRange(1, 1).getValue();
  if (temperature <= settings["hot_buttom"] && "off" == stat) {
    UrlFetchApp.fetch(`http://a.ze.gs/switchbot-ac/-d/02-202307290753-22894539/-a/${settings["hot"]},5,1,on`);
    UrlFetchApp.fetch(`http://a.ze.gs/switchbot-custom/-d/03-202401251013-58699638/-c/Hot/-d/03-202401251013-58699638/-c/Hot`);
    status.getRange(1, 1).setValue("hot");
  } else if (temperature >= settings["hot_up"] && "hot" == stat) {
    UrlFetchApp.fetch(`http://a.ze.gs/switchbot-ac/-d/02-202307290753-22894539/-a/25,5,1,off`);
    UrlFetchApp.fetch(`http://a.ze.gs/switchbot-custom/-d/03-202401251013-58699638/-c/Off`);
    status.getRange(1, 1).setValue("off");
  } else if (temperature <= settings["cool_buttom"] && "cool" == stat) {
    UrlFetchApp.fetch(`http://a.ze.gs/switchbot-ac/-d/02-202307290753-22894539/-a/25,2,1,off`);
    UrlFetchApp.fetch(`http://a.ze.gs/switchbot-custom/-d/03-202401251013-58699638/-c/Off`);
    status.getRange(1, 1).setValue("off");
  } else if (temperature >= settings["cool_up"] && "off" == stat) {
    UrlFetchApp.fetch(`http://a.ze.gs/switchbot-ac/-d/02-202307290753-22894539/-a/${settings["cool"]},2,1,on`);
    UrlFetchApp.fetch(`http://a.ze.gs/switchbot-custom/-d/03-202401251013-58699638/-c/Cool`);
    status.getRange(1, 1).setValue("cool");
  }
  Logger.log(settings);
  Logger.log(temperature);
}
