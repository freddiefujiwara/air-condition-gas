import fs from "fs";
import path from "path";
import vm from "vm";
import { describe, expect, it, vi } from "vitest";

const codePath = path.resolve(process.cwd(), "src/Code.js");

const loadCode = ({
  conditionsLastRow = 1,
  temperature = 0,
  conditionsValues = null,
  conditionsHeaders = ["Row"],
  conditionsRowValues = [1],
  settingsValues = null,
  statusValue = "off",
} = {}) => {
  const outputState = { content: "", mimeType: "" };
  const output = {
    setMimeType: vi.fn((mime) => {
      outputState.mimeType = mime;
    }),
    setContent: vi.fn((content) => {
      outputState.content = content;
    }),
    getContent: () => outputState.content,
    getMimeType: () => outputState.mimeType,
  };

  const conditionsRangeCalls = [];
  const conditionsRanges = new Map();
  const makeRangeKey = (row, col, numRows = 1, numCols = 1) =>
    `${row}:${col}:${numRows}:${numCols}`;
  const conditionsSheet = {
    getLastRow: vi.fn(() => conditionsLastRow),
    getLastColumn: vi.fn(() => conditionsHeaders.length),
    getRange: vi.fn((row, col, numRows = 1, numCols = 1) => {
      const key = makeRangeKey(row, col, numRows, numCols);
      conditionsRangeCalls.push([row, col, numRows, numCols]);
      if (!conditionsRanges.has(key)) {
        const range = {
          setValue: vi.fn(),
          getValues: vi.fn(() => {
            if (row === 1) {
              return [conditionsHeaders];
            }
            if (row === conditionsLastRow) {
              return [conditionsRowValues];
            }
            return [[]];
          }),
          getValue: vi.fn(() => {
            if (row === conditionsLastRow && col === 2) {
              return temperature;
            }
            return undefined;
          }),
        };
        conditionsRanges.set(key, range);
      }
      return conditionsRanges.get(key);
    }),
    getDataRange: vi.fn(() => ({
      getValues: vi.fn(() => conditionsValues ?? []),
    })),
  };

  const statusRange = {
    getValue: vi.fn(() => statusValue),
    setValue: vi.fn(),
  };
  const statusSheet = {
    getRange: vi.fn(() => statusRange),
    getDataRange: vi.fn(() => ({
      getValues: vi.fn(() => [[statusValue]]),
    })),
  };

  const settingSheet = {
    getDataRange: vi.fn(() => ({
      getValues: vi.fn(() => settingsValues ?? []),
    })),
  };

  const SpreadsheetApp = {
    openById: vi.fn(() => ({
      getSheetByName: vi.fn((name) => {
        if (name === "conditions") {
          return conditionsSheet;
        }
        if (name === "status") {
          return statusSheet;
        }
        return settingSheet;
      }),
    })),
  };

  const ContentService = {
    MimeType: {
      TEXT: "text/plain",
      JAVASCRIPT: "application/javascript",
    },
    createTextOutput: vi.fn(() => output),
  };

  const UrlFetchApp = {
    fetch: vi.fn(),
  };

  const Logger = {
    log: vi.fn(),
  };

  const sandbox = {
    SpreadsheetApp,
    ContentService,
    UrlFetchApp,
    Logger,
  };

  const context = vm.createContext(sandbox);
  const code = fs.readFileSync(codePath, "utf8");
  vm.runInContext(code, context, { filename: "Code.js" });

  return {
    context,
    output,
    outputState,
    conditionsSheet,
    statusRange,
    settingSheet,
    UrlFetchApp,
    Logger,
    conditionsRangeCalls,
  };
};

const buildSettings = (overrides = {}) => {
  const defaults = {
    hot_buttom: 20,
    hot_up: 28,
    cool_buttom: 18,
    cool_up: 26,
    hot: 22,
    cool: 24,
  };
  return { ...defaults, ...overrides };
};

const settingsValuesFrom = (settings) => {
  const headers = Object.keys(settings);
  const values = headers.map((key) => settings[key]);
  return [headers, values];
};

describe("doGet", () => {
  it("writes status and exits when status param is provided", () => {
    const { context, statusRange, outputState, output } = loadCode();

    const result = context.doGet({
      parameter: {
        s: "status",
        t: "hot",
      },
    });

    expect(statusRange.setValue).toHaveBeenCalledWith("hot");
    expect(outputState.content).toBe("OK");
    expect(result).toBe(output);
  });

  it("appends condition data when no callback is set", () => {
    const { context, conditionsSheet, outputState } = loadCode({
      conditionsLastRow: 2,
    });

    const result = context.doGet({
      parameter: {
        t: 24,
        h: 60,
      },
    });

    expect(conditionsSheet.getRange).toHaveBeenCalledWith(3, 1);
    expect(conditionsSheet.getRange).toHaveBeenCalledWith(3, 2);
    expect(conditionsSheet.getRange).toHaveBeenCalledWith(3, 3);
    expect(outputState.content).toBe("OK");
    expect(result).not.toBeUndefined();
  });

  it("returns JSONP when callback is provided", () => {
    const { context, conditionsSheet, outputState, statusRange } = loadCode({
      conditionsLastRow: 4,
      conditionsHeaders: ["Date", "Temperature", "Humid"],
      conditionsRowValues: [new Date(0), 24.5, 38],
    });

    context.doGet({
      parameter: {
        callback: "cb",
      },
    });

    const expected = {
      conditions: [
        { Date: "1/1/1970 9:00:00 AM", Temperature: 24.5, Humid: 38 },
      ],
      status: "off",
    };
    const expectedContent = `cb&&cb(${JSON.stringify(expected)});`;

    expect(outputState.content).toBe(expectedContent);
    expect(conditionsSheet.getLastRow).toHaveBeenCalled();
    expect(conditionsSheet.getLastColumn).toHaveBeenCalled();
    expect(conditionsSheet.getRange).toHaveBeenCalledWith(1, 1, 1, 3);
    expect(conditionsSheet.getRange).toHaveBeenCalledWith(4, 1, 1, 3);
    expect(conditionsSheet.getDataRange).not.toHaveBeenCalled();
    expect(statusRange.getValue).toHaveBeenCalled();
  });
});

describe("turnOnAC", () => {
  it("turns on hot mode when temperature is below hot_buttom", () => {
    const settings = buildSettings({ hot_buttom: 19, hot: 23 });
    const { context, UrlFetchApp, statusRange, Logger } = loadCode({
      temperature: 18,
      settingsValues: settingsValuesFrom(settings),
      statusValue: "off",
    });

    context.turnOnAC();

    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      "http://a.ze.gs/switchbot-ac/-d/02-202307290753-22894539/-a/23,5,1,on"
    );
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      "http://a.ze.gs/switchbot-custom/-d/03-202401251013-58699638/-c/Hot"
    );
    expect(statusRange.setValue).toHaveBeenCalledWith("hot");
    expect(Logger.log).toHaveBeenCalledWith(settings);
    expect(Logger.log).toHaveBeenCalledWith(18);
  });

  it("turns off hot mode when temperature exceeds hot_up", () => {
    const settings = buildSettings({ hot_up: 27 });
    const { context, UrlFetchApp, statusRange, Logger } = loadCode({
      temperature: 28,
      settingsValues: settingsValuesFrom(settings),
      statusValue: "hot",
    });

    context.turnOnAC();

    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      "http://a.ze.gs/switchbot-ac/-d/02-202307290753-22894539/-a/25,5,1,off"
    );
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      "http://a.ze.gs/switchbot-custom/-d/03-202401251013-58699638/-c/Off"
    );
    expect(statusRange.setValue).toHaveBeenCalledWith("off");
    expect(Logger.log).toHaveBeenCalledWith(settings);
    expect(Logger.log).toHaveBeenCalledWith(28);
  });

  it("turns off cool mode when temperature is below cool_buttom", () => {
    const settings = buildSettings({ cool_buttom: 17 });
    const { context, UrlFetchApp, statusRange, Logger } = loadCode({
      temperature: 16,
      settingsValues: settingsValuesFrom(settings),
      statusValue: "cool",
    });

    context.turnOnAC();

    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      "http://a.ze.gs/switchbot-ac/-d/02-202307290753-22894539/-a/25,2,1,off"
    );
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      "http://a.ze.gs/switchbot-custom/-d/03-202401251013-58699638/-c/Off"
    );
    expect(statusRange.setValue).toHaveBeenCalledWith("off");
    expect(Logger.log).toHaveBeenCalledWith(settings);
    expect(Logger.log).toHaveBeenCalledWith(16);
  });

  it("turns on cool mode when temperature exceeds cool_up", () => {
    const settings = buildSettings({ cool_up: 26, cool: 25 });
    const { context, UrlFetchApp, statusRange, Logger } = loadCode({
      temperature: 27,
      settingsValues: settingsValuesFrom(settings),
      statusValue: "off",
    });

    context.turnOnAC();

    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      "http://a.ze.gs/switchbot-ac/-d/02-202307290753-22894539/-a/25,2,1,on"
    );
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      "http://a.ze.gs/switchbot-custom/-d/03-202401251013-58699638/-c/Cool"
    );
    expect(statusRange.setValue).toHaveBeenCalledWith("cool");
    expect(Logger.log).toHaveBeenCalledWith(settings);
    expect(Logger.log).toHaveBeenCalledWith(27);
  });
});
