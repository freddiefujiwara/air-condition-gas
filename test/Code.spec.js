import { afterEach, describe, expect, it, vi } from "vitest";

const loadCode = async ({
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

  vi.stubGlobal("SpreadsheetApp", sandbox.SpreadsheetApp);
  vi.stubGlobal("ContentService", sandbox.ContentService);
  vi.stubGlobal("UrlFetchApp", sandbox.UrlFetchApp);
  vi.stubGlobal("Logger", sandbox.Logger);
  await vi.resetModules();
  const module = await import("../src/Code.js");

  return {
    module,
    output,
    outputState,
    conditionsSheet,
    conditionsRanges,
    statusRange,
    settingSheet,
    UrlFetchApp,
    Logger,
    conditionsRangeCalls,
  };
};

afterEach(() => {
  vi.unstubAllGlobals();
  vi.resetModules();
});

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
  it("writes status and exits when status param is provided", async () => {
    const { module, statusRange, outputState, output } = await loadCode();

    const result = module.doGet({
      parameter: {
        s: "status",
        t: "hot",
      },
    });

    expect(statusRange.setValue).toHaveBeenCalledWith("hot");
    expect(outputState.content).toBe("OK");
    expect(result).toBe(output);
  });

  it("appends condition data when no callback is set", async () => {
    const { module, conditionsSheet, outputState } = await loadCode({
      conditionsLastRow: 2,
    });

    const result = module.doGet({
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

  it("defaults missing condition values when appending", async () => {
    const { module, conditionsRanges } = await loadCode({
      conditionsLastRow: 1,
    });

    module.doGet({
      parameter: {},
    });

    const dateRange = conditionsRanges.get("2:1:1:1");
    const temperatureRange = conditionsRanges.get("2:2:1:1");
    const humidityRange = conditionsRanges.get("2:3:1:1");

    expect(dateRange.setValue).toHaveBeenCalledWith(expect.anything());
    expect(temperatureRange.setValue).toHaveBeenCalledWith(0);
    expect(humidityRange.setValue).toHaveBeenCalledWith(0);
  });

  it("returns JSONP when callback is provided", async () => {
    const settings = buildSettings({ hot_buttom: 19 });
    const { module, conditionsSheet, outputState, statusRange, settingSheet } =
      await loadCode({
        conditionsLastRow: 4,
        conditionsHeaders: ["Date", "Temperature", "Humid"],
        conditionsRowValues: [new Date(0), 24.5, 38],
        settingsValues: settingsValuesFrom(settings),
    });

    module.doGet({
      parameter: {
        callback: "cb",
      },
    });

    const expected = {
      conditions: [
        { Date: "1/1/1970 9:00:00 AM", Temperature: 24.5, Humid: 38 },
      ],
      status: "off",
      setting: settings,
    };
    const expectedContent = `cb&&cb(${JSON.stringify(expected)});`;

    expect(outputState.content).toBe(expectedContent);
    expect(conditionsSheet.getLastRow).toHaveBeenCalled();
    expect(conditionsSheet.getLastColumn).toHaveBeenCalled();
    expect(conditionsSheet.getRange).toHaveBeenCalledWith(1, 1, 1, 3);
    expect(conditionsSheet.getRange).toHaveBeenCalledWith(4, 1, 1, 3);
    expect(conditionsSheet.getDataRange).not.toHaveBeenCalled();
    expect(settingSheet.getDataRange).toHaveBeenCalled();
    expect(statusRange.getValue).toHaveBeenCalled();
  });
});

describe("settings payload", () => {
  it("returns empty settings when none exist", async () => {
    const { module, outputState, settingSheet } = await loadCode({
      conditionsHeaders: ["Date", "Temperature", "Humid"],
      conditionsRowValues: [new Date(0), 24.5, 38],
      conditionsLastRow: 2,
    });

    module.doGet({
      parameter: {
        callback: "cb",
      },
    });

    const expected = {
      conditions: [
        { Date: "1/1/1970 9:00:00 AM", Temperature: 24.5, Humid: 38 },
      ],
      status: "off",
      setting: {},
    };

    expect(outputState.content).toBe(`cb&&cb(${JSON.stringify(expected)});`);
    expect(settingSheet.getDataRange).toHaveBeenCalled();
  });
});

describe("turnOnAC", () => {
  it("turns on hot mode when temperature is below hot_buttom", async () => {
    const settings = buildSettings({ hot_buttom: 19, hot: 23 });
    const { module, UrlFetchApp, statusRange, Logger } = await loadCode({
      temperature: 18,
      settingsValues: settingsValuesFrom(settings),
      statusValue: "off",
    });

    module.turnOnAC();

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

  it("turns off hot mode when temperature exceeds hot_up", async () => {
    const settings = buildSettings({ hot_up: 27 });
    const { module, UrlFetchApp, statusRange, Logger } = await loadCode({
      temperature: 28,
      settingsValues: settingsValuesFrom(settings),
      statusValue: "hot",
    });

    module.turnOnAC();

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

  it("turns off cool mode when temperature is below cool_buttom", async () => {
    const settings = buildSettings({ cool_buttom: 17 });
    const { module, UrlFetchApp, statusRange, Logger } = await loadCode({
      temperature: 16,
      settingsValues: settingsValuesFrom(settings),
      statusValue: "cool",
    });

    module.turnOnAC();

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

  it("turns on cool mode when temperature exceeds cool_up", async () => {
    const settings = buildSettings({ cool_up: 26, cool: 25 });
    const { module, UrlFetchApp, statusRange, Logger } = await loadCode({
      temperature: 27,
      settingsValues: settingsValuesFrom(settings),
      statusValue: "off",
    });

    module.turnOnAC();

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
