import { beforeEach, describe, expect, it, vi } from "vitest";
import * as Code from "../src/Code.js";

// Abnormal-case tests: the "input from outside" for this logger is
// (a) the HTTP parameters of the write path (readings and status
// values sent by the home server and the ha frontend), (b) the sheet
// cells read back — the temperature cell that drives the automation,
// the status cell, the settings row, and the latest conditions row
// that is served to ha. Every deformed input must fail in isolation:
// a bad reading must not drive the air conditioner, an unknown status
// must not be stored as the automation's state, and one deformed row
// must not kill the whole read payload. Well-formed traffic (numeric
// or numeric-string readings, the four real statuses) is unchanged.
//
// The harness below mirrors test/Code.spec.js but keeps the sheet
// state in one mutable object so each case can deform a single cell.

const WRITE_TOKEN = "dummy-write-token";
const AC_BASE = "http://home-api.example/switchbot-ac/-d/DUMMY-AC-DEVICE";
const CUSTOM_BASE =
  "http://home-api.example/switchbot-custom/-d/DUMMY-CUSTOM-DEVICE";

const DEFAULT_SETTINGS = {
  hot_buttom: 20,
  hot_up: 28,
  cool_buttom: 18,
  cool_up: 26,
  hot: 22,
  cool: 24,
};

const settingsValuesFrom = (settings) => [
  Object.keys(settings),
  Object.values(settings),
];

const createState = () => ({
  lastRow: 2,
  headers: ["Date", "Temperature", "Humid"],
  rowValues: [new Date(0), 24.5, 38],
  temperatureCell: 24.5,
  settingsValues: settingsValuesFrom(DEFAULT_SETTINGS),
  status: "off",
  properties: {
    SWITCHBOT_AC_URL: AC_BASE,
    SWITCHBOT_CUSTOM_URL: CUSTOM_BASE,
    HOME_API_TOKEN: "dummy-home-api-token",
    WRITE_TOKEN,
  },
  conditionWrites: [],
  statusWrites: [],
  fetches: [],
});

let state = createState();

const conditionsSheet = {
  getLastRow: () => state.lastRow,
  getLastColumn: () => state.headers.length,
  getRange: (row, col) => ({
    setValue: (value) => {
      state.conditionWrites.push({ row, col, value });
    },
    getValues: () => {
      if (row === 1) return [state.headers];
      if (row === state.lastRow) return [state.rowValues];
      return [[]];
    },
    getValue: () =>
      row === state.lastRow && col === 2 ? state.temperatureCell : undefined,
  }),
  getDataRange: () => ({
    getValues: () => [state.headers, state.rowValues],
  }),
};

const statusSheet = {
  getRange: () => ({
    getValue: () => state.status,
    setValue: (value) => {
      state.status = value;
      state.statusWrites.push(value);
    },
  }),
};

const settingSheet = {
  getDataRange: () => ({ getValues: () => state.settingsValues.slice() }),
};

vi.stubGlobal("SpreadsheetApp", {
  openById: vi.fn(() => ({
    getSheetByName: (name) =>
      name === "conditions"
        ? conditionsSheet
        : name === "status"
          ? statusSheet
          : settingSheet,
  })),
});

vi.stubGlobal("ContentService", {
  MimeType: {
    TEXT: "text/plain",
    JAVASCRIPT: "application/javascript",
    JSON: "application/json",
  },
  createTextOutput: vi.fn(() => {
    const output = { content: "", mimeType: "" };
    output.setContent = (content) => {
      output.content = content;
    };
    output.setMimeType = (mimeType) => {
      output.mimeType = mimeType;
    };
    output.getContent = () => output.content;
    output.getMimeType = () => output.mimeType;
    return output;
  }),
});

vi.stubGlobal("UrlFetchApp", {
  fetch: vi.fn((url, options) => {
    state.fetches.push({ url, options });
  }),
});

vi.stubGlobal("Logger", { log: vi.fn() });

vi.stubGlobal("PropertiesService", {
  getScriptProperties: vi.fn(() => ({
    getProperty: (name) =>
      Object.prototype.hasOwnProperty.call(state.properties, name)
        ? state.properties[name]
        : null,
  })),
});

beforeEach(() => {
  state = createState();
});

const fetchUrls = () => state.fetches.map((fetch) => fetch.url);

const readPayload = (callback = "cb") => {
  const output = Code.doGet({ parameter: { callback } });
  const content = output.getContent();
  const json = content.slice(content.indexOf("(") + 1, content.lastIndexOf(")"));
  return JSON.parse(json);
};

describe("latest temperature cell (turnOnAC)", () => {
  // A bare Number() invents a plausible measurement out of every one
  // of these: true -> 1, false -> 0, a Date -> its epoch milliseconds,
  // [65] -> 65, "" -> 0, null -> 0, "Infinity" -> Infinity. The
  // automation must treat all of them as "no reading" instead.
  // (Sheets cells are scalars or Dates; the array/object shapes are
  // prevention so a future cell type cannot wedge the automation.)
  it.each([
    ["boolean true (was read as 1 degree)", true],
    ["boolean false (was read as 0 degrees)", false],
    ["a Date (was read as epoch milliseconds)", new Date(2026, 10, 20)],
    ["an array (was read as its element)", [65]],
    ["an empty string (was read as 0 degrees)", ""],
    ["a whitespace string (was read as 0 degrees)", " "],
    ["null (was read as 0 degrees)", null],
    ["an Infinity string", "Infinity"],
    ["an object with valueOf (was read as 65)", { valueOf: () => 65 }],
  ])("does not act on %s", (_label, cell) => {
    state.temperatureCell = cell;
    state.status = "off";

    Code.turnOnAC();

    expect(state.fetches).toEqual([]);
    expect(state.statusWrites).toEqual([]);
  });

  it("does not act on a non-numeric string", () => {
    state.temperatureCell = "abc";
    state.status = "off";

    Code.turnOnAC();

    expect(state.fetches).toEqual([]);
    expect(state.statusWrites).toEqual([]);
  });

  it("does not act on a NaN cell", () => {
    state.temperatureCell = NaN;
    state.status = "off";

    Code.turnOnAC();

    expect(state.fetches).toEqual([]);
    expect(state.statusWrites).toEqual([]);
  });

  it("still acts on a numeric-string cell", () => {
    state.temperatureCell = "10";
    state.status = "off";

    Code.turnOnAC();

    expect(fetchUrls()).toEqual([
      `${AC_BASE}/-a/22,5,1,on`,
      `${CUSTOM_BASE}/-c/Hot`,
    ]);
    expect(state.statusWrites).toEqual(["hot"]);
  });

  it("keeps Number() semantics for other numeric spellings", () => {
    state.temperatureCell = "0x1A"; // 26 degrees
    state.status = "off";

    Code.turnOnAC();

    expect(fetchUrls()).toEqual([
      `${AC_BASE}/-a/24,2,1,on`,
      `${CUSTOM_BASE}/-c/Cool`,
    ]);
    expect(state.statusWrites).toEqual(["cool"]);
  });

  it("still acts on a finite number outside the physical range", () => {
    // The parser checks the value's shape, not its plausibility:
    // a finite number is a reading, however unlikely.
    state.temperatureCell = 1e12;
    state.status = "off";

    Code.turnOnAC();

    expect(fetchUrls()).toEqual([
      `${AC_BASE}/-a/24,2,1,on`,
      `${CUSTOM_BASE}/-c/Cool`,
    ]);
    expect(state.statusWrites).toEqual(["cool"]);
  });
});

describe("status writes (doGet)", () => {
  const writeStatus = (t) =>
    Code.doGet({ parameter: { s: "status", t, token: WRITE_TOKEN } });

  // The status cell is the automation's state. turnOnAC only acts on
  // "off" / "hot" / "cool" and the ha frontend's fourth value
  // "control" parks the A/C under manual control; anything else
  // stored here silently disables the automation and is displayed
  // by ha as the A/C state.
  it.each([
    ["an unknown word", "banana"],
    ["an empty string", ""],
    ["a missing value", undefined],
    ["an uppercase variant", "HOT"],
    ["a number", 42],
  ])("rejects %s instead of storing it as the state", (_label, t) => {
    const output = writeStatus(t);

    expect(output.getContent()).toBe("NG");
    expect(state.statusWrites).toEqual([]);
    expect(state.status).toBe("off");
  });

  it.each(["off", "hot", "cool", "control"])(
    "accepts the real status %s",
    (t) => {
      const output = writeStatus(t);

      expect(output.getContent()).toBe("OK");
      expect(state.statusWrites).toEqual([t]);
      expect(state.status).toBe(t);
    }
  );

  it("keeps the previous status readable after a rejected write", () => {
    state.status = "cool";

    writeStatus("banana");

    expect(readPayload().status).toBe("cool");
  });
});

describe("settings command values (turnOnAC)", () => {
  it("does not send a command when the hot setting is missing", () => {
    const { hot, ...rest } = DEFAULT_SETTINGS;
    state.settingsValues = settingsValuesFrom(rest);
    state.temperatureCell = 10;
    state.status = "off";

    Code.turnOnAC();

    // Without the fix this fetched ".../-a/undefined,5,1,on".
    expect(state.fetches).toEqual([]);
    expect(state.statusWrites).toEqual([]);
  });

  it("does not send a command when the cool setting is blank", () => {
    state.settingsValues = settingsValuesFrom({
      ...DEFAULT_SETTINGS,
      cool: "",
    });
    state.temperatureCell = 30;
    state.status = "off";

    Code.turnOnAC();

    // Without the fix this fetched ".../-a/,2,1,on".
    expect(state.fetches).toEqual([]);
    expect(state.statusWrites).toEqual([]);
  });

  it("still sends the command for a numeric-string hot setting", () => {
    state.settingsValues = settingsValuesFrom({
      ...DEFAULT_SETTINGS,
      hot: "22",
    });
    state.temperatureCell = 10;
    state.status = "off";

    Code.turnOnAC();

    expect(fetchUrls()).toEqual([
      `${AC_BASE}/-a/22,5,1,on`,
      `${CUSTOM_BASE}/-c/Hot`,
    ]);
    expect(state.statusWrites).toEqual(["hot"]);
  });

  it("keeps numeric-string thresholds working", () => {
    state.settingsValues = settingsValuesFrom({
      hot_buttom: "20",
      hot_up: "28",
      cool_buttom: "18",
      cool_up: "26",
      hot: "22",
      cool: "24",
    });
    state.temperatureCell = 10;
    state.status = "off";

    Code.turnOnAC();

    expect(fetchUrls()).toEqual([
      `${AC_BASE}/-a/22,5,1,on`,
      `${CUSTOM_BASE}/-c/Hot`,
    ]);
    expect(state.statusWrites).toEqual(["hot"]);
  });
});

describe("read payload (doGet with callback)", () => {
  it("survives an unconvertible Date cell", () => {
    // Prevention domain (a Sheets cell cannot hold this shape today):
    // new Date() throws on a value with no usable toString, and one
    // such cell must not kill the whole payload ha depends on.
    state.rowValues = [{ toString: [] }, 24.5, 38];

    const payload = readPayload();

    expect(payload.conditions[0].Date).toBe("");
    expect(payload.conditions[0].Temperature).toBe(24.5);
    expect(payload.status).toBe("off");
  });

  it("does not throw on a non-date string in the Date column", () => {
    state.rowValues = ["not a date", 24.5, 38];

    const payload = readPayload();

    expect(typeof payload.conditions[0].Date).toBe("string");
    expect(payload.conditions[0].Temperature).toBe(24.5);
  });

  it("passes a non-string status cell through verbatim", () => {
    state.status = 42;

    expect(readPayload().status).toBe(42);
  });

  it("returns empty settings when the setting sheet is empty", () => {
    state.settingsValues = [];

    expect(readPayload().setting).toEqual({});
  });

  it("ignores settings rows beyond the first data row", () => {
    state.settingsValues = [
      ["hot"],
      [22],
      [99],
    ];

    expect(readPayload().setting).toEqual({ hot: 22 });
  });
});

describe("write auth and event shapes (doGet)", () => {
  it("rejects an array token", () => {
    const output = Code.doGet({
      parameter: { t: "24", h: "60", token: [WRITE_TOKEN] },
    });

    expect(output.getContent()).toBe("NG");
    expect(state.conditionWrites).toEqual([]);
  });

  it("rejects every write when the WRITE_TOKEN property is unset", () => {
    delete state.properties.WRITE_TOKEN;

    const output = Code.doGet({
      parameter: { t: "24", h: "60", token: WRITE_TOKEN },
    });

    expect(output.getContent()).toBe("NG");
    expect(state.conditionWrites).toEqual([]);
  });

  it("handles a null parameter object", () => {
    const output = Code.doGet({ parameter: null });

    expect(output.getContent()).toBe("NG");
  });

  it("falls back to JSON for a non-string callback", () => {
    const output = Code.doGet({ parameter: { callback: 42 } });

    expect(output.getMimeType()).toBe("application/json");
    expect(JSON.parse(output.getContent()).status).toBe("off");
  });

  it("records string readings verbatim", () => {
    // The wire format is query parameters, i.e. always strings; the
    // ingest contract is to store them as sent.
    const output = Code.doGet({
      parameter: { t: "24.5", h: "60", token: WRITE_TOKEN },
    });

    expect(output.getContent()).toBe("OK");
    expect(state.conditionWrites).toHaveLength(3);
    const byColumn = new Map(
      state.conditionWrites.map((write) => [write.col, write.value])
    );
    expect(byColumn.get(1)).toBeInstanceOf(Date);
    expect(byColumn.get(2)).toBe("24.5");
    expect(byColumn.get(3)).toBe("60");
  });
});
