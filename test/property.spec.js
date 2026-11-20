import { beforeEach, describe, expect, it, vi } from "vitest";
import fc from "fast-check";
import * as Code from "../src/Code.js";

// Property-based tests over the same outside inputs as
// test/abnormal.spec.js: sheet cells read back, HTTP parameters, and
// the settings row. The reference implementations below state the
// documented contracts directly; the properties check that the code
// under test agrees with them for every generated input, and never
// throws. The harness mirrors test/abnormal.spec.js (one mutable
// sheet state, one module import).

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

const KNOWN_STATUSES = ["off", "hot", "cool", "control"];

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

// The documented reading contract: a finite number, or a non-blank
// string that Number() parses to a finite number, is a reading.
// Everything else is no reading at all (NaN here).
const strictNumber = (value) => {
  const parsed =
    typeof value === "number"
      ? value
      : typeof value === "string" && value.trim() !== ""
        ? Number(value)
        : NaN;
  return Number.isFinite(parsed) ? parsed : NaN;
};

// The documented automation contract, in AC_ACTIONS order. The
// thresholds are the owner's own configuration and are kept
// well-formed here; the deformed domain is the temperature cell,
// the stored status, and the two command settings (hot / cool),
// whose values are interpolated into a command sent to a physical
// device.
const referenceOutcome = ({ temperature, status, settings }) => {
  const t = strictNumber(temperature);
  if (!Number.isNaN(t)) {
    if (
      t <= settings.hot_buttom &&
      status === "off" &&
      !Number.isNaN(strictNumber(settings.hot))
    ) {
      return {
        fetches: [
          `${AC_BASE}/-a/${settings.hot},5,1,on`,
          `${CUSTOM_BASE}/-c/Hot`,
        ],
        statusWrites: ["hot"],
      };
    }
    if (t >= settings.hot_up && status === "hot") {
      return {
        fetches: [`${AC_BASE}/-a/25,5,1,off`, `${CUSTOM_BASE}/-c/Off`],
        statusWrites: ["off"],
      };
    }
    if (t <= settings.cool_buttom && status === "cool") {
      return {
        fetches: [`${AC_BASE}/-a/25,2,1,off`, `${CUSTOM_BASE}/-c/Off`],
        statusWrites: ["off"],
      };
    }
    if (
      t >= settings.cool_up &&
      status === "off" &&
      !Number.isNaN(strictNumber(settings.cool))
    ) {
      return {
        fetches: [
          `${AC_BASE}/-a/${settings.cool},2,1,on`,
          `${CUSTOM_BASE}/-c/Cool`,
        ],
        statusWrites: ["cool"],
      };
    }
  }
  return { fetches: [], statusWrites: [] };
};

const runTurnOnAC = () => {
  Code.turnOnAC();
  return {
    fetches: state.fetches.map((fetch) => fetch.url),
    statusWrites: state.statusWrites,
  };
};

const scalarOrArray = fc.oneof(
  fc.integer({ min: 15, max: 30 }),
  fc.double({ min: 15, max: 30, noNaN: true }),
  fc.constantFrom("", " ", "abc", "22", "24.5", true, false, null, undefined),
  fc.array(fc.integer({ min: 0, max: 40 }), { maxLength: 2 }),
  fc.record({ valueOf: fc.constant(() => 65) })
);

describe("turnOnAC", () => {
  it("acts exactly as the reference table for arbitrary cells, statuses and command settings", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.oneof(
          fc.double(),
          fc.string(),
          fc.constantFrom(
            null,
            undefined,
            true,
            false,
            NaN,
            Infinity,
            -Infinity,
            new Date(NaN)
          ),
          fc.date({ noInvalidDate: true }),
          fc.array(fc.double({ noNaN: true }), { maxLength: 3 }),
          fc.record({ valueOf: fc.constant(() => 65) })
        ),
        fc.constantFrom("off", "hot", "cool", "control", "banana", "", "HOT"),
        scalarOrArray,
        scalarOrArray,
        async (cell, status, hot, cool) => {
          state = createState();
          const settings = { ...DEFAULT_SETTINGS, hot, cool };
          state.settingsValues = settingsValuesFrom(settings);
          state.temperatureCell = cell;
          state.status = status;

          const expected = referenceOutcome({
            temperature: cell,
            status,
            settings,
          });
          expect(runTurnOnAC()).toEqual(expected);
        }
      ),
      { numRuns: 100 }
    );
  });

  it("matches the reference table for well-formed numeric readings", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.double({ min: -10, max: 45, noNaN: true }),
        fc.constantFrom("off", "hot", "cool"),
        async (cell, status) => {
          state = createState();
          state.temperatureCell = cell;
          state.status = status;

          const expected = referenceOutcome({
            temperature: cell,
            status,
            settings: DEFAULT_SETTINGS,
          });
          expect(runTurnOnAC()).toEqual(expected);
        }
      ),
      { numRuns: 100 }
    );
  });
});

describe("status writes (doGet)", () => {
  it("stores a status if and only if it is one of the four real statuses", async () => {
    await fc.assert(
      fc.asyncProperty(fc.string(), async (t) => {
        state = createState();
        state.status = "cool";

        const output = Code.doGet({
          parameter: { s: "status", t, token: WRITE_TOKEN },
        });

        if (KNOWN_STATUSES.includes(t)) {
          expect(output.getContent()).toBe("OK");
          expect(state.status).toBe(t);
          expect(state.statusWrites).toEqual([t]);
        } else {
          expect(output.getContent()).toBe("NG");
          expect(state.status).toBe("cool");
          expect(state.statusWrites).toEqual([]);
        }
      }),
      { numRuns: 100 }
    );
  });
});

describe("read payload (doGet with callback)", () => {
  // JSON-representable scalars: what a sheet cell can hold, minus the
  // shapes JSON cannot round-trip (NaN, Infinity). "__proto__" is
  // excluded from headers: a plain-object accumulator drops it
  // silently, and no setting by that name is ever read, so it is
  // outside this conservation property's domain.
  const jsonScalar = fc.oneof(
    fc.integer(),
    // -0 is filtered out: the payload crosses a JSON boundary, and
    // JSON.stringify(-0) is "0", so a negative zero cannot survive
    // the wire format by construction. That is the format's
    // contract, not a shaping bug.
    fc
      .double({ noNaN: true, noDefaultInfinity: true })
      .filter((value) => !Object.is(value, -0)),
    fc.string(),
    fc.boolean(),
    fc.constant(null)
  );
  const headerName = fc
    .string({ minLength: 1, maxLength: 12 })
    .filter((name) => name !== "Date" && name !== "__proto__");

  it("conserves the latest row and the settings map", async () => {
    // Header/value pairs with unique headers, so the header row and
    // the data row always have the same width — as they do in the
    // real read path, which reads both at the sheet's last column.
    const headerValuePairs = (minLength, maxLength) =>
      fc.uniqueArray(fc.tuple(headerName, jsonScalar), {
        minLength,
        maxLength,
        selector: (pair) => pair[0],
      });
    // The Date cell is generated within 2000–2100: formatJapanTime
    // round-trips through a locale string in Asia/Tokyo, and a date
    // within hours of the Date range extremes (year ±275760) tips
    // over that conversion and formats as "Invalid Date". The sheet's
    // Date column only ever holds new Date() stamps written at
    // ingest time, so that fringe is outside the reachable domain
    // and the generator stays out of it deliberately.
    const readingDate = fc.date({
      min: new Date(2000, 0, 1),
      max: new Date(2100, 0, 1),
      noInvalidDate: true,
    });
    await fc.assert(
      fc.asyncProperty(
        headerValuePairs(0, 4),
        readingDate,
        headerValuePairs(0, 4),
        async (extraPairs, dateCell, settingPairs) => {
          state = createState();
          const headers = ["Date", ...extraPairs.map((pair) => pair[0])];
          const rowValues = [dateCell, ...extraPairs.map((pair) => pair[1])];
          state.headers = headers;
          state.rowValues = rowValues;
          state.settingsValues = [
            settingPairs.map((pair) => pair[0]),
            settingPairs.map((pair) => pair[1]),
          ];

          const output = Code.doGet({ parameter: { callback: "cb" } });
          const content = output.getContent();
          const payload = JSON.parse(
            content.slice(content.indexOf("(") + 1, content.lastIndexOf(")"))
          );

          const condition = payload.conditions[0];
          expect(Object.keys(condition).sort()).toEqual([...headers].sort());
          expect(typeof condition.Date).toBe("string");
          expect(condition.Date).not.toContain("Invalid");
          extraPairs.forEach(([header, value]) => {
            expect(condition[header]).toEqual(value);
          });

          const expectedSetting = {};
          settingPairs.forEach(([header, value]) => {
            expectedSetting[header] = value;
          });
          expect(payload.setting).toEqual(expectedSetting);
        }
      ),
      { numRuns: 100 }
    );
  });

  it("emits JSONP only for identifier callbacks and JSON otherwise", async () => {
    const CALLBACK_PATTERN =
      /^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/;
    await fc.assert(
      fc.asyncProperty(fc.string(), async (callback) => {
        state = createState();

        const output = Code.doGet({ parameter: { callback } });
        const content = output.getContent();

        if (CALLBACK_PATTERN.test(callback)) {
          expect(output.getMimeType()).toBe("application/javascript");
          expect(content.startsWith(`${callback}&&${callback}(`)).toBe(true);
        } else {
          expect(output.getMimeType()).toBe("application/json");
          expect(() => JSON.parse(content)).not.toThrow();
        }
      }),
      { numRuns: 100 }
    );
  });
});

describe("condition writes (doGet)", () => {
  it("appends exactly the readings it was sent, plus a fresh Date", async () => {
    await fc.assert(
      fc.asyncProperty(fc.string(), fc.string(), async (t, h) => {
        state = createState();

        const output = Code.doGet({
          parameter: { t, h, token: WRITE_TOKEN },
        });

        expect(output.getContent()).toBe("OK");
        expect(state.conditionWrites).toHaveLength(3);
        const byColumn = new Map(
          state.conditionWrites.map((write) => [write.col, write.value])
        );
        expect(byColumn.get(1)).toBeInstanceOf(Date);
        expect(byColumn.get(2)).toBe(t);
        expect(byColumn.get(3)).toBe(h);
      }),
      { numRuns: 100 }
    );
  });
});
