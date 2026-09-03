import { Window } from "happy-dom";
import { vi } from "vitest";

// Ensure happy-dom DOMParser and XMLSerializer in Node/happy-dom environments
if (typeof globalThis.DOMParser === "undefined") {
  const win = new Window();
  globalThis.DOMParser = win.DOMParser;
  globalThis.XMLSerializer =
    win.XMLSerializer as unknown as typeof XMLSerializer;
}

if (typeof (globalThis as any).DOMMatrix === "undefined") {
  (globalThis as any).DOMMatrix = class DOMMatrix {};
}

export class MockOfficeExtensionError extends Error {
  code?: string;
  debugInfo?: {
    errorLocation?: string;
    statement?: string;
    surroundingStatements?: string[];
  };

  constructor(
    message: string,
    code?: string,
    debugInfo?: {
      errorLocation?: string;
      statement?: string;
      surroundingStatements?: string[];
    },
  ) {
    super(message);
    this.name = "OfficeExtension.Error";
    this.code = code;
    this.debugInfo = debugInfo;
    Object.setPrototypeOf(this, MockOfficeExtensionError.prototype);
  }
}

export function setupOfficeExtensionMock() {
  const g = globalThis as unknown as {
    OfficeExtension?: { Error: typeof MockOfficeExtensionError };
  };
  if (!g.OfficeExtension) {
    g.OfficeExtension = {
      Error: MockOfficeExtensionError,
    };
  } else {
    g.OfficeExtension.Error = MockOfficeExtensionError;
  }
}

export function setupMockOfficeSettings() {
  const g = globalThis as unknown as Record<string, unknown>;
  const store = new Map<string, unknown>();
  if (!g.Office) {
    g.Office = {};
  }
  const office = g.Office as Record<string, unknown>;
  if (!office.context) {
    office.context = {};
  }
  const ctx = office.context as Record<string, unknown>;
  if (!ctx.document) {
    ctx.document = {};
  }
  const doc = ctx.document as Record<string, unknown>;
  if (!doc.settings) {
    doc.settings = {
      get: (key: string) => store.get(key) ?? null,
      set: (key: string, value: unknown) => {
        store.set(key, value);
      },
      refreshAsync: (callback: () => void) => {
        if (typeof callback === "function") callback();
      },
      saveAsync: (callback: () => void) => {
        if (typeof callback === "function") callback();
      },
    };
  }
}

// Ensure OfficeExtension.Error and Office settings are mocked immediately on import
setupOfficeExtensionMock();
setupMockOfficeSettings();

export interface MockOfficeHostOptions<T = unknown> {
  runImpl?: (callback: (context: T) => Promise<unknown>) => Promise<unknown>;
}

export interface MockOfficeHostResult<T = unknown> {
  host: Record<string, unknown>;
  run: ReturnType<typeof vi.fn>;
  context: T;
  restore: () => void;
}

export function mockOfficeHost<T = any>(
  name: "Excel" | "Word" | "PowerPoint" | string,
  context?: T,
  options?: MockOfficeHostOptions<T>,
): MockOfficeHostResult<T> {
  setupOfficeExtensionMock();

  const defaultSync = vi.fn().mockResolvedValue(undefined);
  const defaultContext = (context ?? { sync: defaultSync }) as T;

  const runMock = vi.fn(async (callback: (ctx: T) => Promise<unknown>) => {
    if (options?.runImpl) {
      return options.runImpl(callback);
    }
    return callback(defaultContext);
  });

  const g = globalThis as unknown as Record<string, unknown>;
  const previous = g[name];

  const defaultExcelProps: Record<string, unknown> =
    name === "Excel"
      ? {
          RangeCopyType: {
            all: "All",
            formats: "Formats",
            formulas: "Formulas",
            values: "Values",
          },
        }
      : {};

  const hostObj: Record<string, unknown> = {
    ...defaultExcelProps,
    ...((typeof previous === "object" && previous !== null
      ? previous
      : {}) as Record<string, unknown>),
    run: runMock,
  };

  g[name] = hostObj;

  return {
    host: hostObj,
    run: runMock,
    context: defaultContext,
    restore: () => {
      if (previous !== undefined) {
        g[name] = previous;
      } else {
        delete g[name];
      }
    },
  };
}
