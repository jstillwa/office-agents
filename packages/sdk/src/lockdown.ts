/**
 * SPIKE EVALUATION: Eager Lockdown vs. Lazy Lockdown
 *
 * Decision: NO-GO for eager ensureLockdown() at taskpane init. Keep lazy lockdown as default.
 *
 * Evaluation & Findings:
 * 1. SES `lockdown()` freezes intrinsic JavaScript prototypes (Object.prototype,
 *    Function.prototype, Array.prototype, etc.).
 * 2. Office.js initialization (prior to Office.onReady / Office.initialize resolution)
 *    relies on mutating prototypes and adding dynamic validation stubs and helper properties
 *    to Function and Object constructors across desktop WebView2 (Windows), WebKit (macOS),
 *    and Office on the web frames.
 * 3. Eager execution of lockdown() before Office.onReady() completes leads to fatal
 *    TypeError exceptions ("Cannot add property ... object is not extensible" or
 *    "Cannot assign to read-only property") during Office.js API registration, breaking
 *    the host-to-taskpane handshake and causing the add-in taskpane to fail to load.
 * 4. Desktop platforms also exhibit platform-specific timing where scripts injected by
 *    the host require dynamic reflection before the runtime stabilizes.
 *
 * Recommendation:
 * Keep lazy lockdown as default. Calling ensureLockdown() lazily upon the first
 * sandboxedEval call ensures Office.js finishes initialization safely, while still
 * hardening the JS realm before any untrusted user/model-generated code evaluates.
 */

import "ses";

/* global lockdown */

const PRESERVE_FUNCTION_PROPS = new Set(["length", "name", "prototype"]);
const OFFICE_FUNCTION_STUBS = new Set([
  "_validateParams",
  "_validateParameterCount",
  "_validateParameter",
  "_validateParameterType",
]);

function saveFunctionProperties(): Map<string, unknown> {
  const saved = new Map<string, unknown>();
  for (const key of Object.getOwnPropertyNames(Function)) {
    if (!PRESERVE_FUNCTION_PROPS.has(key)) {
      saved.set(key, (Function as unknown as Record<string, unknown>)[key]);
    }
  }
  return saved;
}

function restoreFunctionProperties(saved: Map<string, unknown>) {
  if (saved.size === 0) return;

  const noop = () => null;
  for (const key of OFFICE_FUNCTION_STUBS) {
    if (saved.has(key)) saved.set(key, noop);
  }

  const fn = globalThis.Function;

  // Fast path: if Function is still extensible, define properties directly
  if (Object.isExtensible(fn)) {
    try {
      for (const [key, value] of saved) {
        Object.defineProperty(fn, key, {
          value,
          writable: false,
          enumerable: false,
          configurable: false,
        });
      }
      return;
    } catch {
      // Fall through to proxy path
    }
  }

  // Slow path: Function is frozen by lockdown, wrap it in a Proxy
  const desc = Object.getOwnPropertyDescriptor(globalThis, "Function");
  if (!desc || (!desc.writable && !desc.configurable)) {
    console.warn(
      "[lockdown] Cannot restore Function properties — Function is neither writable nor configurable",
    );
    return;
  }

  const proxy = new Proxy(fn, {
    get(target, prop, receiver) {
      return saved.has(prop as string)
        ? saved.get(prop as string)
        : Reflect.get(target, prop, receiver);
    },
    has(target, prop) {
      return saved.has(prop as string) ? true : Reflect.has(target, prop);
    },
    getOwnPropertyDescriptor(target, prop) {
      if (saved.has(prop as string)) {
        return {
          value: saved.get(prop as string),
          writable: false,
          enumerable: false,
          configurable: true,
        };
      }
      return Reflect.getOwnPropertyDescriptor(target, prop);
    },
  });

  Object.defineProperty(globalThis, "Function", {
    value: proxy,
    writable: desc.writable ?? false,
    configurable: desc.configurable ?? false,
  });
}

let locked = false;

export function ensureLockdown() {
  if (locked) return;
  try {
    const savedFnProps = saveFunctionProperties();
    lockdown({
      errorTaming: "unsafe",
      consoleTaming: "unsafe",
      overrideTaming: "severe",
      stackFiltering: "verbose",
    });
    locked = true;
    restoreFunctionProperties(savedFnProps);
  } catch (e) {
    if (
      e instanceof TypeError &&
      String(e).includes("SES_ALREADY_LOCKED_DOWN")
    ) {
      locked = true;
    } else {
      throw e;
    }
  }
}
