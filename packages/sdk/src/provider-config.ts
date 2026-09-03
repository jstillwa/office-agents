import type { Api, Model } from "@earendil-works/pi-ai";
import type { StorageNamespace } from "./context";
import { loadOAuthCredentials } from "./oauth";

export type ThinkingLevel = "none" | "low" | "medium" | "high";

export interface ProviderConfig {
  provider: string;
  apiKey: string;
  model: string;
  useProxy: boolean;
  proxyUrl: string;
  thinking: ThinkingLevel;
  followMode: boolean;
  expandToolCalls: boolean;
  apiType?: string;
  customBaseUrl?: string;
  authMethod?: "apikey" | "oauth" | "sso";
}

function getEnv(key: string): string | undefined {
  const metaEnv = (
    import.meta as unknown as { env?: Record<string, string | undefined> }
  ).env;
  if (metaEnv && metaEnv[key] !== undefined) {
    return metaEnv[key];
  }
  const proc = (
    globalThis as unknown as {
      process?: { env?: Record<string, string | undefined> };
    }
  ).process;
  return proc?.env?.[key];
}

export function isEnterprise(): boolean {
  return getEnv("VITE_APP_MODE") === "enterprise";
}

export const ENTERPRISE_GATEWAY_URL =
  getEnv("VITE_ENTERPRISE_GATEWAY_URL") ||
  "https://gateway.enterprise.internal/v1";

export const ENTERPRISE_DEFAULT_MODEL =
  getEnv("VITE_ENTERPRISE_MODEL") || "corporate-default";

export const ENTERPRISE_DEFAULT_API_TYPE =
  getEnv("VITE_ENTERPRISE_API_TYPE") || "openai-completions";

function storageKey(ns: StorageNamespace): string {
  return `${ns.localStoragePrefix}-provider-config`;
}

export const THINKING_LEVELS: { value: ThinkingLevel; label: string }[] = [
  { value: "none", label: "None" },
  { value: "low", label: "Low" },
  { value: "medium", label: "Medium" },
  { value: "high", label: "High" },
];

export const API_TYPES = [
  {
    id: "openai-completions",
    name: "OpenAI Completions",
    hint: "Most compatible — Ollama, vLLM, LMStudio, etc.",
  },
  {
    id: "openai-responses",
    name: "OpenAI Responses",
    hint: "Newer OpenAI API format",
  },
  { id: "anthropic-messages", name: "Anthropic Messages", hint: "Claude API" },
  {
    id: "google-generative-ai",
    name: "Google Generative AI",
    hint: "Gemini API",
  },
  {
    id: "azure-openai-responses",
    name: "Azure OpenAI Responses",
    hint: "Azure-hosted OpenAI",
  },
  {
    id: "openai-codex-responses",
    name: "OpenAI Codex Responses",
    hint: "ChatGPT subscription models",
  },
  {
    id: "google-gemini-cli",
    name: "Google Gemini CLI",
    hint: "Cloud Code Assist",
  },
  { id: "google-vertex", name: "Google Vertex AI", hint: "Vertex AI endpoint" },
];

const VALID_API_TYPES = new Set(API_TYPES.map((t) => t.id));

export function isValidApiType(type: unknown): type is string {
  return typeof type === "string" && (type === "" || VALID_API_TYPES.has(type));
}

function isValidConfigObject(obj: unknown): obj is Record<string, unknown> {
  return typeof obj === "object" && obj !== null && !Array.isArray(obj);
}

export function loadSavedConfig(ns: StorageNamespace): ProviderConfig | null {
  try {
    const saved = localStorage.getItem(storageKey(ns));
    if (saved) {
      const parsed = JSON.parse(saved);
      if (isValidConfigObject(parsed)) {
        const config: ProviderConfig = {
          provider:
            typeof parsed.provider === "string" ? parsed.provider : "custom",
          apiKey: typeof parsed.apiKey === "string" ? parsed.apiKey : "",
          model: typeof parsed.model === "string" ? parsed.model : "",
          useProxy:
            typeof parsed.useProxy === "boolean" ? parsed.useProxy : true,
          proxyUrl: typeof parsed.proxyUrl === "string" ? parsed.proxyUrl : "",
          thinking: (["none", "low", "medium", "high"].includes(
            parsed.thinking as string,
          )
            ? parsed.thinking
            : "none") as ThinkingLevel,
          followMode:
            typeof parsed.followMode === "boolean" ? parsed.followMode : true,
          expandToolCalls:
            typeof parsed.expandToolCalls === "boolean"
              ? parsed.expandToolCalls
              : false,
          apiType: isValidApiType(parsed.apiType)
            ? (parsed.apiType as string)
            : "openai-completions",
          customBaseUrl:
            typeof parsed.customBaseUrl === "string"
              ? parsed.customBaseUrl
              : "",
          authMethod: (["apikey", "oauth", "sso"].includes(
            parsed.authMethod as string,
          )
            ? parsed.authMethod
            : "apikey") as "apikey" | "oauth" | "sso",
        };

        if (isEnterprise()) {
          config.apiKey = "";
          config.useProxy = false;
          config.proxyUrl = "";
          config.authMethod = "sso";
          config.provider = "custom";
          config.customBaseUrl = config.customBaseUrl || ENTERPRISE_GATEWAY_URL;
          config.model = config.model || ENTERPRISE_DEFAULT_MODEL;
          config.apiType = config.apiType || ENTERPRISE_DEFAULT_API_TYPE;
          return config;
        }

        if (config.authMethod === "oauth") {
          const creds = loadOAuthCredentials(ns, config.provider);
          if (creds) config.apiKey = creds.access;
        }
        return config;
      }
    }
  } catch {}

  if (isEnterprise()) {
    return {
      provider: "custom",
      apiKey: "",
      model: ENTERPRISE_DEFAULT_MODEL,
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
      apiType: ENTERPRISE_DEFAULT_API_TYPE,
      customBaseUrl: ENTERPRISE_GATEWAY_URL,
      authMethod: "sso",
    };
  }

  return null;
}

export function saveConfig(ns: StorageNamespace, config: ProviderConfig) {
  if (isEnterprise()) {
    const safeConfig: ProviderConfig = {
      provider: "custom",
      apiKey: "",
      model: config.model || ENTERPRISE_DEFAULT_MODEL,
      useProxy: false,
      proxyUrl: "",
      thinking: config.thinking || "none",
      followMode: config.followMode ?? true,
      expandToolCalls: config.expandToolCalls ?? false,
      apiType:
        isValidApiType(config.apiType) && config.apiType
          ? config.apiType
          : ENTERPRISE_DEFAULT_API_TYPE,
      customBaseUrl: config.customBaseUrl || ENTERPRISE_GATEWAY_URL,
      authMethod: "sso",
    };
    const { apiKey: _, ...withoutKey } = safeConfig;
    localStorage.setItem(storageKey(ns), JSON.stringify(withoutKey));
    return;
  }
  localStorage.setItem(storageKey(ns), JSON.stringify(config));
}

export function buildCustomModel(config: ProviderConfig): Model<Api> | null {
  if (!config.apiType || !config.customBaseUrl || !config.model) return null;
  if (!isValidApiType(config.apiType)) return null;
  return {
    id: config.model,
    name: config.model,
    api: config.apiType as Api,
    provider: "custom",
    baseUrl: config.customBaseUrl,
    reasoning: true,
    input: ["text", "image"] as ("text" | "image")[],
    cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
    contextWindow: 128000,
    maxTokens: 32000,
  };
}

export function applyProxyToModel(
  model: Model<Api>,
  config: ProviderConfig,
): Model<Api> {
  if (!config.useProxy || !config.proxyUrl || !model.baseUrl) return model;
  return {
    ...model,
    baseUrl: `${config.proxyUrl}/?url=${encodeURIComponent(model.baseUrl)}`,
  };
}
