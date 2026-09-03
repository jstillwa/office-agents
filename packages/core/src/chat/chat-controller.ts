import type { Api, Model } from "@earendil-works/pi-ai";
import {
  type AgentContext,
  AgentRuntime,
  emitTelemetry,
  type ProviderConfig,
  type RuntimeState,
  type SessionStats,
} from "@office-agents/sdk";
import { get, type Writable, writable } from "svelte/store";
import type { AppAdapter } from "./app-adapter";

export interface DiagnosticError {
  message: string;
  stack?: string;
  timestamp: number;
}

export interface DiagnosticsBundle {
  timestamp: string;
  hostPlatform: {
    appName?: string;
    appVersion?: string;
    userAgent?: string;
    office?: {
      host?: string;
      platform?: string;
      version?: string;
    } | null;
  };
  sessionStats: {
    sessionId?: string;
    sessionName?: string;
    messageCount: number;
    stats: SessionStats;
    provider?: string;
    model?: string;
  };
  recentErrors: DiagnosticError[];
}

export class ChatController {
  readonly context: AgentContext;
  readonly state: Writable<RuntimeState>;
  adapter: AppAdapter;
  #runtime: AgentRuntime;
  #unsubscribe: (() => void) | null = null;
  #recentErrors: DiagnosticError[] = [];
  #boundWindowError: ((event: ErrorEvent) => void) | null = null;
  #boundUnhandledRejection: ((event: PromiseRejectionEvent) => void) | null =
    null;

  constructor(adapter: AppAdapter, context: AgentContext) {
    this.adapter = adapter;
    this.context = context;
    this.#runtime = new AgentRuntime(adapter, this.context);
    this.state = writable(this.#runtime.getState());
    this.#unsubscribe = this.#runtime.subscribe((next) => this.state.set(next));

    this.#runtime.init().catch((err: unknown) => {
      this.handleUncaughtError(err);
    });

    if (typeof window !== "undefined") {
      this.#boundWindowError = (event: ErrorEvent) => {
        const error = event.error || new Error(event.message || "Window error");
        this.handleUncaughtError(error);
      };
      this.#boundUnhandledRejection = (event: PromiseRejectionEvent) => {
        const reason = event.reason;
        const error =
          reason instanceof Error ? reason : new Error(String(reason));
        this.handleUncaughtError(error);
      };
      window.addEventListener("error", this.#boundWindowError);
      window.addEventListener(
        "unhandledrejection",
        this.#boundUnhandledRejection,
      );
    }
  }

  get recentErrors(): DiagnosticError[] {
    return [...this.#recentErrors];
  }

  handleUncaughtError(error: unknown) {
    const err = error instanceof Error ? error : new Error(String(error));
    const message = err.message || "An unexpected error occurred";
    this.#runtime.setError(message);
    this.#recentErrors.push({
      message,
      stack: err.stack,
      timestamp: Date.now(),
    });
    if (this.#recentErrors.length > 50) {
      this.#recentErrors.shift();
    }
    emitTelemetry(this.adapter.telemetrySink, {
      type: "uncaught_error",
      message,
      stack: err.stack,
      timestamp: Date.now(),
    });
  }

  exportDiagnostics(): DiagnosticsBundle {
    const snapshot = this.snapshot;
    let officeDiag: {
      host?: string;
      platform?: string;
      version?: string;
    } | null = null;
    if (typeof Office !== "undefined" && Office?.context?.diagnostics) {
      const diag = Office.context.diagnostics;
      officeDiag = {
        host: diag.host != null ? String(diag.host) : undefined,
        platform: diag.platform != null ? String(diag.platform) : undefined,
        version: diag.version != null ? String(diag.version) : undefined,
      };
    }

    return {
      timestamp: new Date().toISOString(),
      hostPlatform: {
        appName: this.adapter.appName,
        appVersion: this.adapter.appVersion,
        userAgent:
          typeof navigator !== "undefined" ? navigator.userAgent : undefined,
        office: officeDiag,
      },
      sessionStats: {
        sessionId: snapshot.currentSession?.id,
        sessionName: snapshot.currentSession?.name,
        messageCount: snapshot.messages.length,
        stats: snapshot.sessionStats,
        provider: snapshot.providerConfig?.provider,
        model: snapshot.providerConfig?.model,
      },
      recentErrors: [...this.#recentErrors],
    };
  }

  get snapshot() {
    return get(this.state);
  }

  get availableProviders() {
    return this.#runtime.getAvailableProviders();
  }

  setAdapter(adapter: AppAdapter) {
    this.adapter = adapter;
    this.#runtime.setAdapter(adapter);
  }

  dispose() {
    if (typeof window !== "undefined") {
      if (this.#boundWindowError) {
        window.removeEventListener("error", this.#boundWindowError);
        this.#boundWindowError = null;
      }
      if (this.#boundUnhandledRejection) {
        window.removeEventListener(
          "unhandledrejection",
          this.#boundUnhandledRejection,
        );
        this.#boundUnhandledRejection = null;
      }
    }
    this.#unsubscribe?.();
    this.#unsubscribe = null;
    this.#runtime.dispose();
  }

  getUserId() {
    return this.#runtime.getUserId();
  }

  getModelsForProvider(provider: string): Model<Api>[] {
    return this.#runtime.getModelsForProvider(provider);
  }

  sendMessage(content: string, attachments?: string[]) {
    return this.#runtime.sendMessage(content, attachments);
  }

  setProviderConfig(config: ProviderConfig) {
    this.#runtime.setProviderConfig(config);
  }

  reloadMcpTools() {
    return this.#runtime.reloadMcpTools();
  }

  clearMessages() {
    this.#runtime.clearMessages();
  }

  abort() {
    this.#runtime.abort();
  }

  newSession() {
    return this.#runtime.newSession();
  }

  switchSession(sessionId: string) {
    return this.#runtime.switchSession(sessionId);
  }

  deleteCurrentSession() {
    return this.#runtime.deleteCurrentSession();
  }

  getName(id: number) {
    return this.#runtime.getName(id);
  }

  toggleFollowMode() {
    this.#runtime.toggleFollowMode();
  }

  toggleExpandToolCalls() {
    this.#runtime.toggleExpandToolCalls();
  }

  async processFiles(files: File[]) {
    if (files.length === 0) return;

    const inputs = await Promise.all(
      files.map(async (file) => ({
        name: file.name,
        size: file.size,
        data: new Uint8Array(await file.arrayBuffer()),
      })),
    );
    await this.#runtime.uploadFiles(inputs);
  }

  removeUpload(name: string) {
    return this.#runtime.removeUpload(name);
  }

  async installSkill(files: File[]) {
    if (files.length === 0) return;

    const inputs = await Promise.all(
      files.map(async (file) => {
        const fullPath = file.webkitRelativePath || file.name;
        const parts = fullPath.split("/");
        const path = parts.length > 1 ? parts.slice(1).join("/") : parts[0];
        return { path, data: new Uint8Array(await file.arrayBuffer()) };
      }),
    );

    await this.#runtime.installSkill(inputs);
  }

  uninstallSkill(name: string) {
    return this.#runtime.uninstallSkill(name);
  }
}
