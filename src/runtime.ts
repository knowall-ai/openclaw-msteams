import type { PluginRuntime } from "openclaw/plugin-sdk";

let runtime: PluginRuntime | null = null;

export function setRuntime(next: PluginRuntime) {
  runtime = next;
}

export function getRuntime(): PluginRuntime {
  if (!runtime) {
    throw new Error("msteams-user runtime not initialized");
  }
  return runtime;
}
