import type { OpenClawPluginApi } from "openclaw/plugin-sdk";
import { emptyPluginConfigSchema } from "openclaw/plugin-sdk";
import { msteamsUserPlugin } from "./src/channel.js";
import { setRuntime } from "./src/runtime.js";

const plugin = {
  id: "msteams-user",
  name: "Microsoft Teams (User)",
  description: "Microsoft Teams channel plugin (User Account via Graph API)",
  configSchema: emptyPluginConfigSchema(),
  register(api: OpenClawPluginApi) {
    setRuntime(api.runtime);
    api.registerChannel({ plugin: msteamsUserPlugin });
  },
};

export default plugin;
