import fs from 'fs';
import https from 'https';
import path from 'path';
import { ExpressAdapter, App } from '@microsoft/teams.apps';
import { ConsoleLogger } from '@microsoft/teams.common/logging';
import { DevtoolsPlugin } from '@microsoft/teams.dev';

const sslOptions = {
  key: process.env.SSL_KEY_FILE ? fs.readFileSync(process.env.SSL_KEY_FILE) : void 0,
  cert: process.env.SSL_CRT_FILE ? fs.readFileSync(process.env.SSL_CRT_FILE) : void 0
};
const adapter = new ExpressAdapter();
if (sslOptions.cert && sslOptions.key) {
  const httpsServer = https.createServer(sslOptions, adapter.express);
  adapter.server = httpsServer;
}
const plugins = [];
if (process.env.SSL_KEY_FILE) {
  plugins.push(new DevtoolsPlugin());
}
const app = new App({
  logger: new ConsoleLogger("tab", { level: "debug" }),
  plugins,
  httpServerAdapter: adapter,
  skipAuth: true
});
app.tab("home", path.join(__dirname, "./client"));
(async () => {
  await app.start(process.env.PORT || 3978);
})();
//# sourceMappingURL=index.mjs.map
//# sourceMappingURL=index.mjs.map