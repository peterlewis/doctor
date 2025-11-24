const packageJSON = require("../../package.json");
// disable automatic third-party instrumentation for Application Insights
// speeds up execution by preventing loading unnecessary dependencies
process.env.APPLICATION_INSIGHTS_NO_DIAGNOSTIC_CHANNEL = "none";
import * as appInsights from "applicationinsights";
import * as crypto from "crypto";
import * as fs from "fs";
import * as path from "path";

const noopClient = {
  trackEvent: () => {},
  trackMetric: () => {},
  trackException: () => {},
  trackTrace: () => {},
  flush: (cb?: () => void) => {
    if (cb) cb();
  },
};

// Only enable telemetry if the caller provided a connection string or key.
const connectionString =
  process.env.APPLICATIONINSIGHTS_CONNECTION_STRING ||
  (process.env.APPLICATIONINSIGHTS_INSTRUMENTATIONKEY
    ? `InstrumentationKey=${process.env.APPLICATIONINSIGHTS_INSTRUMENTATIONKEY}`
    : null);

let client: appInsights.TelemetryClient | typeof noopClient = noopClient;

if (connectionString) {
  try {
    const config = appInsights.setup(connectionString);
    config.setInternalLogging(false, false);
    appInsights.start();

    if (appInsights.defaultClient) {
      // append -dev to the version number when ran locally to distinguish production and dev CLI
      const version: string = `${packageJSON.version}${
        fs.existsSync(path.join(__dirname, `..${path.sep}..${path.sep}src`))
          ? "-dev"
          : ""
      }`;

      appInsights.defaultClient.commonProperties = {
        version: version,
        node: process.version,
      };
      appInsights.defaultClient.context.tags["ai.session.id"] =
        crypto.randomBytes(24).toString("base64");
      delete appInsights.defaultClient.context.tags["ai.cloud.roleInstance"];
      delete appInsights.defaultClient.context.tags["ai.cloud.role"];

      client = appInsights.defaultClient;
    }
  } catch (err) {
    // If telemetry init fails, continue with no-op client to avoid crashing the CLI.
    client = noopClient;
  }
}

export default client;
