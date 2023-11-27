// disable automatic third-party instrumentation for Application Insights
// speeds up execution by preventing loading unnecessary dependencies
process.env.APPLICATION_INSIGHTS_NO_DIAGNOSTIC_CHANNEL = 'none';
// prevents tests from hanging
process.env.APPLICATION_INSIGHTS_NO_STATSBEAT = 'true';
import * as appInsights from 'applicationinsights';
import crypto from 'crypto';
import fs from 'fs';
import path from 'path';
import url from 'url';
import { app } from './utils/app.js';

const __dirname = url.fileURLToPath(new URL('.', import.meta.url));
const config = appInsights.setup('6b908c80-d09f-4cf6-8274-e54349a0149a');
config.setInternalLogging(false, false);
// append -dev to the version number when ran locally
// to distinguish production and dev version of the CLI
// in the telemetry
const version: string = `${app.packageJson().version}${fs.existsSync(path.join(__dirname, `..${path.sep}src`)) ? '-dev' : ''}`;
const env: string = process.env.CLIMICROSOFT365_ENV !== undefined ? process.env.CLIMICROSOFT365_ENV : '';
const appInsightsClient: appInsights.TelemetryClient = ((appInsights as any).default as typeof appInsights).defaultClient;
appInsightsClient.commonProperties = {
  version: version,
  node: process.version,
  env: env,
  ci: Boolean(process.env.CI).toString()
};

appInsightsClient.config.proxyHttpUrl = process.env.HTTP_PROXY ?? '';
appInsightsClient.config.proxyHttpsUrl = process.env.HTTPS_PROXY ?? '';

appInsightsClient.context.tags['ai.cloud.roleInstance'] = crypto.createHash('sha256').update(appInsightsClient.context.tags['ai.cloud.roleInstance']).digest('hex');
delete appInsightsClient.context.tags['ai.cloud.role'];
delete appInsightsClient.context.tags['ai.cloud.roleName'];

export default appInsightsClient;