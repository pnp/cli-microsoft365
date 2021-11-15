const packageJSON = require('../package.json');
// disable automatic third-party instrumentation for Application Insights
// speeds up execution by preventing loading unnecessary dependencies
process.env.APPLICATION_INSIGHTS_NO_DIAGNOSTIC_CHANNEL = 'none';
// prevents tests from hanging
process.env.APPLICATION_INSIGHTS_NO_STATSBEAT = 'true';
import * as appInsights from 'applicationinsights';
import * as crypto from 'crypto';
import * as fs from 'fs';
import * as path from 'path';

const config = appInsights.setup('6b908c80-d09f-4cf6-8274-e54349a0149a');
config.setInternalLogging(false, false);
// append -dev to the version number when ran locally
// to distinguish production and dev version of the CLI
// in the telemetry
const version: string = `${packageJSON.version}${fs.existsSync(path.join(__dirname, `..${path.sep}src`)) ? '-dev' : ''}`;
const env: string = process.env.CLIMICROSOFT365_ENV !== undefined ? process.env.CLIMICROSOFT365_ENV : '';
appInsights.defaultClient.commonProperties = {
  version: version,
  node: process.version,
  env: env
};
appInsights.defaultClient.context.tags['ai.session.id'] = crypto.randomBytes(24).toString('base64');
delete appInsights.defaultClient.context.tags['ai.cloud.roleInstance'];
delete appInsights.defaultClient.context.tags['ai.cloud.role'];

export default appInsights.defaultClient;