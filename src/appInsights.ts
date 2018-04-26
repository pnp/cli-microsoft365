const packageJSON = require('../package.json');
import * as fs from 'fs';
import * as path from 'path';
import * as crypto from 'crypto';

import * as appInsights from 'applicationinsights';
const config = appInsights.setup('6b908c80-d09f-4cf6-8274-e54349a0149a');
config.setInternalLogging(false, false);
appInsights.start();
// append -dev to the version number when ran locally
// to distinguish production and dev version of the CLI
// in the telemetry
const version: string = `${packageJSON.version}${fs.existsSync(path.join(__dirname, `..${path.sep}src`)) ? '-dev' : ''}`;
appInsights.defaultClient.commonProperties = {
  version: version
};
appInsights.defaultClient.context.tags['ai.session.id'] = crypto.randomBytes(24).toString('base64');
appInsights.defaultClient.context.tags['ai.user.id'] = getUserId();

export function getUserId(): string {
  const filePath: string = path.join(__dirname, `..${path.sep}.user`);
  let userId: string = '';

  try {
    if (fs.existsSync(filePath)) {
      userId = fs.readFileSync(filePath, 'utf-8');
    }
  }
  catch { }

  if (!userId) {
    userId = crypto.randomBytes(24).toString('base64');
    try {
      fs.writeFileSync(filePath, userId);
    }
    catch { }
  }

  return userId;
}

export default appInsights.defaultClient;