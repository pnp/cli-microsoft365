const packageJSON = require('../package.json');

import * as appInsights from 'applicationinsights';
const config = appInsights.setup('6b908c80-d09f-4cf6-8274-e54349a0149a');
config.setInternalLogging(false, false);
appInsights.start();
appInsights.defaultClient.commonProperties = {
  version: packageJSON.version
};

export default appInsights.defaultClient;