import appInsights from './appInsights.js';
import { cli } from './cli/cli.js';
import { settingsNames } from './settingsNames.js';
import { pid } from './utils/pid.js';
import { session } from './utils/session.js';

async function trackTelemetry(object: any): Promise<void> {
  try {
    const { commandName, properties, exception } = object;

    appInsights.commonProperties.shell = pid.getProcessName(process.ppid) || '';
    appInsights.context.tags[appInsights.context.keys.sessionId] = session.getId(process.ppid);

    if (exception) {
      appInsights.trackException({
        exception
      });
    }
    else {
      appInsights.trackEvent({
        name: commandName,
        properties
      });
    }
    await appInsights.flush();
  }
  catch { }
}

export const telemetry = {
  trackEvent: async (commandName: string, properties: any, exception?: any): Promise<void> => {
    if (cli.getSettingWithDefaultValue<boolean>(settingsNames.disableTelemetry, false)) {
      return;
    }

    await trackTelemetry({
      commandName,
      properties,
      exception
    });
  }
};