import { cli } from './cli/cli.js';
import { googleAnalytics } from './googleAnalytics.js';
import { settingsNames } from './settingsNames.js';
import { pid } from './utils/pid.js';
import { session } from './utils/session.js';

async function trackTelemetry(object: any): Promise<void> {
  try {
    const { commandName, properties, exception } = object;
    const context = {
      shell: pid.getProcessName(process.ppid) || '',
      sessionId: session.getId(process.ppid)
    };

    if (exception) {
      await googleAnalytics.trackException(exception, context);
    }
    else {
      await googleAnalytics.trackEvent(commandName, properties, context);
    }
  }
  catch {
    // Do nothing
  }
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