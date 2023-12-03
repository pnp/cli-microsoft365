import child_process from 'child_process';
import path from 'path';
import url from 'url';
import { Cli } from './cli/Cli.js';
import { settingsNames } from './settingsNames.js';
import { pid } from './utils/pid.js';
import { session } from './utils/session.js';

const __dirname = url.fileURLToPath(new URL('.', import.meta.url));

function trackTelemetry(object: any): void {
  try {
    const child = child_process.spawn('node', [path.join(__dirname, 'telemetryRunner.js')], {
      stdio: ['pipe', 'ignore', 'ignore'],
      detached: true
    });
    child.unref();

    object.shell = pid.getProcessName(process.ppid) || '';
    object.session = session.getId(process.ppid);

    child.stdin.write(JSON.stringify(object));
    child.stdin.end();
  }
  catch { }
}

export const telemetry = {
  trackEvent: (commandName: string, properties: any): void => {
    if (Cli.getInstance().getSettingWithDefaultValue<boolean>(settingsNames.disableTelemetry, false)) {
      return;
    }

    trackTelemetry({
      commandName,
      properties
    });
  }
};