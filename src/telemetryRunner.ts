import appInsights from './appInsights.js';
import * as process from 'process';
import * as fs from 'fs';

process.stdin.setEncoding('utf8');

try {
  // read from stdin
  const input = fs.readFileSync(0, 'utf-8');
  const data = JSON.parse(input);
  const { commandName, properties, exception, shell, session } = data;

  appInsights.commonProperties.shell = shell;
  appInsights.context.tags[appInsights.context.keys.sessionId] = session;

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