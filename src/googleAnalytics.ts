import Axios from 'axios';
import Configstore from 'configstore';
import crypto from 'crypto';
import fs from 'fs';
import os from 'os';
import path from 'path';
import url from 'url';
import config from './config.js';
import { app } from './utils/app.js';

const measurementId = 'G-4BNT8MQCYT';
const apiSecret = 'mm_3WD_TRuO-9MKsuZnhDQ';
const endpoint = 'https://www.google-analytics.com/mp/collect';
const clientIdSetting = 'telemetryClientId';
const maxParameterValueLength = 100;
const requestTimeout = 1000;
const __dirname = url.fileURLToPath(new URL('.', import.meta.url));

export interface TelemetryContext {
  sessionId: string;
  shell: string;
}

interface GoogleAnalyticsEvent {
  name: string;
  params: Record<string, string | number>;
}

const commonProperties: Record<string, string> = {
  version: `${app.packageJson().version}${fs.existsSync(path.join(__dirname, `..${path.sep}src`)) ? '-dev' : ''}`,
  node: process.version,
  os: os.platform(),
  env: process.env.CLIMICROSOFT365_ENV ?? '',
  ci: Boolean(process.env.CI).toString()
};

function getClientId(): string {
  const store = new Configstore(config.configstoreName);
  let clientId = store.get(clientIdSetting) as string | undefined;

  if (!clientId) {
    clientId = crypto.randomUUID();
    store.set(clientIdSetting, clientId);
  }

  return clientId;
}

function getSessionId(sessionId: string): number {
  return parseInt(crypto.createHash('sha256').update(sessionId).digest('hex').substring(0, 12), 16);
}

function toEventValue(value: unknown): string | number {
  if (typeof value === 'number') {
    return value;
  }

  return String(value);
}

function toOptionName(name: string): string {
  if (name.length <= maxParameterValueLength) {
    return name;
  }

  const hash = crypto.createHash('sha256').update(name).digest('hex').substring(0, 8);
  return `${name.substring(0, maxParameterValueLength - hash.length - 1)}_${hash}`;
}

function getEvents(commandName: string, properties: Record<string, unknown>, context: TelemetryContext): GoogleAnalyticsEvent[] {
  const sessionId = getSessionId(context.sessionId);
  const events: GoogleAnalyticsEvent[] = [{
    name: 'command_used',
    params: {
      ['command_name']: commandName,
      ...commonProperties,
      shell: context.shell,
      ['session_id']: sessionId,
      ['engagement_time_msec']: 1
    }
  }];

  Object.entries(properties).forEach(([name, value]) => {
    if (typeof value === 'undefined') {
      return;
    }

    events.push({
      name: 'command_option_used',
      params: {
        ['command_name']: commandName,
        ['option_name']: toOptionName(name),
        ['option_value']: toEventValue(value),
        ['session_id']: sessionId,
        ['engagement_time_msec']: 1
      }
    });
  });

  return events;
}

async function sendEvents(events: GoogleAnalyticsEvent[]): Promise<void> {
  const clientId = getClientId();
  await Promise.all(events.map(async event => {
    const response = await Axios.post(`${endpoint}?measurement_id=${measurementId}&api_secret=${apiSecret}`, {
      ['client_id']: clientId,
      events: [event]
    }, {
      headers: {
        'Content-Type': 'application/json'
      },
      timeout: requestTimeout,
      validateStatus: () => true
    });

    if (response.status < 200 || response.status >= 300) {
      throw new Error(`Google Analytics returned ${response.status}`);
    }
  }));
}

export const googleAnalytics = {
  commonProperties,
  trackEvent: async (commandName: string, properties: Record<string, unknown>, context: TelemetryContext): Promise<void> => {
    await sendEvents(getEvents(commandName, properties, context));
  },
  trackException: async (exception: unknown, context: TelemetryContext): Promise<void> => {
    const description = exception instanceof Error ? exception.message : String(exception);
    await sendEvents([{
      name: 'exception',
      params: {
        description: description.substring(0, 100),
        ...commonProperties,
        shell: context.shell,
        ['session_id']: getSessionId(context.sessionId),
        ['engagement_time_msec']: 1
      }
    }]);
  }
};