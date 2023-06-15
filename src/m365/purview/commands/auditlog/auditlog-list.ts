import Auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { accessToken } from '../../../../utils/accessToken';
import { validation } from '../../../../utils/validation';
import O365MgmtCommand from '../../../base/O365MgmtCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  contentType: string;
  startTime?: string;
  endTime?: string;
}

class PurviewAuditLogListCommand extends O365MgmtCommand {
  private readonly contentTypeOptions = ['AzureActiveDirectory', 'Exchange', 'SharePoint', 'General', 'DLP'];

  public get name(): string {
    return commands.AUDITLOG_LIST;
  }

  public get description(): string {
    return 'List audit logs within your tenant';
  }

  public defaultProperties(): string[] | undefined {
    return ['CreationTime', 'UserId', 'Operation', 'ObjectId'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        startTime: typeof args.options.startTime !== 'undefined',
        endTime: typeof args.options.endTime !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--contentType <contentType>',
        autocomplete: this.contentTypeOptions
      },
      {
        option: '--startTime [startTime]'
      },
      {
        option: '--endTime [endTime]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (this.contentTypeOptions.indexOf(args.options.contentType) === -1) {
          return `'${args.options.contentType}' is not a valid contentType value. Allowed values: ${this.contentTypeOptions}.`;
        }

        if (args.options.startTime) {
          if (!validation.isValidISODateTime(args.options.startTime)) {
            return `'${args.options.startTime}' is not a valid ISO date time string.`;
          }

          const lowerDateLimit = new Date();
          lowerDateLimit.setDate(lowerDateLimit.getDate() - 7);
          lowerDateLimit.setHours(lowerDateLimit.getHours() - 1); // Min date is 7 days ago, however there seems to be an 1h margin
          if (new Date(args.options.startTime) < lowerDateLimit) {
            return 'startTime value cannot be more than 7 days in the past.';
          }
        }

        if (args.options.endTime) {
          if (!validation.isValidISODateTime(args.options.endTime)) {
            return `'${args.options.endTime}' is not a valid ISO date time string.`;
          }

          if (new Date(args.options.endTime) > new Date()) {
            return 'endTime value cannot be in the future.';
          }
        }

        if (args.options.startTime && args.options.endTime && new Date(args.options.startTime) >= new Date(args.options.endTime)) {
          return 'startTime value must be before endTime.';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    // If we don't create a now object, start and end date can be an few extra ms apart due to execution time between code lines
    const now = new Date();
    try {
      let startTime: Date;
      if (args.options.startTime) {
        startTime = new Date(args.options.startTime);
      }
      else {
        startTime = new Date(now);
        startTime.setDate(startTime.getDate() - 1);
      }
      const endTime = args.options.endTime ? new Date(args.options.endTime) : new Date(now);

      if (this.verbose) {
        logger.logToStderr(`Getting audit logs for content type '${args.options.contentType}' within a time frame from '${startTime.toISOString()}' to '${endTime.toISOString()}'.`);
      }

      const tenantId = accessToken.getTenantIdFromAccessToken(Auth.service.accessTokens[Auth.defaultResource].accessToken);
      const contentTypeValue = args.options.contentType === 'DLP' ? 'DLP.All' : 'Audit.' + args.options.contentType;

      await this.ensureSubscription(tenantId, contentTypeValue);
      if (this.verbose) {
        logger.logToStderr(`'${args.options.contentType}' subscription is active.`);
      }

      const contentUris: string[] = [];
      for (const time: Date = startTime; time < endTime; time.setDate(time.getDate() + 1)) {
        const differenceInMs = endTime.getTime() - time.getTime();
        const endTimeBatch = new Date(time.getTime() + Math.min(differenceInMs, 1000 * 60 * 60 * 24)); // ms difference cannot be greater than 1 day

        if (this.verbose) {
          logger.logToStderr(`Get content URIs for date range from '${time.toISOString()}' to '${endTimeBatch.toISOString()}'.`);
        }

        const contentUrisBatch = await this.getContentUris(tenantId, contentTypeValue, time, endTimeBatch);
        contentUris.push(...contentUrisBatch);
      }

      if (this.verbose) {
        logger.logToStderr(`Get content from ${contentUris.length} content URIs.`);
      }

      const logs = await this.getContent(logger, contentUris);
      const sortedLogs = logs.sort(this.auditLogsCompare);

      logger.log(sortedLogs);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async ensureSubscription(tenantId: string, contentType: string): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/api/v1.0/${tenantId}/activity/feed/subscriptions/list`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };
    const subscriptions = await request.get<{ contentType: string; status: string }[]>(requestOptions);

    if (subscriptions.some(s => s.contentType === contentType && s.status === 'enabled')) {
      return;
    }

    requestOptions.url = `${this.resource}/api/v1.0/${tenantId}/activity/feed/subscriptions/start?contentType=${contentType}`;
    const subscription = await request.post<{ status: string }>(requestOptions);

    if (subscription.status !== 'enabled') {
      throw `Unable to start subscription '${contentType}'`;
    }
  }

  private async getContentUris(tenantId: string, contentType: string, startTime: Date, endTime: Date): Promise<string[]> {
    const contentUris: string[] = [];
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/api/v1.0/${tenantId}/activity/feed/subscriptions/content?contentType=${contentType}&startTime=${startTime.toISOString()}&endTime=${endTime.toISOString()}`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json',
      fullResponse: true
    };

    do {
      const response = await request.get<{ headers: any, data: { contentUri: string }[] }>(requestOptions);

      const uris = response.data.map(d => d.contentUri);
      contentUris.push(...uris);

      requestOptions.url = response.headers.nextpageuri;
    } while (requestOptions.url);

    return contentUris;
  }

  private async getContent(logger: Logger, contentUris: string[]): Promise<any[]> {
    const logs: any[] = [];

    const batchSize = 30;
    for (let i = 0; i < contentUris.length; i += batchSize) {
      const contentUrisBatch = contentUris.slice(i, i + batchSize);

      if (this.verbose) {
        logger.logToStderr(`Retrieving content from next ${contentUrisBatch.length} content URIs. Progress: ${Math.round(i / contentUris.length * 100)}%`);
      }

      const batchResult = await Promise.all(
        contentUrisBatch.map(uri => request.get<any[]>(
          {
            url: uri,
            headers: {
              accept: 'application/json'
            },
            responseType: 'json'
          }
        ))
      );

      batchResult.forEach(res => logs.push(...res));
    }

    return logs;
  }

  private auditLogsCompare(a: any, b: any): -1 | 0 | 1 {
    if (a.CreationTime < b.CreationTime) {
      return -1;
    }
    if (a.CreationTime > b.CreationTime) {
      return 1;
    }
    return 0;
  }
}

module.exports = new PurviewAuditLogListCommand();