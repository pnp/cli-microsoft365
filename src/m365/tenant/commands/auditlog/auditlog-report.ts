import * as chalk from 'chalk';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  contentType: string;
  startTime?: string;
  endTime?: string;
}

interface ActivityfeedSubscription {
  contentType: string;
  status: string;
  webhook: string;
}

interface AuditContentList {
  contentType: string;
  contentId: string;
  contentUri: string;
  contentCreated: string;
  contentExpiration: string;
}

interface AuditlogReport {
  CreationTime: string;
  Id: string;
  Workload: string;
  Operation: string;
  UserId: string;
}

enum AuditContentTypes {
  AzureActiveDirectory = "Audit.AzureActiveDirectory",
  Exchange = "Audit.Exchange",
  SharePoint = "Audit.SharePoint",
  General = "Audit.General",
  DLP = "DLP.All"
}

class TenantAuditlogReportCommand extends Command {
  private serviceUrl: string = 'https://manage.office.com/api/v1.0';
  private tenantId: string | undefined;
  private completeAuditReports: AuditlogReport[] = [];

  public get name(): string {
    return `${commands.TENANT_AUDITLOG_REPORT}`;
  }

  public get description(): string {
    return 'Gets audit logs from the Office 365 Management API';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.startTime = typeof args.options.startTime !== 'undefined';
    telemetryProps.endTime = typeof args.options.endTime !== 'undefined';
    telemetryProps.contentType = args.options.contentType;
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['CreationTime', 'Operation', 'ClientIP', 'UserId', 'Workload'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      logger.logToStderr(`Start retrieving Audit Log Report`);
    }

    this.tenantId = Utils.getTenantIdFromAccessToken(auth.service.accessTokens[auth.defaultResource].value);
    this
      .getCompleteAuditReports(args, logger)
      .then((res: AuditlogReport[]): void => {
        logger.log(res);

        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getCompleteAuditReports(args: CommandArgs, logger: Logger): Promise<AuditlogReport[]> {
    return this
      .startContentSubscriptionIfNotActive(args, logger)
      .then((): Promise<AuditContentList[]> => this.getAuditContentList(args, logger))
      .then((auditContentLists: AuditContentList[]): Promise<Promise<AuditlogReport[]>[][]> => this.getBatchedPromises(auditContentLists, 10))
      .then((batchedPromise: Promise<AuditlogReport[]>[][]): Promise<void> => {
        return new Promise<void>((resolve: () => void, reject: (err: any) => void): void => {
          if (batchedPromise.length > 0) {
            this.getBatchedAuditlogData(logger, batchedPromise, 0, resolve, reject);
          }
          else {
            resolve();
          }
        });
      })
      .then(_ => this.completeAuditReports);
  }

  private startContentSubscriptionIfNotActive(args: CommandArgs, logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Checking if subscription is active...`);
    }

    const subscriptionListEndpoint: string = 'activity/feed/subscriptions/list';
    const requestOptions: any = {
      url: `${this.serviceUrl}/${this.tenantId}/${subscriptionListEndpoint}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<ActivityfeedSubscription[]>(requestOptions)
      .then((subscriptionLists: ActivityfeedSubscription[]): boolean => {
        return subscriptionLists.some(subscriptionList =>
          subscriptionList.contentType === (<any>AuditContentTypes)[args.options.contentType] &&
            subscriptionList.status === 'enabled');
      })
      .then((hasActiveSubscription: boolean): Promise<void> => {
        if (hasActiveSubscription) {
          return Promise.resolve();
        }

        if (this.verbose) {
          logger.logToStderr(`Starting subscription since subscription is not active for the content type`);
        }

        const startSubscriptionEndPoint: string = `activity/feed/subscriptions/start?contentType=${(<any>AuditContentTypes)[args.options.contentType]}&PublisherIdentifier=${this.tenantId}`;
        const requestOptions: any = {
          url: `${this.serviceUrl}/${this.tenantId}/${startSubscriptionEndPoint}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      });
  }

  private getAuditContentList(args: CommandArgs, logger: Logger): Promise<AuditContentList[]> {
    if (this.verbose) {
      logger.logToStderr(`Start listing Audit Content URL`);
    }

    let subscriptionListEndpoint: string = `activity/feed/subscriptions/content?contentType=${(<any>AuditContentTypes)[args.options.contentType]}&PublisherIdentifier=${this.tenantId}`;

    if (typeof args.options.startTime !== 'undefined') {
      if (typeof args.options.endTime !== 'undefined') {
        subscriptionListEndpoint += `&starttime=${escape(args.options.startTime)}&endTime=${escape(args.options.endTime)}`;
      }
      else {
        const parsedEndDate: Date = new Date(args.options.startTime);
        parsedEndDate.setDate(parsedEndDate.getDate() + 1);
        subscriptionListEndpoint += `&starttime=${escape(args.options.startTime)}&endTime=${escape(parsedEndDate.toISOString())}`;
      }
    }

    const requestOptions: any = {
      url: `${this.serviceUrl}/${this.tenantId}/${subscriptionListEndpoint}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<AuditContentList[]>(requestOptions)
  }

  private getBatchedPromises(auditContentLists: AuditContentList[], batchSize: number): Promise<Promise<AuditlogReport[]>[][]> {
    const batchedPromises: Promise<AuditlogReport[]>[][] = [];

    for (let i: number = 0; i < auditContentLists.length; i += batchSize) {
      const promiseRequestBatch: Promise<AuditlogReport[]>[] = auditContentLists
        .slice(i, i + batchSize < auditContentLists.length ? i + batchSize : auditContentLists.length)
        .map((AuditContentList: AuditContentList) => this.getAuditLogReportForSingleContentUrl(AuditContentList.contentUri))

      batchedPromises.push(promiseRequestBatch);
    }

    return Promise.resolve(batchedPromises);
  }

  private getBatchedAuditlogData(logger: Logger, batchedPromiseList: Promise<AuditlogReport[]>[][], batchNumber: number, resolve: () => void, reject: (err: any) => void): void {
    if (this.verbose) {
      logger.logToStderr(`Starting Batch : ${batchNumber}`);
    }

    Promise
      .all(batchedPromiseList[batchNumber])
      .then((data: AuditlogReport[][]) => {
        data.forEach(d1 => {
          d1.forEach(d2 => {
            this.completeAuditReports.push(d2);
          });
        });

        if (batchNumber < batchedPromiseList.length - 1) {
          this.getBatchedAuditlogData(logger, batchedPromiseList, ++batchNumber, resolve, reject)
        }
        else {
          resolve();
        }
      }, (err: any): void => reject(err));
  }

  private getAuditLogReportForSingleContentUrl(auditURL: string): Promise<AuditlogReport[]> {
    const requestOptions: any = {
      url: auditURL,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    return request.get<AuditlogReport[]>(requestOptions);
  }

  protected get auditContentTypeLists(): string[] {
    const result: string[] = [];

    for (let auditContentType in AuditContentTypes) {
      result.push(auditContentType);
    }

    return result;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-c, --contentType <contentType>',
        description: `Audit content type of logs to be retrieved, should be one of the following: ${this.auditContentTypeLists.join(', ')}`,
        autocomplete: this.auditContentTypeLists
      },
      {
        option: '-s, --startTime [startTime]',
        description: 'Start time of logs to be retrieved. Start time and end time must both be specified (or both omitted) and must be less than or equal to 24 hours apart.'
      },
      {
        option: '-e, --endTime [endTime]',
        description: 'End time of logs to be retrieved. Start time and end time must both be specified (or both omitted) and must be less than or equal to 24 hours apart.'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if ((<any>AuditContentTypes)[args.options.contentType] === undefined) {
      return `${args.options.contentType} is not a valid value for the contentType option. Allowed values are ${this.auditContentTypeLists.join(' | ')}`;
    }

    if (args.options.startTime || args.options.endTime) {
      if (!args.options.startTime) {
        return `Please specify startTime`;
      }

      const parsedStartTime = Date.parse(args.options.startTime);
      if (isNaN(parsedStartTime)) {
        return `${args.options.startTime} is not a valid startTime. Provide the date in one of the following formats:
'YYYY-MM-DD'
'YYYY-MM-DDThh:mm'
'YYYY-MM-DDThh:mmZ'
'YYYY-MM-DDThh:mm±hh:mm'`;
      }

      const startdateTodayDifference: number = (new Date().getTime() - parsedStartTime) / (1000 * 60 * 60 * 24);
      if (startdateTodayDifference > 7) {
        return `Start time should be no more than 7 days in the past`;
      }

      if (args.options.endTime) {
        const parsedEndTime: number = Date.parse(args.options.endTime);
        if (isNaN(parsedEndTime)) {
          return `${args.options.endTime} is not a valid endTime. Provide the date in one of the following formats:
'YYYY-MM-DD'
'YYYY-MM-DDThh:mm'
'YYYY-MM-DDThh:mmZ'
'YYYY-MM-DDThh:mm±hh:mm'`;
        }

        if (parsedStartTime && parsedEndTime) {
          const startEndDateDifference: number = (parsedEndTime - parsedStartTime) / (1000 * 60 * 60 * 24);
          if (startEndDateDifference < 0 || startEndDateDifference > 1) {
            return `startTime and endTime must be less than or equal to 24 hours apart, with the start time prior to end time.`;
          }
        }
      }

    }

    return true;
  }
}

module.exports = new TenantAuditlogReportCommand();