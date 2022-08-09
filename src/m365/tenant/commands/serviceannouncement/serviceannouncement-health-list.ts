import { ServiceHealth } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  issues?: boolean;
}

class TenantServiceAnnouncementHealthListCommand extends GraphCommand {
  public get name(): string {
    return commands.SERVICEANNOUNCEMENT_HEALTH_LIST;
  }

  public get description(): string {
    return 'This operation provides the health report of all subscribed services for a tenant';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'status', 'service'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        issues: typeof args.options.issues !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-i, --issues' }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this
      .listServiceHealth(args.options)
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private listServiceHealth(options: Options): Promise<ServiceHealth[]> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/admin/serviceAnnouncement/healthOverviews${options.issues && (!options.output || options.output.toLocaleLowerCase() === 'json') ? '?$expand=issues' : ''}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: ServiceHealth[] }>(requestOptions)
      .then(response => {
        const serviceHealthList: ServiceHealth[] | undefined = response.value;

        if (!serviceHealthList) {
          return Promise.reject(`Error fetching service health`);
        }

        return Promise.resolve(serviceHealthList);
      });
  }
}

module.exports = new TenantServiceAnnouncementHealthListCommand();