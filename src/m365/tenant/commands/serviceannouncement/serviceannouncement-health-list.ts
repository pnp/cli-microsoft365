import { ServiceHealth } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.issues = typeof args.options.issues !== 'undefined';
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'status', 'service'];
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '-i, --issues' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new TenantServiceAnnouncementHealthListCommand();