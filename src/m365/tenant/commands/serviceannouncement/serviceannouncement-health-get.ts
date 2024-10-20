import { ServiceHealth } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  serviceName: string;
  issues?: boolean;
}

class TenantServiceAnnouncementHealthGetCommand extends GraphCommand {
  public get name(): string {
    return commands.SERVICEANNOUNCEMENT_HEALTH_GET;
  }

  public get description(): string {
    return 'This operation provides the health information of a specified service for a tenant';
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
      { option: '-s, --serviceName <serviceName>' },
      { option: '-i, --issues' }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const res: any = await this.getServiceHealth(args.options);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getServiceHealth(options: Options): Promise<ServiceHealth> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/admin/serviceAnnouncement/healthOverviews/${options.serviceName}${options.issues && (!options.output || options.output.toLocaleLowerCase() === 'json') ? '?$expand=issues' : ''}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<ServiceHealth>(requestOptions);
  }
}

export default new TenantServiceAnnouncementHealthGetCommand();