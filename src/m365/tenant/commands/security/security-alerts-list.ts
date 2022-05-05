import { Alert } from '@microsoft/microsoft-graph-types';
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
  vendor?: string;
}

class TenantSecurityAlertsListCommand extends GraphCommand {
  public get name(): string {
    return commands.SECURITY_ALERTS_LIST;
  }

  public get description(): string {
    return 'Gets the security alerts for a tenant';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.vendor = typeof args.options.vendor !== 'undefined';
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'severity'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this
      .listAlert(args.options)
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private listAlert(options: Options): Promise<Alert[]> {
    let queryFilter: string = '';
    if (options.vendor) {
      let vendorName = options.vendor;

      switch (options.vendor.toLowerCase()) {
        case 'azure security center':
          vendorName = 'ASC';
          break;
        case 'microsoft cloud app security':
          vendorName = 'MCAS';
          break;
        case 'azure active directory identity protection':
          vendorName = 'IPC';
      }

      queryFilter = `?$filter=vendorInformation/provider eq '${encodeURIComponent(vendorName)}'`;
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/security/alerts${queryFilter}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: Alert[] }>(requestOptions)
      .then(response => {
        const alertList: Alert[] | undefined = response.value;

        if (!alertList) {
          return Promise.reject(`Error fetching security alerts`);
        }

        return Promise.resolve(alertList);
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '--vendor [vendor]' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new TenantSecurityAlertsListCommand();