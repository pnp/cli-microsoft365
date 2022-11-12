import { Alert } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import GraphCommand from '../../../base/GraphCommand';
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

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'severity'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        vendor: typeof args.options.vendor !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--vendor [vendor]' }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const res: any = await this.listAlert(args.options);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
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

      queryFilter = `?$filter=vendorInformation/provider eq '${formatting.encodeQueryParameter(vendorName)}'`;
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
}

module.exports = new TenantSecurityAlertsListCommand();