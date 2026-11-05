import { Alert } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  vendor: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
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

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const res: any = await this.listAlert(args.options);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async listAlert(options: Options): Promise<Alert[]> {
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

    const response: any = await request.get<{ value: Alert[] }>(requestOptions);
    const alertList: Alert[] | undefined = response.value;

    if (!alertList) {
      throw `Error fetching security alerts`;
    }

    return alertList;
  }
}

export default new TenantSecurityAlertsListCommand();