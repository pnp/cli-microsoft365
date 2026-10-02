import { ServiceHealth } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  issues: z.boolean().optional().alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
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

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const response: any = await this.listServiceHealth(args.options);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async listServiceHealth(options: Options): Promise<ServiceHealth[]> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/admin/serviceAnnouncement/healthOverviews${options.issues && (!options.output || options.output.toLocaleLowerCase() === 'json') ? '?$expand=issues' : ''}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: ServiceHealth[] }>(requestOptions);
    const serviceHealthList: ServiceHealth[] | undefined = response.value;

    if (!serviceHealthList) {
      throw `Error fetching service health`;
    }

    return serviceHealthList;
  }
}

export default new TenantServiceAnnouncementHealthListCommand();