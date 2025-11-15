import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { CliRequestOptions } from '../../../../request.js';
import { Organization } from '@microsoft/microsoft-graph-types';
import { odata } from '../../../../utils/odata.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  properties: z.string().optional().alias('p')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraOrganizationListCommand extends GraphCommand {
  public get name(): string {
    return commands.ORGANIZATION_LIST;
  }

  public get description(): string {
    return 'Lists all Microsoft Entra ID organizations';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'tenantType'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let url = `${this.resource}/v1.0/organization`;
      if (args.options.properties) {
        url += `?$select=${args.options.properties}`;
      }
      const requestOptions: CliRequestOptions = {
        url: url,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      if (args.options.verbose) {
        await logger.logToStderr(`Retrieving organizations...`);
      }

      const res = await odata.getAllItems<Organization>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraOrganizationListCommand();