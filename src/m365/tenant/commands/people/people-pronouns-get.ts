import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = globalOptionsZod.strict();

class TenantPeoplePronounsGetCommand extends GraphCommand {
  public get name(): string {
    return commands.PEOPLE_PRONOUNS_GET;
  }

  public get description(): string {
    return 'Retrieves information about pronouns settings for an organization';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr('Retrieving information about pronouns settings...');
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/admin/people/pronouns`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const pronouns = await request.get<any>(requestOptions);

      await logger.log(pronouns);

    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }  
  }
}

export default new TenantPeoplePronounsGetCommand();