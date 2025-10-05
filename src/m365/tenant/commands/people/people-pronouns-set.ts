import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  enabled: z.boolean().alias('e')
});

export declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TenantPeoplePronounsSetCommand extends GraphCommand {
  public get name(): string {
    return commands.PEOPLE_PRONOUNS_SET;
  }

  public get description(): string {
    return 'Manage pronouns settings for an organization';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr('Updating pronouns settings...');
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/admin/people/pronouns`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        data: {
          isEnabledInOrganization: args.options.enabled
        },
        responseType: 'json'
      };

      const pronouns = await request.patch<any>(requestOptions);

      await logger.log(pronouns);

    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TenantPeoplePronounsSetCommand();