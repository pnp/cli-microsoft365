import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.guid().optional().alias('i'),
  displayName: z.string().optional().alias('n')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraM365GroupRenewCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_RENEW;
  }

  public get description(): string {
    return `Renews Microsoft 365 group's expiration`;
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.displayName].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: id or displayName.',
        params: {
          customCode: 'optionSet',
          options: ['id', 'displayName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Renewing Microsoft 365 group's expiration: ${args.options.id || args.options.displayName}...`);
    }

    try {
      let groupId = args.options.id;

      if (args.options.displayName) {
        groupId = await entraGroup.getGroupIdByDisplayName(args.options.displayName);
      }
      const isUnifiedGroup = await entraGroup.isUnifiedGroup(groupId!);

      if (!isUnifiedGroup) {
        throw Error(`Specified group with id '${groupId}' is not a Microsoft 365 group.`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/groups/${groupId}/renew/`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        }
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraM365GroupRenewCommand();