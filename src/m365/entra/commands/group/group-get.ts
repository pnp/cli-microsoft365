import { Group } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.uuid().optional().alias('i'),
  displayName: z.string().optional().alias('n'),
  properties: z.string().optional().alias('p')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraGroupGetCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUP_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Entra group';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.id, options.displayName].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: id or displayName.',
        params: {
          customCode: 'optionSet',
          options: ['id', 'displayName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let group: Group;

    try {
      if (args.options.id) {
        group = await entraGroup.getGroupById(args.options.id, args.options.properties);
      }
      else {
        group = await entraGroup.getGroupByDisplayName(args.options.displayName!, args.options.properties);
      }

      await logger.log(group);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraGroupGetCommand();