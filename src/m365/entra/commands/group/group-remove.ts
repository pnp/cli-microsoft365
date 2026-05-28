import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.uuid().optional().alias('i'),
  displayName: z.string().optional().alias('n'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraGroupRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUP_REMOVE;
  }

  public get description(): string {
    return 'Removes an Entra group';
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
    const removeGroup = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing group ${args.options.id || args.options.displayName}...`);
      }

      try {
        let groupId = args.options.id;

        if (args.options.displayName) {
          groupId = await entraGroup.getGroupIdByDisplayName(args.options.displayName);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/groups/${groupId}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          }
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeGroup();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove group '${args.options.id || args.options.displayName}'?` });

      if (result) {
        await removeGroup();
      }
    }
  }
}

export default new EntraGroupRemoveCommand();