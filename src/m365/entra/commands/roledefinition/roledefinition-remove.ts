import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import { validation } from '../../../../utils/validation.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { cli } from '../../../../cli/cli.js';

const options = globalOptionsZod
  .extend({
    id: zod.alias('i', z.string().optional()),
    displayName: zod.alias('n', z.string().optional()),
    force: zod.alias('f', z.boolean().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraRoleDefinitionRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.ROLEDEFINITION_REMOVE;
  }

  public get description(): string {
    return 'Removes a specific Microsoft Entra ID role definition';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => !options.id !== !options.displayName, {
        message: 'Specify either id or displayName, but not both'
      })
      .refine(options => options.id || options.displayName, {
        message: 'Specify either id or displayName'
      })
      .refine(options => (!options.id && !options.displayName) || options.displayName || (options.id && validation.isValidGuid(options.id)), options => ({
        message: `The '${options.id}' must be a valid GUID`,
        path: ['id']
      }));
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeRoleDefinition = async (): Promise<void> => {
      try {
        let roleDefinitionId = args.options.id;

        if (args.options.displayName) {
          roleDefinitionId = (await roleDefinition.getRoleDefinitionByDisplayName(args.options.displayName, 'id')).id;
        }

        if (args.options.verbose) {
          await logger.logToStderr(`Removing role definition with ID ${roleDefinitionId}...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/roleManagement/directory/roleDefinitions/${roleDefinitionId}`,
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
      await removeRoleDefinition();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove role definition '${args.options.id || args.options.displayName}'?` });

      if (result) {
        await removeRoleDefinition();
      }
    }
  }
}

export default new EntraRoleDefinitionRemoveCommand();