import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import { validation } from '../../../../utils/validation.js';
import { UnifiedRoleDefinition } from '@microsoft/microsoft-graph-types';

const options = globalOptionsZod
  .extend({
    id: zod.alias('i', z.string().optional()),
    displayName: zod.alias('n', z.string().optional()),
    newDisplayName: z.string().optional(),
    allowedResourceActions: zod.alias('a', z.string().optional()),
    description: zod.alias('d', z.string().optional()),
    enabled: zod.alias('e', z.boolean().optional()),
    version: zod.alias('v', z.string().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraRoleDefinitionSetCommand extends GraphCommand {
  public get name(): string {
    return commands.ROLEDEFINITION_SET;
  }

  public get description(): string {
    return 'Updates a custom Microsoft Entra ID role definition';
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
      }))
      .refine(options => Object.values([options.newDisplayName, options.description, options.allowedResourceActions, options.enabled, options.version]).filter(v => typeof v !== 'undefined').length > 0, {
        message: 'Provide value for at least one of the following parameters: newDisplayName, description, allowedResourceActions, enabled or version'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let roleDefinitionId = args.options.id;

      if (args.options.displayName) {
        roleDefinitionId = (await roleDefinition.getRoleDefinitionByDisplayName(args.options.displayName, 'id')).id;
      }

      if (args.options.verbose) {
        await logger.logToStderr(`Updating custom role definition with ID ${roleDefinitionId}...`);
      }

      const data: UnifiedRoleDefinition = {
        displayName: args.options.newDisplayName,
        description: args.options.description,
        isEnabled: args.options.enabled,
        version: args.options.version
      };

      if (args.options.allowedResourceActions) {
        data['rolePermissions'] = [
          {
            allowedResourceActions: args.options.allowedResourceActions.split(',')
          }
        ];
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/roleManagement/directory/roleDefinitions/${roleDefinitionId}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        data: data,
        responseType: 'json'
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraRoleDefinitionSetCommand();