import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { UnifiedRoleDefinition } from '@microsoft/microsoft-graph-types';

const options = globalOptionsZod
  .extend({
    displayName: zod.alias('n', z.string()),
    allowedResourceActions: zod.alias('a', z.string().transform((value) => value.split(',').map(String))),
    description: zod.alias('d', z.string().optional()),
    enabled: zod.alias('e', z.boolean().optional()),
    version: zod.alias('v', z.string().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraRoleDefinitionAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ROLEDEFINITION_ADD;
  }

  public get description(): string {
    return 'Creates a custom Microsoft Entra ID role definition';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.verbose) {
      await logger.logToStderr(`Creating custom role definition with name ${args.options.displayName}...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/roleManagement/directory/roleDefinitions`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      data: {
        displayName: args.options.displayName,
        rolePermissions: [
          {
            allowedResourceActions: args.options.allowedResourceActions
          }
        ],
        description: args.options.description,
        isEnabled: args.options.enabled !== undefined ? args.options.enabled : true,
        version: args.options.version
      },
      responseType: 'json'
    };

    try {
      const result = await request.post<UnifiedRoleDefinition>(requestOptions);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraRoleDefinitionAddCommand();