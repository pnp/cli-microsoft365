import { UnifiedRoleDefinition } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  properties: z.string().optional().alias('p'),
  filter: z.string().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraRoleDefinitionListCommand extends GraphCommand {
  public get name(): string {
    return commands.ROLEDEFINITION_LIST;
  }

  public get description(): string {
    return 'Lists all Microsoft Entra ID role definitions';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'isBuiltIn', 'isEnabled'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Getting Microsoft Entra ID role definitions...');
    }

    try {
      const queryParameters: string[] = [];

      if (args.options.properties) {
        queryParameters.push(`$select=${args.options.properties}`);
      }

      if (args.options.filter) {
        queryParameters.push(`$filter=${args.options.filter}`);
      }

      const queryString = queryParameters.length > 0
        ? `?${queryParameters.join('&')}`
        : '';

      const results = await odata.getAllItems<UnifiedRoleDefinition>(`${this.resource}/v1.0/roleManagement/directory/roleDefinitions${queryString}`);
      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraRoleDefinitionListCommand();