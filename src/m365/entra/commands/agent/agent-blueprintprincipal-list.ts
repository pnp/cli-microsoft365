import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  properties: z.string().optional().alias('p')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraAgentBlueprintprincipalListCommand extends GraphCommand {
  public get name(): string {
    return commands.AGENT_BLUEPRINTPRINCIPAL_LIST;
  }

  public get description(): string {
    return 'Retrieves a list of agent blueprint principals';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'appDisplayName', 'appId'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const queryParameters: string[] = [];

    if (args.options.properties) {
      const allProperties = args.options.properties.split(',').map(prop => prop.trim()).filter(prop => prop.length > 0);
      const selectProperties = [...new Set(allProperties.map(prop => prop.split('/')[0]))];

      if (selectProperties.length > 0) {
        queryParameters.push(`$select=${selectProperties}`);
      }

      if (cli.commandToExecute) {
        cli.commandToExecute.defaultProperties = [
          ...new Set<string>(allProperties.map(prop => prop.split('/')[0]))
        ];
      }
    }

    const queryString = queryParameters.length > 0
      ? `?${queryParameters.join('&')}`
      : '';

    try {
      const results = await odata.getAllItems<any>(`${this.resource}/v1.0/servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal${queryString}`);
      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraAgentBlueprintprincipalListCommand();