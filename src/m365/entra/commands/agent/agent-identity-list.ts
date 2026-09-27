import { DirectoryObject, NullableOption } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
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

/**
 * Agent identity as returned by the Microsoft Graph
 * `servicePrincipals/microsoft.graph.agentIdentity` endpoint.
 */
interface AgentIdentity extends DirectoryObject {
  displayName?: NullableOption<string>;
  createdDateTime?: NullableOption<string>;
  createdByAppId?: NullableOption<string>;
  agentIdentityBlueprintId?: NullableOption<string>;
  accountEnabled?: NullableOption<boolean>;
  disabledByMicrosoftStatus?: NullableOption<string>;
  servicePrincipalType?: NullableOption<string>;
  tags?: NullableOption<string[]>;
}

class EntraAgentIdentityListCommand extends GraphCommand {
  public get name(): string {
    return commands.AGENT_IDENTITY_LIST;
  }

  public get description(): string {
    return 'Retrieves a list of agent identities';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const queryParameters: string[] = [];

    if (args.options.properties) {
      const allProperties = args.options.properties
        .split(',')
        .map(prop => prop.replace(/['"]/g, '').trim())
        .filter(prop => prop.length > 0);
      const selectProperties = allProperties.filter(prop => !prop.includes('/'));

      if (selectProperties.length > 0) {
        queryParameters.push(`$select=${formatting.encodeQueryParameter(selectProperties.join(','))}`);
      }
    }

    const queryString = queryParameters.length > 0
      ? `?${queryParameters.join('&')}`
      : '';

    try {
      const results = await odata.getAllItems<AgentIdentity>(`${this.resource}/v1.0/servicePrincipals/microsoft.graph.agentIdentity${queryString}`);
      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraAgentIdentityListCommand();
