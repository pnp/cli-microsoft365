import { Group } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { CliRequestOptions } from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const GroupType = {
  microsoft365: 'microsoft365',
  security: 'security',
  distribution: 'distribution',
  mailEnabledSecurity: 'mailEnabledSecurity'
} as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  type: zod.coercedEnum(GroupType).optional(),
  properties: z.string().optional().alias('p')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface ExtendedGroup extends Group {
  groupType?: string;
}

class EntraGroupListCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUP_LIST;
  }

  public get description(): string {
    return 'Lists all groups defined in Entra ID.';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'groupType'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let requestUrl: string = `${this.resource}/v1.0/groups`;
      let useConsistencyLevelHeader = false;

      if (args.options.type) {
        switch (args.options.type) {
          case 'microsoft365':
            requestUrl += `?$filter=groupTypes/any(c:c+eq+'Unified')`;
            break;
          case 'security':
            requestUrl += '?$filter=securityEnabled eq true and mailEnabled eq false';
            break;
          case 'distribution':
            useConsistencyLevelHeader = true;
            requestUrl += `?$filter=securityEnabled eq false and mailEnabled eq true and not(groupTypes/any(t:t eq 'Unified'))&$count=true`;
            break;
          case 'mailEnabledSecurity':
            useConsistencyLevelHeader = true;
            requestUrl += `?$filter=securityEnabled eq true and mailEnabled eq true and not(groupTypes/any(t:t eq 'Unified'))&$count=true`;
            break;
        }
      }

      const queryParameters: string[] = [];

      if (args.options.properties) {
        const allProperties = args.options.properties.split(',');
        const selectProperties = allProperties.filter(prop => !prop.includes('/'));

        if (selectProperties.length > 0) {
          queryParameters.push(`$select=${selectProperties}`);
        }
      }

      const queryString = queryParameters.length > 0
        ? `?${queryParameters.join('&')}`
        : '';

      requestUrl += queryString;

      let groups: Group[] = [];

      if (useConsistencyLevelHeader) {
        // While using not() function in the filter, we need to specify the ConsistencyLevel header.
        const requestOptions: CliRequestOptions = {
          url: requestUrl,
          headers: {
            accept: 'application/json;odata.metadata=none',
            ConsistencyLevel: 'eventual'
          },
          responseType: 'json'
        };

        groups = await odata.getAllItems<Group>(requestOptions);
      }
      else {
        groups = await odata.getAllItems<Group>(requestUrl);
      }

      if (cli.shouldTrimOutput(args.options.output)) {
        groups.forEach((group: ExtendedGroup) => {
          if (group.groupTypes && group.groupTypes.length > 0 && group.groupTypes.includes('Unified')) {
            group.groupType = 'Microsoft 365';
          }
          else if (group.mailEnabled && group.securityEnabled) {
            group.groupType = 'Mail enabled security';
          }
          else if (group.securityEnabled) {
            group.groupType = 'Security';
          }
          else if (group.mailEnabled) {
            group.groupType = 'Distribution';
          }
        });
      }

      await logger.log(groups);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraGroupListCommand();