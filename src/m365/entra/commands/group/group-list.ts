import { Group } from '@microsoft/microsoft-graph-types';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { CliRequestOptions } from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type?: string;
}

interface ExtendedGroup extends Group {
  groupType?: string;
}

class EntraGroupListCommand extends GraphCommand {
  private static readonly groupTypes: string[] = ['microsoft365', 'security', 'distribution', 'mailEnabledSecurity'];

  public get name(): string {
    return commands.GROUP_LIST;
  }

  public get description(): string {
    return 'Lists all groups defined in Entra ID.';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'groupType'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        type: typeof args.options.type !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--type [type]',
        autocomplete: EntraGroupListCommand.groupTypes
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.type && EntraGroupListCommand.groupTypes.every(g => g.toLowerCase() !== args.options.type?.toLowerCase())) {
          return `${args.options.type} is not a valid type value. Allowed values microsoft365|security|distribution|mailEnabledSecurity.`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let requestUrl: string = `${this.resource}/v1.0/groups`;
      let useConsistencyLevelHeader = false;

      if (args.options.type) {
        const groupType = EntraGroupListCommand.groupTypes.find(g => g.toLowerCase() === args.options.type?.toLowerCase());

        switch (groupType) {
          case 'microsoft365':
            requestUrl += `?$filter=groupTypes/any(c:c+eq+'Unified')`;
            break;
          case 'security':
            requestUrl += '?$filter=securityEnabled eq true and mailEnabled eq false';
            break;
          case 'distribution':
            requestUrl += '?$filter=securityEnabled eq false and mailEnabled eq true';
            break;
          case 'mailEnabledSecurity':
            useConsistencyLevelHeader = true;
            requestUrl += `?$filter=securityEnabled eq true and mailEnabled eq true and not(groupTypes/any(t:t eq 'Unified'))&$count=true`;
            break;
        }
      }

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
          if (group.groupTypes && group.groupTypes.length > 0 && group.groupTypes[0] === 'Unified') {
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