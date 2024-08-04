import { Group } from '@microsoft/microsoft-graph-types';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { CliRequestOptions } from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import aadCommands from '../../aadCommands.js';
import { validation } from '../../../../utils/validation.js';
import { formatting } from '../../../../utils/formatting.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type?: string;
  joined?: boolean;
  associated?: boolean;
  userId?: string;
  userName?: string;
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

  public alias(): string[] | undefined {
    return [aadCommands.GROUP_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'groupType'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
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
      },
      {
        option: '-j, --joined'
      },
      {
        option: '-a, --associated'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.type && EntraGroupListCommand.groupTypes.every(g => g.toLowerCase() !== args.options.type?.toLowerCase())) {
          return `${args.options.type} is not a valid type value. Allowed values microsoft365|security|distribution|mailEnabledSecurity.`;
        }

        if (args.options.userId !== undefined && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID for userId.`;
        }

        if (args.options.userName !== undefined && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid UPN for userName.`;
        }

        if ((args.options.userId !== undefined || args.options.userName !== undefined) && !args.options.joined && !args.options.associated) {
          return 'You must specify either joined or associated when specifying userId or userName.';
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['joined', 'associated'],
        runsWhen: (args: CommandArgs) => !!args.options.joined || !!args.options.associated
      },
      {
        options: ['userId', 'userName'],
        runsWhen: (args: CommandArgs) => args.options.userId !== undefined || args.options.userName !== undefined
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('userId', 'userName');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.GROUP_LIST, commands.GROUP_LIST);

    if (this.verbose) {
      if (!args.options.joined && !args.options.associated) {
        await logger.logToStderr(`Retrieving Microsoft Teams in the tenant...`);
      }
      else {
        const user = args.options.userId || args.options.userName || 'me';
        await logger.logToStderr(`Retrieving Microsoft Teams ${args.options.joined ? 'joined by' : 'associated with'} ${user}...`);
      }
    }

    try {
      let endpoint = `${this.resource}/v1.0`;
      if (args.options.joined || args.options.associated) {
        if (!args.options.userId && !args.options.userName && accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken)) {
          throw `You must specify either userId or userName when using application only permissions and specifying the ${args.options.joined ? 'joined' : 'associated'} option`;
        }

        endpoint += args.options.userId || args.options.userName ? `/users/${args.options.userId || formatting.encodeQueryParameter(args.options.userName!)}` : '/me';
        endpoint += args.options.joined ? '/memberOf' : '/transitiveMemberOf';
        endpoint += '/microsoft.graph.group';
      }
      else {
        // Get all team groups within the tenant
        endpoint += `/groups`;
      }
      let useConsistencyLevelHeader = false;

      if (args.options.type) {
        const groupType = EntraGroupListCommand.groupTypes.find(g => g.toLowerCase() === args.options.type?.toLowerCase());

        switch (groupType) {
          case 'microsoft365':
            endpoint += `?$filter=groupTypes/any(c:c+eq+'Unified')`;
            break;
          case 'security':
            useConsistencyLevelHeader = true;
            endpoint += '?$filter=securityEnabled eq true and mailEnabled eq false&$count=true';
            break;
          case 'distribution':
            useConsistencyLevelHeader = true;
            endpoint += '?$filter=securityEnabled eq false and mailEnabled eq true&$count=true';
            break;
          case 'mailEnabledSecurity':
            useConsistencyLevelHeader = true;
            endpoint += `?$filter=securityEnabled eq true and mailEnabled eq true and not(groupTypes/any(t:t eq 'Unified'))&$count=true`;
            break;
        }
      }

      let groups: Group[] = [];

      if (useConsistencyLevelHeader) {
        // While using not() function in the filter, we need to specify the ConsistencyLevel header.
        const requestOptions: CliRequestOptions = {
          url: endpoint,
          headers: {
            accept: 'application/json;odata.metadata=none',
            ConsistencyLevel: 'eventual'
          },
          responseType: 'json'
        };

        groups = await odata.getAllItems<Group>(requestOptions);
      }
      else {
        groups = await odata.getAllItems<Group>(endpoint);
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