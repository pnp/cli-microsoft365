import { Channel, Group } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import { aadGroup } from '../../../../utils/aadGroup';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface ExtendedGroup extends Group {
  resourceProvisioningOptions: string[];
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId?: string;
  teamName?: string;
  type?: string;
}

class TeamsChannelListCommand extends GraphCommand {
  public get name(): string {
    return commands.CHANNEL_LIST;
  }

  public get description(): string {
    return 'Lists channels in the specified Microsoft Teams team';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        teamId: typeof args.options.teamId !== 'undefined',
        teamName: typeof args.options.teamName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --teamId [teamId]'
      },
      {
        option: '--teamName [teamName]'
      },
      {
        option: '--type [type]',
        autocomplete: ['standard', 'private', 'shared']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (args.options.type && ['standard', 'private', 'shared'].indexOf(args.options.type.toLowerCase()) === -1) {
          return `${args.options.type} is not a valid type value. Allowed values standard|private|shared`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['teamId', 'teamName'] });
  }

  private getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return Promise.resolve(args.options.teamId);
    }

    return aadGroup
      .getGroupByDisplayName(args.options.teamName!)
      .then(group => {
        if ((group as ExtendedGroup).resourceProvisioningOptions.indexOf('Team') === -1) {
          return Promise.reject(`The specified team does not exist in the Microsoft Teams`);
        }

        return group.id!;
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const teamId: string = await this.getTeamId(args);
      let endpoint: string = `${this.resource}/v1.0/teams/${teamId}/channels`;

      if (args.options.type) {
        endpoint += `?$filter=membershipType eq '${args.options.type}'`;
      }

      const items = await odata.getAllItems<Channel>(endpoint);
      logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsChannelListCommand();