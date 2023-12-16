import { Channel, Group } from '@microsoft/microsoft-graph-types';
import GlobalOptions from '../../../../GlobalOptions.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface ExtendedGroup extends Group {
  resourceProvisioningOptions: string[];
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  teamId?: string;
  teamName?: string;
  force?: boolean;
}

class TeamsChannelRemoveCommand extends GraphCommand {
  private teamId: string = "";

  public get name(): string {
    return commands.CHANNEL_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified channel in the Microsoft Teams team';
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
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        teamId: typeof args.options.teamId !== 'undefined',
        teamName: typeof args.options.teamName !== 'undefined',
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--teamId [teamId]'
      },
      {
        option: '--teamName [teamName]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidTeamsChannelId(args.options.id)) {
          return `${args.options.id} is not a valid Teams channel id`;
        }

        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['id', 'name'] },
      { options: ['teamId', 'teamName'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeChannel = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Removing channel ${args.options.id || args.options.name} from team ${args.options.teamId || args.options.teamName}`);
        }

        this.teamId = await this.getTeamId(args);
        const channelId: string = await this.getChannelId(args);

        const requestOptionsDelete: any = {
          url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(this.teamId)}/channels/${formatting.encodeQueryParameter(channelId)}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        await request.delete(requestOptionsDelete);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeChannel();
    }
    else {
      const channel = args.options.name ? args.options.name : args.options.id;
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the channel ${channel} from team ${args.options.teamId || args.options.teamName}?` });

      if (result) {
        await removeChannel();
      }
    }
  }

  private async getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return args.options.teamId;
    }

    const group: Group = await aadGroup.getGroupByDisplayName(args.options.teamName!);

    if ((group as ExtendedGroup).resourceProvisioningOptions.indexOf('Team') === -1) {
      throw 'The specified team does not exist';
    }
    else {
      return group.id!;
    }
  }

  private async getChannelId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    const channelRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(this.teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.name!)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res: { value: Channel[] } = await request.get<{ value: Channel[] }>(channelRequestOptions);
    const channelItem: Channel | undefined = res.value[0];

    if (!channelItem) {
      throw 'The specified channel does not exist in this Microsoft Teams team';
    }

    return channelItem.id!;
  }
}

export default new TeamsChannelRemoveCommand();