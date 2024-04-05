import { Channel, Group } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
import { formatting } from '../../../../utils/formatting.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';

interface ExtendedGroup extends Group {
  resourceProvisioningOptions: string[];
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  teamName: string;
  channelId: string;
  channelName: string;
  id: string;
}

class TeamsMessageRestoreCommand extends GraphCommand {
  public get name(): string {
    return commands.MESSAGE_RESTORE;
  }

  public get description(): string {
    return 'Restores a deleted message from a channel in a Microsoft Teams team';
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
        teamName: typeof args.options.teamName !== 'undefined',
        channelId: typeof args.options.channelId !== 'undefined',
        channelName: typeof args.options.channelName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --teamId [teamId]'
      },
      {
        option: '--teamName [teamName]'
      },
      {
        option: '-c, --channelId [channelId]'
      },
      {
        option: '--channelName [channelName]'
      },
      {
        option: '-i, --id <id>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (args.options.channelId && !validation.isValidTeamsChannelId(args.options.channelId as string)) {
          return `${args.options.channelId} is not a valid Teams ChannelId`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['teamId', 'teamName'] }, { options: ['channelId', 'channelName'] });
  }

  private async getTeamId(options: Options): Promise<string> {
    if (options.teamId) {
      return options.teamId;
    }

    const group = await entraGroup.getGroupByDisplayName(options.teamName!);

    if ((group as ExtendedGroup).resourceProvisioningOptions.indexOf('Team') === -1) {
      throw 'The specified team does not exist in the Microsoft Teams';
    }

    return group.id!;
  }

  private async getChannelId(options: Options, teamId: string): Promise<string> {
    if (options.channelId) {
      return options.channelId;
    }

    const channelRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(options.channelName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: Channel[] }>(channelRequestOptions);
    const channelItem: Channel | undefined = response.value[0];

    if (!channelItem) {
      throw `The specified channel does not exist in the Microsoft Teams team`;
    }

    return channelItem.id!;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken)) {
        throw 'This command currently does not support app only permissions.';
      }

      const teamId: string = await this.getTeamId(args.options);
      const channelId: string = await this.getChannelId(args.options, teamId);
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/teams/${teamId}/channels/${channelId}/messages/${args.options.id}/undoSoftDelete`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res: any = await request.post(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsMessageRestoreCommand();