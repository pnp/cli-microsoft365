import { Group } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import { aadGroup } from '../../../../utils/aadGroup';
import GraphCommand from '../../../base/GraphCommand';
import { Channel } from '../../Channel';
import commands from '../../commands';

interface ExtendedGroup extends Group {
  resourceProvisioningOptions: string[];
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  channelId?: string;
  channelName?: string;
  teamId?: string;
  teamName?: string;
  description?: string
  newChannelName?: string;
}

class TeamsChannelSetCommand extends GraphCommand {
  private teamId: string = "";

  public get name(): string {
    return commands.CHANNEL_SET;
  }
  public get description(): string {
    return 'Updates properties of the specified channel in the given Microsoft Teams team';
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
        channelId: typeof args.options.channelId !== 'undefined',
        channelName: typeof args.options.channelName !== 'undefined',
        teamId: typeof args.options.teamId !== 'undefined',
        teamName: typeof args.options.teamName !== 'undefined',
        newChannelName: typeof args.options.newChannelName !== 'undefined',
        description: typeof args.options.description !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-c, --channelId [channelId]'
      },
      {
        option: '-n, --channelName [channelName]'
      },
      {
        option: '-i, --teamId [teamId]'
      },
      {
        option: '--teamName [teamName]'
      },
      {
        option: '--newChannelName [newChannelName]'
      },
      {
        option: '--description [description]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (args.options.channelName && args.options.channelName.toLowerCase() === "general") {
          return 'General channel cannot be updated';
        }

        if (args.options.channelId && !validation.isValidTeamsChannelId(args.options.channelId)) {
          return `${args.options.channelId} is not a valid Teams Channel Id`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      ['channelId', 'channelName'],
      ['teamId', 'teamName']
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      this.teamId = await this.getTeamId(args);
      const channelId: string = await this.getChannelId(args);

      const data: any = this.mapRequestBody(args.options);
      const requestOptionsPatch: any = {
        url: `${this.resource}/v1.0/teams/${encodeURIComponent(this.teamId)}/channels/${encodeURIComponent(channelId)}`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: data
      };

      await request.patch(requestOptionsPatch);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};

    if (options.newChannelName) {
      requestBody.displayName = options.newChannelName;
    }

    if (options.description) {
      requestBody.description = options.description;
    }

    return requestBody;
  }

  private async getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return args.options.teamId;
    }

    const group: Group = await aadGroup.getGroupByDisplayName(args.options.teamName!);

    if ((group as ExtendedGroup).resourceProvisioningOptions.indexOf('Team') === -1) {
      throw 'The specified team does not exist in the Microsoft Teams';
    }
    else {
      return group.id!;
    }
  }

  private async getChannelId(args: CommandArgs): Promise<string> {
    if (args.options.channelId) {
      return args.options.channelId;
    }

    const channelRequestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(this.teamId)}/channels?$filter=displayName eq '${encodeURIComponent(args.options.channelName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res: { value: Channel[] } = await request.get<{ value: Channel[] }>(channelRequestOptions);
    const channelItem: Channel | undefined = res.value[0];

    if (!channelItem) {
      throw `The specified channel does not exist in the Microsoft Teams team`;
    }

    return channelItem.id;
  }
}

module.exports = new TeamsChannelSetCommand();