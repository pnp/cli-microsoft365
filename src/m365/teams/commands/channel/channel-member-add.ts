import { Channel, Group } from '@microsoft/microsoft-graph-types';
import * as os from 'os';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import { aadGroup } from '../../../../utils/aadGroup';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { formatting } from '../../../../utils/formatting';

interface ExtendedGroup extends Group {
  resourceProvisioningOptions: string[];
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId?: string;
  teamName?: string;
  channelId?: string;
  channelName?: string;
  userId?: string;
  userDisplayName?: string;
  owner: boolean;
}

class TeamsChannelMemberAddCommand extends GraphCommand {
  public get name(): string {
    return commands.CHANNEL_MEMBER_ADD;
  }

  public get description(): string {
    return 'Adds a specified member in the specified Microsoft Teams private or shared team channel';
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
        channelName: typeof args.options.channelName !== 'undefined',
        userId: typeof args.options.userId !== 'undefined',
        userDisplayName: typeof args.options.userDisplayName !== 'undefined',
        owner: args.options.owner
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
        option: '-c, --channelId [channelId]'
      },
      {
        option: '--channelName [channelName]'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--userDisplayName [userDisplayName]'
      },
      {
        option: '--owner'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId as string)) {
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
    this.optionSets.push(
      { options: ['teamId', 'teamName'] },
      { options: ['channelId', 'channelName'] },
      { options: ['userId', 'userDisplayName'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const teamId: string = await this.getTeamId(args);
      const channelId: string = await this.getChannelId(teamId, args);
      const userIds: string[] = await this.getUserId(args);
      const endpoint: string = `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels/${formatting.encodeQueryParameter(channelId)}/members`;
      const roles: string[] = args.options.owner ? ["owner"] : [];
      const tasks: Promise<void>[] = [];

      for (const userId of userIds) {
        tasks.push(this.addUser(userId, endpoint, roles));
      }

      await Promise.all(tasks);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async addUser(userId: string, endpoint: string, roles: string[]): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: endpoint,
      headers: {
        'content-type': 'application/json;odata=nometadata',
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        'roles': roles,
        'user@odata.bind': `${this.resource}/v1.0/users('${userId}')`
      }
    };

    return request.post(requestOptions);
  }

  private async getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return args.options.teamId;
    }

    const group = await aadGroup.getGroupByDisplayName(args.options.teamName!);
    if ((group as ExtendedGroup).resourceProvisioningOptions.indexOf('Team') === -1) {
      throw 'The specified team does not exist in the Microsoft Teams';
    }

    return group.id!;
  }

  private async getChannelId(teamId: string, args: CommandArgs): Promise<string> {
    if (args.options.channelId) {
      return args.options.channelId;
    }

    const channelRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.channelName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: Channel[] }>(channelRequestOptions);
    const channelItem: Channel | undefined = response.value[0];

    if (!channelItem) {
      throw `The specified channel '${args.options.channelName}' does not exist in the Microsoft Teams team with ID '${teamId}'`;
    }

    if (channelItem.membershipType !== "private") {
      throw `The specified channel is not a private channel`;
    }

    return channelItem.id!;
  }

  private async getUserId(args: CommandArgs): Promise<string[]> {
    if (args.options.userId) {
      return args.options.userId.split(',').map(u => u.trim());
    }

    const tasks: Promise<string>[] = [];
    const userDisplayNames: any | undefined = args.options.userDisplayName && args.options.userDisplayName.split(',').map(u => u.trim());

    for (const userName of userDisplayNames) {
      tasks.push(this.getSingleUser(userName));
    }

    return Promise.all(tasks);
  }

  private async getSingleUser(userDisplayName: string): Promise<string> {
    const userRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/users?$filter=displayName eq '${formatting.encodeQueryParameter(userDisplayName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: any[] }>(userRequestOptions);
    const userItem: any | undefined = response.value[0];

    if (!userItem) {
      throw `The specified user '${userDisplayName}' does not exist`;
    }

    if (response.value.length > 1) {
      throw `Multiple users with display name '${userDisplayName}' found. Please disambiguate:${os.EOL}${response.value.map(x => `- ${x.id}`).join(os.EOL)}`;
    }

    return userItem.id;
  }
}

module.exports = new TeamsChannelMemberAddCommand();