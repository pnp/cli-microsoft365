import { Channel, Group } from '@microsoft/microsoft-graph-types';
import * as os from 'os';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
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
    return 'Adds a conversation member in a private channel.';
  }

  public alias(): string[] | undefined {
    return [commands.CONVERSATIONMEMBER_ADD];
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
      ['teamId', 'teamName'],
      ['channelId', 'channelName'],
      ['userId', 'userDisplayName']
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this.showDeprecationWarning(logger, commands.CONVERSATIONMEMBER_ADD, commands.CHANNEL_MEMBER_ADD);

    let teamId: string = '';
    let channelId: string = '';

    this
      .getTeamId(args)
      .then((_teamId: string): Promise<string> => {
        teamId = _teamId;
        return this.getChannelId(_teamId, args);
      })
      .then((_channelId: string): Promise<string[]> => {
        channelId = _channelId;
        return this.getUserId(args);
      })
      .then((userIds: string[]): Promise<void[]> => {
        const endpoint: string = `${this.resource}/v1.0/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/members`;
        const roles: string[] = args.options.owner ? ["owner"] : [];
        const tasks: Promise<void>[] = [];

        for (const userId of userIds) {
          tasks.push(this.addUser(userId, endpoint, roles));
        }

        return Promise.all(tasks);
      })
      .then(_ => cb(),
        (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private addUser(userId: string, endpoint: string, roles: string[]): Promise<void> {
    const requestOptions: any = {
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

  private getChannelId(teamId: string, args: CommandArgs): Promise<string> {
    if (args.options.channelId) {
      return Promise.resolve(args.options.channelId);
    }

    const channelRequestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(teamId)}/channels?$filter=displayName eq '${encodeURIComponent(args.options.channelName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: Channel[] }>(channelRequestOptions)
      .then(response => {
        const channelItem: Channel | undefined = response.value[0];

        if (!channelItem) {
          return Promise.reject(`The specified channel '${args.options.channelName}' does not exist in the Microsoft Teams team with ID '${teamId}'`);
        }

        if (channelItem.membershipType !== "private") {
          return Promise.reject(`The specified channel is not a private channel`);
        }

        return Promise.resolve(channelItem.id!);
      });
  }

  private getUserId(args: CommandArgs): Promise<string[]> {
    if (args.options.userId) {
      return Promise.resolve(args.options.userId.split(',').map(u => u.trim()));
    }

    const tasks: Promise<string>[] = [];
    const userDisplayNames: any | undefined = args.options.userDisplayName && args.options.userDisplayName.split(',').map(u => u.trim());

    for (const userName of userDisplayNames) {
      tasks.push(this.getSingleUser(userName));
    }

    return Promise.all(tasks);
  }

  private getSingleUser(userDisplayName: string): Promise<string> {
    const userRequestOptions: any = {
      url: `${this.resource}/v1.0/users?$filter=displayName eq '${encodeURIComponent(userDisplayName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: any[] }>(userRequestOptions)
      .then(response => {
        const userItem: any | undefined = response.value[0];

        if (!userItem) {
          return Promise.reject(`The specified user '${userDisplayName}' does not exist`);
        }

        if (response.value.length > 1) {
          return Promise.reject(`Multiple users with display name '${userDisplayName}' found. Please disambiguate:${os.EOL}${response.value.map(x => `- ${x.id}`).join(os.EOL)}`);
        }

        return Promise.resolve(userItem.id);
      }, err => { return Promise.reject(err); });
  }
}

module.exports = new TeamsChannelMemberAddCommand();