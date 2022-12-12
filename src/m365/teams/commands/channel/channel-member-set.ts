import { Channel, Group } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import { aadGroup } from '../../../../utils/aadGroup';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { ConversationMember } from '../../ConversationMember';
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
  userName?: string;
  userId?: string;
  id?: string;
  role: string;
}

class TeamsChannelMemberSetCommand extends GraphCommand {
  private teamId: string = '';
  private channelId: string = '';

  public get name(): string {
    return commands.CHANNEL_MEMBER_SET;
  }

  public get description(): string {
    return 'Updates the role of the specified member in the specified Microsoft Teams private or shared team channel';
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
        userName: typeof args.options.userName !== 'undefined',
        userId: typeof args.options.userId !== 'undefined',
        id: typeof args.options.id !== 'undefined'
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
        option: '--channelId [channelId]'
      },
      {
        option: '--channelName [channelName]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--id [id]'
      },
      {
        option: '-r, --role <role>',
        autocomplete: ['owner', 'member']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (args.options.channelId && !validation.isValidTeamsChannelId(args.options.channelId)) {
          return `${args.options.channelId} is not a valid Teams Channel ID`;
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (['owner', 'member'].indexOf(args.options.role) === -1) {
          return `${args.options.role} is not a valid role value. Allowed values owner|member`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['teamId', 'teamName'] },
      { options: ['channelId', 'channelName'] },
      { options: ['userName', 'userId', 'id'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      this.teamId = await this.getTeamId(args);
      this.channelId = await this.getChannelId(args);
      const memberId: string = await this.getMemberId(args);
      const requestOptions: any = {
        url: `${this.resource}/v1.0/teams/${this.teamId}/channels/${this.channelId}/members/${memberId}`,
        headers: {
          'accept': 'application/json;odata.metadata=none',
          'Prefer': 'return=representation'
        },
        responseType: 'json',
        data: {
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          roles: [args.options.role]
        }
      };

      const member = await request.patch(requestOptions);
      logger.log(member);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
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

  private getChannelId(args: CommandArgs): Promise<string> {
    if (args.options.channelId) {
      return Promise.resolve(args.options.channelId);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(this.teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.channelName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: Channel[] }>(requestOptions)
      .then(response => {
        const channelItem: Channel | undefined = response.value[0];

        if (!channelItem) {
          return Promise.reject(`The specified channel does not exist in the Microsoft Teams team`);
        }

        if (channelItem.membershipType !== "private") {
          return Promise.reject(`The specified channel is not a private channel`);
        }

        return Promise.resolve(channelItem.id!);
      });
  }

  private getMemberId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return Promise.resolve(args.options.id);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${this.teamId}/channels/${this.channelId}/members`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: ConversationMember[] }>(requestOptions)
      .then(response => {
        const conversationMembers = response.value.filter(x =>
          args.options.userId && x.userId?.toLocaleLowerCase() === args.options.userId.toLocaleLowerCase() ||
          args.options.userName && x.email?.toLocaleLowerCase() === args.options.userName.toLocaleLowerCase()
        );

        const conversationMember: ConversationMember | undefined = conversationMembers[0];

        if (!conversationMember) {
          return Promise.reject(`The specified member does not exist in the Microsoft Teams channel`);
        }

        if (conversationMembers.length > 1) {
          return Promise.reject(`Multiple Microsoft Teams channel members with name ${args.options.userName} found: ${response.value.map(x => x.userId)}`);
        }

        return Promise.resolve(conversationMember.id);
      });
  }
}

module.exports = new TeamsChannelMemberSetCommand();