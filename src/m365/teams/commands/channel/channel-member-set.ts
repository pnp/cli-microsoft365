import { Channel } from '../../Channel';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import { validation } from '../../../../utils';
import GlobalOptions from '../../../../GlobalOptions';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import request from '../../../../request';
import { ConversationMember } from '../../ConversationMember';

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
    return 'Updates the role of the specified member in the specified Microsoft Teams private team channel';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.teamId = typeof args.options.teamId !== 'undefined';
    telemetryProps.teamName = typeof args.options.teamName !== 'undefined';
    telemetryProps.channelId = typeof args.options.channelId !== 'undefined';
    telemetryProps.channelName = typeof args.options.channelName !== 'undefined';
    telemetryProps.userName = typeof args.options.userName !== 'undefined';
    telemetryProps.userId = typeof args.options.userId !== 'undefined';
    telemetryProps.id = typeof args.options.id !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getTeamId(args)
      .then((teamId: string): Promise<string> => {
        this.teamId = teamId;
        return this.getChannelId(args);
      })
      .then((channelId: string): Promise<string> => {
        this.channelId = channelId;
        return this.getMemberId(args);
      })
      .then((memberId: string) => {
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

        return request.patch(requestOptions);
      })
      .then((member): void => {
        logger.log(member);
        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
  
  private getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return Promise.resolve(args.options.teamId);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groups?$filter=displayName eq '${encodeURIComponent(args.options.teamName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: [{ id: string, resourceProvisioningOptions: string[] }] }>(requestOptions)
      .then(response => {
        const groupItem: { id: string, resourceProvisioningOptions: string[] } | undefined = response.value[0];

        if (!groupItem) {
          return Promise.reject(`The specified team does not exist in the Microsoft Teams`);
        }

        if (groupItem.resourceProvisioningOptions.indexOf('Team') === -1) {
          return Promise.reject(`The specified team does not exist in the Microsoft Teams`);
        }

        if (response.value.length > 1) {
          return Promise.reject(`Multiple Microsoft Teams teams with name ${args.options.teamName} found: ${response.value.map(x => x.id)}`);
        }

        return Promise.resolve(groupItem.id);
      });
  }

  private getChannelId(args: CommandArgs): Promise<string> {
    if (args.options.channelId) {
      return Promise.resolve(args.options.channelId);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(this.teamId)}/channels?$filter=displayName eq '${encodeURIComponent(args.options.channelName as string)}'`,
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

        return Promise.resolve(channelItem.id);
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
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
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.teamId && args.options.teamName) {
      return 'Specify either teamId or teamName, but not both';
    }

    if (!args.options.teamId && !args.options.teamName) {
      return 'Specify teamId or teamName, one is required';
    }

    if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
      return `${args.options.teamId} is not a valid GUID`;
    }

    if (args.options.channelId && args.options.channelName) {
      return 'Specify either channelId or channelName, but not both';
    }

    if (!args.options.channelId && !args.options.channelName) {
      return 'Specify channelId or channelName, one is required';
    }

    if (args.options.channelId && !validation.isValidTeamsChannelId(args.options.channelId)) {
      return `${args.options.channelId} is not a valid Teams Channel ID`;
    }

    if ((args.options.userName && args.options.userId) || 
      (args.options.userName && args.options.id) || 
      (args.options.userId && args.options.id)) {
      return 'Specify either userName, userId or id, but not multiple.';
    }

    if (!args.options.userName && !args.options.userId && !args.options.id) {
      return 'Specify either userName, userId or id, one is required';
    }

    if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
      return `${args.options.userId} is not a valid GUID`;
    }

    if (['owner', 'member'].indexOf(args.options.role) === -1) {
      return `${args.options.role} is not a valid role value. Allowed values owner|member`;
    } 

    return true;
  }
}

module.exports = new TeamsChannelMemberSetCommand();