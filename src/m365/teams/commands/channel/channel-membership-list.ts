import { Channel } from '../../Channel';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import { odata, validation } from '../../../../utils';
import { ConversationMember } from '@microsoft/microsoft-graph-types';
import GlobalOptions from '../../../../GlobalOptions';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import request from '../../../../request';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId?: string;
  teamName?: string;
  channelId?: string;
  channelName?: string;
  role?: string;
}

class TeamsChannelMembershipListCommand extends GraphCommand {
  private teamId: string = '';

  public get name(): string {
    return commands.CHANNEL_MEMBERSHIP_LIST;
  }

  public get description(): string {
    return 'Lists memberships in the specified Microsoft Teams team channel';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'roles', 'displayName', 'userId', 'email'];
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.teamId = typeof args.options.teamId !== 'undefined';
    telemetryProps.teamName = typeof args.options.teamName !== 'undefined';
    telemetryProps.channelId = typeof args.options.channelId !== 'undefined';
    telemetryProps.channelName = typeof args.options.channelName !== 'undefined';
    telemetryProps.role = typeof args.options.role;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getTeamId(args)
      .then((teamId: string): Promise<string> => {
        this.teamId = teamId;
        return this.getChannelId(args);
      })
      .then((channelId: string): Promise<ConversationMember[]> => {
        const endpoint = `${this.resource}/v1.0/teams/${this.teamId}/channels/${channelId}/members`;
        return odata.getAllItems<ConversationMember>(endpoint, logger);
      })
      .then((memberships): void => {
        if (args.options.role) {
          if (args.options.role === 'member') {
            // Members have no role value
            memberships = memberships.filter(i => i.roles!.length === 0);
          }
          else {
            memberships = memberships.filter(i => i.roles!.indexOf(args.options.role!) !== -1);
          }
        }

        logger.log(memberships);
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

    const channelRequestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(this.teamId)}/channels?$filter=displayName eq '${encodeURIComponent(args.options.channelName as string)}'`,
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
          return Promise.reject(`The specified channel does not exist in the Microsoft Teams team`);
        }

        return Promise.resolve(channelItem.id);
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
        option: '-r, --role [role]',
        autocomplete: ['owner', 'member', 'guest']
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

    if (args.options.role) {
      if (['owner', 'member', 'guest'].indexOf(args.options.role) === -1) {
        return `${args.options.role} is not a valid role value. Allowed values owner|member|guest`;
      }
    }

    return true;
  }
}

module.exports = new TeamsChannelMembershipListCommand();