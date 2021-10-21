import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import { Channel } from '../../Channel';
import commands from '../../commands';
import { ConversationMember } from '../../ConversationMember';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId?: string;
  teamName?: string;
  channelId?: string;
  channelName?: string;
}

class TeamsConversationMemberListCommand extends GraphItemsListCommand<any> {
  private teamId: string = "";

  public get name(): string {
    return commands.CONVERSATIONMEMBER_LIST;
  }

  public get description(): string {
    return 'Lists members of a channel in Microsoft Teams in the current tenant.';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.teamId = typeof args.options.teamId !== 'undefined';
    telemetryProps.teamName = typeof args.options.teamName !== 'undefined';
    telemetryProps.channelId = typeof args.options.channelId !== 'undefined';
    telemetryProps.channelName = typeof args.options.channelName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getTeamId(args)
      .then((teamId: string) => {
        this.teamId = teamId;
        return this.getChannelId(teamId, args);
      }).then((channelId: string) => {
        const endpoint: string = `${this.resource}/v1.0/teams/${encodeURIComponent(this.teamId)}/channels/${encodeURIComponent(channelId)}/members`;
        return this.getAllItems(endpoint, logger, true);
      }).then((): void => {
        if (args.options.output === 'json') {
          logger.log(this.items);
        }
        else {
          logger.log(this.items.map((c: ConversationMember) => {
            return {
              id: c.id,
              displayName: c.displayName,
              userId: c.userId,
              email: c.email
            };
          }));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
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
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.teamId && args.options.teamName) {
      return 'Specify either teamId or teamName, but not both.';
    }

    if (!args.options.teamId && !args.options.teamName) {
      return 'Specify teamId or teamName, one is required';
    }

    if (args.options.teamId && !Utils.isValidGuid(args.options.teamId as string)) {
      return `${args.options.teamId} is not a valid GUID`;
    }

    if (args.options.channelId && args.options.channelName) {
      return 'Specify either channelId or channelName, but not both.';
    }

    if (!args.options.channelId && !args.options.channelName) {
      return 'Specify channelId or channelName, one is required';
    }

    if (args.options.channelId && !Utils.isValidTeamsChannelId(args.options.channelId as string)) {
      return `${args.options.channelId} is not a valid Teams ChannelId`;
    }

    return true;
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
        const filteredResponseByTeam: { id: string, resourceProvisioningOptions: string[] }[] = response.value.filter(t => t.resourceProvisioningOptions.includes('Team'));
        const groupItem: { id: string } | undefined = filteredResponseByTeam[0];

        if (!groupItem) {
          return Promise.reject(`The specified team '${args.options.teamName}' does not exist in the Microsoft Teams`);
        }

        if (filteredResponseByTeam.length > 1) {
          return Promise.reject(`Multiple Microsoft Teams teams with name '${args.options.teamName}' found: ${filteredResponseByTeam.map(x => x.id)}`);
        }

        return Promise.resolve(groupItem.id);
      });
  }

  private getChannelId(teamId: string, args: CommandArgs): Promise<string> {
    if (args.options.channelId) {
      const channelIdRequestOptions: any = {
        url: `${this.resource}/v1.0/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(args.options.channelId as string)}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      return new Promise<string>((resolve: (channelId: string) => void, reject: (error: any) => void): void => {
        request
          .get<Channel>(channelIdRequestOptions)
          .then((response: Channel) => {
            const channelItem: Channel | undefined = response;
            return resolve(channelItem.id);
          }, (err: any) => {
            if (err.error && err.error.code === "NotFound") {
              return reject(`The specified channel '${args.options.channelId}' does not exist or is invalid in the Microsoft Teams team with ID '${teamId}'`);
            }
            else {
              return reject(err);
            }
          });
      });
    }
    else {
      const channelRequestOptions: any = {
        url: `${this.resource}/v1.0/teams/${encodeURIComponent(teamId)}/channels?$filter=displayName eq '${encodeURIComponent(args.options.channelName as string)}'`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      return new Promise<string>((resolve: (channelId: string) => void, reject: (error: any) => void): void => {
        request
          .get<{ value: Channel[] }>(channelRequestOptions)
          .then(response => {
            const channelItem: Channel | undefined = response.value[0];

            if (!channelItem) {
              return reject(`The specified channel '${args.options.channelName}' does not exist in the Microsoft Teams team with ID '${teamId}'`);
            }

            return resolve(channelItem.id);
          }, err => reject(err));
      });
    }
  }
}

module.exports = new TeamsConversationMemberListCommand();