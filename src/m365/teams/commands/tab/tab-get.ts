import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import { Channel } from '../../Channel';
import commands from '../../commands';
import { Tab } from '../../Tab';
import { Team } from '../../Team';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId?: string;
  teamName?: string;
  channelId?: string;
  channelName?: string;
  tabId?: string;
  tabName?: string;
}

class TeamsTabGetCommand extends GraphCommand {
  private teamId: string = "";
  private channelId: string = "";

  public get name(): string {
    return commands.TEAMS_TAB_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Teams tab';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.teamId = typeof args.options.teamId !== 'undefined';
    telemetryProps.teamName = typeof args.options.teamName !== 'undefined';
    telemetryProps.channelId = typeof args.options.channelId !== 'undefined';
    telemetryProps.channelName = typeof args.options.channelName !== 'undefined';
    telemetryProps.tabId = typeof args.options.tabId !== 'undefined';
    telemetryProps.tabName = typeof args.options.tabName !== 'undefined';
    return telemetryProps;
  }

  private getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return Promise.resolve(args.options.teamId);
    }

    const teamRequestOptions: any = {
      url: `${this.resource}/v1.0/me/joinedTeams?$filter=displayName eq '${encodeURIComponent(args.options.teamName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: Team[] }>(teamRequestOptions)
      .then(response => {
        const teamItem: Team | undefined = response.value[0];

        if (!teamItem) {
          return Promise.reject(`The specified team does not exist in the Microsoft Teams`);
        }

        if (response.value.length > 1) {
          return Promise.reject(`Multiple Microsoft Teams teams with name ${args.options.teamName} found: ${response.value.map(x => x.id)}`);
        }

        return Promise.resolve(teamItem.id);
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

  private getTabId(args: CommandArgs): Promise<string> {
    if (args.options.tabId) {
      return Promise.resolve(args.options.tabId);
    }

    const tabRequestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(this.teamId)}/channels/${encodeURIComponent(this.channelId)}/tabs?$filter=displayName eq '${encodeURIComponent(args.options.tabName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: Tab[] }>(tabRequestOptions)
      .then(response => {
        const tabItem: Tab | undefined = response.value[0];

        if (!tabItem) {
          return Promise.reject(`The specified tab does not exist in the Microsoft Teams team channel`);
        }

        return Promise.resolve(tabItem.id);
      });
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
        return this.getTabId(args);
      })
      .then((tabId: string): Promise<Tab> => {
        const endpoint: string = `${this.resource}/v1.0/teams/${encodeURIComponent(this.teamId)}/channels/${encodeURIComponent(this.channelId)}/tabs/${encodeURIComponent(tabId)}`;

        const requestOptions: any = {
          url: endpoint,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        return request.get<Tab>(requestOptions);
      })
      .then((res: Tab): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
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
        option: '--tabId [tabId]'
      },
      {
        option: '--tabName [tabName]'
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

    if (args.options.tabId && args.options.tabName) {
      return 'Specify either tabId or tabName, but not both.';
    }

    if (!args.options.tabId && !args.options.tabName) {
      return 'Specify tabId or tabName, one is required';
    }

    if (args.options.tabId && !Utils.isValidGuid(args.options.tabId as string)) {
      return `${args.options.tabId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new TeamsTabGetCommand();