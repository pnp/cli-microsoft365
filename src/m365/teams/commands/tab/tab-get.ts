import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import Utils from '../../../../Utils';
import request from '../../../../request';
import { Team } from '../../Team';
import { Channel } from '../../Channel';
import { Tab } from '../../Tab';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  teamName: string;
  channelId: string;
  channelName: string;
  tabId: string;
  tabName: string;
}

class TeamsTabGetCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_TAB_GET}`;
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

    return new Promise<string>((resolve: (result: string) => void, reject: (error: any) => void): void => {
      const teamRequestOptions: any = {
        url: `${this.resource}/v1.0/me/joinedTeams?$filter=displayName eq '${encodeURIComponent(args.options.teamName)}'`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        json: true
      }

      request
        .get<{ value: Team[] }>(teamRequestOptions)
        .then((res: { value: Team[] }): Promise<string> => {
          const teamItem: Team | undefined = res.value[0];

          if (res.value.length > 1) {
            return Promise.reject(`Multiple Microsoft Teams teams with name ${args.options.tabName} found: ${res.value.map(x => x.id)}`);
          }

          if (!teamItem) {
            return Promise.reject(`The specified team does not exist in the Microsoft Teams`);
          }

          const teamId: string = res.value[0].id;
          return Promise.resolve(teamId);
        })
        .then((teamId: string): void => {
          resolve(teamId);
        })
        .catch((error?: string): void => {
          reject(error);
        });
    });
  }

  private getChannelId(args: CommandArgs): Promise<string> {
    if (args.options.channelId) {
      return Promise.resolve(args.options.channelId);
    }

    return new Promise((resolve: (result: string) => void, reject: (error: any) => void): void => {
      const channelRequestOptions: any = {
        url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels?$filter=displayName eq '${encodeURIComponent(args.options.channelName)}'`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        json: true
      }

      request
        .get<{ value: Channel[] }>(channelRequestOptions)
        .then((res: { value: Channel[] }): Promise<string> => {
          const channelItem: Channel | undefined = res.value[0];

          if (!channelItem) {
            return Promise.reject(`The specified channel does not exist in the Microsoft Teams team`);
          }

          const channelId: string = res.value[0].id;
          return Promise.resolve(channelId);
        })
        .then((channelId: string): void => {
          resolve(channelId);
        })
        .catch((error?: string): void => {
          reject(error);
        });
    });
  }

  private getTabId(args: CommandArgs): Promise<string> {
    if (args.options.tabId) {
      return Promise.resolve(args.options.tabId);
    }

    return new Promise((resolve: (result: string) => void, reject: (error: any) => void): void => {
      const channelRequestOptions: any = {
        url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels/${encodeURIComponent(args.options.channelId)}/tabs?$filter=displayName eq '${encodeURIComponent(args.options.tabName)}'`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        json: true
      }

      request
        .get<{ value: Tab[] }>(channelRequestOptions)
        .then((res: { value: Tab[] }): Promise<string> => {
          const tabItem: Tab | undefined = res.value[0];

          if (!tabItem) {
            return Promise.reject(`The specified tab does not exist in the Microsoft Teams team channel`);
          }

          const tabId: string = res.value[0].id;
          return Promise.resolve(tabId);
        })
        .then((tabId: string): void => {
          resolve(tabId);
        })
        .catch((error?: string): void => {
          reject(error);
        });
    });
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getTeamId(args)
      .then((_teamId: string) => {
        args.options.teamId = _teamId;  
        return this.getChannelId(args);
      })
      .then((_channelId: string) => {
        args.options.channelId = _channelId;
        return this.getTabId(args);
      })
      .then((tabId: string) => {
        
        const endpoint: string = `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels/${encodeURIComponent(args.options.channelId)}/tabs/${encodeURIComponent(tabId)}`;

        const requestOptions: any = {
          url: endpoint,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          json: true
        }

        request
          .get<Tab>(requestOptions)
          .then((res: Tab): void => {
            cmd.log(res.webUrl);

            if (this.verbose) {
              cmd.log(vorpal.chalk.green('DONE'));
            }

            cb();
          }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
      })
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId [teamId]',
        description: 'The ID of the Microsoft Teams team where the tab is located. Specify either teamId or teamName but not both'
      },
      {
        option: '--teamName [teamName]',
        description: 'The display name of the Microsoft Teams team where the tab is located. Specify either teamId or teamName but not both'
      },
      {
        option: '-c, --channelId [channelId]',
        description: 'The ID of the Microsoft Teams channel where the tab is located. Specify either channelId or channelName but not both'
      },
      {
        option: '--channelName [channelName]',
        description: 'The display name of the Microsoft Teams channel where the tab is located. Specify either channelId or channelName but not both'
      },
      {
        option: '-t, --tabId [tabId]',
        description: 'The ID of the Microsoft Teams tab. Specify either tabId or tabName but not both'
      },
      {
        option: '--tabName [tabName]',
        description: 'The display name of the Microsoft Teams tab. Specify either tabId or tabName but not both'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.teamId && args.options.teamName) {
        return 'Specify either "teamId" or "teamName", but not both.';
      }

      if (!args.options.teamId && !args.options.teamName) {
        return 'Specify teamId or teamName, one is required';
      }

      if (args.options.teamId && !Utils.isValidGuid(args.options.teamId as string)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      if (args.options.channelId && args.options.channelName) {
        return 'Specify either "channelId" or "channelName", but not both.';
      }

      if (!args.options.channelId && !args.options.channelName) {
        return 'Specify channelId or channelName, one is required';
      }

      if (args.options.channelId && !Utils.isValidTeamsChannelId(args.options.channelId as string)) {
        return `${args.options.channelId} is not a valid Teams ChannelId`;
      }

      if (args.options.tabId && args.options.tabName) {
        return 'Specify either "tabId" or "tabName", but not both.';
      }

      if (!args.options.tabId && !args.options.tabName) {
        return 'Specify tabId or tabName, one is required';
      }

      if (args.options.tabId && !Utils.isValidGuid(args.options.tabId as string)) {
        return `${args.options.tabId} is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    You can only retrieve tabs for teams of which you are a member.

  Examples:
  
    Get url of a Microsoft Teams Tab with id 1432c9da-8b9c-4602-9248-e0800f3e3f07
      ${this.name} --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype --tabId 1432c9da-8b9c-4602-9248-e0800f3e3f07

    Get url of a Microsoft Teams Tab with name "Tab Name"
      ${this.name} --teamName "Team Name" --channelName "Channel Name" --tabName "Tab Name"
    `);
  }
}

module.exports = new TeamsTabGetCommand();