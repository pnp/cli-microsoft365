import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate, CommandError
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import Utils from '../../../../Utils';
import request from '../../../../request';
import { Team } from '../../Team';
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
    telemetryProps.teamId = args.options.teamId;
    telemetryProps.teamName = args.options.teamName;
    telemetryProps.channelId = args.options.channelId;
    telemetryProps.channelName = args.options.channelName;
    telemetryProps.tabId = args.options.tabId;
    telemetryProps.tabName = args.options.tabName;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const endpoint: string = `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels/${encodeURIComponent(args.options.channelId)}/tabs/${encodeURIComponent(args.options.tabId)}`;

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
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the Microsoft Teams team where the tab is located. Specify either teamId or teamName but not both'
      },
      {
        option: '--teamName <teamName>',
        description: 'The display name of the Microsoft Teams team where the tab is located. Specify either teamId or teamName but not both'
      },
      {
        option: '-c, --channelId <channelId>',
        description: 'The ID of the Microsoft Teams channel where the tab is located. Specify either channelId or channelName but not both'
      },
      {
        option: '--channelName <channelName>',
        description: 'The display name of the Microsoft Teams channel where the tab is located. Specify either channelId or channelName but not both'
      },
      {
        option: '-t, --tabId <tabId>',
        description: 'The ID of the Microsoft Teams tab. Specify either tabId or tabName but not both'
      },
      {
        option: '--tabName <tabName>',
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