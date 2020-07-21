import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import Utils from '../../../../Utils';
import request from '../../../../request';
import { Tab } from '../../Tab';

const vorpal: Vorpal = require('../../../../vorpal-init');

export enum TabTypeOptions {
  Static = "Static",
  Configurable = "Configurable"
}

export interface Deeplink {
  deeplink: string;
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  channelId: string;
  tabId: string;
  label: string;
  tabType: string;
}

class TeamsDeeplinkTabGenerateCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_DEEPLINK_TAB_GENERATE}`;
  }

  public get description(): string {
    return 'Generates a Microsoft Teams deep link from an existing Tab in a Channel';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels/${encodeURIComponent(args.options.channelId)}/tabs/${encodeURIComponent(args.options.tabId)}?$expand=teamsApp`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      json: true
    }

    request
      .get<Tab>(requestOptions)
      .then((res: Tab): void => {
        let appId: string = res.teamsApp.id;
        let entityId: string = res.configuration.entityId;
        let contentUrl: string = encodeURIComponent(res.webUrl);
        let deeplink: Deeplink = { deeplink: "" };

        let tabTypeInput: string = args.options.tabType ? args.options.tabType.trim() : TabTypeOptions.Static;

        if (TabTypeOptions[(tabTypeInput as keyof typeof TabTypeOptions)].valueOf() == TabTypeOptions.Configurable) {         
            let context: string = `{"channelId": "${args.options.channelId}"}`;
            deeplink = { deeplink: `https://teams.microsoft.com/l/entity/${appId}/${entityId}?webUrl=${contentUrl}&label=${args.options.label}&context=${context}` };
            cmd.log(deeplink);
        }
        else if (TabTypeOptions[(tabTypeInput as keyof typeof TabTypeOptions)].valueOf() == TabTypeOptions.Static) {
            deeplink = { deeplink: `https://teams.microsoft.com/l/entity/${appId}/${entityId}?webUrl=${contentUrl}&label=${args.options.label}` };
            cmd.log(deeplink);
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  /**
   * Maps the base TabTypeOptions enum to string array so it can 
   * more easily be used in validation or descriptions.
   */
  protected get tabTypeMap(): string[] {
    const result: string[] = [];

    for (let tabType in TabTypeOptions) {
      result.push(tabType);
    }
    return result;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the team where the tab exists'
      },
      {
        option: '-c, --channelId <channelId>',
        description: 'The ID of the channel where the tab exists'
      },
      {
        option: '-t, --tabId <tabId>',
        description: 'The ID of the tab to generate the deep link from'
      },
      {
        option: '-l, --label <label>',
        description: 'The label to use in the deep link'
      },
      {
        option: '-m, --tabType <TabTypeOptions>',
        description: `The tab type. Allowed values ${this.tabTypeMap.join('|')}. Default ${this.tabTypeMap[0]}`,
        autocomplete: this.tabTypeMap
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.teamId) {
        return 'Required parameter teamId missing';
      }

      if (!args.options.channelId) {
        return 'Required parameter channelId missing';
      }

      if (!args.options.tabId) {
        return 'Required parameter tabId missing';
      }

      if (!args.options.label) {
        return 'Required parameter label missing';
      }

      if (!Utils.isValidGuid(args.options.teamId as string)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      if (!Utils.isValidTeamsChannelId(args.options.channelId as string)) {
        return `${args.options.channelId} is not a valid Teams ChannelId`;
      }

      if (!Utils.isValidGuid(args.options.tabId as string)) {
        return `${args.options.tabId} is not a valid GUID`;
      }

      if (args.options.tabType) {
        const tabTypeOption: TabTypeOptions = TabTypeOptions[(args.options.tabType.trim() as keyof typeof TabTypeOptions)];

        if (!tabTypeOption) {
          return `${args.options.tabType} is not a valid tabType value`;
        }
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    You can only retrieve deeplink to tabs for teams of which you are a member.

    Tabs 'Conversations' and 'Files' are present in every team and therefore not
    included in the list of available tabs to retrieve deeplink.

    Examples:
    Generates a Microsoft Teams deep link from an existing Tab in a Channel
      Get deeplink for tab with id, for a configurable tab
      ${chalk.grey('1432c9da-8b9c-4602-9248-e0800f3e3f07')}
        ${this.name} --teamId '00000000-0000-0000-0000-000000000000' --channelId '19:00000000000000000000000000000000@thread.skype' --tabId 1432c9da-8b9c-4602-9248-e0800f3e3f07 --label MyLabel --tabType Configurable
      
      Get deeplink for tab with id, for a static tab
      ${chalk.grey('1432c9da-8b9c-4602-9248-e0800f3e3f07')}
        ${this.name} --teamId '00000000-0000-0000-0000-000000000000' --channelId '19:00000000000000000000000000000000@thread.skype' --tabId 1432c9da-8b9c-4602-9248-e0800f3e3f07 --label MyLabel --tabType Static
    `);
  }
}

module.exports = new TeamsDeeplinkTabGenerateCommand();