import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import { Tab } from './Tab';
import { GraphItemsListCommand } from '../GraphItemsListCommand';


const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  joined?: boolean;
  teamId: string;
  channelId: string;
}

class TabListCommand extends GraphItemsListCommand<Tab> {
  public get name(): string {
    return `${commands.TEAMS_TAB_LIST}`;
  }

  public get description(): string {
    return 'Lists Tabs in a Microsoft Teams team.';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.joined = args.options.joined;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {

    let endpoint: string = `${auth.service.resource}/V1.0/teams/${args.options.teamId}/channels/${args.options.channelId}/tabs?$expand=teamsApp`;

    this
      .getAllItems(endpoint, cmd, true)
      .then((): void => {
        if (args.options.output === 'json') {
          cmd.log(this.items);
        }
        else {
          cmd.log(this.items.map(t => {
            return {
              id: t.id,   
              displayName: t.displayName,                         
              teamsAppTabId: t.teamsApp.id,
            }
          }));
        }

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
        description: 'The ID of the team to list the tab of'
      },
      {
        option: '-c, --channelId <channelId>',
        description: 'The ID of the channel to list the tab of'
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

      if (!Utils.isValidGuid(args.options.teamId as string)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      if (!args.options.channelId) {
        return 'Required parameter channelId missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to the Microsoft Graph
    using the ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:

    To list available tabs in a specific Microsoft Teams team, you have to first log in to
    the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

    You can only see the tab list of a team you are a member of.

    The tabs Conversations and Files are present in every team and therefor not provided as a
    tab in the response from the graph call. 

  Examples:
  
    List all tabs in a Microsoft Teams team
      ${chalk.grey(config.delimiter)} ${this.name} --teamId <teamId> --channelId <channelId>

    Include the all the values from the tab configuration and associated teams app
      ${chalk.grey(config.delimiter)} ${this.name} --teamId <teamId> --channelId <channelId> --output json

  Details:
    
    The command uses Microsoft Graph to retrive the tab information. More details on the underlying
    graph endpoint can be found at:
    https://docs.microsoft.com/en-us/graph/api/teamstab-list?view=graph-rest-1.0
`);
  }
}

module.exports = new TabListCommand();