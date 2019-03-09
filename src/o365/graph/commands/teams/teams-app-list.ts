import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption } from '../../../../Command';
import { TeamsApp } from './TeamsApp'
import { GraphItemsListCommand } from '../GraphItemsListCommand'

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  all?: boolean;
}

class GraphTeamsAppListCommand extends GraphItemsListCommand<TeamsApp> {
  public get name(): string {
    return `${commands.TEAMS_APP_LIST}`;
  }

  public get description(): string {
    return 'Lists apps from the Microsoft Teams app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.all = args.options.all || false;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let endpoint: string = `${auth.service.resource}/v1.0/appCatalogs/teamsApps`;
    if (!args.options.all) {
      endpoint += `?$filter=distributionMethod eq 'organization'`;
    }

    this
      .getAllItems(endpoint, cmd, true)
      .then((): void => {
        if (args.options.output === 'json') {
          cmd.log(this.items);
        }
        else {
          cmd.log(this.items.map(i => {
            return {
              id: i.id,
              displayName: i.displayName,
              distributionMethod: i.distributionMethod
            };
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
        option: '-a, --all',
        description: 'Get apps from your organization\'s app catalog and the Microsoft Teams store'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());

    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to the Microsoft Graph
    using the ${chalk.blue(commands.LOGIN)} command.

  Remarks:

    To list apps in the Microsoft Teams app catalog, you have to first log in to
    the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:

    List all Microsoft Teams apps from your organization's app catalog only
      ${chalk.grey(config.delimiter)} ${commands.TEAMS_APP_LIST}
         
    List all apps from the Microsoft Teams app catalog and the Microsoft Teams store
      ${chalk.grey(config.delimiter)} ${commands.TEAMS_APP_LIST} --all
`);
  }
}

module.exports = new GraphTeamsAppListCommand();