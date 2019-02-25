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

class TeamsAppListCommand extends GraphItemsListCommand<TeamsApp> {
  public get name(): string {
    return `${commands.TEAMS_APP_LIST}`;
  }

  public get description(): string {
    return 'List apps from the Microsoft Teams app catalog';
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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let endpoint: string = `${auth.service.resource}/v1.0/appCatalogs/teamsApps?$filter=distributionMethod eq 'organization'`;

    if (args.options.all) {
      endpoint = `${auth.service.resource}/v1.0/appCatalogs/teamsApps`;
    }
    this
      .getAllItems(endpoint, cmd, true)
      .then((): Promise<any> => {
        return Promise.resolve();
      })
      .then((res?: TeamsApp[]): void => {
        if (res) {
          this.items = res;
        }

        cmd.log(this.items);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());

    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to the Microsoft Graph
      using the ${chalk.blue(commands.LOGIN)} command.

      Examples:

      List all Microsoft Teams Apps from your organization's app catalog only
      ${chalk.grey(config.delimiter)} ${commands.TEAMS_APP_LIST}
         
      List all Apps from the Microsoft Teams app catalog and the Microsoft Teams store
      ${chalk.grey(config.delimiter)} ${commands.TEAMS_APP_LIST} --all

      `);
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.all = args.options.all || false;
    return telemetryProps;
  }
}

module.exports = new TeamsAppListCommand();