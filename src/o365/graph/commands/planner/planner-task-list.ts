import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import { GraphItemsListCommand } from '../GraphItemsListCommand';
import { Task } from './Task';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string
}

class GraphPlannerTaskListCommand extends GraphItemsListCommand<Task> {
  public get name(): string {
    return `${commands.PLANNER_TASK_LIST}`;
  }

  public get description(): string {
    return 'Lists Planner tasks of the user.';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.userId = typeof args.options.userId !== 'undefined';
    telemetryProps.userName = typeof args.options.userName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let endpoint: string = "/me/planner/tasks";

    if (args.options.userId || args.options.userName)
      endpoint = `/users/${encodeURIComponent(args.options.userId ? args.options.userId : args.options.userName as string)}/planner/tasks`;

    endpoint = `${auth.service.resource}/v1.0${endpoint}`;
    
    this
    .getAllItems(endpoint, cmd, true)
    .then((): void => {
      if (args.options.output === 'json') {
        cmd.log(this.items);
      }
      else {
        cmd.log(this.items.map(t => {
          const task: any = {
            id: t.id,
            title: t.title,
            startDateTime: t.startDateTime,
            dueDateTime: t.dueDateTime,
            completedDateTime: t.completedDateTime
          };
          return task;
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
        option: '--userId [userId]',
        description: 'The ID of the user to retrieve information for. Specify userId or userName but not both. If none of them are specified, current user tasks will be returned.'
      },
      {
        option: '--userName [userName]',
        description: 'The username of the user to retrieve information for. Specify userId or userName but not both. If none of them are specified, current user tasks will be returned.'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.userId && args.options.userName) {
        return 'Specify either userId or userName but not both. If none of them are specified, current user tasks will be returned.';
      }

      if (args.options.userId &&
        !Utils.isValidGuid(args.options.userId)) {
        return `${ args.options.userId } is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${ chalk.yellow('Important:') } before using this command, log in to the Microsoft Graph
    using the ${ chalk.blue(commands.LOGIN) } command.

      Remarks:

    To list planner tasks of a user, you have to first log in to
    the Microsoft Graph using the ${ chalk.blue(commands.LOGIN) } command,
      eg.${ chalk.grey(`${config.delimiter} ${commands.LOGIN}`) }.

    You can retrieve information about a user, either by specifying that user's
    userId or user name(${ chalk.grey(`userPrincipalName`) }), but not both.If none of them are specified, current user tasks will be returned.

    If current user don't have permission to retrieve tasks of the specified userId or username, you will get
    ${ chalk.grey(`You do not have the required permissions to access this item.`) }
    error.

      Examples:

    List tasks of the current logged in user
    ${ chalk.grey(config.delimiter) } ${ this.name }

    List tasks of the user with userId ${ chalk.grey(`1caf7dcd-7e83-4c3a-94f7-932a1299c844`) }
    ${ chalk.grey(config.delimiter) } ${ this.name } --userId 1caf7dcd-7e83-4c3a-94f7-932a1299c844

    List tasks of the user with username ${ chalk.grey(`AarifS@contoso.onmicrosoft.com`) }
    ${ chalk.grey(config.delimiter) } ${ this.name } --userName AarifS@contoso.onmicrosoft.com
      `);
  }
}

module.exports = new GraphPlannerTaskListCommand();