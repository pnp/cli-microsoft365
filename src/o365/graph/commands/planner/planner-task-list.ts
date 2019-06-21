import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import { GraphItemsListCommand } from '../GraphItemsListCommand';
import { Task } from './Task';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {

}

class GraphPlannerTaskListCommand extends GraphItemsListCommand<Task> {
  public get name(): string {
    return `${commands.PLANNER_TASK_LIST}`;
  }

  public get description(): string {
    return 'Lists Planner tasks of the current logged in user.';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let endpoint: string = "/me/planner/tasks";

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

    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean  => {
      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to the Microsoft Graph
    using the ${ chalk.blue(commands.LOGIN)} command.

      Remarks:

    To list planner tasks of a current logged in user, you have to first log in to
    the Microsoft Graph using the ${ chalk.blue(commands.LOGIN)} command,
      eg.${ chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

      Examples:

    List tasks of the current logged in user
    ${ chalk.grey(config.delimiter)} ${this.name}`
    );
  }
}

module.exports = new GraphPlannerTaskListCommand();