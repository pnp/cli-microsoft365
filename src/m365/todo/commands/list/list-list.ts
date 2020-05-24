import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import { ToDoList } from '../../ToDoList';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions { }

class TodoListListCommand extends GraphItemsListCommand<ToDoList> {
  public get name(): string {
    return `${commands.LIST_LIST}`;
  }

  public get description(): string {
    return 'Returns a list of Microsoft To Do task lists';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    this
      .getAllItems(`${this.resource}/beta/me/todo/lists`, cmd, true)
      .then((): void => {
        if (args.options.output === 'json') {
          cmd.log(this.items);
        }
        else {
          cmd.log(this.items.map(i => {
            return {
              displayName: i.displayName,
              id: i.id
            };
          }));
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(`  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently in preview
    and is subject to change once the API reached general availability.
    
  Examples:

    Get the list of Microsoft To Do task lists
      ${this.name}
`);
  }
}

module.exports = new TodoListListCommand();