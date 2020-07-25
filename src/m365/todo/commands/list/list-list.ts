import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import { ToDoList } from '../../ToDoList';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }
}

module.exports = new TodoListListCommand();