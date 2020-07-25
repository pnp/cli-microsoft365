import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
}

class TodoListAddCommand extends GraphCommand {
  public get name(): string {
    return `${commands.LIST_ADD}`;
  }

  public get description(): string {
    return 'Adds a new Microsoft To Do task list';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const body: any = {
      displayName: args.options.name
    };

    const requestOptions: any = {
      url: `${this.resource}/beta/me/todo/lists`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      body,
      json: true
    };

    request
      .post(requestOptions)
      .then((res: any): void => {
        cmd.log(res);

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: `The name of the task list to add`
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new TodoListAddCommand();