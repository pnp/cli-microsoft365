import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
    CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const data: any = {
      displayName: args.options.name
    };

    const requestOptions: any = {
      url: `${this.resource}/beta/me/todo/lists`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      data,
      responseType: 'json'
    };

    request
      .post(requestOptions)
      .then((res: any): void => {
        logger.log(res);

        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new TodoListAddCommand();