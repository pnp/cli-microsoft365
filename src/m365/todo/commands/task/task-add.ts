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
  title: string;
  listName?: string;
  listId?: string;
}

class TodoTaskAddCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TASK_ADD}`;
  }

  public get description(): string {
    return 'Add a task to a Microsoft To Do task list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listName = typeof args.options.listName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0`;

    let id: string | undefined;

    const data: any = {
      title: args.options.title
    };

    ((): Promise<any> => {
      if (args.options.listName) {
        const requestOptions: any = {
          url: `${endpoint}/me/todo/lists?$filter=displayName eq '${args.options.listName}'`,
          headers: {
            accept: 'application/json',
            'content-type': 'application/json'
          },
          responseType: 'json'
        };
        return request.get(requestOptions);
      }

      id = args.options.listId;
      return Promise.resolve(undefined as any);
    })().then((res?: { value: [{ id: string }] }) => {
      if (res && res.value) {
        if (!res.value.length) {
          return Promise.reject('No tasks list found with the specified name');
        }

        if (res.value[0] && res.value[0].id) {
          id = res.value[0].id;
        }
      }

      const requestOptions: any = {
        url: `${endpoint}/me/todo/lists/${id}/tasks`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'Content-Type': 'application/json'
        },
        data,
        responseType: 'json'
      };

      return request.post(requestOptions);
    }).then((res): void => {
      logger.log(res);

      if (this.verbose) {
        logger.log(chalk.green('DONE'));
      }
      cb();
    }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-t, --title <title>',
        description: `The title of the task`
      },
      {
        option: '--listName [listName]',
        description: 'The name of the task list in which to create the task in. Specify either listName or listId, not both'
      },
      {
        option: '--listId [listId]',
        description: 'The id of the task list in which to create the task in. Specify either listName or listId, not both'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.listId && args.options.listName) {
      return 'Specify listId or listName, but not both';
    }

    if (!args.options.listId && !args.options.listName) {
      return 'Specify listId or listName';
    }

    return true;
  }
}

module.exports = new TodoTaskAddCommand();