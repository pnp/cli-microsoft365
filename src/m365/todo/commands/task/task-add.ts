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

    this
      .getTodoListId(args)
      .then((listId: string): Promise<any> => {
        const requestOptions: any = {
          url: `${endpoint}/me/todo/lists/${listId}/tasks`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'Content-Type': 'application/json'
          },
          data: {
            title: args.options.title
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then((res): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getTodoListId(args: CommandArgs): Promise<string> {
    if (args.options.listId) {
      return Promise.resolve(args.options.listId);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${escape(args.options.listName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<{ value: [{ id: string }] }>(requestOptions)
      .then(response => {
        const taskList: { id: string } | undefined = response.value[0];

        if (!taskList) {
          return Promise.reject(`The specified task list does not exist`);
        }

        return Promise.resolve(taskList.id);
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-t, --title <title>'
      },
      {
        option: '--listName [listName]'
      },
      {
        option: '--listId [listId]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.listId && args.options.listName) {
      return 'Specify listId or listName but not both';
    }

    if (!args.options.listId && !args.options.listName) {
      return 'Specify listId or listName';
    }

    return true;
  }
}

module.exports = new TodoTaskAddCommand();