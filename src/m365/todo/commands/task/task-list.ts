import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { odata } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { ToDoTask } from '../../ToDoTask';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  listName?: string;
  listId?: string;
}

class TodoTaskListCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_LIST;
  }

  public get description(): string {
    return 'List tasks from a Microsoft To Do task list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listName = typeof args.options.listName !== 'undefined';
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'status', 'createdDateTime', 'lastModifiedDateTime'];
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getTodoListId(args)
      .then((listId: string): Promise<any> => {
        const endpoint: string = `${this.resource}/v1.0/me/todo/lists/${listId}/tasks`;
        return odata.getAllItems(endpoint, logger);
      })
      .then((items: ToDoTask[]): void => {
        if (args.options.output === 'json') {
          logger.log(items);
        }
        else {
          logger.log(items.map(m => {
            return {
              id: m.id,
              title: m.title,
              status: m.status,
              createdDateTime: m.createdDateTime,
              lastModifiedDateTime: m.lastModifiedDateTime
            };
          }));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
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
      return 'Specify listId or listName, but not both';
    }

    if (!args.options.listId && !args.options.listName) {
      return 'Specify listId or listName';
    }

    return true;
  }
}

module.exports = new TodoTaskListCommand();