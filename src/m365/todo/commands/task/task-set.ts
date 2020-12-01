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
  id:string;  
  listName?: string;
  listId?: string;
  title?: string;
  status?: string; 
}

class TodoTaskSetCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TASK_SET}`;
  }

  public get description(): string {
    return 'Update a task in a Microsoft To Do task list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listName = typeof args.options.listName !== 'undefined';
    telemetryProps.status = typeof args.options.status !== 'undefined';
    telemetryProps.title = typeof args.options.title !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0`;
    const data = this.mapRequestBody(args.options);
    this
      .getTodoListId(args)
      .then((listId: string): Promise<any> => {
        const requestOptions: any = {
          url: `${endpoint}/me/todo/lists/${listId}/tasks/${encodeURIComponent(args.options.id)}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'Content-Type': 'application/json'
          },
          data: data,
          responseType: 'json'
        };

        return request.patch(requestOptions);
      })
      .then((res): void => {
        logger.log(res);

        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }
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
        option: '-i, --id <id>',
        description: `The id of the task to update`
      },
      {
        option: '-t, --title [title]',
        description: `Sets the task title`
      },
      {
        option: '-s, --status [status]',
        description: `Set the task title. Allowed values are notStarted|inProgress|completed|waitingOnOthers|deferred`,
        autocomplete: ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred']
      },
      {
        option: '--listName [listName]',
        description: 'The name of the list in which the task exists. Specify either listName or listId but not both'
      },
      {
        option: '--listId [listId]',
        description: 'The id of the list in which the task exists. Specify either listName or listId but not both'
      }    
      
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.id) {
      return 'Specify the id of the task to update';
    }
    if (args.options.listId && args.options.listName) {
      return 'Specify listId or listName but not both';
    }
    if (!args.options.listId && !args.options.listName) {
      return 'Specify listId or listName';
    }
    
    if (args.options.status &&
      args.options.status !== 'notStarted' &&
      args.options.status !== 'inProgress' &&
      args.options.status !== 'completed' &&
      args.options.status !== 'waitingOnOthers' &&
      args.options.status !== 'deferred') {
      return `${args.options.status} is not a valid value. Allowed values are notStarted|inProgress|completed|waitingOnOthers|deferred`;
    }

    return true;
  }
  private mapRequestBody(options: Options): any {
    const requestBody: any = {};

    if (options.status) {
      requestBody.status = options.status;
    }

    if (options.title) {
      requestBody.title = options.title;
    }

    return requestBody;
  }
}


module.exports = new TodoTaskSetCommand();