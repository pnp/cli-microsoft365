import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';
import { ToDoTask } from '../../ToDoTask.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().alias('i'),
  listName: z.string().optional(),
  listId: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TodoTaskGetCommand extends GraphDelegatedCommand {
  public get name(): string {
    return commands.TASK_GET;
  }

  public get description(): string {
    return 'Gets a specific task from a Microsoft To Do task list';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema.refine(opts => [opts.listId, opts.listName].filter(x => x !== undefined).length === 1, {
      message: `Specify either 'listId' or 'listName', but not both.`,
      params: { customCode: 'optionSet', options: ['listId', 'listName'] }
    });
  }

  private async getTodoListId(args: CommandArgs): Promise<string> {
    if (args.options.listId) {
      return args.options.listId;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.listName!)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: [{ id: string }] }>(requestOptions);

    const taskList = response.value[0];
    if (!taskList) {
      throw `The specified task list does not exist`;
    }

    return taskList.id;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const listId: string = await this.getTodoListId(args);
      const requestOptions: any = {
        url: `${this.resource}/v1.0/me/todo/lists/${listId}/tasks/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const item: ToDoTask = await request.get(requestOptions);

      if (!cli.shouldTrimOutput(args.options.output)) {
        await logger.log(item);
      }
      else {
        await logger.log({
          id: item.id,
          title: item.title,
          status: item.status,
          createdDateTime: item.createdDateTime,
          lastModifiedDateTime: item.lastModifiedDateTime
        });
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TodoTaskGetCommand();