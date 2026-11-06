import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().alias('i'),
  listName: z.string().optional(),
  listId: z.string().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TodoTaskRemoveCommand extends GraphDelegatedCommand {
  public get name(): string {
    return commands.TASK_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified Microsoft To Do task';
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeToDoTask(args.options);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the task ${args.options.id} from list ${args.options.listId || args.options.listName}?` });

      if (result) {
        await this.removeToDoTask(args.options);
      }
    }
  }

  private async getToDoListId(options: Options): Promise<string | undefined> {
    if (options.listName) {
      // Search list by its name
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(options.listName)}'`,
        headers: {
          accept: "application/json;odata.metadata=none"
        },
        responseType: 'json'
      };
      const response: { value: { id: string }[] } = await request.get<{ value: { id: string }[] }>(requestOptions);

      return response.value && response.value.length === 1 ? response.value[0].id : undefined;
    }

    return options.listId as string;
  }

  private async removeToDoTask(options: Options): Promise<void> {
    try {
      const toDoListId: string | undefined = await this.getToDoListId(options);

      if (!toDoListId) {
        throw `The list ${options.listName} cannot be found`;
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/me/todo/lists/${toDoListId}/tasks/${options.id}`,
        headers: {
          accept: "application/json;odata.metadata=none"
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TodoTaskRemoveCommand();