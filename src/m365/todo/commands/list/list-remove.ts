import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.string().optional().alias('n'),
  id: z.string().optional().alias('i'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TodoListRemoveCommand extends GraphDelegatedCommand {
  public get name(): string {
    return commands.LIST_REMOVE;
  }

  public get description(): string {
    return 'Removes a Microsoft To Do task list';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema.refine(opts => [opts.id, opts.name].filter(x => x !== undefined).length === 1, {
      message: `Specify either 'id' or 'name', but not both.`,
      params: { customCode: 'optionSet', options: ['id', 'name'] }
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeList(args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the task list ${args.options.id || args.options.name}?` });

      if (result) {
        await this.removeList(args);
      }
    }
  }

  private async getListId(args: CommandArgs): Promise<string | undefined> {
    if (args.options.id) {
      return args.options.id as string;
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.name!)}'`,
      headers: {
        accept: "application/json;odata.metadata=none"
      },
      responseType: 'json'
    };

    const response: any = await request.get(requestOptions);

    return response.value && response.value.length === 1 ? response.value[0].id : undefined;
  }

  private async removeList(args: CommandArgs): Promise<void> {
    try {
      const listId = await this.getListId(args);

      if (!listId) {
        throw `The list ${args.options.name} cannot be found`;
      }

      const requestOptions: any = {
        url: `${this.resource}/v1.0/me/todo/lists/${listId}`,
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

export default new TodoListRemoveCommand();