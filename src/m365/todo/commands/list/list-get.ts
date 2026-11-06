import { Logger } from '../../../../cli/Logger.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';
import { ToDoList } from '../../ToDoList.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().optional().alias('i'),
  name: z.string().optional().alias('n')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TodoListGetCommand extends GraphDelegatedCommand {
  public get name(): string {
    return commands.LIST_GET;
  }

  public get description(): string {
    return 'Gets a specific list of Microsoft To Do task lists';
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
    try {
      const item = await this.getList(args.options);
      await logger.log(item);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getList(options: Options): Promise<any> {
    const requestOptions: CliRequestOptions = {
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    if (options.id) {
      requestOptions.url = `${this.resource}/v1.0/me/todo/lists/${options.id}`;
      const result = await request.get<ToDoList>(requestOptions);
      return result;
    }

    requestOptions.url = `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(options.name!)}'`;
    const result = await request.get<{ value: ToDoList[] }>(requestOptions);

    if (result.value.length === 0) {
      throw `The specified list '${options.name}' does not exist.`;
    }

    return result.value[0];
  }
}

export default new TodoListGetCommand();