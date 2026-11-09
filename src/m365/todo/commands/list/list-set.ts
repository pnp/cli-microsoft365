import { Logger } from '../../../../cli/Logger.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().optional().alias('i'),
  name: z.string().optional().alias('n'),
  newName: z.string()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TodoListSetCommand extends GraphDelegatedCommand {
  public get name(): string {
    return commands.LIST_SET;
  }

  public get description(): string {
    return 'Updates a Microsoft To Do task list';
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
    const endpoint: string = `${this.resource}/v1.0`;
    const data: any = {
      displayName: args.options.newName
    };

    try {
      const listId: string = await this.getListId(args);

      if (!listId) {
        throw `The list ${args.options.name} cannot be found`;
      }

      const requestOptions: any = {
        url: `${endpoint}/me/todo/lists/${listId}`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        data,
        responseType: 'json'
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getListId(args: CommandArgs): Promise<string> {
    const endpoint: string = `${this.resource}/v1.0`;
    if (args.options.id) {
      return args.options.id;
    }

    const requestOptions: any = {
      url: `${endpoint}/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.name!)}'`,
      headers: {
        accept: "application/json;odata.metadata=none"
      },
      responseType: 'json'
    };

    return request
      .get(requestOptions)
      .then((response: any) => response.value && response.value.length === 1 ? response.value[0].id : null);
  }
}

export default new TodoListSetCommand();