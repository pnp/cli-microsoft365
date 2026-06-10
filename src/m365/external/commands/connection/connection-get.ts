import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().optional().alias('i'),
  name: z.string().optional().alias('n')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class ExternalConnectionGetCommand extends GraphCommand {
  public get name(): string {
    return commands.CONNECTION_GET;
  }

  public get description(): string {
    return 'Gets a specific external connection';
  }

  public alias(): string[] | undefined {
    return [commands.EXTERNALCONNECTION_GET];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.id, options.name].filter(x => x !== undefined).length === 1, {
        message: `Specify either 'id' or 'name', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'name']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let url: string = `${this.resource}/v1.0/external/connections`;
    if (args.options.id) {
      url += `/${formatting.encodeQueryParameter(args.options.id as string)}`;
    }
    else {
      url += `?$filter=name eq '${formatting.encodeQueryParameter(args.options.name as string)}'`;
    }

    const requestOptions: any = {
      url: url,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      let res = await request.get<any>(requestOptions);

      if (args.options.name) {
        if (res.value.length === 0) {
          throw `External connection with name '${args.options.name}' not found`;
        }

        res = res.value[0];
      }

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new ExternalConnectionGetCommand();