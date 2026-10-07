import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.coerce.number().optional(),
  email: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class VivaEngageUserGetCommand extends VivaEngageCommand {
  public get name(): string {
    return commands.ENGAGE_USER_GET;
  }

  public get description(): string {
    return 'Retrieves the current user or searches for a user by ID or e-mail';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => !(opts.id !== undefined && opts.email !== undefined), {
        message: 'You are only allowed to search by ID or e-mail but not both',
        params: {
          customCode: 'optionSet',
          options: ['id', 'email']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let endPoint = `${this.resource}/v1/users/current.json`;

    if (args.options.id) {
      endPoint = `${this.resource}/v1/users/${args.options.id}.json`;
    }
    else if (args.options.email) {
      endPoint = `${this.resource}/v1/users/by_email.json?email=${formatting.encodeQueryParameter(args.options.email)}`;
    }

    const requestOptions: CliRequestOptions = {
      url: endPoint,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const res: any = await request.get(requestOptions);

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new VivaEngageUserGetCommand();