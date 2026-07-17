import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.uuid().optional(),
  userName: z.string().optional(),
  ids: z.string().refine(ids => !ids.split(',').some(e => !validation.isValidGuid(e)), {
    error: e => `'${e.input}' contains one or more invalid GUIDs.`
  })
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraUserLicenseAddCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_LICENSE_ADD;
  }

  public get description(): string {
    return 'Assigns a license to a user';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.userId, options.userName].filter(o => o !== undefined).length === 1, {
        error: `Specify either 'userId' or 'userName'.`,
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const addLicenses = args.options.ids.split(',').map(x => { return { "disabledPlans": [], "skuId": x }; });
    const requestBody = { "addLicenses": addLicenses, "removeLicenses": [] };

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/users/${formatting.encodeQueryParameter(args.options.userId ?? args.options.userName!)}/assignLicense`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      data: requestBody,
      responseType: 'json'
    };

    try {
      const response = await request.post(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraUserLicenseAddCommand();