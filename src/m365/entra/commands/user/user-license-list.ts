import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.uuid().optional(),
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid user principal name (UPN).`
  }).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraUserLicenseListCommand extends GraphDelegatedCommand {
  public get name(): string {
    return commands.USER_LICENSE_LIST;
  }

  public get description(): string {
    return 'Lists the license details for a given user';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'skuId', 'skuPartNumber'];
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => !options.userId || !options.userName, {
        error: `Specify either 'userId' or 'userName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving licenses from user: ${args.options.userId || args.options.userName || 'current user'}.`);
    }

    let requestUrl: string = `${this.resource}/v1.0/`;
    if (args.options.userId || args.options.userName) {
      requestUrl += `users/${formatting.encodeQueryParameter(args.options.userId || args.options.userName as string)}`;
    }
    else {
      requestUrl += 'me';
    }
    requestUrl += '/licenseDetails';

    try {
      const items = await odata.getAllItems<any>(requestUrl);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraUserLicenseListCommand();