import { SignIn } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid userName.`
  }).optional().alias('n'),
  userId: z.uuid().optional(),
  appDisplayName: z.string().optional(),
  appId: z.uuid().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraUserSigninListCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_SIGNIN_LIST;
  }

  public get description(): string {
    return 'Retrieves the Entra ID user sign-ins for the tenant';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => !(options.userId && options.userName), {
        error: `Specify either 'userId' or 'userName', but not both.`
      })
      .refine(options => !(options.appId && options.appDisplayName), {
        error: `Specify either 'appId' or 'appDisplayName', but not both.`
      });
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'userPrincipalName', 'appId', 'appDisplayName', 'createdDateTime'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let endpoint: string = `${this.resource}/v1.0/auditLogs/signIns`;
      let filter: string = "";
      if (args.options.userName || args.options.userId) {
        filter = args.options.userId ?
          `?$filter=userId eq '${formatting.encodeQueryParameter(args.options.userId as string)}'` :
          `?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(args.options.userName as string)}'`;
      }
      if (args.options.appId || args.options.appDisplayName) {
        filter += filter ? " and " : "?$filter=";
        filter += args.options.appId ?
          `appId eq '${formatting.encodeQueryParameter(args.options.appId)}'` :
          `appDisplayName eq '${formatting.encodeQueryParameter(args.options.appDisplayName as string)}'`;
      }
      endpoint += filter;

      const signins = await odata.getAllItems<SignIn>(endpoint);
      await logger.log(signins);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraUserSigninListCommand();
