import { z } from 'zod';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.string().refine(val => validation.isValidGuid(val), {
    error: 'The value must be a valid GUID.'
  }).optional(),
  userName: z.string().refine(val => validation.isValidUserPrincipalName(val), {
    error: 'The value must be a valid user principal name (UPN).'
  }).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PurviewSensitivityLabelListCommand extends GraphCommand {
  public get name(): string {
    return commands.SENSITIVITYLABEL_LIST;
  }

  public get description(): string {
    return 'Get a list of sensitivity labels';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'isActive'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
    if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName) {
      this.handleError(`Specify at least 'userId' or 'userName' when using application permissions.`);
    }

    const requestUrl: string = args.options.userId || args.options.userName
      ? `${this.resource}/beta/users/${args.options.userId || args.options.userName}/security/informationProtection/sensitivityLabels`
      : `${this.resource}/beta/me/security/informationProtection/sensitivityLabels`;

    try {
      const items = await odata.getAllItems(requestUrl);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PurviewSensitivityLabelListCommand();