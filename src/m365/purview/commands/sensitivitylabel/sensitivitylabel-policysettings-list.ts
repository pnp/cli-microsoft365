import { z } from 'zod';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.string().refine(val => validation.isValidGuid(val), {
    message: 'The value must be a valid GUID.'
  }).optional(),
  userName: z.string().refine(val => validation.isValidUserPrincipalName(val), {
    message: 'The value must be a valid user principal name (UPN).'
  }).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PurviewSensitivityLabelPolicySettingsListCommand extends GraphCommand {
  public get name(): string {
    return commands.SENSITIVITYLABEL_POLICYSETTINGS_LIST;
  }

  public get description(): string {
    return 'Get a list of policy settings for a sensitivity label';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
    if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName) {
      this.handleError(`Specify at least 'userId' or 'userName' when using application permissions.`);
    }

    const requestUrl: string = args.options.userId || args.options.userName
      ? `${this.resource}/beta/users/${args.options.userId || args.options.userName}/security/informationProtection/labelPolicySettings`
      : `${this.resource}/beta/me/security/informationProtection/labelPolicySettings`;

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const res: any = await request.get<any>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PurviewSensitivityLabelPolicySettingsListCommand();