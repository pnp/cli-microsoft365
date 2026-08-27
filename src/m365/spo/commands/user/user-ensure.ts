import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { Group } from '@microsoft/microsoft-graph-types';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(webUrl => validation.isValidSharePointUrl(webUrl) === true, {
    error: e => validation.isValidSharePointUrl(e.input as string).toString()
  }).alias('u'),
  entraId: z.string().refine(id => validation.isValidGuid(id), { error: e => `${e.input} is not a valid GUID.` }).optional(),
  userName: z.string().refine(userName => validation.isValidUserPrincipalName(userName), { error: e => `${e.input} is not a valid userName.` }).optional(),
  loginName: z.string().optional(),
  entraGroupId: z.string().refine(id => validation.isValidGuid(id), { error: e => `${e.input} is not a valid GUID for option 'entraGroupId'.` }).optional(),
  entraGroupName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoUserEnsureCommand extends SpoCommand {
  public get name(): string {
    return commands.USER_ENSURE;
  }

  public get description(): string {
    return 'Ensures that a user is available on a specific site';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema.refine(opts => [opts.entraId, opts.userName, opts.loginName, opts.entraGroupId, opts.entraGroupName].filter(value => value !== undefined).length === 1, {
      error: 'Specify one of the following options: entraId, userName, loginName, entraGroupId, entraGroupName.',
      params: {
        customCode: 'optionSet',
        options: ['entraId', 'userName', 'loginName', 'entraGroupId', 'entraGroupName']
      }
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Ensuring user ${args.options.entraId || args.options.userName || args.options.loginName || args.options.entraGroupId || args.options.entraGroupName} at site ${args.options.webUrl}`);
    }

    try {
      const requestBody = {
        logonName: await this.getUpn(args.options)
      };

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/ensureuser`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        data: requestBody,
        responseType: 'json'
      };

      const response = await request.post(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUpn(options: Options): Promise<string> {
    if (options.userName) {
      return options.userName;
    }

    if (options.entraId) {
      return entraUser.getUpnByUserId(options.entraId);
    }

    if (options.loginName) {
      return options.loginName;
    }

    let upn: string = '';
    if (options.entraGroupId || options.entraGroupName) {
      const entraGroup = await this.getEntraGroup(options.entraGroupId, options.entraGroupName);
      upn = entraGroup.mailEnabled ? `c:0o.c|federateddirectoryclaimprovider|${entraGroup.id}` : `c:0t.c|tenant|${entraGroup.id}`;
    }

    return upn;
  }

  private async getEntraGroup(entraGroupId?: string, entraGroupName?: string): Promise<Group> {
    if (entraGroupId) {
      return entraGroup.getGroupById(entraGroupId);
    }

    return entraGroup.getGroupByDisplayName(entraGroupName!);
  }
}

export default new SpoUserEnsureCommand();
