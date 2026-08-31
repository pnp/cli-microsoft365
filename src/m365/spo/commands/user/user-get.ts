import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { Group } from '@microsoft/microsoft-graph-types';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(webUrl => validation.isValidSharePointUrl(webUrl) === true, {
    error: e => validation.isValidSharePointUrl(e.input as string).toString()
  }).alias('u'),
  id: z.number().int().positive().optional().alias('i'),
  email: z.string().refine(email => validation.isValidUserPrincipalName(email), { error: e => `${e.input} is not a valid email.` }).optional(),
  loginName: z.string().optional(),
  userName: z.string().refine(userName => validation.isValidUserPrincipalName(userName), { error: e => `${e.input} is not a valid userName.` }).optional(),
  entraGroupId: z.string().refine(id => validation.isValidGuid(id), { error: e => `${e.input} is not a valid GUID.` }).optional(),
  entraGroupName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface SpoUser {
  Id: number;
  IsHiddenInUI: boolean;
  Title: string;
  PrincipalType: number;
  Email: string;
  Expiration: string;
  IsEmailAuthenticationGuestUser: boolean;
  IsShareByEmailGuestUser: boolean;
  IsSiteAdmin: boolean;
  UserId: {
    NameId: string;
    NameIdIssuer: string;
    urn: string;
  };
  UserPrincipalName: string;
}

interface CommandArgs {
  options: Options;
}

class SpoUserGetCommand extends SpoCommand {
  public get name(): string {
    return commands.USER_GET;
  }

  public get description(): string {
    return 'Gets a site user within specific web';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema.refine(opts => [opts.id, opts.email, opts.loginName, opts.userName, opts.entraGroupId, opts.entraGroupName].filter(value => value !== undefined).length <= 1, {
      error: 'Specify no more than one of the following options: id, email, loginName, userName, entraGroupId, entraGroupName.',
      params: {
        customCode: 'optionSet',
        options: ['id', 'email', 'loginName', 'userName', 'entraGroupId', 'entraGroupName']
      }
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving information for user in site '${args.options.webUrl}'...`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web/`;

    if (args.options.id) {
      requestUrl += `siteusers/GetById('${formatting.encodeQueryParameter(args.options.id.toString())}')`;
    }
    else if (args.options.email) {
      requestUrl += `siteusers/GetByEmail('${formatting.encodeQueryParameter(args.options.email)}')`;
    }
    else if (args.options.loginName) {
      requestUrl += `siteusers/GetByLoginName('${formatting.encodeQueryParameter(args.options.loginName)}')`;
    }
    else if (args.options.userName) {
      const user = await this.getUser(requestUrl, args.options.userName);
      requestUrl += `siteusers/GetById('${formatting.encodeQueryParameter(user.Id.toString())}')`;
    }
    else if (args.options.entraGroupId || args.options.entraGroupName) {
      const entraGroup = await this.getEntraGroup(args.options.entraGroupId!, args.options.entraGroupName!);

      // For entra groups, M365 groups have an associated email and security groups don't
      if (entraGroup?.mail) {
        requestUrl += `siteusers/GetByEmail('${formatting.encodeQueryParameter(entraGroup.mail)}')`;
      }
      else {
        requestUrl += `siteusers/GetByLoginName('c:0t.c|tenant|${entraGroup?.id}')`;
      }
    }
    else {
      requestUrl += `currentuser`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      method: 'GET',
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const userInstance = await request.get(requestOptions);
      await logger.log(userInstance);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUser(baseUrl: string, userName: string): Promise<SpoUser> {
    const requestUrl: string = `${baseUrl}siteusers?$filter=UserPrincipalName eq ('${formatting.encodeQueryParameter(userName)}')`;
    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const userInstance = await request.get(requestOptions);
    const userInstanceValue = (userInstance as {
      value: SpoUser[];
    }).value[0];

    if (!userInstanceValue) {
      throw `User not found: ${userName}`;
    }

    return userInstanceValue;
  }

  private async getEntraGroup(entraGroupId: string, entraGroupName: string): Promise<Group> {
    if (entraGroupId) {
      return entraGroup.getGroupById(entraGroupId);
    }

    return entraGroup.getGroupByDisplayName(entraGroupName);
  }
}

export default new SpoUserGetCommand();