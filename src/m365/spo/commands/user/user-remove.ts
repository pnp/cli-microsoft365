import { Group } from '@microsoft/microsoft-graph-types';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { spo } from '../../../../utils/spo.js';
import request, { CliRequestOptions } from '../../../../request.js';
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
  id: z.number().optional().alias('i'),
  loginName: z.string().optional(),
  email: z.string().refine(email => validation.isValidUserPrincipalName(email), { error: e => `${e.input} is not a valid email.` }).optional(),
  userName: z.string().refine(userName => validation.isValidUserPrincipalName(userName), { error: e => `${e.input} is not a valid userName.` }).optional(),
  entraGroupId: z.string().refine(id => validation.isValidGuid(id), { error: e => `${e.input} is not a valid GUID.` }).optional(),
  entraGroupName: z.string().optional(),
  force: z.boolean().optional().alias('f')
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
};
interface CommandArgs {
  options: Options;
}
class SpoUserRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.USER_REMOVE;
  }

  public get description(): string {
    return 'Removes user from specific web';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema.refine(opts => [opts.id, opts.loginName, opts.email, opts.userName, opts.entraGroupId, opts.entraGroupName].filter(value => value !== undefined).length === 1, {
      error: 'Specify one of the following options: id, loginName, email, userName, entraGroupId, entraGroupName.',
      params: {
        customCode: 'optionSet',
        options: ['id', 'loginName', 'email', 'userName', 'entraGroupId', 'entraGroupName']
      }
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeUser(logger, args.options);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove specified user from the site ${args.options.webUrl}?` });

      if (result) {
        await this.removeUser(logger, args.options);
      }
    }
  }

  private async removeUser(logger: Logger, options: Options): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing user from  subsite ${options.webUrl} ...`);
    }
    try {
      let requestUrl: string = `${encodeURI(options.webUrl)}/_api/web/siteusers/`;
      if (options.id) {
        requestUrl += `removebyid(${options.id})`;
      }
      else if (options.loginName) {
        requestUrl += `removeByLoginName('${formatting.encodeQueryParameter(options.loginName as string)}')`;
      }
      else if (options.email) {
        const user = await spo.getUserByEmail(options.webUrl, options.email, logger, this.verbose);
        requestUrl += `removebyid(${user.Id})`;
      }
      else if (options.userName) {
        const user = await this.getUser(options);

        if (!user) {
          throw new Error(`User not found: ${options.userName}`);
        }

        if (this.verbose) {
          await logger.logToStderr(`Removing user ${user.Title} ...`);
        }
        requestUrl += `removebyid(${user.Id})`;
      }
      else if (options.entraGroupId || options.entraGroupName) {
        const entraGroup = await this.getEntraGroup(options);
        if (this.verbose) {
          await logger.logToStderr(`Removing entra group ${entraGroup?.displayName} ...`);
        }
        //for entra groups, M365 groups have an associated email and security groups don't
        if (entraGroup?.mail) {
          //M365 group is prefixed with c:0o.c|federateddirectoryclaimprovider
          requestUrl += `removeByLoginName('c:0o.c|federateddirectoryclaimprovider|${entraGroup.id}')`;
        }
        else {
          //security group is prefixed with c:0t.c|tenant
          requestUrl += `removeByLoginName('c:0t.c|tenant|${entraGroup?.id}')`;
        }
      }

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUser(options: Options): Promise<any> {
    const requestUrl: string = `${options.webUrl}/_api/web/siteusers?$filter=UserPrincipalName eq ('${formatting.encodeQueryParameter(options.userName!)}')`;
    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const userInstance = await request.get(requestOptions);
    return (userInstance as {
      value: SpoUser[];
    }).value[0];
  }

  private async getEntraGroup(options: Options): Promise<Group> {
    return options.entraGroupId ? await entraGroup.getGroupById(options.entraGroupId) : await entraGroup.getGroupByDisplayName(options.entraGroupName!);
  }
}

export default new SpoUserRemoveCommand();
