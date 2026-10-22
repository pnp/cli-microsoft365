import { z } from 'zod';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.looseObject({
  ...globalOptionsZod.shape,
  id: z.uuid().optional().alias('i'),
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid userName.`
  }).optional().alias('n'),
  accountEnabled: z.boolean().optional(),
  resetPassword: z.boolean().optional(),
  forceChangePasswordNextSignIn: z.boolean().optional(),
  forceChangePasswordNextSignInWithMfa: z.boolean().optional(),
  currentPassword: z.string().optional(),
  newPassword: z.string().optional(),
  displayName: z.string().optional(),
  firstName: z.string().max(64, { error: `The max length for the firstName option is 64 characters.` }).optional(),
  lastName: z.string().max(64, { error: `The max length for the lastName option is 64 characters.` }).optional(),
  usageLocation: z.string().regex(/^[a-zA-Z]{2}$/, { error: e => `'${e.input}' is not a valid usageLocation.` }).optional(),
  officeLocation: z.string().optional(),
  jobTitle: z.string().max(128, { error: `The max length for the jobTitle option is 128 characters.` }).optional(),
  companyName: z.string().max(64, { error: `The max length for the companyName option is 64 characters.` }).optional(),
  department: z.string().max(64, { error: `The max length for the department option is 64 characters.` }).optional(),
  preferredLanguage: z.string().min(2, { error: e => `'${e.input}' is not a valid preferredLanguage.` }).optional(),
  managerUserId: z.uuid().optional(),
  managerUserName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid user principal name.`
  }).optional(),
  removeManager: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraUserSetCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_SET;
  }

  public get description(): string {
    return 'Updates information about the specified user';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.id, options.userName].filter(o => o !== undefined).length === 1, {
        error: `Specify either 'id' or 'userName'.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'userName']
        }
      })
      .refine(options => {
        if (!options.managerUserId && !options.managerUserName && !options.removeManager) {
          return true;
        }
        return [options.managerUserId, options.managerUserName, options.removeManager].filter(o => o !== undefined).length === 1;
      }, {
        error: `Specify either 'managerUserId', 'managerUserName', or 'removeManager'.`,
        params: {
          customCode: 'optionSet',
          options: ['managerUserId', 'managerUserName', 'removeManager']
        }
      })
      .refine(options => !(!options.resetPassword && ((options.currentPassword && !options.newPassword) || (options.newPassword && !options.currentPassword))), {
        error: `Specify both currentPassword and newPassword when you want to change your password.`
      })
      .refine(options => !(options.resetPassword && options.currentPassword), {
        error: `When resetting a user's password, don't specify the current password.`
      })
      .refine(options => !(options.resetPassword && !options.newPassword), {
        error: `When resetting a user's password, specify the new password to set for the user, using the newPassword option.`
      })
      .refine(options => !(options.forceChangePasswordNextSignIn && !options.resetPassword), {
        error: `The option forceChangePasswordNextSignIn can only be used in combination with the resetPassword option.`
      })
      .refine(options => !(options.forceChangePasswordNextSignInWithMfa && !options.resetPassword), {
        error: `The option forceChangePasswordNextSignInWithMfa can only be used in combination with the resetPassword option.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.currentPassword) {
        if (args.options.id && args.options.id !== accessToken.getUserIdFromAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken)) {
          throw `You can only change your own password. Please use --id @meId to reference to your own userId`;
        }
        else if (args.options.userName && args.options.userName.toLowerCase() !== accessToken.getUserNameFromAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken).toLowerCase()) {
          throw 'You can only change your own password. Please use --userName @meUserName to reference to your own user principal name';
        }
      }

      if (this.verbose) {
        await logger.logToStderr(`Updating user ${args.options.userName || args.options.id}`);
      }

      const requestUrl = `${this.resource}/v1.0/users/${formatting.encodeQueryParameter(args.options.id ? args.options.id : args.options.userName as string)}`;
      const manifest: any = this.mapRequestBody(args.options);

      if (Object.keys(manifest).some(k => manifest[k] !== undefined)) {
        if (this.verbose) {
          await logger.logToStderr(`Setting the updated properties for user ${args.options.userName || args.options.id}`);
        }
        const requestOptions: CliRequestOptions = {
          url: requestUrl,
          headers: {
            accept: 'application/json'
          },
          responseType: 'json',
          data: manifest
        };

        await request.patch(requestOptions);
      }

      if (args.options.currentPassword) {
        await this.changePassword(requestUrl, args.options, logger);
      }

      if (args.options.managerUserId || args.options.managerUserName) {
        if (this.verbose) {
          await logger.logToStderr(`Updating the manager to ${args.options.managerUserId || args.options.managerUserName}`);
        }
        await this.updateManager(args.options);
      }
      else if (args.options.removeManager) {
        if (this.verbose) {
          await logger.logToStderr('Removing the manager');
        }
        const user = args.options.id || args.options.userName;
        await this.removeManager(user!);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {
      displayName: options.displayName,
      givenName: options.firstName,
      surname: options.lastName,
      usageLocation: options.usageLocation,
      officeLocation: options.officeLocation,
      jobTitle: options.jobTitle,
      companyName: options.companyName,
      department: options.department,
      preferredLanguage: options.preferredLanguage,
      accountEnabled: options.accountEnabled
    };

    this.addUnknownOptionsToPayloadZod(requestBody, options);

    if (options.resetPassword) {
      requestBody.passwordProfile = {
        forceChangePasswordNextSignIn: options.forceChangePasswordNextSignIn || false,
        forceChangePasswordNextSignInWithMfa: options.forceChangePasswordNextSignInWithMfa || false,
        password: options.newPassword
      };
    }

    // Replace every empty string with null
    for (const key in requestBody) {
      if (typeof requestBody[key] === 'string' && requestBody[key].trim() === '') {
        requestBody[key] = null;
      }
    }

    return requestBody;
  }

  private async changePassword(requestUrl: string, options: Options, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Changing password for user ${options.userName || options.id}`);
    }

    const requestBody = {
      currentPassword: options.currentPassword,
      newPassword: options.newPassword
    };
    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}/changePassword`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: requestBody
    };
    await request.post(requestOptions);
  }

  private async updateManager(options: Options): Promise<void> {
    const managerRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/users/${options.id || options.userName}/manager/$ref`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      data: {
        '@odata.id': `${this.resource}/v1.0/users/${options.managerUserId || options.managerUserName}`
      }
    };
    await request.put(managerRequestOptions);
  }

  private async removeManager(userId: string): Promise<void> {
    const managerRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/users/${userId}/manager/$ref`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      }
    };
    await request.delete(managerRequestOptions);
  }
}

export default new EntraUserSetCommand();