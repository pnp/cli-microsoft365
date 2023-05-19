import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { accessToken } from '../../../../utils/accessToken';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  objectId?: string;
  userPrincipalName?: string;
  accountEnabled?: boolean;
  resetPassword?: boolean;
  forceChangePasswordNextSignIn?: boolean;
  forceChangePasswordNextSignInWithMfa?: boolean;
  currentPassword?: string;
  newPassword?: string;
  displayName?: string;
  firstName?: string;
  lastName?: string;
  usageLocation?: string;
  officeLocation?: string;
  jobTitle?: string;
  companyName?: string;
  department?: string;
  preferredLanguage?: string;
  managerUserId?: string;
  managerUserName?: string;
  removeManager?: boolean;
}

class AadUserSetCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_SET;
  }

  public get description(): string {
    return 'Updates information about the specified user';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initTypes();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        objectId: typeof args.options.objectId !== 'undefined',
        userPrincipalName: typeof args.options.userPrincipalName !== 'undefined',
        accountEnabled: !!args.options.accountEnabled,
        resetPassword: !!args.options.resetPassword,
        forceChangePasswordNextSignIn: !!args.options.forceChangePasswordNextSignIn,
        currentPassword: typeof args.options.currentPassword !== 'undefined',
        newPassword: typeof args.options.newPassword !== 'undefined',
        displayName: typeof args.options.displayName !== 'undefined',
        firstName: typeof args.options.firstName !== 'undefined',
        lastName: typeof args.options.lastName !== 'undefined',
        forceChangePasswordNextSignInWithMfa: !!args.options.forceChangePasswordNextSignInWithMfa,
        usageLocation: typeof args.options.usageLocation !== 'undefined',
        officeLocation: typeof args.options.officeLocation !== 'undefined',
        jobTitle: typeof args.options.jobTitle !== 'undefined',
        companyName: typeof args.options.companyName !== 'undefined',
        department: typeof args.options.department !== 'undefined',
        preferredLanguage: typeof args.options.preferredLanguage !== 'undefined',
        managerUserId: typeof args.options.managerUserId !== 'undefined',
        managerUserName: typeof args.options.managerUserName !== 'undefined',
        removeManager: typeof args.options.removeManager !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --objectId [objectId]'
      },
      {
        option: '-n, --userPrincipalName [userPrincipalName]'
      },
      {
        option: '--accountEnabled [accountEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--resetPassword'
      },
      {
        option: '--forceChangePasswordNextSignIn'
      },
      {
        option: '--currentPassword [currentPassword]'
      },
      {
        option: '--newPassword [newPassword]'
      },
      {
        option: '--displayName [displayName]'
      },
      {
        option: '--firstName [firstName]'
      },
      {
        option: '--lastName [lastName]'
      },
      {
        option: '--forceChangePasswordNextSignInWithMfa'
      },
      {
        option: '--usageLocation [usageLocation]'
      },
      {
        option: '--officeLocation [officeLocation]'
      },
      {
        option: '--jobTitle [jobTitle]'
      },
      {
        option: '--companyName [companyName]'
      },
      {
        option: '--department [department]'
      },
      {
        option: '--preferredLanguage [preferredLanguage]'
      },
      {
        option: '--managerUserId [managerUserId]'
      },
      {
        option: '--managerUserName [managerUserName]'
      },
      {
        option: '--removeManager'
      }
    );
  }

  #initTypes(): void {
    this.types.boolean.push('accountEnabled');
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.objectId &&
          !validation.isValidGuid(args.options.objectId)) {
          return `${args.options.objectId} is not a valid GUID`;
        }

        if (args.options.userPrincipalName && !validation.isValidUserPrincipalName(args.options.userPrincipalName)) {
          return `${args.options.userPrincipalName} is not a valid userPrincipalName`;
        }

        if (!args.options.resetPassword && ((args.options.currentPassword && !args.options.newPassword) || (args.options.newPassword && !args.options.currentPassword))) {
          return `Specify both currentPassword and newPassword when you want to change your password`;
        }

        if (args.options.resetPassword && args.options.currentPassword) {
          return `When resetting a user's password, don't specify the current password`;
        }

        if (args.options.resetPassword && !args.options.newPassword) {
          return `When resetting a user's password, specify the new password to set for the user, using the newPassword option`;
        }

        if (args.options.firstName && args.options.firstName.length > 64) {
          return `The max lenght for the firstName option is 64 characters`;
        }

        if (args.options.lastName && args.options.lastName.length > 64) {
          return `The max lenght for the lastName option is 64 characters`;
        }

        if (args.options.forceChangePasswordNextSignIn && !args.options.resetPassword) {
          return `The option forceChangePasswordNextSignIn can only be used in combination with the resetPassword option`;
        }

        if (args.options.forceChangePasswordNextSignInWithMfa && !args.options.resetPassword) {
          return `The option forceChangePasswordNextSignInWithMfa can only be used in combination with the resetPassword option`;
        }

        if (args.options.usageLocation) {
          const regex = new RegExp('^[a-zA-Z]{2}$');
          if (!regex.test(args.options.usageLocation)) {
            return `'${args.options.usageLocation}' is not a valid usageLocation.`;
          }
        }

        if (args.options.jobTitle && args.options.jobTitle.length > 128) {
          return `The max lenght for the jobTitle option is 128 characters`;
        }

        if (args.options.companyName && args.options.companyName.length > 64) {
          return `The max lenght for the companyName option is 64 characters`;
        }

        if (args.options.department && args.options.department.length > 64) {
          return `The max lenght for the department option is 64 characters`;
        }

        if (args.options.preferredLanguage && args.options.preferredLanguage.length < 2) {
          return `'${args.options.preferredLanguage}' is not a valid preferredLanguage`;
        }

        if (args.options.managerUserName && !validation.isValidUserPrincipalName(args.options.managerUserName)) {
          return `'${args.options.managerUserName}' is not a valid user principal name`;
        }

        if (args.options.managerUserId && !validation.isValidGuid(args.options.managerUserId)) {
          return `'${args.options.managerUserId}' is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['objectId', 'userPrincipalName']
      },
      {
        options: ['managerUserId', 'managerUserName', 'removeManager'],
        runsWhen: (args) => args.options.managerUserId || args.options.managerUserName || args.options.removeManager
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.currentPassword) {
        if (args.options.objectId && args.options.objectId !== accessToken.getUserIdFromAccessToken(auth.service.accessTokens[auth.defaultResource].accessToken)) {
          throw `You can only change your own password. Please use --objectId @meId to reference to your own userId`;
        }
        else if (args.options.userPrincipalName && args.options.userPrincipalName.toLowerCase() !== accessToken.getUserNameFromAccessToken(auth.service.accessTokens[auth.defaultResource].accessToken).toLowerCase()) {
          throw 'You can only change your own password. Please use --userPrincipalName @meUserName to reference to your own user principal name';
        }
      }

      if (this.verbose) {
        logger.logToStderr(`Updating user ${args.options.userPrincipalName || args.options.objectId}`);
      }

      const requestUrl = `${this.resource}/v1.0/users/${formatting.encodeQueryParameter(args.options.objectId ? args.options.objectId : args.options.userPrincipalName as string)}`;
      const manifest: any = this.mapRequestBody(args.options);

      if (Object.keys(manifest).some(k => manifest[k] !== undefined)) {
        if (this.verbose) {
          logger.logToStderr(`Setting the updated properties for user ${args.options.userPrincipalName || args.options.objectId}`);
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
          logger.logToStderr(`Updating the manager to ${args.options.managerUserId || args.options.managerUserName}`);
        }
        await this.updateManager(args.options);
      }
      else if (args.options.removeManager) {
        if (this.verbose) {
          logger.logToStderr('Removing the manager');
        }
        const user = args.options.objectId || args.options.userPrincipalName;
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

    this.addUnknownOptionsToPayload(requestBody, options);

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
      logger.logToStderr(`Changing password for user ${options.userPrincipalName || options.objectId}`);
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
      url: `${this.resource}/v1.0/users/${options.objectId || options.userPrincipalName}/manager/$ref`,
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

module.exports = new AadUserSetCommand();