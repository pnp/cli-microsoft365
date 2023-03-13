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
  currentPassword?: string;
  newPassword?: string;
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
        newPassword: typeof args.options.newPassword !== 'undefined'
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

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['objectId', 'userPrincipalName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr(`Updating user ${args.options.userPrincipalName || args.options.objectId}`);
      }

      if (args.options.currentPassword) {
        if (args.options.objectId && args.options.objectId !== accessToken.getUserIdFromAccessToken(auth.service.accessTokens[auth.defaultResource].accessToken)) {
          throw `You can only change your own password. Please use --objectId @meId to reference to your own userId`;
        }
        else if (args.options.userPrincipalName && args.options.userPrincipalName.toLowerCase() !== accessToken.getUserNameFromAccessToken(auth.service.accessTokens[auth.defaultResource].accessToken).toLowerCase()) {
          throw 'You can only change your own password. Please use --userPrincipalName @meUserName to reference to your own user principal name';
        }
      }

      const requestUrl = `${this.resource}/v1.0/users/${formatting.encodeQueryParameter(args.options.objectId ? args.options.objectId : args.options.userPrincipalName as string)}`;
      const manifest: any = this.mapRequestBody(args.options);

      if (Object.keys(manifest).length > 0) {
        if (this.verbose) {
          logger.logToStderr(`Setting the updated properties for the user ${args.options.userPrincipalName || args.options.objectId}`);
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
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};

    const excludeOptions: string[] = [
      'debug',
      'verbose',
      'output',
      'objectId',
      'i',
      'userPrincipalName',
      'n',
      'resetPassword',
      'accountEnabled',
      'currentPassword',
      'newPassword',
      'forceChangePasswordNextSignIn'
    ];

    if (options.accountEnabled !== undefined) {
      requestBody['AccountEnabled'] = options.accountEnabled;
    }

    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
        requestBody[key] = `${(<any>options)[key]}`;
      }
    });

    if (options.resetPassword) {
      requestBody.passwordProfile = {
        forceChangePasswordNextSignIn: options.forceChangePasswordNextSignIn || false,
        password: options.newPassword
      };
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
}

module.exports = new AadUserSetCommand();