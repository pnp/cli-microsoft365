import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface ExtendedUser extends User {
  password: string;
}

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  displayName: string;
  userName: string;
  accountEnabled?: boolean;
  mailNickname?: string;
  password?: string;
  firstName?: string;
  lastName?: string;
  forceChangePasswordNextSignIn: boolean;
  forceChangePasswordNextSignInWithMfa: boolean;
  usageLocation?: string;
  jobTitle?: string;
  companyName?: string;
  department?: string;
  preferredLanguage?: string;
  managerUserId?: string;
  managerUserName?: string;
}

class AadUserAddCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_ADD;
  }

  public get description(): string {
    return 'Creates a new user';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        accountEnabled: typeof args.options.accountEnabled !== 'undefined',
        mailNickname: typeof args.options.mailNickname !== 'undefined',
        password: typeof args.options.password !== 'undefined',
        firstName: typeof args.options.firstName !== 'undefined',
        lastName: typeof args.options.lastName !== 'undefined',
        forceChangePasswordNextSignIn: !!args.options.forceChangePasswordNextSignIn,
        forceChangePasswordNextSignInWithMfa: !!args.options.forceChangePasswordNextSignInWithMfa,
        usageLocation: typeof args.options.usageLocation !== 'undefined',
        jobTitle: typeof args.options.jobTitle !== 'undefined',
        companyName: typeof args.options.companyName !== 'undefined',
        department: typeof args.options.department !== 'undefined',
        preferredLanguage: typeof args.options.preferredLanguage !== 'undefined',
        managerUserId: typeof args.options.managerUserId !== 'undefined',
        managerUserName: typeof args.options.managerUserName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--displayName <displayName>'
      },
      {
        option: '--userName <userName>'
      },
      {
        option: '--accountEnabled [accountEnabled]'
      },
      {
        option: '--mailNickname [mailNickname]'
      },
      {
        option: '--password [password]'
      },
      {
        option: '--firstName [firstName]'
      },
      {
        option: '--lastName [lastName]'
      },
      {
        option: '--forceChangePasswordNextSignIn'
      },
      {
        option: '--forceChangePasswordNextSignInWithMfa'
      },
      {
        option: '--usageLocation [usageLocation]'
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
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid userName`;
        }

        if (args.options.usageLocation) {
          const regex = new RegExp('^[A-Z]{2}$');
          if (!regex.test(args.options.usageLocation)) {
            return `'${args.options.usageLocation}' is not a valid usageLocation.`;
          }
        }

        if (args.options.firstName && args.options.firstName.length > 64) {
          return `The maximum amount of characters for 'firstName' is 64.`;
        }

        if (args.options.lastName && args.options.lastName.length > 64) {
          return `The maximum amount of characters for 'lastName' is 64.`;
        }

        if (args.options.jobTitle && args.options.jobTitle.length > 128) {
          return `The maximum amount of characters for 'jobTitle' is 128.`;
        }

        if (args.options.companyName && args.options.companyName.length > 64) {
          return `The maximum amount of characters for 'companyName' is 64.`;
        }

        if (args.options.department && args.options.department.length > 64) {
          return `The maximum amount of characters for 'department' is 64.`;
        }

        if (args.options.managerUserId && args.options.managerUserName) {
          return `Specify either 'managerUserId' or 'managerUserName', but not both.`;
        }

        if (args.options.managerUserName && !validation.isValidUserPrincipalName(args.options.managerUserName)) {
          return `'${args.options.managerUserName}' is not a valid user principal name.`;
        }

        if (args.options.managerUserId && !validation.isValidGuid(args.options.managerUserId)) {
          return `'${args.options.managerUserId}' is not a valid GUID.`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.boolean.push('accountEnabled');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Adding user to AAD with displayName ${args.options.displayName} and userPrincipalName ${args.options.userName}`);
    }

    try {
      const requestBody = {
        accountEnabled: args.options.accountEnabled || true,
        displayName: args.options.displayName,
        userPrincipalName: args.options.userName,
        mailNickName: args.options.mailNickname || args.options.userName.split('@')[0],
        passwordProfile: {
          forceChangePasswordNextSignIn: args.options.forceChangePasswordNextSignIn,
          forceChangePasswordNextSignInWithMfa: args.options.forceChangePasswordNextSignInWithMfa,
          password: args.options.password || this.generatePassword()
        },
        givenName: args.options.firstName,
        surName: args.options.lastName,
        usageLocation: args.options.usageLocation,
        jobTitle: args.options.jobTitle,
        companyName: args.options.companyName,
        department: args.options.department,
        preferredLanguage: args.options.preferredLanguage
      };

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/users`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: requestBody
      };
      const user = await request.post<ExtendedUser>(requestOptions);
      user.password = requestBody.passwordProfile.password;

      if (args.options.managerUserId || args.options.managerUserName) {
        requestOptions.url = `${this.resource}/v1.0/users/${user.id}/manager/$ref`;
        requestOptions.data = {
          '@odata.id': `${this.resource}/v1.0/users/${args.options.managerUserId || args.options.managerUserName}`
        };
        await request.put(requestOptions);
      }

      logger.log(user);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private generatePassword(): string {
    return Math.random().toString(36).slice(-12);
  }
}

module.exports = new AadUserAddCommand();