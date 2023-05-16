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
  forceChangePasswordNextSignIn?: boolean;
  forceChangePasswordNextSignInWithMfa?: boolean;
  usageLocation?: string;
  officeLocation?: string;
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

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
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
        officeLocation: typeof args.options.officeLocation !== 'undefined',
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
        option: '--accountEnabled [accountEnabled]',
        autocomplete: ['true', 'false']
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
          const regex = new RegExp('^[a-zA-Z]{2}$');
          if (!regex.test(args.options.usageLocation)) {
            return `'${args.options.usageLocation}' is not a valid usageLocation.`;
          }
        }

        if (args.options.preferredLanguage && args.options.preferredLanguage.length < 2) {
          return `'${args.options.preferredLanguage}' is not a valid preferredLanguage`;
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

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['managerUserId', 'managerUserName'],
        runsWhen: (args) => args.options.managerId || args.options.managerUserName
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
      const manifest: any = this.mapRequestBody(args.options);
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/users`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: manifest
      };

      const user = await request.post<ExtendedUser>(requestOptions);
      user.password = requestOptions.data.passwordProfile.password;

      if (args.options.managerUserId || args.options.managerUserName) {
        const managerRequestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/users/${user.id}/manager/$ref`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          data: {
            '@odata.id': `${this.resource}/v1.0/users/${args.options.managerUserId || args.options.managerUserName}`
          }
        };
        await request.put(managerRequestOptions);
      }

      logger.log(user);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {
      accountEnabled: options.accountEnabled ?? true,
      displayName: options.displayName,
      userPrincipalName: options.userName,
      mailNickName: options.mailNickname ?? options.userName.split('@')[0],
      passwordProfile: {
        forceChangePasswordNextSignIn: options.forceChangePasswordNextSignIn || false,
        forceChangePasswordNextSignInWithMfa: options.forceChangePasswordNextSignInWithMfa || false,
        password: options.password ?? this.generatePassword()
      },
      givenName: options.firstName,
      surName: options.lastName,
      usageLocation: options.usageLocation,
      officeLocation: options.officeLocation,
      jobTitle: options.jobTitle,
      companyName: options.companyName,
      department: options.department,
      preferredLanguage: options.preferredLanguage
    };

    this.addUnknownOptionsToPayload(requestBody, options);

    // Replace every empty string with null
    for (const key in requestBody) {
      if (typeof requestBody[key] === 'string' && requestBody[key].trim() === '') {
        requestBody[key] = null;
      }
    }

    return requestBody;
  }

  /**
   * Generate a password with at least: one digit, one lowercase chracter, one uppercase character and special character.
   */
  private generatePassword(): string {
    const numberChars = '0123456789';
    const upperChars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    const lowerChars = 'abcdefghijklmnopqrstuvwxyz';
    const specialChars = '-_@%$#*&';
    const allChars = numberChars + upperChars + lowerChars + specialChars;
    let randPasswordArray = Array(15);

    randPasswordArray[0] = numberChars;
    randPasswordArray[1] = upperChars;
    randPasswordArray[2] = lowerChars;
    randPasswordArray[3] = specialChars;
    randPasswordArray = randPasswordArray.fill(allChars, 4);

    const randomCharacterArray = randPasswordArray.map((charSet: string) => charSet[Math.floor(Math.random() * charSet.length)]);
    return this.shuffleArray(randomCharacterArray).join('');
  }

  private shuffleArray(characterArray: string[]): string[] {
    for (let i = characterArray.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      const temp = characterArray[i];
      characterArray[i] = characterArray[j];
      characterArray[j] = temp;
    }
    return characterArray;
  }
}

module.exports = new AadUserAddCommand();