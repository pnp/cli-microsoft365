import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { odata } from '../../../../utils/odata.js';
import { UserRegistrationDetails } from '@microsoft/microsoft-graph-types';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import { formatting } from '../../../../utils/formatting.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  isAdmin?: boolean;
  userType?: string;
  userPreferredMethodForSecondaryAuthentication?: string;
  systemPreferredAuthenticationMethods?: string;
  isSelfServicePasswordResetRegistered?: boolean;
  isSelfServicePasswordResetEnabled?: boolean;
  isSelfServicePasswordResetCapable?: boolean;
  isMfaRegistered?: boolean;
  isMfaCapable?: boolean;
  isPasswordlessCapable?: boolean;
  isSystemPreferredAuthenticationMethodEnabled?: boolean;
  methodsRegistered?: string;
  userIds?: string;
  userPrincipalNames?: string;
  properties?: string;
}

const authenticationMethods = ['push', 'oath', 'voiceMobile', 'voiceAlternateMobile', 'voiceOffice', 'sms', 'none'];
const methodsRegistered = ['mobilePhone', 'email', 'fido2', 'microsoftAuthenticatorPush', 'softwareOneTimePasscode'];

class EntraUserRegistrationDetailsListCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_REGISTRATIONDETAILS_LIST;
  }
  public get description(): string {
    return 'Retrieves a list of the authentication methods registered for users';
  }

  public defaultProperties(): string[] | undefined {
    return ['userPrincipalName', 'methodsRegistered', 'lastUpdatedDateTime'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        isAdmin: !!args.options.isAdmin,
        userType: typeof args.options.userType !== 'undefined',
        userPreferredMethodForSecondaryAuthentication: typeof args.options.userPreferredMethodForSecondaryAuthentication !== 'undefined',
        systemPreferredAuthenticationMethods: typeof args.options.systemPreferredAuthenticationMethods !== 'undefined',
        isSelfServicePasswordResetRegistered: !!args.options.isSelfServicePasswordResetRegistered,
        isSelfServicePasswordResetEnabled: !!args.options.isSelfServicePasswordResetEnabled,
        isSelfServicePasswordResetCapable: !!args.options.isSelfServicePasswordResetCapable,
        isMfaRegistered: !!args.options.isMfaRegistered,
        isMfaCapable: !!args.options.isMfaCapable,
        isPasswordlessCapable: !!args.options.isPasswordlessCapable,
        isSystemPreferredAuthenticationMethodEnabled: !!args.options.isSystemPreferredAuthenticationMethodEnabled,
        methodsRegistered: typeof args.options.methodsRegistered !== 'undefined',
        userIds: typeof args.options.userIds !== 'undefined',
        userPrincipalNames: typeof args.options.userPrincipalNames !== 'undefined',
        properties: typeof args.options.properties !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--isAdmin [isAdmin]'
      },
      {
        option: '--userType [userType]',
        autocomplete: ['member', 'guest']
      },
      {
        option: '--userPreferredMethodForSecondaryAuthentication [userPreferredMethodForSecondaryAuthentication ]',
        autocomplete: authenticationMethods
      },
      {
        option: '--systemPreferredAuthenticationMethods [systemPreferredAuthenticationMethods ]',
        autocomplete: authenticationMethods
      },
      {
        option: '--isSelfServicePasswordResetRegistered [isSelfServicePasswordResetRegistered]'
      },
      {
        option: '--isSelfServicePasswordResetEnabled [isSelfServicePasswordResetEnabled]'
      },
      {
        option: '--isSelfServicePasswordResetCapable [isSelfServicePasswordResetCapable]'
      },
      {
        option: '--isMfaRegistered [isMfaRegistered]'
      },
      {
        option: '--isMfaCapable [isMfaCapable]'
      },
      {
        option: '--isPasswordlessCapable [isPasswordlessCapable]'
      },
      {
        option: '--isSystemPreferredAuthenticationMethodEnabled [isSystemPreferredAuthenticationMethodEnabled]'
      },
      {
        option: '--methodsRegistered [methodsRegistered]',
        autocomplete: methodsRegistered
      },
      {
        option: '--userIds [userIds]'
      },
      {
        option: '--userPrincipalNames [userPrincipalNames]'
      },
      {
        option: '-p, --properties [properties]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userType) {
          if (['member', 'guest'].every(type => type !== args.options.userType)) {
            return `'${args.options.userType}' is not a valid userType value. Allowed values member, guest`;
          }
        }

        if (args.options.userPreferredMethodForSecondaryAuthentication) {
          const methods = args.options.userPreferredMethodForSecondaryAuthentication.split(',').map(m => m.trim());
          const invalidMethods = methods.filter(m => !authenticationMethods.includes(m));
          if (invalidMethods.length > 0) {
            return `'${args.options.userPreferredMethodForSecondaryAuthentication}' is not a valid userPreferredMethodForSecondaryAuthentication value. Invalid values: ${invalidMethods.join(',')}. Allowed values ${authenticationMethods.join(', ')}`;
          }
        }

        if (args.options.systemPreferredAuthenticationMethods) {
          const methods = args.options.systemPreferredAuthenticationMethods.split(',').map(m => m.trim());
          const invalidMethods = methods.filter(m => !authenticationMethods.includes(m));
          if (invalidMethods.length > 0) {
            return `'${args.options.systemPreferredAuthenticationMethods}' is not a valid systemPreferredAuthenticationMethods value. Invalid values: ${invalidMethods.join(',')}. Allowed values ${authenticationMethods.join(', ')}`;
          }
        }

        if (args.options.methodsRegistered) {
          const methods = args.options.methodsRegistered.split(',').map(m => m.trim());
          const invalidMethods = methods.filter(m => !methodsRegistered.includes(m));
          if (invalidMethods.length > 0) {
            return `'${args.options.methodsRegistered}' is not a valid methodsRegistered value. Invalid values: ${invalidMethods.join(',')}. Allowed values ${methodsRegistered.join(', ')}`;
          }
        }

        if (args.options.userIds) {
          const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.userIds);
          if (isValidGUIDArrayResult !== true) {
            return `The following GUIDs are invalid for the option 'userIds': ${isValidGUIDArrayResult}.`;
          }
        }

        if (args.options.userPrincipalNames) {
          const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.userPrincipalNames);
          if (isValidUPNArrayResult !== true) {
            return `The following user principal names are invalid for the option 'userPrincipalNames': ${isValidUPNArrayResult}.`;
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let userUpns: string[] = [];

      if (args.options.userIds) {
        const ids = args.options.userIds.split(',').map(m => m.trim());
        userUpns = await Promise.all(ids.map(id => entraUser.getUpnByUserId(id)));
      }

      const requestUrl = this.getRequestUrl(args.options, userUpns);
      const result = await odata.getAllItems<UserRegistrationDetails>(requestUrl);

      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getRequestUrl(options: Options, userUpns: string[]): string {
    const queryParameters: string[] = [];

    if (options.properties) {
      queryParameters.push(`$select=${options.properties}`);
    }

    const filters: string[] = [];
    if (options.isAdmin !== undefined) {
      filters.push(`isAdmin eq ${options.isAdmin}`);
    }

    if (options.isMfaCapable !== undefined) {
      filters.push(`isMfaCapable eq ${options.isMfaCapable}`);
    }

    if (options.isMfaRegistered !== undefined) {
      filters.push(`isMfaRegistered eq ${options.isMfaRegistered}`);
    }

    if (options.isPasswordlessCapable !== undefined) {
      filters.push(`isPasswordlessCapable eq ${options.isPasswordlessCapable}`);
    }

    if (options.isSelfServicePasswordResetCapable !== undefined) {
      filters.push(`isSelfServicePasswordResetCapable eq ${options.isSelfServicePasswordResetCapable}`);
    }

    if (options.isSelfServicePasswordResetEnabled !== undefined) {
      filters.push(`isSelfServicePasswordResetEnabled eq ${options.isSelfServicePasswordResetEnabled}`);
    }

    if (options.isSelfServicePasswordResetRegistered !== undefined) {
      filters.push(`isSelfServicePasswordResetRegistered eq ${options.isSelfServicePasswordResetRegistered}`);
    }

    if (options.isSystemPreferredAuthenticationMethodEnabled !== undefined) {
      filters.push(`isSystemPreferredAuthenticationMethodEnabled eq ${options.isSystemPreferredAuthenticationMethodEnabled}`);
    }

    const methodsRegistered = options.methodsRegistered?.split(',').map(method => `methodsRegistered/any(m:m eq '${method.trim()}')`);
    const methodsRegisteredFilter = methodsRegistered?.join(' or ');
    if (methodsRegisteredFilter) {
      filters.push(`(${methodsRegisteredFilter})`);
    }

    const systemPreferredAuthenticationMethods = options.systemPreferredAuthenticationMethods?.split(',').map(method => `systemPreferredAuthenticationMethods/any(m:m eq '${method.trim()}')`);
    const systemPreferredAuthenticationMethodsFilter = systemPreferredAuthenticationMethods?.join(' or ');

    if (systemPreferredAuthenticationMethodsFilter) {
      filters.push(`(${systemPreferredAuthenticationMethodsFilter})`);
    }

    const userUPNs: string[] = [];
    if (userUpns.length > 0) {
      userUpns.forEach(upn => {
        userUPNs.push(`userPrincipalName eq '${formatting.encodeQueryParameter(upn)}'`);
      });
    }

    if (options.userPrincipalNames) {
      const upns = options.userPrincipalNames.split(',').map(m => m.trim());
      upns.forEach(upn => {
        userUPNs.push(`userPrincipalName eq '${formatting.encodeQueryParameter(upn)}'`);
      });
    }

    if (userUPNs.length > 0) {
      filters.push(`(${userUPNs.join(' or ')})`);
    }

    const userPreferredMethodForSecondaryAuthentication = options.userPreferredMethodForSecondaryAuthentication?.split(',').map(method => `userPreferredMethodForSecondaryAuthentication eq '${method.trim()}'`);
    const userPreferredMethodForSecondaryAuthenticationFilter = userPreferredMethodForSecondaryAuthentication?.join(' or ');

    if (userPreferredMethodForSecondaryAuthenticationFilter) {
      filters.push(`(${userPreferredMethodForSecondaryAuthenticationFilter})`);
    }

    if (options.userType) {
      filters.push(`userType eq '${options.userType}'`);
    }

    if (filters.length > 0) {
      queryParameters.push(`$filter=${filters.join(' and ')}`);
    }

    const queryString = queryParameters.length > 0
      ? `?${queryParameters.join('&')}`
      : '';

    return `${this.resource}/v1.0/reports/authenticationMethods/userRegistrationDetails${queryString}`;
  }
}

export default new EntraUserRegistrationDetailsListCommand();