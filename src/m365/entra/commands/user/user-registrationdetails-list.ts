import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { odata } from '../../../../utils/odata.js';
import { UserRegistrationDetails } from '@microsoft/microsoft-graph-types';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import { formatting } from '../../../../utils/formatting.js';

const authenticationMethodValues = ['push', 'oath', 'voiceMobile', 'voiceAlternateMobile', 'voiceOffice', 'sms', 'none'] as const;
const methodsRegisteredValues = ['mobilePhone', 'email', 'fido2', 'microsoftAuthenticatorPush', 'softwareOneTimePasscode'] as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  isAdmin: z.boolean().optional(),
  userType: z.enum(['member', 'guest']).optional(),
  userPreferredMethodForSecondaryAuthentication: z.string().refine(val => {
    const methods = val.split(',').map(m => m.trim());
    return methods.every(m => (authenticationMethodValues as readonly string[]).includes(m));
  }, {
    error: e => `'${e.input}' is not a valid userPreferredMethodForSecondaryAuthentication value. Allowed values ${authenticationMethodValues.join(', ')}.`
  }).optional(),
  systemPreferredAuthenticationMethods: z.string().refine(val => {
    const methods = val.split(',').map(m => m.trim());
    return methods.every(m => (authenticationMethodValues as readonly string[]).includes(m));
  }, {
    error: e => `'${e.input}' is not a valid systemPreferredAuthenticationMethods value. Allowed values ${authenticationMethodValues.join(', ')}.`
  }).optional(),
  isSelfServicePasswordResetRegistered: z.boolean().optional(),
  isSelfServicePasswordResetEnabled: z.boolean().optional(),
  isSelfServicePasswordResetCapable: z.boolean().optional(),
  isMfaRegistered: z.boolean().optional(),
  isMfaCapable: z.boolean().optional(),
  isPasswordlessCapable: z.boolean().optional(),
  isSystemPreferredAuthenticationMethodEnabled: z.boolean().optional(),
  methodsRegistered: z.string().refine(val => {
    const methods = val.split(',').map(m => m.trim());
    return methods.every(m => (methodsRegisteredValues as readonly string[]).includes(m));
  }, {
    error: e => `'${e.input}' is not a valid methodsRegistered value. Allowed values ${methodsRegisteredValues.join(', ')}.`
  }).optional(),
  userIds: z.string().refine(val => validation.isValidGuidArray(val) === true, {
    error: e => `The following GUIDs are invalid for the option 'userIds': ${validation.isValidGuidArray(e.input as string)}.`
  }).optional(),
  userPrincipalNames: z.string().refine(val => validation.isValidUserPrincipalNameArray(val) === true, {
    error: e => `The following user principal names are invalid for the option 'userPrincipalNames': ${validation.isValidUserPrincipalNameArray(e.input as string)}.`
  }).optional(),
  properties: z.string().optional().alias('p')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

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

  public get schema(): z.ZodTypeAny | undefined {
    return options;
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