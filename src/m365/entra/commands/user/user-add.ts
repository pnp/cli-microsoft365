import { User } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface ExtendedUser extends User {
  password: string;
}

export const options = z.looseObject({
  ...globalOptionsZod.shape,
  displayName: z.string(),
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid userName.`
  }),
  accountEnabled: z.boolean().optional(),
  mailNickname: z.string().optional(),
  password: z.string().optional(),
  firstName: z.string().max(64, { error: `The maximum amount of characters for 'firstName' is 64.` }).optional(),
  lastName: z.string().max(64, { error: `The maximum amount of characters for 'lastName' is 64.` }).optional(),
  forceChangePasswordNextSignIn: z.boolean().optional(),
  forceChangePasswordNextSignInWithMfa: z.boolean().optional(),
  usageLocation: z.string().regex(/^[a-zA-Z]{2}$/, { error: e => `'${e.input}' is not a valid usageLocation.` }).optional(),
  officeLocation: z.string().optional(),
  jobTitle: z.string().max(128, { error: `The maximum amount of characters for 'jobTitle' is 128.` }).optional(),
  companyName: z.string().max(64, { error: `The maximum amount of characters for 'companyName' is 64.` }).optional(),
  department: z.string().max(64, { error: `The maximum amount of characters for 'department' is 64.` }).optional(),
  preferredLanguage: z.string().min(2, { error: e => `'${e.input}' is not a valid preferredLanguage.` }).optional(),
  managerUserId: z.uuid().optional(),
  managerUserName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid user principal name.`
  }).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraUserAddCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_ADD;
  }

  public get description(): string {
    return 'Creates a new user';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => !(options.managerUserId && options.managerUserName), {
        error: `Specify either 'managerUserId' or 'managerUserName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['managerUserId', 'managerUserName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding user to AAD with displayName ${args.options.displayName} and userPrincipalName ${args.options.userName}`);
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

      await logger.log(user);
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

    this.addUnknownOptionsToPayloadZod(requestBody, options);

    return requestBody;
  }

  /**
   * Generate a password with at least: one digit, one lowercase character, one uppercase character, and a special character.
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

export default new EntraUserAddCommand();