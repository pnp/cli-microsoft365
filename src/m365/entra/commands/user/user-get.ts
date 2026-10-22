import { User } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.uuid().optional().alias('i'),
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid userName.`
  }).optional().alias('n'),
  email: z.string().optional(),
  properties: z.string().optional().alias('p'),
  withManager: z.boolean().optional()
});

export declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraUserGetCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_GET;
  }

  public get description(): string {
    return 'Gets information about the specified user';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.id, options.userName, options.email].filter(o => o !== undefined).length === 1, {
        error: `Specify either 'id', 'userName', or 'email'.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'userName', 'email']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let userIdOrPrincipalName = args.options.id;

      if (args.options.userName) {
        // single user can be retrieved also by user principal name
        userIdOrPrincipalName = formatting.encodeQueryParameter(args.options.userName);
      }
      else if (args.options.email) {
        userIdOrPrincipalName = await entraUser.getUserIdByEmail(args.options.email);
      }

      const requestUrl: string = this.getRequestUrl(userIdOrPrincipalName!, args.options);

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const user = await request.get<User>(requestOptions);
      await logger.log(user);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getRequestUrl(userIdOrPrincipalName: string, options: Options): string {
    const queryParameters: string[] = [];

    if (options.properties) {
      const allProperties = options.properties.split(',');
      const selectProperties = allProperties.filter(prop => !prop.includes('/'));

      if (selectProperties.length > 0) {
        queryParameters.push(`$select=${selectProperties}`);
      }
    }

    if (options.withManager) {
      queryParameters.push('$expand=manager($select=displayName,userPrincipalName,id,mail)');
    }

    const queryString = queryParameters.length > 0
      ? `?${queryParameters.join('&')}`
      : '';

    // user principal name can start with $ but it violates the OData URL convention, so it must be enclosed in parenthesis and single quotes
    return userIdOrPrincipalName.startsWith('%24')
      ? `${this.resource}/v1.0/users('${userIdOrPrincipalName}')${queryString}`
      : `${this.resource}/v1.0/users/${userIdOrPrincipalName}${queryString}`;
  }
}

export default new EntraUserGetCommand();
