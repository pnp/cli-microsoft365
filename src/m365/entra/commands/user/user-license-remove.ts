import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { formatting } from '../../../../utils/formatting.js';
import { cli } from '../../../../cli/cli.js';
import GraphCommand from '../../../base/GraphCommand.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.uuid().optional(),
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid user principal name (UPN).`
  }).optional(),
  ids: z.string().refine(ids => !ids.split(',').some(e => !validation.isValidGuid(e)), {
    error: e => `'${e.input}' contains one or more invalid GUIDs.`
  }),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraUserLicenseRemoveCommand extends GraphCommand {

  public get name(): string {
    return commands.USER_LICENSE_REMOVE;
  }

  public get description(): string {
    return 'Removes a license from a user';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.userId, options.userName].filter(o => o !== undefined).length === 1, {
        error: `Specify either 'userId' or 'userName'.`,
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName']
        }
      });
  }

  public async commandAction(logger: Logger, args: any): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing the licenses for the user '${args.options.userId || args.options.userName}'...`);
    }

    if (args.options.force) {
      await this.deleteUserLicenses(args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the licenses for the user '${args.options.userId || args.options.userName}'?` });

      if (result) {
        await this.deleteUserLicenses(args);
      }
    }
  }

  private async deleteUserLicenses(args: CommandArgs): Promise<void> {
    const removeLicenses = args.options.ids.split(',');
    const requestBody = { "addLicenses": [], "removeLicenses": removeLicenses };

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/users/${formatting.encodeQueryParameter(args.options.userId || args.options.userName as string)}/assignLicense`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      data: requestBody,
      responseType: 'json'
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraUserLicenseRemoveCommand();