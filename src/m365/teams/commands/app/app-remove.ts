import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().refine(val => validation.isValidGuid(val), {
    message: 'The value must be a valid GUID.'
  }).optional().alias('i'),
  name: z.string().optional().alias('n'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;
interface CommandArgs {
  options: Options;
}

class TeamsAppRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_REMOVE;
  }

  public get description(): string {
    return 'Removes a Teams app from the organization\'s app catalog';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.name].filter(x => x !== undefined).length === 1, {
        message: 'Specify either id or name, but not both.',
        params: {
          customCode: 'optionSet',
          options: ['id', 'name']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeApp = async (): Promise<void> => {
      try {
        const appId: string = await this.getAppId(args.options, logger);

        if (this.verbose) {
          await logger.logToStderr(`Removing app with ID ${appId}`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/appCatalogs/teamsApps/${appId}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          }
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeApp();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the Teams app ${args.options.id || args.options.name} from the app catalog?` });

      if (result) {
        await removeApp();
      }
    }
  }

  private async getAppId(options: Options, logger: Logger): Promise<any> {
    if (options.id) {
      return options.id;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving app Id...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/appCatalogs/teamsApps?$filter=displayName eq '${formatting.encodeQueryParameter(options.name as string)}'&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: { id: string; }[] }>(requestOptions);
    const app: { id: string; } | undefined = response.value[0];

    if (!app) {
      throw `The specified Teams app does not exist`;
    }

    if (response.value.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', response.value);
      const result = await cli.handleMultipleResultsFound<{ id: string; }>(`Multiple Teams apps with name '${options.name}' found.`, resultAsKeyValuePair);
      return result.id;
    }

    return app.id;
  }
}

export default new TeamsAppRemoveCommand();