import fs from 'fs';
import path from 'path';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().refine(val => validation.isValidGuid(val), {
    message: 'The value must be a valid GUID.'
  }).optional().alias('i'),
  name: z.string().optional().alias('n'),
  filePath: z.string()
    .refine(val => {
      const fullPath = path.resolve(val);
      if (!fs.existsSync(fullPath)) {
        return false;
      }
      if (fs.lstatSync(fullPath).isDirectory()) {
        return false;
      }
      return true;
    }, {
      message: 'Specified file does not exist or points to a directory.'
    })
    .alias('p')
});

declare type Options = z.infer<typeof options>;
interface CommandArgs {
  options: Options;
}

class TeamsAppUpdateCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_UPDATE;
  }

  public get description(): string {
    return 'Updates Teams app in the organization\'s app catalog';
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
    const { filePath } = args.options;

    try {
      const appId: string = await this.getAppId(args.options);
      const fullPath: string = path.resolve(filePath);
      if (this.verbose) {
        await logger.logToStderr(`Updating app with id '${appId}' and file '${fullPath}' in the app catalog...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/appCatalogs/teamsApps/${appId}`,
        headers: {
          "content-type": "application/zip"
        },
        data: fs.readFileSync(fullPath)
      };

      await request.put(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppId(options: Options): Promise<string> {
    if (options.id) {
      return options.id;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/appCatalogs/teamsApps?$filter=displayName eq '${formatting.encodeQueryParameter(options.name as string)}'`,
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
      const result = await cli.handleMultipleResultsFound<{ id: string; }>(`Multiple Teams apps with name ${options.name} found.`, resultAsKeyValuePair);
      return result.id;
    }

    return app.id;
  }
}

export default new TeamsAppUpdateCommand();