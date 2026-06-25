import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { Publisher } from './Solution.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  environmentName: z.string().alias('e'),
  id: z.string().refine(val => validation.isValidGuid(val), {
    message: 'The value must be a valid GUID.'
  }).optional().alias('i'),
  name: z.string().optional().alias('n'),
  asAdmin: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpSolutionPublisherGetCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.SOLUTION_PUBLISHER_GET;
  }

  public get description(): string {
    return 'Gets information about the specified publisher in a given environment.';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.name].filter(x => x !== undefined).length === 1, {
        message: `Specify either 'id' or 'name', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'name']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving a specific publisher '${args.options.id || args.options.name}'...`);
    }

    const res = await this.getSolutionPublisher(args);
    await logger.log(res);
  }

  private async getSolutionPublisher(args: CommandArgs): Promise<any> {
    const requestOptions: CliRequestOptions = {
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      if (args.options.id) {
        requestOptions.url = `${dynamicsApiUrl}/api/data/v9.0/publishers(${args.options.id})?$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&api-version=9.1`;

        const result = await request.get<Publisher>(requestOptions);
        return result;
      }

      requestOptions.url = `${dynamicsApiUrl}/api/data/v9.0/publishers?$filter=friendlyname eq '${args.options.name}'&$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&api-version=9.1`;
      const result = await request.get<{ value: Publisher[] }>(requestOptions);

      if (result.value.length === 0) {
        throw `The specified publisher '${args.options.name}' does not exist.`;
      }

      return result.value[0];
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpSolutionPublisherGetCommand();