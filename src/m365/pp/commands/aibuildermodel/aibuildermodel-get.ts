import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

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

class PpAiBuilderModelGetCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.AIBUILDERMODEL_GET;
  }

  public get description(): string {
    return 'Gets a specific AI builder model in the specified Power Platform environment';
  }

  public get schema(): z.ZodType {
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
      await logger.logToStderr(`Retrieving an AI builder model '${args.options.id || args.options.name}'...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const res = await this.getAiBuilderModel(dynamicsApiUrl, args.options);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAiBuilderModel(dynamicsApiUrl: string, options: Options): Promise<any> {
    const requestOptions: CliRequestOptions = {
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    if (options.id) {
      requestOptions.url = `${dynamicsApiUrl}/api/data/v9.1/msdyn_aimodels(${options.id})?$filter=iscustomizable/Value eq true`;
      const result = await request.get<any>(requestOptions);
      return result;
    }

    requestOptions.url = `${dynamicsApiUrl}/api/data/v9.1/msdyn_aimodels?$filter=msdyn_name eq '${options.name}' and iscustomizable/Value eq true`;
    const result = await request.get<{ value: any[] }>(requestOptions);

    if (result.value.length === 0) {
      throw `The specified AI builder model '${options.name}' does not exist.`;
    }

    if (result.value.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('msdyn_aimodelid', result.value);
      return await cli.handleMultipleResultsFound(`Multiple AI builder models with name '${options.name}' found.`, resultAsKeyValuePair);
    }

    return result.value[0];
  }
}

export default new PpAiBuilderModelGetCommand();