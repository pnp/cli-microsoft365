import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command, { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import ppAiBuilderModelGetCommand from './aibuildermodel-get.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  environmentName: z.string().alias('e'),
  id: z.string().refine(val => validation.isValidGuid(val), {
    error: 'The value must be a valid GUID.'
  }).optional().alias('i'),
  name: z.string().optional().alias('n'),
  asAdmin: z.boolean().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpAiBuilderModelRemoveCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.AIBUILDERMODEL_REMOVE;
  }

  public get description(): string {
    return 'Removes an AI builder model in the specified Power Platform environment.';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.name].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'id' or 'name', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'name']
        }
      });
  }

  public async commandAction(logger: Logger, args: any): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing AI builder model '${args.options.id || args.options.name}'...`);
    }

    if (args.options.force) {
      await this.deleteAiBuilderModel(args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove AI builder model '${args.options.id || args.options.name}'?` });

      if (result) {
        await this.deleteAiBuilderModel(args);
      }
    }
  }

  private async getAiBuilderModelId(args: CommandArgs): Promise<any> {
    if (args.options.id) {
      return args.options.id;
    }

    const options = {
      environmentName: args.options.environmentName,
      name: args.options.name,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await cli.executeCommandWithOutput(ppAiBuilderModelGetCommand as Command, { options: { ...options, _: [] } });
    const getAiBuilderModelOutput = JSON.parse(output.stdout);
    return getAiBuilderModelOutput.msdyn_aimodelid;
  }

  private async deleteAiBuilderModel(args: CommandArgs): Promise<void> {
    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const aiBuilderModelId = await this.getAiBuilderModelId(args);
      const requestOptions: CliRequestOptions = {
        url: `${dynamicsApiUrl}/api/data/v9.1/msdyn_aimodels(${aiBuilderModelId})`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpAiBuilderModelRemoveCommand();