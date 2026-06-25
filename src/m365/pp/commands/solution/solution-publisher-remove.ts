import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command, { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import ppSolutionPublisherGetCommand, { options as ppSolutionPublisherGetOptions } from './solution-publisher-get.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  environmentName: z.string().alias('e'),
  id: z.string().refine(val => validation.isValidGuid(val), {
    message: 'The value must be a valid GUID.'
  }).optional().alias('i'),
  name: z.string().optional().alias('n'),
  asAdmin: z.boolean().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpSolutionPublisherRemoveCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.SOLUTION_PUBLISHER_REMOVE;
  }

  public get description(): string {
    return 'Removes a specific publisher in the specified Power Platform environment.';
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
      await logger.logToStderr(`Removes a publisher '${args.options.id || args.options.name}'...`);
    }

    if (args.options.force) {
      await this.deletePublisher(args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove publisher '${args.options.id || args.options.name}'?` });

      if (result) {
        await this.deletePublisher(args);
      }
    }
  }

  private async getPublisherId(args: CommandArgs): Promise<any> {
    if (args.options.id) {
      return args.options.id;
    }

    const options: z.infer<typeof ppSolutionPublisherGetOptions> = {
      environmentName: args.options.environmentName,
      name: args.options.name,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await cli.executeCommandWithOutput(ppSolutionPublisherGetCommand as Command, { options: { ...options, _: [] } });
    const getPublisherOutput = JSON.parse(output.stdout);
    return getPublisherOutput.publisherid;
  }

  private async deletePublisher(args: CommandArgs): Promise<void> {
    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const publisherId = await this.getPublisherId(args);
      const requestOptions: CliRequestOptions = {
        url: `${dynamicsApiUrl}/api/data/v9.1/publishers(${publisherId})`,
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

export default new PpSolutionPublisherRemoveCommand();