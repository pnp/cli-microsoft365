import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
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
  asAdmin: z.boolean().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpSolutionRemoveCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.SOLUTION_REMOVE;
  }

  public get description(): string {
    return 'Removes a specific solution in the specified Power Platform environment.';
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
      await logger.logToStderr(`Removing solution '${args.options.id || args.options.name}'...`);
    }

    if (args.options.force) {
      await this.deleteSolution(args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove solution '${args.options.id || args.options.name}'?` });

      if (result) {
        await this.deleteSolution(args);
      }
    }
  }

  private async getSolutionId(args: CommandArgs, dynamicsApiUrl: string): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    const solution = await powerPlatform.getSolutionByName(dynamicsApiUrl, args.options.name!);
    return solution.solutionid;
  }

  private async deleteSolution(args: CommandArgs): Promise<void> {
    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const solutionId = await this.getSolutionId(args, dynamicsApiUrl);
      const requestOptions: CliRequestOptions = {
        url: `${dynamicsApiUrl}/api/data/v9.1/solutions(${solutionId})`,
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

export default new PpSolutionRemoveCommand();