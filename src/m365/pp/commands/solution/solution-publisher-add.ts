import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  environmentName: z.string().alias('e'),
  name: z.string()
    .refine(val => /^[a-zA-Z_][A-Za-z0-9_]+$/.test(val), {
      message: 'Option name may only consist of alphanumeric characters and underscores. The first character cannot be a number.'
    }).alias('n'),
  displayName: z.string(),
  prefix: z.string()
    .refine(val => /^(?!mscrm.*$)[a-zA-Z][A-Za-z0-9]{1,7}$/.test(val), {
      message: `Option prefix may only consist of alphanumeric characters. The first character cannot be a number and may not start with 'mscrm'. It must be between 2 and 8 characters long.`
    }),
  choiceValuePrefix: z.string()
    .refine(val => {
      const num = Number(val);
      return Number.isInteger(num) && num >= 10000 && num <= 99999;
    }, {
      message: 'Option choiceValuePrefix should be an integer between 10000 and 99999.'
    }),
  asAdmin: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpSolutionPublisherAddCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.SOLUTION_PUBLISHER_ADD;
  }

  public get description(): string {
    return 'Adds a specified publisher in a given environment';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding new publisher '${args.options.name}'...`);
    }
    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const requestOptions: CliRequestOptions = {
        url: `${dynamicsApiUrl}/api/data/v9.0/publishers`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          uniquename: args.options.name,
          friendlyname: args.options.displayName,
          customizationprefix: args.options.prefix,
          customizationoptionvalueprefix: Number(args.options.choiceValuePrefix)
        }
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpSolutionPublisherAddCommand();