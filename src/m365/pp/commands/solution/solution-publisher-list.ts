import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { Publisher } from './Solution.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  environmentName: z.string().alias('e'),
  withMicrosoftPublishers: z.boolean().optional(),
  asAdmin: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpSolutionPublisherListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.SOLUTION_PUBLISHER_LIST;
  }

  public get description(): string {
    return 'Lists publishers in a given environment.';
  }

  public defaultProperties(): string[] | undefined {
    return ['publisherid', 'uniquename', 'friendlyname'];
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of publishers...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const requestOptions: CliRequestOptions = {
        url: `${dynamicsApiUrl}/api/data/v9.0/publishers?$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix${!args.options.withMicrosoftPublishers ? `&$filter=publisherid ne 'd21aab70-79e7-11dd-8874-00188b01e34f'` : ''}&api-version=9.1`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get<{ value: Publisher[] }>(requestOptions);
      await logger.log(res.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpSolutionPublisherListCommand();