import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  environmentName: z.string().alias('e'),
  asAdmin: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpAiBuilderModelListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.AIBUILDERMODEL_LIST;
  }

  public get description(): string {
    return 'Lists available AI builder models in the specified Power Platform environment.';
  }

  public defaultProperties(): string[] | undefined {
    return ['msdyn_name', 'msdyn_aimodelid', 'createdon', 'modifiedon'];
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving available AI Builder models`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const aimodels = await odata.getAllItems<any>(`${dynamicsApiUrl}/api/data/v9.0/msdyn_aimodels?$filter=iscustomizable/Value eq true`);
      await logger.log(aimodels);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpAiBuilderModelListCommand();