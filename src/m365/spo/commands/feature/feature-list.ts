import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { Feature } from './Feature.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).alias('u'),
  scope: z.enum(['Site', 'Web']).optional().alias('s')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFeatureListCommand extends SpoCommand {
  public get name(): string {
    return commands.FEATURE_LIST;
  }

  public get description(): string {
    return 'Lists Features activated in the specified site or site collection';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const scope: string = (args.options.scope) ? args.options.scope : 'Web';

    try {
      const features = await odata.getAllItems<Feature>(`${args.options.webUrl}/_api/${scope}/Features?$select=DisplayName,DefinitionId`);
      await logger.log(features);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFeatureListCommand();