import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { NavigationNode } from './NavigationNode.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  location: z.enum(['QuickLaunch', 'TopNavigationBar']).alias('l')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoNavigationNodeListCommand extends SpoCommand {
  public get name(): string {
    return commands.NAVIGATION_NODE_LIST;
  }

  public get description(): string {
    return 'Lists nodes from the specified site navigation';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'Title', 'Url'];
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving navigation nodes...`);
    }

    try {
      const res = await odata.getAllItems<NavigationNode>(`${args.options.webUrl}/_api/web/navigation/${args.options.location.toLowerCase()}?$expand=Children,Children/Children,Children/Children/Children`);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoNavigationNodeListCommand();