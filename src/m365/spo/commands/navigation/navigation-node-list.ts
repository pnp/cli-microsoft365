import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { MenuState } from './NavigationNode.js';

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
    return ['Key', 'Title', 'SimpleUrl'];
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving navigation nodes...`);
    }

    try {
      const menuState: MenuState = args.options.location === 'TopNavigationBar'
        ? await spo.getTopNavigationMenuState(args.options.webUrl)
        : await spo.getQuickLaunchMenuState(args.options.webUrl);
      await logger.log(menuState.Nodes);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoNavigationNodeListCommand();