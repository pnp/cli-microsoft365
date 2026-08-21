import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).alias('u'),
  scope: z.enum(['All', 'Site', 'Web']).optional().alias('s')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoCommandSetListCommand extends SpoCommand {
  public get name(): string {
    return commands.COMMANDSET_LIST;
  }

  public get description(): string {
    return 'Gets a list of ListView Command Sets that are added to a site.';
  }

  public defaultProperties(): string[] | undefined {
    return ['Name', 'Location', 'Scope', 'Id'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Attempt to get commandsets...`);
      }

      const commandsets = await spo.getCustomActions(args.options.webUrl, args.options.scope, `startswith(Location,'ClientSideExtension.ListViewCommandSet')`);

      await logger.log(commandsets);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

export default new SpoCommandSetListCommand();