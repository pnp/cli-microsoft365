import auth from '../../../Auth.js';
import { globalOptionsZod } from '../../../Command.js';
import { Logger } from '../../../cli/Logger.js';
import { urlUtil } from '../../../utils/urlUtil.js';
import { validation } from '../../../utils/validation.js';
import SpoCommand from '../../base/SpoCommand.js';
import commands from '../commands.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  url: z.string().alias('u').refine(url => validation.isValidSharePointUrl(url) === true, {
    message: 'Specify a valid SharePoint URL'
  })
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSetCommand extends SpoCommand {
  public get name(): string {
    return commands.SET;
  }

  public get description(): string {
    return 'Sets the URL of the root SharePoint site collection for use in SPO commands';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    auth.connection.spoUrl = urlUtil.removeTrailingSlashes(args.options.url);

    try {
      await auth.storeConnectionInfo();
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSetCommand();