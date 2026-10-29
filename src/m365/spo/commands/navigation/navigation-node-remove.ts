import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    message: 'The value of the option \'webUrl\' is not a valid SharePoint Online site URL.'
  }).alias('u'),
  location: z.string().pipe(z.enum(['QuickLaunch', 'TopNavigationBar'])).alias('l'),
  id: z.int().positive().alias('i'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoNavigationNodeRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.NAVIGATION_NODE_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified navigation node';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeNode(logger, args.options);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the node ${args.options.id} from the navigation?` });

      if (result) {
        await this.removeNode(logger, args.options);
      }
    }
  }

  private async removeNode(logger: Logger, options: Options): Promise<void> {
    try {
      const res = await spo.getRequestDigest(options.webUrl);

      if (this.verbose) {
        await logger.logToStderr(`Removing navigation node...`);
      }

      const requestOptions: any = {
        url: `${options.webUrl}/_api/web/navigation/${options.location.toLowerCase()}/getbyid(${options.id})`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'X-RequestDigest': res.FormDigestValue
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

export default new SpoNavigationNodeRemoveCommand();