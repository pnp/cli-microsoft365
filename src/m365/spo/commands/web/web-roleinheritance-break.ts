import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `${e.input} is not a valid SharePoint Online site URL.`
  }).alias('u'),
  clearExistingPermissions: z.boolean().optional().alias('c'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoWebRoleInheritanceBreakCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_ROLEINHERITANCE_BREAK;
  }

  public get description(): string {
    return 'Break role inheritance of subsite';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Break role inheritance of subsite with URL ${args.options.webUrl}...`);
    }

    if (args.options.force) {
      await this.breakRoleInheritance(args.options);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to break the role inheritance of subsite ${args.options.webUrl}?` });

      if (result) {
        await this.breakRoleInheritance(args.options);
      }
    }
  }

  private async breakRoleInheritance(options: Options): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${options.webUrl}/_api/web/breakroleinheritance(${!options.clearExistingPermissions})`,
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoWebRoleInheritanceBreakCommand();
