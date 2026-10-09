import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { ContextInfo, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().refine(val => validation.isValidGuid(val), {
    message: 'The value must be a valid GUID.'
  }).alias('i'),
  force: z.boolean().optional().alias('f')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteScriptRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.SITESCRIPT_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified site script';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeSiteScript(logger, args.options.id);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the site script ${args.options.id}?` });

      if (result) {
        await this.removeSiteScript(logger, args.options.id);
      }
    }
  }

  private async removeSiteScript(logger: Logger, id: string): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const formDigest: ContextInfo = await spo.getRequestDigest(spoUrl);
      const requestOptions: any = {
        url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteScript`,
        headers: {
          'X-RequestDigest': formDigest.FormDigestValue,
          'content-type': 'application/json;charset=utf-8',
          accept: 'application/json;odata=nometadata'
        },
        data: { id: id },
        responseType: 'json'
      };
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteScriptRemoveCommand();