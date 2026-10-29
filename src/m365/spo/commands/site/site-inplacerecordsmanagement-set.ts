import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  siteUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  enabled: z.boolean()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteInPlaceRecordsManagementSetCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_INPLACERECORDSMANAGEMENT_SET;
  }

  public get description(): string {
    return 'Activates or deactivates in-place records management for a site collection';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: any = {
      url: `${args.options.siteUrl}/_api/site/features/${args.options.enabled ? 'add' : 'remove'}`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      data: {
        featureId: 'da2e115b-07e4-49d9-bb2c-35e93bb9fca9',
        force: true
      },
      responseType: 'json'
    };

    if (this.verbose) {
      await logger.logToStderr(`${args.options.enabled ? 'Activating' : 'Deactivating'} in-place records management for site ${args.options.siteUrl}`);
    }

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteInPlaceRecordsManagementSetCommand();