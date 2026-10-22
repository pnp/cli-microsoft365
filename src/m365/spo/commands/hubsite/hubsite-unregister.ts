import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  url: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).alias('u'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoHubSiteUnregisterCommand extends SpoCommand {
  public get name(): string {
    return commands.HUBSITE_UNREGISTER;
  }

  public get description(): string {
    return 'Unregisters the specified site collection as a hub site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const unregisterHubSite = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Unregistering site collection ${args.options.url} as a hub site...`);
        }

        const res = await spo.getRequestDigest(args.options.url);

        const requestOptions: CliRequestOptions = {
          url: `${args.options.url}/_api/site/UnregisterHubSite`,
          headers: {
            'X-RequestDigest': res.FormDigestValue,
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await unregisterHubSite();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to unregister the site collection ${args.options.url} as a hub site?` });

      if (result) {
        await unregisterHubSite();
      }
    }
  }
}

export default new SpoHubSiteUnregisterCommand();