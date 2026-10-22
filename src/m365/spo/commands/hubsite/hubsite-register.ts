import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  siteUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).alias('u')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoHubSiteRegisterCommand extends SpoCommand {
  public get name(): string {
    return commands.HUBSITE_REGISTER;
  }

  public get description(): string {
    return 'Registers the specified site collection as a hub site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const reqDigest = await spo.getRequestDigest(args.options.siteUrl);

      const requestOptions: CliRequestOptions = {
        url: `${args.options.siteUrl}/_api/site/RegisterHubSite`,
        headers: {
          'X-RequestDigest': reqDigest.FormDigestValue,
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const res = await request.post(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoHubSiteRegisterCommand();