import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: 'Specify valid userName.'
  }).alias('n'),
  apiKey: z.string(),
  domain: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraUserHibpCommand extends AnonymousCommand {
  public get name(): string {
    return commands.USER_HIBP;
  }

  public get description(): string {
    return 'Allows you to retrieve all accounts that have been pwned with the specified username';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `https://haveibeenpwned.com/api/v3/breachedaccount/${formatting.encodeQueryParameter(args.options.userName)}${(args.options.domain ? `?domain=${formatting.encodeQueryParameter(args.options.domain)}` : '')}`,
        headers: {
          'accept': 'application/json',
          'hibp-api-key': args.options.apiKey,
          'x-anonymous': true
        },
        responseType: 'json'
      };

      const res = await request.get(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      if ((err && err.response !== undefined && err.response.status === 404) && (this.debug || this.verbose)) {
        await logger.log('No pwnage found');
        return;
      }
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraUserHibpCommand();