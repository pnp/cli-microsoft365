import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = globalOptionsZod
  .extend({
    grantId: z.string().alias('i'),
    scope: z.string().alias('s')
  }).strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraOAuth2GrantSetCommand extends GraphCommand {
  public get name(): string {
    return commands.OAUTH2GRANT_SET;
  }

  public get description(): string {
    return 'Updates OAuth2 permissions for the service principal';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Updating OAuth2 permissions...`);
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/oauth2PermissionGrants/${formatting.encodeQueryParameter(args.options.grantId)}`,
        headers: {
          'content-type': 'application/json'
        },
        responseType: 'json',
        data: {
          "scope": args.options.scope
        }
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraOAuth2GrantSetCommand();