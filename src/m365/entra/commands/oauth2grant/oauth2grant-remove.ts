import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = globalOptionsZod
  .extend({
    grantId: z.string().alias('i'),
    force: z.boolean().optional().alias('f')
  }).strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraOAuth2GrantRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.OAUTH2GRANT_REMOVE;
  }

  public get description(): string {
    return 'Removes specified service principal OAuth2 permissions';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeOauth2Grant: () => Promise<void> = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing OAuth2 permissions...`);
      }

      try {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/oauth2PermissionGrants/${formatting.encodeQueryParameter(args.options.grantId)}`,
          headers: {
            'accept': 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeOauth2Grant();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the OAuth2 permissions for ${args.options.grantId}?` });

      if (result) {
        await removeOauth2Grant();
      }
    }
  }
}

export default new EntraOAuth2GrantRemoveCommand();