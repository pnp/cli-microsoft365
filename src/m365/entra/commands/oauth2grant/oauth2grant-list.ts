import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = globalOptionsZod
  .extend({
    spObjectId: z.uuid().alias('i')
  }).strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraOAuth2GrantListCommand extends GraphCommand {
  public get name(): string {
    return commands.OAUTH2GRANT_LIST;
  }

  public get description(): string {
    return 'Lists OAuth2 permission grants for the specified service principal';
  }

  public defaultProperties(): string[] | undefined {
    return ['objectId', 'resourceId', 'scope'];
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of OAuth grants for the service principal...`);
    }

    try {
      const res = await odata.getAllItems<any>(`${this.resource}/v1.0/oauth2PermissionGrants?$filter=clientId eq '${formatting.encodeQueryParameter(args.options.spObjectId)}'`);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraOAuth2GrantListCommand();